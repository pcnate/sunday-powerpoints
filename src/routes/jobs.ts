import { Router, Request, Response } from 'express';
import { Server as SocketIOServer } from 'socket.io';
import { RowDataPacket, ResultSetHeader } from 'mysql2/promise';
import { getPool } from '../db';
import {
  Job, JobLog, CreateJobRequest, ClaimJobRequest, HeartbeatRequest,
  CompleteJobRequest, FailJobRequest, JobLogEntry, JobListQuery, WorkerInfo
} from '../types/job';

const JOB_TYPES = [ 'video-alignment', 'transcription', 'claude-processing', 'transcode' ];
const JOB_STATUSES = [ 'pending', 'queued', 'processing', 'completed', 'failed', 'cancelled' ];

/** Stale threshold — workers not seen in this many ms are marked stale. */
const WORKER_STALE_MS = 120_000;

/**
 * In-memory tracker for connected workers.
 * Keyed by worker_id, updated on claim, heartbeat, complete, and fail.
 */
const workers = new Map<string, {
  types: string[];
  current_job_id: number | null;
  current_job_type: string | null;
  last_seen: Date;
  first_seen: Date;
  jobs_completed: number;
  jobs_failed: number;
}>();


/**
 * Record that a worker was seen (claim or heartbeat).
 *
 * @param worker_id - the worker identifier
 * @param types - job types this worker handles (from claim)
 * @param job_id - current job id if executing
 * @param job_type - current job type if executing
 */
function touchWorker( worker_id: string, types?: string[], job_id?: number | null, job_type?: string | null ): void {
  const existing = workers.get( worker_id );
  if ( existing ) {
    existing.last_seen = new Date();
    if ( types ) existing.types = types;
    if ( job_id !== undefined ) existing.current_job_id = job_id;
    if ( job_type !== undefined ) existing.current_job_type = job_type;
  } else {
    workers.set( worker_id, {
      types: types || [],
      current_job_id: job_id ?? null,
      current_job_type: job_type ?? null,
      last_seen: new Date(),
      first_seen: new Date(),
      jobs_completed: 0,
      jobs_failed: 0,
    } );
  }
}


/**
 * Record a job completion for a worker.
 *
 * @param worker_id - the worker identifier
 */
function workerJobCompleted( worker_id: string ): void {
  const w = workers.get( worker_id );
  if ( w ) {
    w.jobs_completed++;
    w.current_job_id = null;
    w.current_job_type = null;
    w.last_seen = new Date();
  }
}


/**
 * Record a job failure for a worker.
 *
 * @param worker_id - the worker identifier
 */
function workerJobFailed( worker_id: string ): void {
  const w = workers.get( worker_id );
  if ( w ) {
    w.jobs_failed++;
    w.current_job_id = null;
    w.current_job_type = null;
    w.last_seen = new Date();
  }
}


/**
 * Get all tracked workers with computed status.
 *
 * @returns array of WorkerInfo objects
 */
export function getWorkers(): WorkerInfo[] {
  const now = Date.now();
  const result: WorkerInfo[] = [];

  for ( const [ id, w ] of workers ) {
    const msSinceSeen = now - w.last_seen.getTime();
    let status: WorkerInfo[ 'status' ] = 'idle';
    if ( msSinceSeen > WORKER_STALE_MS ) {
      status = 'stale';
    } else if ( w.current_job_id !== null ) {
      status = 'executing';
    }

    result.push({
      id,
      types: w.types,
      current_job_id: w.current_job_id,
      current_job_type: w.current_job_type,
      last_seen: w.last_seen,
      first_seen: w.first_seen,
      jobs_completed: w.jobs_completed,
      jobs_failed: w.jobs_failed,
      status,
    });
  }

  return result;
}

/**
 * Create the Express Router for all job API endpoints.
 *
 * @param io - Socket.IO server instance for real-time events
 * @returns Express Router with job routes mounted
 */
export function createJobRoutes( io: SocketIOServer ): Router {
  const router = Router();

  // POST /api/jobs - Create a new job
  router.post( '/', async ( req: Request, res: Response ) => {
    try {
      const body: CreateJobRequest = req.body;

      if ( !body.type || !JOB_TYPES.includes( body.type ) ) {
        res.status( 400 ).json({ error: `Invalid type. Must be one of: ${ JOB_TYPES.join( ', ' ) }` });
        return;
      }
      if ( !body.sunday_date || !/^\d{8}$/.test( body.sunday_date ) ) {
        res.status( 400 ).json({ error: 'Missing or invalid sunday_date (YYYYMMDD required)' });
        return;
      }
      if ( !body.input_path ) {
        res.status( 400 ).json({ error: 'input_path is required' });
        return;
      }

      const pool = getPool();

      // Check for duplicate pending/processing job of same type + sunday_date
      const [ existing ] = await pool.query<RowDataPacket[]>(
        `SELECT id FROM jobs WHERE type = ? AND sunday_date = ? AND status IN ('pending', 'queued', 'processing')`,
        [ body.type, body.sunday_date ]
      );
      if ( existing.length > 0 ) {
        res.status( 409 ).json({ error: 'A job of this type already exists for this date', existing_id: existing[ 0 ].id });
        return;
      }

      const [ result ] = await pool.query<ResultSetHeader>(
        `INSERT INTO jobs (type, status, priority, sunday_date, input_path, output_path, metadata, parent_job_id, max_retries)
         VALUES (?, 'pending', ?, ?, ?, ?, ?, ?, ?)`,
        [
          body.type,
          body.priority ?? 5,
          body.sunday_date,
          body.input_path,
          body.output_path ?? null,
          body.metadata ? JSON.stringify( body.metadata ) : null,
          body.parent_job_id ?? null,
          body.max_retries ?? 3,
        ]
      );

      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ result.insertId ] );
      const job = parseJobRow( rows[ 0 ] );

      io.emit( 'job:created', job );
      console.log( `[JOBS] Created job #${ job.id } (${ job.type }) for ${ job.sunday_date }` );

      res.status( 201 ).json( job );
    } catch ( err ) {
      console.error( '[JOBS] Error creating job:', err );
      res.status( 500 ).json({ error: 'Failed to create job' });
    }
  } );

  // GET /api/jobs - List/filter jobs
  router.get( '/', async ( req: Request, res: Response ) => {
    try {
      const pool = getPool();
      const query: JobListQuery = {
        status: req.query.status as string,
        type: req.query.type as string,
        sunday_date: req.query.sunday_date as string,
        limit: parseInt( req.query.limit as string ) || 50,
        offset: parseInt( req.query.offset as string ) || 0,
      };

      let where = 'WHERE 1=1';
      const params: unknown[] = [];

      if ( query.status ) {
        const statuses = query.status.split( ',' ).filter( s => JOB_STATUSES.includes( s ) );
        if ( statuses.length > 0 ) {
          where += ` AND status IN (${ statuses.map( () => '?' ).join( ',' ) })`;
          params.push( ...statuses );
        }
      }
      if ( query.type ) {
        const types = query.type.split( ',' ).filter( t => JOB_TYPES.includes( t ) );
        if ( types.length > 0 ) {
          where += ` AND type IN (${ types.map( () => '?' ).join( ',' ) })`;
          params.push( ...types );
        }
      }
      if ( query.sunday_date ) {
        where += ' AND sunday_date = ?';
        params.push( query.sunday_date );
      }

      const [ countRows ] = await pool.query<RowDataPacket[]>(
        `SELECT COUNT(*) as total FROM jobs ${ where }`, params
      );
      const total = countRows[ 0 ].total;

      const [ rows ] = await pool.query<RowDataPacket[]>(
        `SELECT * FROM jobs ${ where } ORDER BY priority ASC, created_at DESC LIMIT ? OFFSET ?`,
        [ ...params, query.limit, query.offset ]
      );

      res.json({
        jobs: rows.map( parseJobRow ),
        total,
        limit: query.limit,
        offset: query.offset,
      });
    } catch ( err ) {
      console.error( '[JOBS] Error listing jobs:', err );
      res.status( 500 ).json({ error: 'Failed to list jobs' });
    }
  } );

  // GET /api/jobs/workers - List connected workers
  // (must be before /:id to avoid treating "workers" as a job ID)
  router.get( '/workers', ( _req: Request, res: Response ) => {
    res.json({ workers: getWorkers() });
  } );

  // GET /api/jobs/:id - Get job details with recent logs
  router.get( '/:id', async ( req: Request, res: Response ) => {
    try {
      const pool = getPool();
      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );

      if ( rows.length === 0 ) {
        res.status( 404 ).json({ error: 'Job not found' });
        return;
      }

      const [ logRows ] = await pool.query<RowDataPacket[]>(
        'SELECT * FROM job_logs WHERE job_id = ? ORDER BY created_at DESC LIMIT 100',
        [ req.params.id ]
      );

      res.json({
        ...parseJobRow( rows[ 0 ] ),
        logs: logRows as JobLog[],
      });
    } catch ( err ) {
      console.error( '[JOBS] Error getting job:', err );
      res.status( 500 ).json({ error: 'Failed to get job' });
    }
  } );

  // PUT /api/jobs/:id - Update job (priority, cancel)
  router.put( '/:id', async ( req: Request, res: Response ) => {
    try {
      const pool = getPool();
      const { priority, status } = req.body;

      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      if ( rows.length === 0 ) {
        res.status( 404 ).json({ error: 'Job not found' });
        return;
      }

      const updates: string[] = [];
      const params: unknown[] = [];

      if ( priority !== undefined ) {
        if ( typeof priority !== 'number' || priority < 1 || priority > 10 ) {
          res.status( 400 ).json({ error: 'Priority must be a number between 1 and 10' });
          return;
        }
        updates.push( 'priority = ?' );
        params.push( priority );
      }

      if ( status === 'cancelled' ) {
        const currentStatus = rows[ 0 ].status;
        if ( currentStatus === 'completed' || currentStatus === 'cancelled' ) {
          res.status( 400 ).json({ error: `Cannot cancel a job with status '${ currentStatus }'` });
          return;
        }
        updates.push( 'status = ?' );
        params.push( 'cancelled' );
      } else if ( status !== undefined ) {
        res.status( 400 ).json({ error: 'Only status change to "cancelled" is allowed via PUT' });
        return;
      }

      if ( updates.length === 0 ) {
        res.status( 400 ).json({ error: 'No valid fields to update' });
        return;
      }

      params.push( req.params.id );
      await pool.query( `UPDATE jobs SET ${ updates.join( ', ' ) } WHERE id = ?`, params );

      const [ updated ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      const job = parseJobRow( updated[ 0 ] );

      io.emit( 'job:updated', job );
      res.json( job );
    } catch ( err ) {
      console.error( '[JOBS] Error updating job:', err );
      res.status( 500 ).json({ error: 'Failed to update job' });
    }
  } );

  // POST /api/jobs/next - Worker claims next available job (atomic)
  router.post( '/next', async ( req: Request, res: Response ) => {
    try {
      const body: ClaimJobRequest = req.body;

      if ( !body.worker_id ) {
        res.status( 400 ).json({ error: 'worker_id is required' });
        return;
      }
      if ( !body.types || !Array.isArray( body.types ) || body.types.length === 0 ) {
        res.status( 400 ).json({ error: 'types array is required and must not be empty' });
        return;
      }

      const validTypes = body.types.filter( t => JOB_TYPES.includes( t ) );
      if ( validTypes.length === 0 ) {
        res.status( 400 ).json({ error: `Invalid types. Must be from: ${ JOB_TYPES.join( ', ' ) }` });
        return;
      }

      const pool = getPool();
      const connection = await pool.getConnection();

      try {
        await connection.beginTransaction();

        const placeholders = validTypes.map( () => '?' ).join( ',' );
        const [ rows ] = await connection.query<RowDataPacket[]>(
          `SELECT * FROM jobs
           WHERE status = 'pending' AND type IN (${ placeholders })
           ORDER BY priority ASC, created_at ASC
           LIMIT 1
           FOR UPDATE SKIP LOCKED`,
          validTypes
        );

        if ( rows.length === 0 ) {
          await connection.commit();
          touchWorker( body.worker_id, validTypes, null, null );
          res.status( 204 ).send();
          return;
        }

        const jobRow = rows[ 0 ];
        await connection.query(
          `UPDATE jobs SET status = 'processing', worker_id = ?, started_at = NOW(), heartbeat_at = NOW() WHERE id = ?`,
          [ body.worker_id, jobRow.id ]
        );

        await connection.commit();

        const [ updated ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ jobRow.id ] );
        const job = parseJobRow( updated[ 0 ] );

        touchWorker( body.worker_id, validTypes, job.id, job.type );
        io.emit( 'job:updated', job );
        console.log( `[JOBS] Worker '${ body.worker_id }' claimed job #${ job.id } (${ job.type })` );

        res.json({ job });
      } catch ( err ) {
        await connection.rollback();
        throw err;
      } finally {
        connection.release();
      }
    } catch ( err ) {
      console.error( '[JOBS] Error claiming job:', err );
      res.status( 500 ).json({ error: 'Failed to claim job' });
    }
  } );

  // POST /api/jobs/:id/heartbeat - Worker sends heartbeat
  router.post( '/:id/heartbeat', async ( req: Request, res: Response ) => {
    try {
      const body: HeartbeatRequest = req.body;
      const pool = getPool();

      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      if ( rows.length === 0 ) {
        res.status( 404 ).json({ error: 'Job not found' });
        return;
      }

      const job = rows[ 0 ];
      if ( job.worker_id !== body.worker_id ) {
        res.status( 403 ).json({ error: 'Worker does not own this job' });
        return;
      }

      const updateFields = [ 'heartbeat_at = NOW()' ];
      const params: unknown[] = [];

      if ( body.progress !== undefined ) {
        updateFields.push( "metadata = JSON_SET( COALESCE(metadata, '{}'), '$.progress', ? )" );
        params.push( body.progress );
      }

      params.push( req.params.id );
      await pool.query( `UPDATE jobs SET ${ updateFields.join( ', ' ) } WHERE id = ?`, params );

      touchWorker( body.worker_id, undefined, job.id, job.type );

      // Tell the worker to stop if the job was cancelled
      const shouldContinue = job.status === 'processing';

      res.json({ continue: shouldContinue });
    } catch ( err ) {
      console.error( '[JOBS] Error processing heartbeat:', err );
      res.status( 500 ).json({ error: 'Failed to process heartbeat' });
    }
  } );

  // POST /api/jobs/:id/complete - Worker marks job as completed
  router.post( '/:id/complete', async ( req: Request, res: Response ) => {
    try {
      const body: CompleteJobRequest = req.body;
      const pool = getPool();

      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      if ( rows.length === 0 ) {
        res.status( 404 ).json({ error: 'Job not found' });
        return;
      }

      const job = rows[ 0 ];
      if ( job.worker_id !== body.worker_id ) {
        res.status( 403 ).json({ error: 'Worker does not own this job' });
        return;
      }
      if ( job.status !== 'processing' ) {
        res.status( 400 ).json({ error: `Cannot complete a job with status '${ job.status }'` });
        return;
      }

      const updateFields = [ "status = 'completed'", 'completed_at = NOW()' ];
      const params: unknown[] = [];

      if ( body.output_path ) {
        updateFields.push( 'output_path = ?' );
        params.push( body.output_path );
      }

      if ( body.metadata ) {
        // Merge new metadata with existing
        updateFields.push( "metadata = JSON_MERGE_PATCH( COALESCE(metadata, '{}'), ? )" );
        params.push( JSON.stringify( body.metadata ) );
      }

      params.push( req.params.id );
      await pool.query( `UPDATE jobs SET ${ updateFields.join( ', ' ) } WHERE id = ?`, params );

      const [ updated ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      const completedJob = parseJobRow( updated[ 0 ] );

      workerJobCompleted( body.worker_id );
      io.emit( 'job:updated', completedJob );
      console.log( `[JOBS] Job #${ completedJob.id } (${ completedJob.type }) completed` );

      // Trigger job chaining
      await handleJobChaining( completedJob, io );

      res.json({ success: true });
    } catch ( err ) {
      console.error( '[JOBS] Error completing job:', err );
      res.status( 500 ).json({ error: 'Failed to complete job' });
    }
  } );

  // POST /api/jobs/:id/fail - Worker marks job as failed
  router.post( '/:id/fail', async ( req: Request, res: Response ) => {
    try {
      const body: FailJobRequest = req.body;
      const pool = getPool();

      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      if ( rows.length === 0 ) {
        res.status( 404 ).json({ error: 'Job not found' });
        return;
      }

      const job = rows[ 0 ];
      if ( job.worker_id !== body.worker_id ) {
        res.status( 403 ).json({ error: 'Worker does not own this job' });
        return;
      }

      const willRetry = job.retry_count < job.max_retries;

      if ( willRetry ) {
        await pool.query(
          `UPDATE jobs SET status = 'pending', worker_id = NULL, started_at = NULL, heartbeat_at = NULL,
           retry_count = retry_count + 1, error_message = ? WHERE id = ?`,
          [ body.error_message, req.params.id ]
        );
      } else {
        await pool.query(
          `UPDATE jobs SET status = 'failed', error_message = ?, completed_at = NOW() WHERE id = ?`,
          [ body.error_message, req.params.id ]
        );
      }

      const [ updated ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ req.params.id ] );
      const failedJob = parseJobRow( updated[ 0 ] );

      workerJobFailed( body.worker_id );
      io.emit( 'job:updated', failedJob );
      console.log( `[JOBS] Job #${ failedJob.id } (${ failedJob.type }) failed: ${ body.error_message }${ willRetry ? ` (will retry, attempt ${ failedJob.retry_count }/${ failedJob.max_retries })` : ' (no retries left)' }` );

      res.json({ success: true, will_retry: willRetry, retry_count: failedJob.retry_count });
    } catch ( err ) {
      console.error( '[JOBS] Error failing job:', err );
      res.status( 500 ).json({ error: 'Failed to fail job' });
    }
  } );

  // POST /api/jobs/:id/logs - Worker sends log entries
  router.post( '/:id/logs', async ( req: Request, res: Response ) => {
    try {
      const entries: JobLogEntry[] = req.body.entries;
      if ( !Array.isArray( entries ) || entries.length === 0 ) {
        res.status( 400 ).json({ error: 'entries array is required' });
        return;
      }

      const pool = getPool();

      const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT id FROM jobs WHERE id = ?', [ req.params.id ] );
      if ( rows.length === 0 ) {
        res.status( 404 ).json({ error: 'Job not found' });
        return;
      }

      const values = entries.map( e => [ req.params.id, e.level || 'info', e.message ] );
      await pool.query(
        'INSERT INTO job_logs (job_id, level, message) VALUES ?',
        [ values ]
      );

      io.emit( 'job:log', { job_id: parseInt( req.params.id ), entries } );
      res.json({ success: true, count: entries.length });
    } catch ( err ) {
      console.error( '[JOBS] Error writing job logs:', err );
      res.status( 500 ).json({ error: 'Failed to write job logs' });
    }
  } );

  return router;
}

// --- Job Chaining ---

/**
 * Handle automatic job chaining after a job completes.
 * Transcription completion triggers claude-processing creation.
 *
 * @param job - the completed job
 * @param io - Socket.IO server for emitting events
 */
async function handleJobChaining( job: Job, io: SocketIOServer ): Promise<void> {
  const pool = getPool();

  switch ( job.type ) {
    case 'video-alignment': {
      // The alignment job produces scene_changes. The server will generate a Kdenlive project.
      // No auto-chain — user must manually edit + render the production MP4.
      // scanFolders() will detect the production MP4 and create a transcription job.
      console.log( `[JOBS] Video alignment complete for ${ job.sunday_date }. Kdenlive project should be generated.` );
      break;
    }

    case 'transcode': {
      // Transcode complete → production MP4 exists → create transcription job
      const outputPath = job.output_path;
      if ( !outputPath ) {
        console.error( `[JOBS] Cannot chain transcription: no output_path for transcode job #${ job.id }` );
        return;
      }

      const [ existingTranscription ] = await pool.query<RowDataPacket[]>(
        `SELECT id FROM jobs WHERE type = 'transcription' AND sunday_date = ? AND status IN ('pending', 'queued', 'processing')`,
        [ job.sunday_date ]
      );

      if ( existingTranscription.length > 0 ) {
        console.log( `[JOBS] Transcription job already exists for ${ job.sunday_date }, skipping chain` );
        return;
      }

      const vttOutput = outputPath.replace( /\.mp4$/i, '.vtt' );

      const [ transcodeResult ] = await pool.query<ResultSetHeader>(
        `INSERT INTO jobs (type, status, priority, sunday_date, input_path, output_path, metadata, parent_job_id)
         VALUES ('transcription', 'pending', ?, ?, ?, ?, ?, ?)`,
        [
          job.priority,
          job.sunday_date,
          outputPath,
          vttOutput,
          JSON.stringify({ whisper_model: 'large-v3', language: 'en' }),
          job.id,
        ]
      );

      const [ transcodeRows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ transcodeResult.insertId ] );
      const chainedJob = parseJobRow( transcodeRows[ 0 ] );
      io.emit( 'job:created', chainedJob );
      console.log( `[JOBS] Chained transcription job #${ chainedJob.id } from transcode job #${ job.id }` );
      break;
    }

    case 'transcription': {
      // Auto-create claude-processing job
      const metadata = job.metadata || {};
      const vttPath = ( metadata as Record<string, unknown> ).vtt_path || job.output_path;

      if ( !vttPath ) {
        console.error( `[JOBS] Cannot chain claude-processing: no VTT path for job #${ job.id }` );
        return;
      }

      // Determine output path for notes
      const outputDirectory = process.env.OUTPUT_DIRECTORY || '/output';
      const notesPath = require( 'path' ).join( outputDirectory, job.sunday_date, `${ job.sunday_date } Notes.txt` );

      const [ existing ] = await pool.query<RowDataPacket[]>(
        `SELECT id FROM jobs WHERE type = 'claude-processing' AND sunday_date = ? AND status IN ('pending', 'queued', 'processing')`,
        [ job.sunday_date ]
      );

      if ( existing.length > 0 ) {
        console.log( `[JOBS] Claude-processing job already exists for ${ job.sunday_date }, skipping chain` );
        return;
      }

      const [ result ] = await pool.query<ResultSetHeader>(
        `INSERT INTO jobs (type, status, priority, sunday_date, input_path, output_path, metadata, parent_job_id)
         VALUES ('claude-processing', 'pending', ?, ?, ?, ?, ?, ?)`,
        [
          job.priority,
          job.sunday_date,
          vttPath,
          notesPath,
          JSON.stringify({ prompt_template: 'sermon-notes' }),
          job.id,
        ]
      );

      const [ newRows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ result.insertId ] );
      const newJob = parseJobRow( newRows[ 0 ] );
      io.emit( 'job:created', newJob );
      console.log( `[JOBS] Chained claude-processing job #${ newJob.id } from transcription job #${ job.id }` );
      break;
    }

    case 'claude-processing': {
      // Terminal job — no auto-chain
      console.log( `[JOBS] Claude processing complete for ${ job.sunday_date }` );
      break;
    }
  }
}

// --- Stale Job Recovery ---

/**
 * Recover jobs with stale heartbeats by retrying or marking as failed.
 *
 * @param io - Socket.IO server for emitting events
 */
export async function recoverStaleJobs( io: SocketIOServer ): Promise<void> {
  try {
    const pool = getPool();
    const timeoutMs = parseInt( process.env.HEARTBEAT_TIMEOUT_MS || '120000', 10 );
    const cutoff = new Date( Date.now() - timeoutMs );

    const [ staleRows ] = await pool.query<RowDataPacket[]>(
      `SELECT id, retry_count, max_retries FROM jobs WHERE status = 'processing' AND heartbeat_at < ?`,
      [ cutoff ]
    );

    for ( const row of staleRows ) {
      if ( row.retry_count < row.max_retries ) {
        await pool.query(
          `UPDATE jobs SET status = 'pending', worker_id = NULL, started_at = NULL, heartbeat_at = NULL,
           retry_count = retry_count + 1,
           error_message = CONCAT( COALESCE(error_message, ''), '\nRecovered from stale heartbeat at ', NOW() )
           WHERE id = ?`,
          [ row.id ]
        );
        console.log( `[JOBS] Recovered stale job #${ row.id } (retry ${ row.retry_count + 1 }/${ row.max_retries })` );
      } else {
        await pool.query(
          `UPDATE jobs SET status = 'failed',
           error_message = CONCAT( COALESCE(error_message, ''), '\nFailed: heartbeat timeout after max retries' ),
           completed_at = NOW()
           WHERE id = ?`,
          [ row.id ]
        );
        console.log( `[JOBS] Job #${ row.id } failed after max retries (heartbeat timeout)` );
      }

      const [ updated ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ row.id ] );
      io.emit( 'job:updated', parseJobRow( updated[ 0 ] ) );
    }
  } catch ( err ) {
    console.error( '[JOBS] Error recovering stale jobs:', err );
  }
}

// --- Auto-detection helper ---

/**
 * Create a job only if no pending/queued/processing job of the same type and date exists.
 *
 * @param data - job creation request data
 * @returns the created job, or null if a duplicate exists
 */
export async function createJobIfNotExists( data: CreateJobRequest ): Promise<Job | null> {
  const pool = getPool();

  const [ existing ] = await pool.query<RowDataPacket[]>(
    `SELECT id FROM jobs WHERE type = ? AND sunday_date = ? AND status IN ('pending', 'queued', 'processing')`,
    [ data.type, data.sunday_date ]
  );

  if ( existing.length > 0 ) return null;

  const [ result ] = await pool.query<ResultSetHeader>(
    `INSERT INTO jobs (type, status, priority, sunday_date, input_path, output_path, metadata, parent_job_id, max_retries)
     VALUES (?, 'pending', ?, ?, ?, ?, ?, ?, ?)`,
    [
      data.type,
      data.priority ?? 5,
      data.sunday_date,
      data.input_path,
      data.output_path ?? null,
      data.metadata ? JSON.stringify( data.metadata ) : null,
      data.parent_job_id ?? null,
      data.max_retries ?? 3,
    ]
  );

  const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT * FROM jobs WHERE id = ?', [ result.insertId ] );
  return parseJobRow( rows[ 0 ] );
}

// --- Helpers ---

/**
 * Parse a raw database row into a typed Job object.
 *
 * @param row - raw database row
 * @returns parsed Job object
 */
function parseJobRow( row: RowDataPacket ): Job {
  return {
    ...row,
    metadata: typeof row.metadata === 'string' ? JSON.parse( row.metadata ) : row.metadata,
  } as Job;
}
