import { getPool, healthCheck } from './db';
import { RowDataPacket } from 'mysql2/promise';
import { VideoRecord } from './types/job';


/**
 * Metadata fields returned by an ffprobe job.
 */
export interface FfprobeMetadata {
  duration_secs?: number;
  duration_timecode?: string;
  framerate?: number;
  width?: number;
  height?: number;
  video_codec?: string;
  audio_codec?: string;
  audio_streams?: number;
  bitrate?: number;
  file_size?: number;
  created_at?: string;
}


/**
 * Check if a video row exists for a given input path (any status).
 *
 * @param inputPath - relative path to the video file
 * @returns the video record, or null if not found
 */
export async function getVideoByPath( inputPath: string ): Promise<VideoRecord | null> {
  if ( !await healthCheck() ) return null;
  const pool = getPool();

  const [ rows ] = await pool.query<RowDataPacket[]>(
    'SELECT * FROM videos WHERE input_path = ? LIMIT 1',
    [ inputPath ]
  );

  return rows.length > 0 ? rows[ 0 ] as VideoRecord : null;
}


/**
 * Insert a placeholder video row with status='pending'.
 *
 * @param sundayDate - YYYYMMDD folder date
 * @param inputPath - relative path to the video file
 * @param filename - just the filename portion
 * @param jobId - the ffprobe job ID
 */
export async function createPendingVideo(
  sundayDate: string,
  inputPath: string,
  filename: string,
  jobId: number
): Promise<void> {
  const pool = getPool();

  await pool.query(
    `INSERT INTO videos (sunday_date, input_path, filename, status, job_id)
     VALUES (?, ?, ?, 'pending', ?)
     ON DUPLICATE KEY UPDATE id = id`,
    [ sundayDate, inputPath, filename, jobId ]
  );
}


/**
 * Update a video row with ffprobe results and set status='completed'.
 *
 * @param inputPath - relative path to the video file
 * @param metadata - ffprobe metadata fields
 */
export async function completeVideo( inputPath: string, metadata: FfprobeMetadata ): Promise<void> {
  if ( !await healthCheck() ) return;
  const pool = getPool();

  await pool.query(
    `UPDATE videos SET
       status = 'completed',
       duration_secs = ?,
       duration_timecode = ?,
       framerate = ?,
       width = ?,
       height = ?,
       video_codec = ?,
       audio_codec = ?,
       audio_streams = ?,
       bitrate = ?,
       file_size = ?,
       file_created_at = ?
     WHERE input_path = ?`,
    [
      metadata.duration_secs ?? null,
      metadata.duration_timecode ?? null,
      metadata.framerate ?? null,
      metadata.width ?? null,
      metadata.height ?? null,
      metadata.video_codec ?? null,
      metadata.audio_codec ?? null,
      metadata.audio_streams ?? null,
      metadata.bitrate ?? null,
      metadata.file_size ?? null,
      metadata.created_at ?? null,
      inputPath,
    ]
  );
}


/**
 * Mark a video row as failed.
 *
 * @param inputPath - relative path to the video file
 */
export async function failVideo( inputPath: string ): Promise<void> {
  if ( !await healthCheck() ) return;
  const pool = getPool();

  await pool.query(
    `UPDATE videos SET status = 'failed' WHERE input_path = ?`,
    [ inputPath ]
  );
}


/**
 * Get completed video records for a list of input paths.
 * Used by the Kdenlive generator to look up durations.
 *
 * @param inputPaths - array of relative paths
 * @returns array of matching completed video records
 */
export async function getVideosByPaths( inputPaths: string[] ): Promise<VideoRecord[]> {
  if ( inputPaths.length === 0 ) return [];
  if ( !await healthCheck() ) return [];
  const pool = getPool();

  const placeholders = inputPaths.map( () => '?' ).join( ',' );
  const [ rows ] = await pool.query<RowDataPacket[]>(
    `SELECT * FROM videos WHERE input_path IN (${ placeholders }) AND status = 'completed'`,
    inputPaths
  );

  return rows as VideoRecord[];
}


/**
 * Get the oldest completed MTS video older than 6 months.
 * Used by the MTS cleanup scheduler to find deletion candidates.
 *
 * @returns the oldest MTS video record, or null if none qualify
 */
export async function getOldestMtsVideo(): Promise<VideoRecord | null> {
  if ( !await healthCheck() ) return null;
  const pool = getPool();

  const [ rows ] = await pool.query<RowDataPacket[]>(
    `SELECT * FROM videos
     WHERE status = 'completed'
       AND LOWER(filename) LIKE '%.mts'
       AND file_created_at < DATE_SUB( NOW(), INTERVAL 6 MONTH )
     ORDER BY file_created_at ASC
     LIMIT 1`
  );

  return rows.length > 0 ? rows[ 0 ] as VideoRecord : null;
}


/**
 * Delete a video row by input path.
 * Used after an MTS file is deleted from the filesystem.
 *
 * @param inputPath - relative path to the video file
 */
export async function deleteVideo( inputPath: string ): Promise<void> {
  if ( !await healthCheck() ) return;
  const pool = getPool();

  await pool.query(
    'DELETE FROM videos WHERE input_path = ?',
    [ inputPath ]
  );
}
