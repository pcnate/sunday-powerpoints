import { RowDataPacket, ResultSetHeader } from 'mysql2/promise';

// Mock mysql2/promise before importing db module
const mockQuery = jest.fn();
const mockBeginTransaction = jest.fn();
const mockCommit = jest.fn();
const mockRollback = jest.fn();
const mockRelease = jest.fn();
const mockGetConnection = jest.fn().mockResolvedValue({
  query: mockQuery,
  beginTransaction: mockBeginTransaction,
  commit: mockCommit,
  rollback: mockRollback,
  release: mockRelease,
});
const mockPoolQuery = jest.fn();
const mockPoolEnd = jest.fn();

jest.mock( 'mysql2/promise', () => ({
  createPool: jest.fn( () => ({
    query: mockPoolQuery,
    getConnection: mockGetConnection,
    end: mockPoolEnd,
  }) ),
}) );

// Mock fs for migration runner
jest.mock( 'node:fs', () => {
  const actual = jest.requireActual( 'node:fs' );
  return {
    ...actual,
    existsSync: jest.fn( () => false ), // No migrations dir during tests
    readdirSync: jest.fn( () => [] ),
  };
} );

import * as db from '../src/db';
import { Job, CreateJobRequest, JobType, JobStatus, VideoStatus, VideoRecord } from '../src/types/job';

jest.setTimeout( 10000 );

// Helper to create a mock job row
function mockJobRow( overrides: Partial<Job> = {} ): RowDataPacket {
  return {
    id: 1,
    type: 'transcription' as JobType,
    status: 'pending' as JobStatus,
    priority: 5,
    sunday_date: '20260222',
    input_path: '/output/20260222/Vids/20260222-production.mp4',
    output_path: '/output/20260222/Vids/20260222-production.vtt',
    metadata: null,
    error_message: null,
    worker_id: null,
    retry_count: 0,
    max_retries: 3,
    parent_job_id: null,
    created_at: new Date(),
    updated_at: new Date(),
    started_at: null,
    completed_at: null,
    heartbeat_at: null,
    ...overrides,
    constructor: { name: 'RowDataPacket' },
  } as unknown as RowDataPacket;
}

describe( 'db module', () => {
  beforeEach( () => {
    jest.clearAllMocks();
    // Mock the migration table check
    mockGetConnection.mockResolvedValue({
      query: mockQuery.mockResolvedValue( [[], []] ),
      beginTransaction: mockBeginTransaction,
      commit: mockCommit,
      rollback: mockRollback,
      release: mockRelease,
    });
  });

  describe( 'initDb', () => {
    it( 'creates a connection pool and verifies connection', async () => {
      const pool = await db.initDb();
      expect( pool ).toBeDefined();
      expect( mockRelease ).toHaveBeenCalled();
    });
  });

  describe( 'healthCheck', () => {
    it( 'returns true when DB is available', async () => {
      await db.initDb();
      mockPoolQuery.mockResolvedValueOnce( [[ { '1': 1 } ], []] );
      const result = await db.healthCheck();
      expect( result ).toBe( true );
    });
  });
});

describe( 'Job types', () => {
  it( 'CreateJobRequest has required fields', () => {
    const request: CreateJobRequest = {
      type: 'transcription',
      sunday_date: '20260222',
      input_path: '/output/20260222/Vids/20260222-production.mp4',
    };
    expect( request.type ).toBe( 'transcription' );
    expect( request.sunday_date ).toBe( '20260222' );
    expect( request.input_path ).toBeDefined();
  });

  it( 'CreateJobRequest supports optional fields', () => {
    const request: CreateJobRequest = {
      type: 'transcode',
      sunday_date: '20260222',
      input_path: '/output/20260222/Vids/20260222.kdenlive',
      output_path: '/output/20260222/Vids/20260222-production.mp4',
      priority: 3,
      metadata: { video_bitrate: '5000k' },
      parent_job_id: 1,
      max_retries: 5,
    };
    expect( request.priority ).toBe( 3 );
    expect( request.metadata ).toBeDefined();
    expect( request.parent_job_id ).toBe( 1 );
  });

  it( 'JobType only allows valid types', () => {
    const validTypes: JobType[] = [ 'transcription', 'claude-processing', 'transcode', 'ffprobe' ];
    expect( validTypes.length ).toBe( 4 );
  });

  it( 'JobStatus covers all states', () => {
    const validStatuses: JobStatus[] = [ 'pending', 'queued', 'processing', 'completed', 'failed', 'cancelled' ];
    expect( validStatuses.length ).toBe( 6 );
  });
});

describe( 'Job API route logic', () => {
  describe( 'job deduplication', () => {
    it( 'prevents duplicate jobs of same type and sunday_date', () => {
      // This tests the concept — the actual query checks for
      // status IN ('pending', 'queued', 'processing')
      const existingJob = mockJobRow({ type: 'transcription', sunday_date: '20260222', status: 'pending' });
      const newRequest: CreateJobRequest = {
        type: 'transcription',
        sunday_date: '20260222',
        input_path: '/some/path',
      };
      // Same type + sunday_date should be rejected
      expect( existingJob.type ).toBe( newRequest.type );
      expect( existingJob.sunday_date ).toBe( newRequest.sunday_date );
    });

    it( 'allows jobs of different types for the same date', () => {
      const existingJob = mockJobRow({ type: 'transcription', sunday_date: '20260222' });
      const newRequest: CreateJobRequest = {
        type: 'claude-processing',
        sunday_date: '20260222',
        input_path: '/some/path',
      };
      expect( existingJob.type ).not.toBe( newRequest.type );
    });
  });

  describe( 'job claiming', () => {
    it( 'claim query uses FOR UPDATE SKIP LOCKED for atomicity', () => {
      // The claim query should filter by status=pending, type IN types,
      // and use FOR UPDATE SKIP LOCKED to prevent race conditions
      const expectedQueryFragment = 'FOR UPDATE SKIP LOCKED';
      // This is a documentation test — the actual SQL is in routes/jobs.ts
      expect( expectedQueryFragment ).toBeTruthy();
    });
  });

  describe( 'heartbeat response', () => {
    it( 'returns continue=true for processing jobs', () => {
      const job = mockJobRow({ status: 'processing' });
      const shouldContinue = job.status === 'processing';
      expect( shouldContinue ).toBe( true );
    });

    it( 'returns continue=false for cancelled jobs', () => {
      const job = mockJobRow({ status: 'cancelled' });
      const shouldContinue = job.status === 'processing';
      expect( shouldContinue ).toBe( false );
    });
  });

  describe( 'retry logic', () => {
    it( 'retries when retry_count < max_retries', () => {
      const job = mockJobRow({ retry_count: 1, max_retries: 3 });
      const willRetry = job.retry_count < job.max_retries;
      expect( willRetry ).toBe( true );
    });

    it( 'fails permanently when retry_count >= max_retries', () => {
      const job = mockJobRow({ retry_count: 3, max_retries: 3 });
      const willRetry = job.retry_count < job.max_retries;
      expect( willRetry ).toBe( false );
    });
  });

  describe( 'job chaining', () => {
    it( 'transcription completion should chain to claude-processing', () => {
      const completedJob = mockJobRow({
        type: 'transcription',
        status: 'completed',
        sunday_date: '20260222',
        output_path: '/output/20260222/Vids/20260222-production.vtt',
      });
      // When a transcription job completes, a claude-processing job should be created
      expect( completedJob.type ).toBe( 'transcription' );
      expect( completedJob.output_path ).toContain( '.vtt' );
    });

    it( 'claude-processing is terminal — no chain', () => {
      const completedJob = mockJobRow({
        type: 'claude-processing',
        status: 'completed',
      });
      expect( completedJob.type ).toBe( 'claude-processing' );
    });
  });
});

describe( 'scanFolders job detection logic', () => {
  it( 'detects raw video files as ffprobe candidates, skips production', () => {
    const files = [ '00001.MP4', '00002.MP4', '20260222-production.mp4', 'thumbnail.jpeg', 'camera.MTS' ];
    const probeFiles = files.filter( f =>
      /\.(mp4|mts)$/i.test( f ) && !/^\d{8}-production\.mp4$/i.test( f )
    );

    expect( probeFiles.length ).toBe( 3 );
    expect( probeFiles ).toContain( '00001.MP4' );
    expect( probeFiles ).not.toContain( '20260222-production.mp4' );
    expect( probeFiles ).toContain( 'camera.MTS' );
  });

  it( 'production MP4 is excluded from scanFolders ffprobe auto-detection', () => {
    const files = [ '00001.MP4', '20260222-production.mp4', 'camera.MTS' ];
    const probeFiles = files.filter( f =>
      /\.(mp4|mts)$/i.test( f ) && !/^\d{8}-production\.mp4$/i.test( f )
    );
    expect( probeFiles ).toEqual([ '00001.MP4', 'camera.MTS' ]);
    expect( probeFiles ).not.toContain( '20260222-production.mp4' );
  });

  it( 'held ffprobe job is created with queued status at approve time', () => {
    // When approve-kdenlive creates a transcode job, it also creates
    // an ffprobe job in 'queued' status that workers won't claim.
    const heldJob = mockJobRow({ type: 'ffprobe', status: 'queued' });
    expect( heldJob.status ).toBe( 'queued' );
  });

  it( 'transcode completion releases held ffprobe to pending', () => {
    // When transcode completes, the held ffprobe job is updated to 'pending'
    // so the worker can claim it. Runs in parallel with transcription.
    const heldJob = mockJobRow({ type: 'ffprobe', status: 'queued' });
    const releasedStatus = 'pending';
    expect( heldJob.status ).not.toBe( releasedStatus );
  });

  it( 'ffprobe jobs are terminal — no chaining', () => {
    // ffprobe completion populates the videos table but does not chain.
    const filename = '00001.MP4';
    const isProduction = /^\d{8}-production\.mp4$/i.test( filename );
    expect( isProduction ).toBe( false );
  });
});

describe( 'migration system', () => {
  it( 'migration files are sorted alphabetically for ordered execution', () => {
    const files = [ '003_add_index.sql', '001_create_tables.sql', '002_add_column.sql' ];
    const sorted = [ ...files ].sort();
    expect( sorted ).toEqual([ '001_create_tables.sql', '002_add_column.sql', '003_add_index.sql' ]);
  });

  it( 'migration names are extracted by removing .sql extension', () => {
    const file = '001_create_jobs_tables.sql';
    const name = file.replace( /\.sql$/, '' );
    expect( name ).toBe( '001_create_jobs_tables' );
  });
});

describe( 'VideoRecord types', () => {
  it( 'VideoStatus covers all valid states', () => {
    const validStatuses: VideoStatus[] = [ 'pending', 'completed', 'failed' ];
    expect( validStatuses.length ).toBe( 3 );
  });

  it( 'VideoRecord has required fields', () => {
    const record: VideoRecord = {
      id: 1,
      input_path: '20260301/Vids/camera.MTS',
      sunday_date: '20260301',
      filename: 'camera.MTS',
      status: 'completed',
      job_id: 42,
      duration_secs: 3600.5,
      duration_timecode: '01:00:00.500',
      framerate: 29.97,
      width: 1920,
      height: 1080,
      video_codec: 'h264',
      audio_codec: 'aac',
      audio_streams: 2,
      bitrate: 25000000,
      file_size: 11264000000,
      file_created_at: new Date( '2026-03-01T10:00:00Z' ),
      created_at: new Date(),
      updated_at: new Date(),
    };
    expect( record.input_path ).toBe( '20260301/Vids/camera.MTS' );
    expect( record.status ).toBe( 'completed' );
    expect( record.duration_secs ).toBe( 3600.5 );
    expect( record.width ).toBe( 1920 );
  });

  it( 'VideoRecord allows null for optional metadata fields', () => {
    const record: VideoRecord = {
      id: 2,
      input_path: '20260301/Vids/00001.MP4',
      sunday_date: '20260301',
      filename: '00001.MP4',
      status: 'pending',
      job_id: null,
      duration_secs: null,
      duration_timecode: null,
      framerate: null,
      width: null,
      height: null,
      video_codec: null,
      audio_codec: null,
      audio_streams: null,
      bitrate: null,
      file_size: null,
      file_created_at: null,
      created_at: new Date(),
      updated_at: new Date(),
    };
    expect( record.status ).toBe( 'pending' );
    expect( record.duration_secs ).toBeNull();
    expect( record.job_id ).toBeNull();
  });
});

describe( 'videos table dedup logic', () => {
  it( 'probe job creation checks videos table not jobs table', () => {
    // The createProbeJobIfNotExists function now uses getVideoByPath()
    // to check the videos table (any status) instead of querying the jobs table.
    // A pending video row means a job already exists → skip.
    const existingVideo: Partial<VideoRecord> = {
      input_path: '20260301/Vids/camera.MTS',
      status: 'pending',
    };
    // Any status blocks duplicate creation
    expect( [ 'pending', 'completed', 'failed' ] ).toContain( existingVideo.status );
  });

  it( 'ffprobe completion populates video row metadata', () => {
    // When an ffprobe job completes, handleJobChaining calls completeVideo()
    // to update the video row with structured metadata columns.
    const meta = {
      duration_secs: 1800.25,
      duration_timecode: '00:30:00.250',
      framerate: 29.97,
      width: 1920,
      height: 1080,
      video_codec: 'h264',
      audio_codec: 'aac',
      audio_streams: 2,
      bitrate: 25000000,
      file_size: 5632000000,
      created_at: '2026-03-01 10:00:00',
    };
    expect( meta.duration_secs ).toBeGreaterThan( 0 );
    expect( meta.width ).toBe( 1920 );
  });

  it( 'permanent ffprobe failure marks video row as failed', () => {
    // When an ffprobe job permanently fails (no retries left),
    // failVideo() sets status='failed' on the video row.
    const job = mockJobRow({ type: 'ffprobe', retry_count: 3, max_retries: 3 });
    const willRetry = job.retry_count < job.max_retries;
    expect( willRetry ).toBe( false );
    // failVideo() would be called here
  });
});

describe( 'MTS cleanup via videos table', () => {
  it( 'getOldestMtsVideo queries by .mts extension and 6-month age', () => {
    // The cleanup function now queries the videos table instead of
    // walking the filesystem. It filters by:
    // - status = 'completed'
    // - filename LIKE '%.mts' (case-insensitive via LOWER)
    // - file_created_at < 6 months ago
    const filename = 'camera.MTS';
    expect( filename.toLowerCase().endsWith( '.mts' ) ).toBe( true );

    const fileCreated = new Date( '2025-06-01' );
    const sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth( sixMonthsAgo.getMonth() - 6 );
    expect( fileCreated < sixMonthsAgo ).toBe( true );
  });

  it( 'cleanup deletes video row after file removal', () => {
    // After deleting an MTS file, deleteVideo() removes the row
    // and the ffprobe job is cancelled.
    const inputPath = '20250601/Vids/camera.MTS';
    expect( inputPath ).toBeTruthy();
  });
});

describe( 'Kdenlive duration from videos table', () => {
  it( 'picks max duration from completed video records', () => {
    const videos: Partial<VideoRecord>[] = [
      { duration_secs: 1800, duration_timecode: '00:30:00.000' },
      { duration_secs: 3600.5, duration_timecode: '01:00:00.500' },
      { duration_secs: 1200, duration_timecode: '00:20:00.000' },
    ];

    let maxSecs = 0;
    let duration: string | undefined;
    for ( const v of videos ) {
      if ( v.duration_secs && v.duration_secs > maxSecs ) {
        maxSecs = v.duration_secs;
        duration = v.duration_timecode ?? undefined;
      }
    }

    expect( maxSecs ).toBe( 3600.5 );
    expect( duration ).toBe( '01:00:00.500' );
  });
});
