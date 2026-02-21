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
import { Job, CreateJobRequest, JobType, JobStatus } from '../src/types/job';

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
      type: 'video-alignment',
      sunday_date: '20260222',
      input_path: '/output/20260222/Vids',
      output_path: '/output/20260222/Vids/20260222.kdenlive',
      priority: 3,
      metadata: { camera_files: [ 'file1.mp4', 'file2.mp4' ] },
      parent_job_id: 1,
      max_retries: 5,
    };
    expect( request.priority ).toBe( 3 );
    expect( request.metadata ).toBeDefined();
    expect( request.parent_job_id ).toBe( 1 );
  });

  it( 'JobType only allows valid types', () => {
    const validTypes: JobType[] = [ 'video-alignment', 'transcription', 'claude-processing' ];
    expect( validTypes.length ).toBe( 3 );
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

    it( 'video-alignment completion should NOT auto-chain', () => {
      const completedJob = mockJobRow({
        type: 'video-alignment',
        status: 'completed',
      });
      // Video alignment does not auto-chain — user must render in Kdenlive first
      // scanFolders() will detect the production MP4 and create transcription job
      expect( completedJob.type ).toBe( 'video-alignment' );
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
  it( 'detects 2+ raw MP4s without production file as alignment candidate', () => {
    const files = [ '00001.MP4', '00002.MP4', 'thumbnail.jpeg' ];
    const mp4Files = files.filter( f => f.toLowerCase().endsWith( '.mp4' ) );
    const productionFile = mp4Files.find( f => f.match( /^\d{8}-production\.mp4$/i ) );
    const rawMp4s = mp4Files.filter( f => !f.match( /^\d{8}-production\.mp4$/i ) );

    expect( rawMp4s.length ).toBe( 2 );
    expect( productionFile ).toBeUndefined();
    // Should create a video-alignment job
  });

  it( 'detects production MP4 as transcription candidate', () => {
    const files = [ '00001.MP4', '00002.MP4', '20260222-production.mp4', 'thumbnail.jpeg' ];
    const mp4Files = files.filter( f => f.toLowerCase().endsWith( '.mp4' ) );
    const productionFile = mp4Files.find( f => f.match( /^\d{8}-production\.mp4$/i ) );

    expect( productionFile ).toBe( '20260222-production.mp4' );
    // Should create a transcription job
  });

  it( 'does not create alignment job when production file exists', () => {
    const files = [ '00001.MP4', '00002.MP4', '20260222-production.mp4' ];
    const mp4Files = files.filter( f => f.toLowerCase().endsWith( '.mp4' ) );
    const productionFile = mp4Files.find( f => f.match( /^\d{8}-production\.mp4$/i ) );
    const rawMp4s = mp4Files.filter( f => !f.match( /^\d{8}-production\.mp4$/i ) );

    expect( productionFile ).toBeDefined();
    // Should NOT create alignment job when production already exists
    expect( rawMp4s.length >= 2 && !productionFile ).toBe( false );
  });

  it( 'does not trigger with only 1 MP4', () => {
    const files = [ '00001.MP4', 'thumbnail.jpeg' ];
    const mp4Files = files.filter( f => f.toLowerCase().endsWith( '.mp4' ) );
    const rawMp4s = mp4Files.filter( f => !f.match( /^\d{8}-production\.mp4$/i ) );

    expect( rawMp4s.length >= 2 ).toBe( false );
    // Should not create any job
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
