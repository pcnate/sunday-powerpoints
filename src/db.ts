import mysql, { Pool, PoolConnection, RowDataPacket, ResultSetHeader } from 'mysql2/promise';
import * as fs from 'fs';
import * as path from 'path';

let pool: Pool | null = null;

export interface DbConfig {
  host: string;
  port: number;
  user: string;
  password: string;
  database: string;
}

export function getDbConfig(): DbConfig {
  return {
    host: process.env.MYSQL_HOST || 'localhost',
    port: parseInt( process.env.MYSQL_PORT || '3306', 10 ),
    user: process.env.MYSQL_USER || 'root',
    password: process.env.MYSQL_PASSWORD || '',
    database: process.env.MYSQL_DATABASE || 'sunday_powerpoints',
  };
}

export async function initDb(): Promise<Pool> {
  if ( pool ) return pool;

  const config = getDbConfig();
  pool = mysql.createPool({
    host: config.host,
    port: config.port,
    user: config.user,
    password: config.password,
    database: config.database,
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    enableKeepAlive: true,
    keepAliveInitialDelay: 10000,
  });

  // Verify connection
  const connection = await pool.getConnection();
  console.log( `[DB] Connected to MySQL at ${ config.host }:${ config.port }/${ config.database }` );
  connection.release();

  // Run pending migrations
  await runMigrations( pool );

  return pool;
}

export function getPool(): Pool {
  if ( !pool ) throw new Error( 'Database not initialized. Call initDb() first.' );
  return pool;
}

export async function closeDb(): Promise<void> {
  if ( pool ) {
    await pool.end();
    pool = null;
    console.log( '[DB] Connection pool closed' );
  }
}

export async function healthCheck(): Promise<boolean> {
  try {
    const p = getPool();
    await p.query( 'SELECT 1' );
    return true;
  } catch {
    return false;
  }
}

// --- Migration Runner ---

async function ensureMigrationsTable( connection: PoolConnection ): Promise<void> {
  await connection.query( `
    CREATE TABLE IF NOT EXISTS migrations (
      id         INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
      name       VARCHAR(255) NOT NULL UNIQUE,
      applied_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
  ` );
}

async function getAppliedMigrations( connection: PoolConnection ): Promise<Set<string>> {
  const [rows] = await connection.query<RowDataPacket[]>( 'SELECT name FROM migrations ORDER BY id' );
  return new Set( rows.map( row => row.name ) );
}

async function runMigrations( p: Pool ): Promise<void> {
  const migrationsDir = path.join( __dirname, 'migrations' );

  if ( !fs.existsSync( migrationsDir ) ) {
    console.log( '[DB] No migrations directory found, skipping migrations' );
    return;
  }

  const files = fs.readdirSync( migrationsDir )
    .filter( f => f.endsWith( '.sql' ) )
    .sort();

  if ( files.length === 0 ) {
    console.log( '[DB] No migration files found' );
    return;
  }

  const connection = await p.getConnection();
  try {
    await ensureMigrationsTable( connection );
    const applied = await getAppliedMigrations( connection );

    for ( const file of files ) {
      const migrationName = file.replace( /\.sql$/, '' );

      if ( applied.has( migrationName ) ) {
        continue;
      }

      console.log( `[DB] Applying migration: ${ migrationName }` );
      const sql = fs.readFileSync( path.join( migrationsDir, file ), 'utf-8' );

      // Split on semicolons to execute multiple statements
      // Filter out empty statements and the migrations table creation (handled above)
      const statements = sql
        .split( /;\s*$/m )
        .map( s => s.trim() )
        .filter( s => s.length > 0 )
        .filter( s => !s.includes( 'CREATE TABLE IF NOT EXISTS migrations' ) );

      await connection.beginTransaction();
      try {
        for ( const statement of statements ) {
          await connection.query( statement );
        }
        await connection.query(
          'INSERT INTO migrations (name) VALUES (?)',
          [migrationName]
        );
        await connection.commit();
        console.log( `[DB] Migration applied: ${ migrationName }` );
      } catch ( err ) {
        await connection.rollback();
        console.error( `[DB] Migration failed: ${ migrationName }`, err );
        throw err;
      }
    }

    console.log( `[DB] All migrations up to date (${ files.length } total, ${ files.length - applied.size } applied)` );
  } finally {
    connection.release();
  }
}
