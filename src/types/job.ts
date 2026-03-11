export type JobType = 'transcription' | 'claude-processing' | 'transcode' | 'ffprobe';
export type JobStatus = 'pending' | 'queued' | 'processing' | 'completed' | 'failed' | 'cancelled';
export type LogLevel = 'info' | 'warn' | 'error' | 'debug';

export interface Job {
  id: number;
  type: JobType;
  status: JobStatus;
  priority: number;
  sunday_date: string;
  input_path: string;
  output_path: string | null;
  metadata: Record<string, unknown> | null;
  error_message: string | null;
  worker_id: string | null;
  retry_count: number;
  max_retries: number;
  parent_job_id: number | null;
  created_at: Date;
  updated_at: Date;
  started_at: Date | null;
  completed_at: Date | null;
  heartbeat_at: Date | null;
}

export interface JobLog {
  id: number;
  job_id: number;
  level: LogLevel;
  message: string;
  created_at: Date;
}

export interface CreateJobRequest {
  type: JobType;
  sunday_date: string;
  input_path: string;
  output_path?: string;
  priority?: number;
  metadata?: Record<string, unknown>;
  parent_job_id?: number;
  max_retries?: number;
}

export interface ClaimJobRequest {
  worker_id: string;
  types: JobType[];
}

export interface HeartbeatRequest {
  worker_id: string;
  progress?: number;
}

export interface CompleteJobRequest {
  worker_id: string;
  output_path?: string;
  metadata?: Record<string, unknown>;
}

export interface FailJobRequest {
  worker_id: string;
  error_message: string;
}

export interface JobLogEntry {
  level: LogLevel;
  message: string;
}

export interface JobListQuery {
  status?: string;
  type?: string;
  sunday_date?: string;
  limit?: number;
  offset?: number;
}

export interface JobListResponse {
  jobs: Job[];
  total: number;
  limit: number;
  offset: number;
}

export interface HeartbeatResponse {
  continue: boolean;
}

export interface WorkerInfo {
  id: string;
  types: string[];
  current_job_id: number | null;
  current_job_type: string | null;
  last_seen: Date;
  first_seen: Date;
  jobs_completed: number;
  jobs_failed: number;
  status: 'idle' | 'executing' | 'stale';
}

export type VideoStatus = 'pending' | 'completed' | 'failed';

export interface VideoRecord {
  id: number;
  input_path: string;
  sunday_date: string;
  filename: string;
  status: VideoStatus;
  job_id: number | null;
  duration_secs: number | null;
  duration_timecode: string | null;
  framerate: number | null;
  width: number | null;
  height: number | null;
  video_codec: string | null;
  audio_codec: string | null;
  audio_streams: number | null;
  bitrate: number | null;
  file_size: number | null;
  file_created_at: Date | null;
  created_at: Date;
  updated_at: Date;
}
