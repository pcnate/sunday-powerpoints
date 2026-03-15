use anyhow::{ Context, Result };
use reqwest::Client;
use serde::{ Deserialize, Serialize };
use std::time::Duration;


/// HTTP client wrapping all Express server API calls.
///
/// The worker communicates exclusively through this client —
/// no direct MySQL access.
pub struct ApiClient {
    client: Client,
    base_url: String,
}


// --- Request/Response types ---

/// Request body for POST /api/jobs/next (claim a job).
#[derive( Debug, Serialize )]
pub struct ClaimJobRequest {
    pub worker_id: String,
    pub types: Vec<String>,
}


/// Response body from POST /api/jobs/next when a job is available.
#[derive( Debug, Deserialize )]
pub struct ClaimJobResponse {
    pub job: Job,
}


/// Request body for POST /api/jobs/:id/heartbeat.
#[derive( Debug, Serialize )]
pub struct HeartbeatRequest {
    pub worker_id: String,
    #[serde( skip_serializing_if = "Option::is_none" )]
    pub progress: Option<f64>,
}


/// Response body from POST /api/jobs/:id/heartbeat.
#[derive( Debug, Deserialize )]
pub struct HeartbeatResponse {
    /// If false, the job has been cancelled and the worker should stop.
    #[serde( rename = "continue" )]
    pub should_continue: bool,
}


/// Request body for POST /api/jobs/:id/complete.
#[derive( Debug, Serialize )]
pub struct CompleteJobRequest {
    pub worker_id: String,
    #[serde( skip_serializing_if = "Option::is_none" )]
    pub output_path: Option<String>,
    #[serde( skip_serializing_if = "Option::is_none" )]
    pub metadata: Option<serde_json::Value>,
}


/// Request body for POST /api/jobs/:id/fail.
#[derive( Debug, Serialize )]
pub struct FailJobRequest {
    pub worker_id: String,
    pub error_message: String,
}


/// A single log entry to send to the server.
#[derive( Debug, Serialize )]
pub struct LogEntry {
    pub level: String,
    pub message: String,
}


/// Request body for POST /api/jobs/:id/logs.
#[derive( Debug, Serialize )]
pub struct SendLogsRequest {
    pub entries: Vec<LogEntry>,
}


/// A job as returned by the Express API.
#[derive( Debug, Clone, Serialize, Deserialize )]
pub struct Job {
    pub id: u32,
    #[serde( rename = "type" )]
    pub job_type: String,
    pub status: String,
    pub priority: u8,
    pub sunday_date: String,
    pub input_path: String,
    pub output_path: Option<String>,
    pub metadata: Option<serde_json::Value>,
    pub error_message: Option<String>,
    pub worker_id: Option<String>,
    pub retry_count: u8,
    pub max_retries: u8,
    pub parent_job_id: Option<u32>,
    pub created_at: String,
    pub updated_at: String,
    pub started_at: Option<String>,
    pub completed_at: Option<String>,
    pub heartbeat_at: Option<String>,
}


impl ApiClient {
    /// Create a new API client.
    ///
    /// @param base_url - the Express server URL (e.g., "http://192.168.1.100:8080")
    pub fn new( base_url: &str ) -> Self {
        let client = Client::builder()
            .timeout( Duration::from_secs( 10 ) )
            .build()
            .expect( "Failed to build HTTP client" );

        Self {
            client,
            base_url: base_url.trim_end_matches( '/' ).to_string(),
        }
    }


    /// Claim the next available job from the queue.
    ///
    /// Returns `Ok(Some(job))` if a job was claimed, `Ok(None)` if no
    /// jobs are available (204 response), or an error on failure.
    ///
    /// @param worker_id - this worker's identifier
    /// @param types - job types this worker can handle
    pub async fn claim_job(
        &self,
        worker_id: &str,
        types: &[String],
    ) -> Result<Option<Job>> {
        let url = format!( "{}/api/jobs/next", self.base_url );
        let body = ClaimJobRequest {
            worker_id: worker_id.to_string(),
            types: types.to_vec(),
        };

        let response = self.client
            .post( &url )
            .json( &body )
            .send()
            .await
            .context( "Failed to send claim request" )?;

        match response.status().as_u16() {
            204 => Ok( None ),
            200 => {
                let claim: ClaimJobResponse = response.json().await
                    .context( "Failed to parse claim response" )?;
                Ok( Some( claim.job ) )
            }
            status => {
                let text = response.text().await.unwrap_or_default();
                anyhow::bail!( "Claim request failed with status {}: {}", status, text );
            }
        }
    }


    /// Claim a specific job by ID.
    ///
    /// Returns `Ok(Some(job))` if claimed, `Ok(None)` if the job is not
    /// pending (e.g., already claimed or completed), or an error on failure.
    ///
    /// @param job_id - the specific job to claim
    /// @param worker_id - this worker's identifier
    pub async fn claim_specific_job(
        &self,
        job_id: u32,
        worker_id: &str,
    ) -> Result<Option<Job>> {
        let url = format!( "{}/api/jobs/{}/claim", self.base_url, job_id );
        let body = serde_json::json!({ "worker_id": worker_id });

        let response = self.client
            .post( &url )
            .json( &body )
            .send()
            .await
            .context( "Failed to send claim request" )?;

        match response.status().as_u16() {
            200 => {
                let claim: ClaimJobResponse = response.json().await
                    .context( "Failed to parse claim response" )?;
                Ok( Some( claim.job ) )
            }
            409 => Ok( None ),
            status => {
                let text = response.text().await.unwrap_or_default();
                anyhow::bail!( "Claim request failed with status {}: {}", status, text );
            }
        }
    }


    /// Send a heartbeat for a running job.
    ///
    /// Returns `true` if the worker should continue, `false` if the
    /// job has been cancelled.
    ///
    /// @param job_id - the job being processed
    /// @param worker_id - this worker's identifier
    /// @param progress - optional progress percentage (0.0–1.0)
    pub async fn heartbeat(
        &self,
        job_id: u32,
        worker_id: &str,
        progress: Option<f64>,
    ) -> Result<bool> {
        let url = format!( "{}/api/jobs/{}/heartbeat", self.base_url, job_id );
        let body = HeartbeatRequest {
            worker_id: worker_id.to_string(),
            progress,
        };

        let response = self.client
            .post( &url )
            .json( &body )
            .send()
            .await
            .context( "Failed to send heartbeat" )?;

        if !response.status().is_success() {
            let text = response.text().await.unwrap_or_default();
            anyhow::bail!( "Heartbeat failed with status: {}", text );
        }

        let hb: HeartbeatResponse = response.json().await
            .context( "Failed to parse heartbeat response" )?;

        Ok( hb.should_continue )
    }


    /// Mark a job as completed.
    ///
    /// @param job_id - the job to complete
    /// @param worker_id - this worker's identifier
    /// @param output_path - optional path to the output file
    /// @param metadata - optional additional metadata to merge
    pub async fn complete_job(
        &self,
        job_id: u32,
        worker_id: &str,
        output_path: Option<String>,
        metadata: Option<serde_json::Value>,
    ) -> Result<()> {
        let url = format!( "{}/api/jobs/{}/complete", self.base_url, job_id );
        let body = CompleteJobRequest {
            worker_id: worker_id.to_string(),
            output_path,
            metadata,
        };

        let response = self.client
            .post( &url )
            .json( &body )
            .send()
            .await
            .context( "Failed to send complete request" )?;

        if !response.status().is_success() {
            let text = response.text().await.unwrap_or_default();
            anyhow::bail!( "Complete request failed: {}", text );
        }

        Ok(())
    }


    /// Mark a job as failed.
    ///
    /// @param job_id - the job that failed
    /// @param worker_id - this worker's identifier
    /// @param error_message - description of what went wrong
    pub async fn fail_job(
        &self,
        job_id: u32,
        worker_id: &str,
        error_message: &str,
    ) -> Result<()> {
        let url = format!( "{}/api/jobs/{}/fail", self.base_url, job_id );
        let body = FailJobRequest {
            worker_id: worker_id.to_string(),
            error_message: error_message.to_string(),
        };

        let response = self.client
            .post( &url )
            .json( &body )
            .send()
            .await
            .context( "Failed to send fail request" )?;

        if !response.status().is_success() {
            let text = response.text().await.unwrap_or_default();
            anyhow::bail!( "Fail request failed: {}", text );
        }

        Ok(())
    }


    /// Send log entries for a job.
    ///
    /// @param job_id - the job to log against
    /// @param entries - log entries to send
    pub async fn send_logs(
        &self,
        job_id: u32,
        entries: Vec<LogEntry>,
    ) -> Result<()> {
        let url = format!( "{}/api/jobs/{}/logs", self.base_url, job_id );
        let body = SendLogsRequest { entries };

        let response = self.client
            .post( &url )
            .json( &body )
            .send()
            .await
            .context( "Failed to send logs" )?;

        if !response.status().is_success() {
            let text = response.text().await.unwrap_or_default();
            anyhow::bail!( "Send logs failed: {}", text );
        }

        Ok(())
    }


    /// Quick health check — try to reach the server.
    ///
    /// Returns true if the server responds to GET /api/jobs?limit=0.
    pub async fn health_check( &self ) -> bool {
        let url = format!( "{}/api/jobs?limit=0", self.base_url );
        match self.client.get( &url ).send().await {
            Ok( resp ) => resp.status().is_success(),
            Err( _ ) => false,
        }
    }
}
