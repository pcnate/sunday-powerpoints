-- Normalize absolute paths in jobs and videos tables to be relative to output directory.
-- Strips everything before the YYYYMMDD folder segment (e.g. C:\...\Sermons\20250420\... → 20250420/...).
-- Also normalizes backslashes to forward slashes for consistency.

-- jobs.input_path: strip prefix up to YYYYMMDD segment, normalize slashes
UPDATE jobs
SET input_path = REGEXP_REPLACE(
  REPLACE(input_path, '\\', '/'),
  '^.*/([0-9]{8}/)',
  '$1'
)
WHERE input_path LIKE '%\\%' OR input_path LIKE '%/%/%/%';

-- jobs.output_path: same treatment
UPDATE jobs
SET output_path = REGEXP_REPLACE(
  REPLACE(output_path, '\\', '/'),
  '^.*/([0-9]{8}/)',
  '$1'
)
WHERE output_path IS NOT NULL
  AND (output_path LIKE '%\\%' OR output_path LIKE '%/%/%/%');

-- jobs.metadata JSON: normalize vtt_path inside metadata
UPDATE jobs
SET metadata = JSON_SET(
  metadata,
  '$.vtt_path',
  REGEXP_REPLACE(
    REPLACE(JSON_UNQUOTE(JSON_EXTRACT(metadata, '$.vtt_path')), '\\', '/'),
    '^.*/([0-9]{8}/)',
    '$1'
  )
)
WHERE metadata IS NOT NULL
  AND JSON_EXTRACT(metadata, '$.vtt_path') IS NOT NULL
  AND (JSON_UNQUOTE(JSON_EXTRACT(metadata, '$.vtt_path')) LIKE '%\\%'
       OR JSON_UNQUOTE(JSON_EXTRACT(metadata, '$.vtt_path')) LIKE '%/%/%/%');

-- videos: delete absolute-path duplicates that would collide after normalization.
-- Keep the row with the shorter (already relative) path; delete the absolute-path row.
DELETE v1 FROM videos v1
INNER JOIN videos v2
  ON REGEXP_REPLACE(REPLACE(v1.input_path, '\\', '/'), '^.*/([0-9]{8}/)', '$1')
   = REGEXP_REPLACE(REPLACE(v2.input_path, '\\', '/'), '^.*/([0-9]{8}/)', '$1')
  AND v1.id > v2.id;

-- videos.input_path: normalize remaining rows
UPDATE videos
SET input_path = REGEXP_REPLACE(
  REPLACE(input_path, '\\', '/'),
  '^.*/([0-9]{8}/)',
  '$1'
)
WHERE input_path LIKE '%\\%' OR input_path LIKE '%/%/%/%';
