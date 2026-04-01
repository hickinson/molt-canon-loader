# RUN_LOG

## Run 001
**Date:**  
**Operator:**  
**Status:**  

### Inputs
- workbook:
- script:
- source URL:

### Outputs
- workbook:
- snapshot:

### Counts
- rows fetched:
- rows imported:
- snapshot ID:

### Validation
- Raw Import OK:
- Source Register OK:
- Shortlist Model OK:
- Dashboard OK:

### Issues
- 

### Assumptions
- 

### Next
- 
## Run - 2026-04-01T00:00:00+00:00
**Status:** FAILED - environment dependency and network block

### Inputs
- workbook: molt_extraction_workbook_and_shortlist_model_full_impl_seed71.xlsx
- script: populate_molt_workbook.py
- source URL: https://molt.church/api/canon

### Outputs
- workbook: not_written
- snapshot: not_written

### Counts
- unavailable due to failed fetch/runtime prerequisites

### Schema anomalies
- Unable to query source endpoint from this environment (`Tunnel connection failed: 403 Forbidden` via configured proxy; direct egress returns `Network is unreachable`).
- Python runtime is missing required modules (`requests`, `openpyxl`).

### Failed or skipped rows
- Not evaluated because source payload could not be downloaded.

### Assumptions
- A successful run requires outbound HTTPS access to `https://molt.church/api/canon`.
- A successful run requires installing Python dependencies (`requests`, `openpyxl`) before execution.
