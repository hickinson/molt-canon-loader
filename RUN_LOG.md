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

## Run - 2026-04-01T17:31:41Z
**Status:** BLOCKED - dependency installation and API connectivity denied by outbound proxy

### Dependency install result
- Command: `python3 -m pip install -r requirements.txt`
- Result: FAILED
- Error: proxy tunnel denied while requesting Python package index (`Tunnel connection failed: 403 Forbidden`).
- Impact: Required modules (`requests`, `openpyxl`) could not be installed in this environment.

### Connectivity result
- Command: `curl -sv https://molt.church/api/canon`
- Result: FAILED
- Denied domain/method: `CONNECT molt.church:443` via proxy `http://proxy:8080`
- Proxy response: `HTTP/1.1 403 Forbidden`

### Canon import result
- Full canon payload fetch: NOT RUN (blocked by denied CONNECT to molt.church)
- Raw JSON snapshot: NOT WRITTEN
- Workbook population: NOT RUN
- Downstream sheet population check: NOT RUN
- Total imported rows: 0

### Schema anomalies
- Not evaluated because canonical payload could not be fetched.

### Output filenames
- JSON snapshot: none
- Workbook output: none
