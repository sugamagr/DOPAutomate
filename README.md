# DOP Agent Portal | RD Installment Automation

> Bulk RD installment payment automation for the India Post DOP Agent Portal (Finacle) using Selenium WebDriver.

Takes a CSV of LOTs (batches of RD account numbers), processes payments in bulk, downloads PDF receipts, and merges them into a single file.

## How It Works

The script runs in three phases:

### Phase 1: Pay Installments
For each LOT in the CSV:
1. Enters RD account numbers into the Deposit Accounts page
2. Clicks **Fetch** and verifies the result count matches the CSV
3. Validates every account's due date falls in the current month
4. Selects all checkboxes (handles pagination for 10+ accounts)
5. Clicks **Save** → **Pay All Saved Installments**
6. Captures the payment Reference ID and saves progress to CSV + XLSX

### Phase 2: Download PDFs
1. Navigates to Reports → Recurring Deposit Installment Report
2. Searches by each LOT's Reference ID
3. Downloads the PDF receipt and renames it to `<LOT>_<RefID>.pdf`

### Phase 3: Merge PDFs
Merges all single-page PDF receipts into one combined file (`Merged_<LOT range>.pdf`). Multi-page PDFs are skipped.

## Requirements

- Python 3.8+
- Google Chrome
- ChromeDriver (matching your Chrome version)

### Python Packages

```bash
pip install selenium openpyxl psutil pypdf flask
```

## Input CSV Format

| LOT | RD Numbers | Count |
|-----|-----------|-------|
| 1   | 020027226090,020027226249,... | 7 |
| 2   | 020029815976,020029816858,... | 10 |

- **LOT** : Batch number
- **RD Numbers** : Comma-separated RD account numbers
- **Count** : Expected number of accounts in this LOT

## Usage

1. Update the `CSV_FILE` and `XLSX_FILE` paths at the top of `dop_automate.py`
2. Run the script:
   ```bash
   python3 dop_automate.py
   ```
3. Chrome opens → **log in manually** to the DOP Agent Portal
4. Navigate to: Accounts → Agent Enquire & Update Screen → Deposit Accounts
5. Press **ENTER** in the terminal to start automation
6. Open `http://127.0.0.1:5555` in your browser for the live dashboard
7. After Phase 1 completes, confirm Phase 2 (PDF downloads) when prompted

## Output

```
~/Downloads/LOT_YYYY-MM-DD/
  ├── 1_C320461082.pdf          # Individual LOT receipts
  ├── 2_C320461083.pdf
  ├── ...
  └── Merged_1-25.pdf           # All single-page PDFs merged
```

The CSV and XLSX files are updated after every LOT with the following status columns:

`Fetch_Status` | `Count_Match` | `Due_Date_Check` | `Selected` | `Selection_Verified` | `Save_Status` | `Pay_Status` | `Reference_ID` | `Remarks`

## Resumability

- LOTs with `Pay_Status=OK` are skipped on re-run
- PDFs already on disk (matching filename) are skipped
- Progress is saved to CSV after every single LOT
- The merged PDF is skipped if it already exists

## Live Dashboard

A web-based dashboard launches automatically at `http://127.0.0.1:5555` when the script starts.

**What you see:**
- Current phase, LOT, and step in real time
- Progress bar with completion percentage
- Per-LOT status table (pending, running, done, failed, skipped)
- Memory usage and elapsed time
- Live scrolling log of all automation output

**What you can control:**
- **Pause / Resume** → freezes automation at the next safe checkpoint
- **Skip LOT** → jumps to the next LOT immediately
- **Stop After Current** → finishes the current LOT, saves progress, and exits
- **Delay sliders** → adjust Short, Medium, Long, and Checkbox delays on the fly
- **Per-LOT skip toggles** → mark pending LOTs to be skipped before they run

The dashboard auto-reconnects if the connection drops and works in any modern browser.

## Safety Features

| Feature | Description |
|---------|-------------|
| Zombie Chrome cleanup | Kills leftover automation Chrome processes on startup |
| Memory watchdog | Monitors RAM usage, saves progress and exits if it exceeds 3.5 GB |
| Global timeout | Auto-exits after 30 minutes to prevent hangs |
| Pagination cap | Page navigation limited to 10 clicks to prevent runaway loops |
| Chrome memory flags | Runs with `--disable-gpu`, `--disable-dev-shm-usage`, `--disable-extensions` |
| Gentle pacing | Deliberate delays between actions to avoid triggering rate limits |

## Configuration

Key constants at the top of `dop_automate.py` (delays can also be changed live via the dashboard):

```python
CSV_FILE            = "/path/to/your/input.csv"
XLSX_FILE           = "/path/to/your/output.xlsx"
DELAY_SHORT         = 1.5       # seconds, after small actions
DELAY_MEDIUM        = 3.0       # seconds, after fetch / page loads
DELAY_LONG          = 5.0       # seconds, between LOTs
DELAY_CHECKBOX      = 0.4       # seconds, between checkbox clicks
GLOBAL_TIMEOUT_MINS = 30
MEMORY_LIMIT_MB     = 3500
```

## Platform Support

Works on **macOS**, **Windows**, and **Linux**. Automatically uses `Cmd+A` on Mac and `Ctrl+A` on Windows/Linux.
