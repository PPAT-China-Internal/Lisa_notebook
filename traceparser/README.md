# Telemetry Parser - Power Analysis Tool

A standalone Python tool for analyzing power telemetry data from QPT (Qualcomm Power Telemetry) traces and UDAS (Universal Data Acquisition System) logs.

## Features

- Parse QPT data from ftrace files
- Calculate average power consumption per channel
- Compare QPT measurements with UDAS data (if available)
- Export results to CSV
- Pretty-printed console output with summary statistics

## Installation

### Install Required Dependencies

```bash
pip install -r requirements.txt
```

### Optional: UDAS Libraries

If you have access to LISA UDAS libraries, install them according to your organization's instructions. The tool will work without UDAS, performing QPT-only analysis.

## Usage

### Method 1: Edit the script directly

Edit the default paths in `TelemetryParser.py`:

```python
TRACE_PATH = r"C:\path\to\your\Default_ftrace.txt"
UDAS_BASE_DIR = r"C:\path\to\your\power\directory"
OUTPUT_FILE = "results.csv"  # Optional
```

Then run:

```bash
python TelemetryParser.py
```

### Method 2: Command Line Arguments

```bash
python TelemetryParser.py <trace_path> [udas_base_dir] [output_csv]
```

**Examples:**

```bash
# QPT analysis only
python TelemetryParser.py "C:\traces\Default_ftrace.txt"

# QPT + UDAS analysis
python TelemetryParser.py "C:\traces\Default_ftrace.txt" "C:\power\1"

# Save results to CSV
python TelemetryParser.py "C:\traces\Default_ftrace.txt" "C:\power\1" "results.csv"
```

## Output

The tool generates:

1. **Console Output**: Pretty-printed table with:
   - Channel names
   - Channel IDs
   - QPT Power (mW)
   - UDAS Power (mW) - if available
   - Power rail mappings
   - Difference calculations

2. **Summary Statistics**:
   - Total QPT Power
   - Total UDAS Power (if available)
   - Total Difference

3. **CSV Export** (optional):
   - All results saved to specified file

### Sample Output

```
================================================================================
POWER ANALYSIS RESULTS
================================================================================

+----------+-------------+----------------+-----------------+----------------+------------+
| Channel  | Channel ID  | QPT Power (mW) | UDAS Power (mW) | Breakdown Rail | UDAS Rail  |
+==========+=============+================+=================+================+============+
| cpu-m    | 0x6b6       | 1234.5         | 1250.3          | APC0_CX        | APC0_CX    |
| cpu-l    | 0x89e       | 987.6          | 995.2           | APC1_CX        | APC1_CX    |
| gpu      | 0x59b       | 543.2          | 550.1           | GFX            | GFX        |
+----------+-------------+----------------+-----------------+----------------+------------+

================================================================================
SUMMARY STATISTICS
================================================================================
Total QPT Power: 2765.3 mW
Total UDAS Power: 2795.6 mW
Total Difference: 30.3 mW
================================================================================
```

## Channel Mapping

The tool includes predefined channel mappings for:
- CPU clusters (gold/prime)
- GPU
- NSP (Neural Signal Processor)
- Debug channels

Mappings can be customized by editing the `channel_mapping` DataFrame in the script.

## Troubleshooting

### "No QPT data events found"
- Verify the trace file path is correct
- Check that the trace file contains `qpt_data_update` events
- Ensure the file is not corrupted

### "UDAS libraries not available"
- This is normal if LISA UDAS is not installed
- The tool will continue with QPT-only analysis

### "File not found" errors
- Check all paths use correct Windows format (backslashes or raw strings)
- Verify directories exist and are accessible

## Requirements

- Python 3.7+
- pandas
- numpy
- tabulate
- LISA UDAS libraries (optional)

## License

Internal tool - check with your organization for usage rights.
