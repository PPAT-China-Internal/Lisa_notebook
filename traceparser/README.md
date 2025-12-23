# QPT Telemetry Parser - Power Analysis Tool

A specialized Python tool for analyzing QPT (Qualcomm Power Telemetry) data from ftrace files. This tool parses power telemetry events, calculates average power consumption per channel, and generates comprehensive analysis reports.

## 🚀 Features

- **QPT Data Parsing**: Extract `qpt_data_update` events from ftrace files
- **Power Calculation**: Calculate average power consumption using energy differential method
- **Channel Mapping**: Automatic mapping of channel IDs to readable names and breakdown rails
- **Multiple Output Formats**: 
  - Pretty-printed console tables
  - CSV export
  - **Excel export with multiple sheets** (NEW!)
- **Comprehensive Statistics**: Total power, channel count, averages, and max power analysis
- **Debug Channel Filtering**: Automatically excludes debug channels from final results

## 📋 Supported Channels

The tool includes predefined mappings for 21 channels:
- **CPU**: APC0_CX, APC0_MX, APC1_CX, APC1_MX (cpu-m, cpu-l clusters)
- **GPU**: GFX, GFX_MXC (gpu, gpu-spare)
- **NSP**: NSP1_CX, NSP2_CX (nsp, nsp-spare)
- **Debug**: debug-0 through debug-12 (filtered from output)

## 🛠️ Installation

### Install Dependencies

```bash
pip install -r requirements.txt
```

## 📖 Usage

### Method 1: Edit Script Configuration

Edit the default path in `TelemetryParserQPT.py`:

```python
TRACE_PATH = r"C:\path\to\your\Default_ftrace.txt"
OUTPUT_FILE = None  # Optional CSV output
```

Then run:

```bash
python TelemetryParserQPT.py
```

### Method 2: Command Line Arguments

```bash
python TelemetryParserQPT.py <trace_path> [output_csv]
```

**Examples:**

```bash
# Basic QPT analysis
python TelemetryParserQPT.py "C:\traces\Default_ftrace.txt"

# Save results to CSV
python TelemetryParserQPT.py "C:\traces\Default_ftrace.txt" "results.csv"
```

## 📊 Output Formats

### 1. Console Output
Pretty-printed table showing:
- Channel names
- QPT Power (mW)
- Summary statistics

```
================================================================================
QPT POWER ANALYSIS RESULTS
================================================================================

┌─────────┬─────────────────┐
│ Channel │ QPT Power (mW)  │
├─────────┼─────────────────┤
│ gpu     │ 1234.5          │
│ cpu-l   │ 567.8           │
│ nsp     │ 123.4           │
└─────────┴─────────────────┘

================================================================================
SUMMARY STATISTICS
================================================================================
Total QPT Power: 1925.7 mW
Number of Channels: 3
Average Power per Channel: 641.9 mW
Max Power Channel: gpu (1234.5 mW)
================================================================================

Excel results saved to: QPT_Power_Analysis_20241210_143052.xlsx
```

### 2. Excel Output (Automatic)
Every analysis automatically generates a timestamped Excel file with:

**Sheet 1 - QPT_Results:**
- Channel names and power values
- Formatted columns with proper widths

**Sheet 2 - Summary:**
- Total QPT Power
- Number of channels
- Average power per channel
- Maximum power channel and value

### 3. CSV Output (Optional)
Simple CSV format for further data processing

## ⚙️ How It Works

### 1. Data Parsing
- Uses regex patterns to extract QPT events from ftrace logs
- Handles both formats: with and without Channel_Name field
- Extracts: timestamp, channel_hex, channel_name, energy_uj, avg_power_uw

### 2. Channel Mapping
- Maps channel IDs to readable names using predefined mapping table
- Fills missing channel names automatically
- Associates channels with breakdown rails (APC0_CX, GFX, etc.)

### 3. Power Calculation
- Uses energy differential method: `Power = (Energy_end - Energy_start) / Time_diff`
- Filters unreasonable values (0-100W range)
- Converts from microjoules to watts

### 4. Result Processing
- Filters out debug channels
- Sorts by power consumption (highest first)
- Generates comprehensive statistics

## 🔧 Customization

### Adding New Channels
Edit the `channel_mapping` DataFrame in the script:

```python
channel_mapping = pd.DataFrame({
    'Channel_name': ['your-channel'],
    'Channel_ID': ['0xabc'],
    'Breakdown_rail': ['YOUR_RAIL']
})
```

### Adjusting Power Limits
Modify the power validation in `calculate_rail_power()`:

```python
if 0 <= power_w <= 100:  # Adjust upper limit as needed
```

## 🐛 Troubleshooting

### "No QPT data events found"
- Verify trace file path is correct
- Check that file contains `qpt_data_update` events
- Ensure file is readable (not corrupted)

### "Could not calculate power"
- Check if trace contains sufficient data points
- Verify timestamps are valid
- Ensure energy values are reasonable

### Excel save errors
- Check write permissions in current directory
- Ensure openpyxl is installed: `pip install openpyxl`

## 📋 Requirements

- Python 3.7+
- pandas >= 1.3.0
- numpy >= 1.20.0
- tabulate >= 0.8.9
- openpyxl >= 3.0.0 (for Excel export)

## 🔄 Version History

- **v1.1**: Added Excel export with multiple sheets
- **v1.0**: Initial QPT-only analysis with console and CSV output

## 📄 License

Internal tool - check with your organization for usage rights.
