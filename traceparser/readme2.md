# TelemetryParserQPT2.py - Enhanced QPT Telemetry Parser

## Overview

TelemetryParserQPT2.py is an enhanced Qualcomm Power Telemetry (QPT) parser that provides comprehensive power consumption analysis with Excel export capabilities. It reads channel names directly from log files without requiring external mapping tables and offers flexible filtering options for different types of channels.

## Features

- **Direct Channel Name Extraction**: Reads channel names directly from log files
- **Flexible Filtering Options**: 
  - Debug channel filtering (default: hidden)
  - Unknown hex-only channel filtering (default: hidden)
- **Excel Export**: Generates comprehensive Excel reports with multiple sheets
- **Interactive Visualizations**: Bokeh-based interactive plots
- **Power Consumption Analysis**: Calculates power consumption using energy differences
- **Batch Processing**: Support for analyzing multiple trace files

## Requirements

```bash
pip install pandas numpy openpyxl bokeh
```

## Quick Start

### Basic Usage

```python
from TelemetryParserQPT2 import analyze_power_telemetry_v2

# Analyze trace file and generate Excel report
trace_path = r"C:\path\to\your\trace_file.txt"
qpt_data, plot, excel_path = analyze_power_telemetry_v2(
    trace_path,
    save_excel=True,  # Generate Excel file
    output_dir="./results"  # Save Excel in results directory
)
```

### Command Line Usage

```bash
python TelemetryParserQPT2.py
```

Update the `trace_path` variable in the script to point to your trace file.

## Main Functions

### `analyze_power_telemetry_v2()`

Main analysis function with comprehensive options:

```python
qpt_data, plot, excel_path = analyze_power_telemetry_v2(
    trace_path,                    # Path to trace file
    show_debug=False,              # Show debug channels (default: False)
    show_unknown_channels=False,   # Show hex-only channels (default: False)
    min_power_w=0.001,            # Minimum power threshold for plotting
    palette='viridis',            # Color palette ('viridis', 'plasma', 'turbo')
    save_excel=True,              # Generate Excel file
    output_dir=None               # Output directory (default: same as trace file)
)
```

### Convenience Functions

```python
# Clean analysis - only named non-debug channels
qpt_data, plot, excel_path = quick_analysis_clean(trace_path, save_excel=True)

# Analysis with debug channels
qpt_data, plot, excel_path = quick_analysis_with_debug(trace_path, save_excel=True)

# Complete analysis - all channels
qpt_data, plot, excel_path = quick_analysis_all(trace_path, save_excel=True)

# Only unknown channels (for debugging)
qpt_data, plot, excel_path = quick_analysis_unknown_only(trace_path, save_excel=True)
```

### Channel Summary

```python
from TelemetryParserQPT2 import get_channel_summary

summary = get_channel_summary(trace_path)
print(f"Total channels: {summary['total']}")
print(f"Named channels: {summary['named']}")
print(f"Unnamed channels: {summary['unnamed']}")
print(f"Debug channels: {summary['debug']}")
```

### Batch Processing

```python
from TelemetryParserQPT2 import batch_analysis_with_excel

trace_files = [
    r"C:\path\to\trace1.txt",
    r"C:\path\to\trace2.txt",
    r"C:\path\to\trace3.txt"
]

excel_files = batch_analysis_with_excel(
    trace_files,
    output_dir=r"C:\results",
    show_debug=False,
    show_unknown_channels=True
)
```

## Excel Output Format

The generated Excel file contains multiple sheets:

### Sheet 1: All Channels
- **Channel**: Channel name (if available) or hex ID (if unnamed)
- **Power (mW)**: Power consumption in milliwatts
- Sorted by power consumption (high to low)

### Sheet 2: Named Channels
- Only channels with actual names
- Same format as All Channels sheet

### Sheet 3: Non-Debug Channels
- All channels excluding debug channels
- Same format as All Channels sheet

### Sheet 4: Clean View
- Only named non-debug channels
- Same format as All Channels sheet

### Sheet 5: Analysis Info
- Trace file information
- Analysis date and time
- Channel statistics
- Total power consumption

## Channel Types

### Named Channels
Channels with actual names extracted from the trace file:
```
Channel_Name:VDDCX_APC0 -> "VDDCX_APC0"
Channel_Name:VDDMX -> "VDDMX"
```

### Unnamed Channels (Hex-only)
Channels without names, identified only by hex ID:
```
Channel:0x12345678 -> "Channel_0x12345678"
```

### Debug Channels
Channels with names starting with "debug-" or "debug_":
```
Channel_Name:debug-rail1 -> "debug-rail1"
```

## Filtering Options

### Default Behavior
- **Debug channels**: Hidden by default
- **Unnamed channels**: Hidden by default
- **Named channels**: Shown by default

### Show All Channels
```python
qpt_data, plot, excel_path = analyze_power_telemetry_v2(
    trace_path,
    show_debug=True,           # Show debug channels
    show_unknown_channels=True, # Show unnamed channels
    save_excel=True
)
```

### Auto-Detection
The script automatically detects if all channels are unnamed and adjusts filtering accordingly:

```python
# If all channels are unnamed, script will automatically show them
if summary['named'] == 0 and summary['unnamed'] > 0:
    show_unknown_channels = True
```

## Power Calculation

Power is calculated using the energy difference method:

```python
power_w = (energy_last - energy_first) / (time_last - time_first) / 1e6
```

- Energy values are in microjoules (uj)
- Time values are in seconds
- Result is in watts, converted to milliwatts for display

## Visualization

Interactive Bokeh plots are generated when available:

- **X-axis**: Time (seconds, relative to start)
- **Y-axis**: Power (milliwatts)
- **Legend**: Channel names with average power
- **Hover tool**: Shows detailed information
- **Interactive controls**: Pan, zoom, reset, save

## File Naming Convention

Excel files are automatically named with timestamp:
```
{trace_filename}_power_analysis_{YYYYMMDD_HHMMSS}.xlsx
```

Example:
```
Default_ftrace_power_analysis_20241210_143052.xlsx
```

## Error Handling

The script handles various error conditions:

- **File not found**: Returns None values
- **No QPT data**: Returns empty DataFrames
- **Parsing errors**: Continues processing, reports errors
- **Excel export errors**: Reports error, continues execution

## Supported Trace Format

The parser expects QPT trace lines in the following formats:

### With Channel Name
```
[12345]     123.456789: qpt_data_update: Channel:0x12345678 Channel_Name:VDDCX_APC0 adc_value:0xabcd energy:1234uj avg_power:5678uw
```

### Without Channel Name
```
[12345]     123.456789: qpt_data_update: Channel:0x12345678 adc_value:0xabcd energy:1234uj avg_power:5678uw
```

## Performance Considerations

- **Large files**: Parser processes files line by line for memory efficiency
- **Excel limits**: Raw data sheet limited to 10,000 rows to avoid Excel limits
- **Filtering**: Applied during parsing to reduce memory usage
- **Batch processing**: Processes files sequentially to avoid memory issues

## Troubleshooting

### No Data Found
```
No QPT data events found.
```
- Check if trace file contains `qpt_data_update` events
- Verify file path is correct
- Check file permissions

### All Data Filtered Out
```
No QPT data found or all data filtered out.
```
- Try enabling unknown channels: `show_unknown_channels=True`
- Try enabling debug channels: `show_debug=True`
- Check channel summary first: `get_channel_summary(trace_path)`

### Excel Export Failed
```
Error saving to Excel: [error message]
```
- Check output directory permissions
- Ensure openpyxl is installed: `pip install openpyxl`
- Check available disk space

## Example Output

```
=== Channel Summary ===
Channel Summary:
Total unique channels: 21
Named channels: 0
Unnamed (hex only) channels: 21
Debug channels: 0

Note: All channels are unnamed (hex only) channels
Will use show_unknown_channels=True for analysis

=== Power Analysis with Excel Export ===
============================================================
Power Telemetry Analysis V2
============================================================
Trace file: C:\Project\telemetry\trace.txt
Show debug channels: False (default: False)
Show unknown hex channels: True (default: False)
Minimum power threshold: 0.001W
------------------------------------------------------------
Parsing QPT data from: C:\Project\telemetry\trace.txt
Parsed 8610 QPT data points
Power Analysis Results Summary:
Total channels displayed: 21
Named channels: 0
Unnamed channels (hex only): 21
Debug channels: Hidden (default)
Unknown hex channels: Shown
--------------------------------------------------

Results saved to Excel file: C:\Project\traceparser\trace_power_analysis_20241210_143052.xlsx
Sheets created: All Channels, Named Channels, Non-Debug Channels, Clean View, Analysis Info

=== Excel File Generated ===
File path: C:\Project\traceparser\trace_power_analysis_20241210_143052.xlsx
Contains the following sheets:
- All Channels: All channels (show name if available, show ID if not)
- Named Channels: Only named channels
- Non-Debug Channels: Non-debug channels
- Clean View: Clean view (named non-debug channels)
- Analysis Info: Analysis information
```

## Version History

- **Version 2.0**: Enhanced Excel export with multiple sheets, improved filtering, batch processing support
- **Version 1.0**: Basic QPT parsing and analysis functionality