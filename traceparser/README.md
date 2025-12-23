# QPT Telemetry Parser - Power Analysis Tool

A specialized Python tool for analyzing QPT (Qualcomm Power Telemetry) data from ftrace files. This tool parses power telemetry events, calculates average power consumption per channel, and generates comprehensive analysis reports.

## 🚀 Features

- **QPT Data Parsing**: Extract `qpt_data_update` events from ftrace files
- **Power Calculation**: Calculate average power consumption using energy differential method
- **Channel Mapping**: Automatic mapping of channel IDs to readable names and breakdown rails
- **Multiple Output Formats**: 
  - Pretty-printed console tables
  - CSV export
  - **Excel export with multiple sheets**
- **Comprehensive Statistics**: Total power, channel count, averages, and max power analysis
- **Debug Channel Filtering**: Automatically excludes debug channels from final results
- **Interactive Visualizations**: Bokeh-based interactive plots (V2 only)
- **Auto-detection of Channel Types**: Automatically detects and handles unnamed channels
- **Configurable Power Units**: Display in W, mW, or μW

## 📋 Supported Channels

The tool includes predefined mappings for 21 channels:
- **CPU**: APC0_CX, APC0_MX, APC1_CX, APC1_MX (cpu-m, cpu-l clusters)
- **GPU**: GFX, GFX_MXC (gpu, gpu-spare)
- **NSP**: NSP1_CX, NSP2_CX (nsp, nsp-spare)
- **Debug**: debug-0 through debug-12 (filtered from output)

## 🔄 Version Comparison

The tool has evolved through three versions, each with different features and capabilities:

| Feature | V1 (TelemetryParserQPT.py) | V2 (TelemetryParserQPTV2.py) | V3 (TelemetryParserQPTV3.py) |
|---------|----------------------------|------------------------------|------------------------------|
| Core Parsing | ✅ Basic parsing | ✅ Enhanced parsing | ✅ Simplified parsing |
| Channel Mapping | ✅ Fixed mapping table | ✅ Auto-detection | ✅ Auto-detection |
| Excel Export | ✅ Basic (2 sheets) | ✅ Enhanced (5+ sheets) | ✅ Simplified (3 sheets) |
| Visualizations | ❌ None | ✅ Interactive Bokeh plots | ❌ None |
| Debug Filtering | ✅ Basic | ✅ Configurable | ✅ Configurable |
| Unknown Channels | ❌ Hidden | ✅ Optional display | ✅ Auto-detection |
| Power Units | ✅ mW only | ✅ Configurable (W/mW/μW) | ✅ Configurable (W/mW/μW) |
| Code Complexity | Low | High | Medium |
| Dependencies | Basic | Advanced (+ Bokeh) | Basic |

## 🛠️ Installation

### Install Dependencies

```bash
# For V1 (basic)
pip install pandas numpy tabulate openpyxl

# For V2 (with visualizations)
pip install pandas numpy tabulate openpyxl bokeh

# For V3 (simplified)
pip install pandas openpyxl
```

## 📖 Usage

### Method 1: Edit Script Configuration

Edit the default path in the script you want to use:

```python
# In TelemetryParserQPT.py (V1)
TRACE_PATH = r"C:\path\to\your\Default_ftrace.txt"
OUTPUT_FILE = None  # Optional CSV output

# In TelemetryParserQPTV2.py or V3 (example usage section)
trace_path = r"C:\path\to\your\Default_ftrace.txt"
```

Then run:

```bash
# For V1
python TelemetryParserQPT.py

# For V2
python TelemetryParserQPTV2.py

# For V3
python TelemetryParserQPTV3.py
```

### Method 2: Command Line Arguments

```bash
# For V1
python TelemetryParserQPT.py <trace_path> [output_csv]

# For V2 and V3
python TelemetryParserQPTV2.py <trace_path> [excel_file]
python TelemetryParserQPTV3.py <trace_path> [excel_file]
```

**Examples:**

```bash
# Basic QPT analysis with V1
python TelemetryParserQPT.py "C:\traces\Default_ftrace.txt"

# Analysis with V2 and custom Excel output
python TelemetryParserQPTV2.py "C:\traces\Default_ftrace.txt" "results.xlsx"

# Simplified analysis with V3
python TelemetryParserQPTV3.py "C:\traces\Default_ftrace.txt"
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

+----------+------------------+
| Channel  | QPT Power (mW)   |
+==========+==================+
| gpu      | 1234.5           |
+----------+------------------+
| cpu-l    | 567.8            |
+----------+------------------+
| nsp      | 123.4            |
+----------+------------------+

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

### 2. Excel Output

#### V1 (Basic)
Generates a timestamped Excel file with:
- **Sheet 1 - QPT_Results**: Channel names and power values
- **Sheet 2 - Summary**: Total power, channel count, averages, max power

#### V2 (Enhanced)
Comprehensive Excel file with multiple sheets:
- **All Channels**: Complete data including debug and unnamed channels
- **Named Channels**: Only channels with proper names
- **Non-Debug Channels**: All channels except debug channels
- **Clean View**: Named non-debug channels only
- **Analysis Info**: Metadata about the analysis

#### V3 (Simplified)
Streamlined Excel file with:
- **All Channels**: Complete data with all channels
- **Named Channels**: Only named non-debug channels
- **Non-Debug Channels**: All channels except debug channels

### 3. Interactive Visualization (V2 only)
Bokeh-based interactive plots with:
- Time-series power data for each channel
- Hover tooltips showing detailed information
- Legend with show/hide functionality
- Pan, zoom, and export capabilities

### 4. CSV Output (V1 only)
Simple CSV format for further data processing

## ⚙️ How It Works

### 1. Data Parsing
- Uses regex patterns to extract QPT events from ftrace logs
- Handles both formats: with and without Channel_Name field
- Extracts: timestamp, channel_hex, channel_name, energy_uj, avg_power_uw

### 2. Channel Mapping
#### V1
- Maps channel IDs to readable names using predefined mapping table
- Fills missing channel names automatically
- Associates channels with breakdown rails (APC0_CX, GFX, etc.)

#### V2 & V3
- Direct channel name extraction from logs
- Auto-detection of channel types
- Configurable display of debug and unnamed channels

### 3. Power Calculation
- Uses energy differential method: `Power = (Energy_end - Energy_start) / Time_diff`
- Filters unreasonable values (configurable range, default 0-100W)
- Converts from microjoules to watts/milliwatts/microwatts

### 4. Result Processing
- Configurable filtering of debug channels
- Optional display of unnamed (hex-only) channels
- Sorts by power consumption (highest first)
- Generates comprehensive statistics

## 🔧 Customization

### V1: Adding New Channels
Edit the `channel_mapping` DataFrame in the script:

```python
channel_mapping = pd.DataFrame({
    'Channel_name': ['your-channel'],
    'Channel_ID': ['0xabc'],
    'Breakdown_rail': ['YOUR_RAIL']
})
```

### V2 & V3: Configuration Options
Modify the configuration dictionary:

```python
config = {
    "show_debug": False,           # Show debug channels
    "show_unknown_channels": True, # Show channels with only hex IDs
    "min_power_w": 0.001,         # Minimum power threshold (W)
    "max_power_w": 100.0,         # Maximum power threshold (W)
    "power_unit": "mW",           # Power unit (W, mW, uW)
    "decimal_places": 1,          # Decimal precision
    "save_excel": True            # Save Excel output
}

# Apply configuration
analyze_power_telemetry_v2(trace_path, config)  # For V2
analyze_power_telemetry_v3(trace_path, **config) # For V3
```

### V2: Visualization Options
```python
# Additional V2 visualization options
config.update({
    "palette": "viridis",        # Color palette (viridis, plasma, turbo)
    "plot_width": 1000,          # Plot width in pixels
    "plot_height": 600,          # Plot height in pixels
    "save_plot": True,           # Save plot to file
    "plot_format": "html"        # Plot format (html, png, svg)
})
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

### Visualization errors (V2)
- Ensure Bokeh is installed: `pip install bokeh`
- For PNG/SVG export: `pip install selenium phantomjs pillow`

### "No channels with average power above threshold"
- Adjust the `min_power_w` parameter to a lower value

## 📋 Requirements

### V1 (Basic)
- Python 3.7+
- pandas >= 1.3.0
- numpy >= 1.20.0
- tabulate >= 0.8.9
- openpyxl >= 3.0.0 (for Excel export)

### V2 (Enhanced)
- All V1 requirements
- bokeh >= 2.4.0 (for visualizations)
- Optional: selenium, phantomjs, pillow (for PNG/SVG export)

### V3 (Simplified)
- Python 3.7+
- pandas >= 1.3.0
- openpyxl >= 3.0.0

## 🔄 Version History

### TelemetryParserQPT.py (V1)
- **v1.0**: Initial QPT-only analysis with console and CSV output
- **v1.1**: Added Excel export with multiple sheets

### TelemetryParserQPTV2.py (V2)
- **v2.0**: Complete rewrite with enhanced features
  - Auto-detection of channel types
  - Interactive Bokeh visualizations
  - Comprehensive Excel export with 5+ sheets
  - Configurable power units and filtering options

### TelemetryParserQPTV3.py (V3)
- **v3.0**: Simplified version with core functionality
  - Streamlined code and dependencies
  - Focused on essential features
  - Improved performance and reliability

## 📄 License

Internal tool - check with your organization for usage rights.

## 🔍 Which Version Should I Use?

- **V1 (TelemetryParserQPT.py)**: Use for basic analysis with fixed channel mapping
- **V2 (TelemetryParserQPTV2.py)**: Use for comprehensive analysis with visualizations and advanced features
- **V3 (TelemetryParserQPTV3.py)**: Use for quick, simplified analysis with minimal dependencies
