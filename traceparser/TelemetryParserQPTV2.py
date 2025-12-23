#!/usr/bin/env python3
"""
TelemetryParserQPTV2.py - Enhanced QPT Telemetry Parser V2

A simplified QPT (Qualcomm Power Telemetry) parser with Excel export capabilities.

Features:
- Direct channel name extraction from logs
- Auto-detection of channel types
- Excel export with multiple sheets
- Interactive Bokeh visualizations
- Simple command line interface

Author: Generated for Power Analysis
Version: 2.0
"""

import os
import re
import warnings
import argparse
from typing import Optional, Dict, Any, Tuple, List
import pandas as pd
import numpy as np
from datetime import datetime

# Suppress warnings
warnings.filterwarnings('ignore')

# Default configuration
DEFAULT_CONFIG = {
    "show_debug": False,
    "show_unknown_channels": False,
    "min_power_w": 0.001,
    "max_power_w": 100.0,
    "min_time_diff_s": 0.001,
    "power_unit": "mW",
    "decimal_places": 1,
    "save_excel": True,
    "verbose": True,
    "auto_detect_channels": True
}


def format_table_output(results_df: pd.DataFrame, power_unit: str = 'mW') -> str:
    """
    Format results as a nice table for CLI output
    
    Parameters:
    results_df: DataFrame with Channel and Power columns
    power_unit: Power unit for display
    
    Returns:
    str: Formatted table string
    """
    if results_df.empty:
        return "No results to display."
    
    # Prepare data
    channels = results_df['Channel'].tolist()
    powers = results_df[f'Power ({power_unit})'].tolist()
    
    # Calculate column widths
    max_channel_width = max(len(str(ch)) for ch in channels)
    max_power_width = max(len(f"{p:.1f}") for p in powers)
    
    # Ensure minimum widths
    channel_width = max(max_channel_width, len('Channel'))
    power_width = max(max_power_width, len(f'QPT Power ({power_unit})'))
    
    # Create table
    lines = []
    
    # Header separator
    lines.append('=' * 80)
    lines.append('QPT POWER ANALYSIS RESULTS')
    lines.append('=' * 80)
    
    # Table header
    header_sep = '+' + '-' * (channel_width + 2) + '+' + '-' * (power_width + 2) + '+'
    header_row = f"| {'Channel':<{channel_width}} | {'QPT Power (' + power_unit + ')':<{power_width}} |"
    header_eq = '+' + '=' * (channel_width + 2) + '+' + '=' * (power_width + 2) + '+'
    
    lines.append(header_sep)
    lines.append(header_row)
    lines.append(header_eq)
    
    # Data rows
    for channel, power in zip(channels, powers):
        power_str = f"{power:.1f}" if power != int(power) else f"{int(power)}"
        row = f"| {channel:<{channel_width}} | {power_str:>{power_width}} |"
        lines.append(row)
        lines.append(header_sep)
    
    return '\n'.join(lines)


def display_results_simple(qpt_results: pd.DataFrame, show_debug: bool = False, 
                          show_unknown_channels: bool = False, power_unit: str = 'mW',
                          decimal_places: int = 1, verbose: bool = True):
    """
    Display results in a simple table format
    
    Parameters:
    qpt_results: Power calculation results
    show_debug: Show debug channels
    show_unknown_channels: Show unknown hex-only channels
    power_unit: Power unit for display
    decimal_places: Number of decimal places
    verbose: Enable verbose output
    """
    if qpt_results.empty:
        if verbose:
            print("No results to display.")
        return

    # Apply filtering
    filtered_results = qpt_results.copy()
    
    if not show_debug:
        filtered_results = filtered_results[
            ~filtered_results['channel_display_name'].str.lower().str.contains('debug', na=False)
        ]
    
    if not show_unknown_channels:
        filtered_results = filtered_results[filtered_results['has_name'] == True]
    
    if filtered_results.empty:
        if verbose:
            print("No results to display after filtering.")
        return

    # Convert power to requested unit
    power_multiplier = {'W': 1.0, 'mW': 1000.0, 'uW': 1000000.0}
    multiplier = power_multiplier.get(power_unit, 1000.0)
    
    filtered_results[f"Power ({power_unit})"] = (filtered_results["power_w"] * multiplier).round(decimal_places)
    
    # Create display dataframe
    display_df = pd.DataFrame({
        "Channel": filtered_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
        f"Power ({power_unit})": filtered_results[f"Power ({power_unit})"]
    })
    
    # Sort by power consumption (descending)
    display_df = display_df.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
    
    # Display formatted table
    table_output = format_table_output(display_df, power_unit)
    print(table_output)
    
    if verbose:
        total_power = display_df[f"Power ({power_unit})"].sum()
        print(f"\nTotal Power: {total_power:.{decimal_places}f} {power_unit}")
        print(f"Channels: {len(display_df)}")


def create_argument_parser():
    """
    Create command line argument parser
    
    Returns:
    argparse.ArgumentParser: Configured argument parser
    """
    parser = argparse.ArgumentParser(
        description="QPT Telemetry Parser V2 - Enhanced power analysis tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python TelemetryParserQPTV2.py trace.txt
  python TelemetryParserQPTV2.py trace.txt output.xlsx
        """
    )
    
    # Required arguments
    parser.add_argument('trace_file', help='Path to trace file')
    parser.add_argument('excel_file', nargs='?', help='Excel output filename (optional)')
    
    return parser

# Check and import visualization libraries
BOKEH_AVAILABLE = False
try:
    from bokeh.plotting import figure, output_notebook, show
    from bokeh.models import ColumnDataSource, HoverTool
    from bokeh.models.tools import (
        PanTool, BoxZoomTool, ResetTool, 
        SaveTool, WheelZoomTool
    )
    from bokeh.palettes import Viridis256, Plasma256, Turbo256
    from bokeh.io import reset_output
    
    # Configure Bokeh for inline plotting
    output_notebook()
    reset_output()
    BOKEH_AVAILABLE = True
    print("Bokeh visualization libraries loaded successfully.")
except ImportError:
    print("Bokeh libraries not available. Only basic analysis will be performed.")


def is_debug_channel(channel_name: str) -> bool:
    """
    Check if channel is a debug channel
    
    Parameters:
    channel_name: Channel name to check
    
    Returns:
    bool: True if channel is a debug channel
    """
    if pd.isna(channel_name):
        return False
    channel_str = str(channel_name).lower()
    return channel_str.startswith('debug-') or channel_str.startswith('debug_')


def parse_qpt_data_update_v2(trace_path: str, show_debug: bool = False, show_unknown_channels: bool = False) -> pd.DataFrame:
    """
    Parse qpt_data_update events from trace file without mapping mechanism.
    
    Parameters:
    trace_path: Path to trace file
    show_debug: Whether to show debug-* channels (default: False - hidden)
    show_unknown_channels: Whether to show channels with only hex IDs (default: False - hidden)
    
    Returns:
    pd.DataFrame: Parsed QPT data with columns:
        - timestamp: Event timestamp
        - channel_hex: Channel hex ID
        - channel_name: Channel identifier
        - channel_display_name: Formatted display name
        - energy_uj: Energy in microjoules
        - avg_power_uw: Average power in microwatts
        - has_name: Boolean indicating if channel has a real name
    """
    # Regex for lines with Channel_Name
    pattern_with_name = re.compile(
        r'.*?\s+\[(\d+)\]\s+[\.\s]+\s+(\d+\.\d+):\s+qpt_data_update:\s+'
        r'Channel:(0x[0-9a-f]+)\s+Channel_Name:([^\s]+)\s+'
        r'adc_value:(0x[0-9a-f]+)\s+energy:(\d+)uj\s+avg_power:(\d+)uw',
        re.IGNORECASE
    )
    # Regex for lines without Channel_Name
    pattern_without_name = re.compile(
        r'.*?\s+\[(\d+)\]\s+[\.\s]+\s+(\d+\.\d+):\s+qpt_data_update:\s+'
        r'Channel:(0x[0-9a-f]+)\s+'
        r'adc_value:(0x[0-9a-f]+)\s+energy:(\d+)uj\s+avg_power:(\d+)uw',
        re.IGNORECASE
    )
    
    data = []
    print(f"Parsing QPT data from: {trace_path}")
    
    try:
        with open(trace_path, 'r', errors='ignore') as f:
            for line in f:
                if 'qpt_data_update' in line:
                    match = pattern_with_name.search(line)
                    if match:
                        channel_name = match.group(4)
                        channel_hex = match.group(3).lower()
                        
                        # Apply filtering logic for debug channels
                        if not show_debug and is_debug_channel(channel_name):
                            continue
                            
                        data.append({
                            'timestamp': float(match.group(2)),
                            'channel_hex': channel_hex,
                            'channel_name': channel_name,
                            'channel_display_name': channel_name,  # Use actual name from log
                            'energy_uj': int(match.group(6)),
                            'avg_power_uw': int(match.group(7)),
                            'has_name': True
                        })
                    else:
                        match = pattern_without_name.search(line)
                        if match:
                            channel_hex = match.group(3).lower()
                            
                            # Apply filtering logic for unknown channels (default: hidden)
                            if not show_unknown_channels:
                                continue
                                
                            data.append({
                                'timestamp': float(match.group(2)),
                                'channel_hex': channel_hex,
                                'channel_name': channel_hex,  # Use hex as identifier
                                'channel_display_name': f"Channel_{channel_hex}",  # Display format
                                'energy_uj': int(match.group(5)),
                                'avg_power_uw': int(match.group(6)),
                                'has_name': False
                            })
                            
        if not data:
            print("No QPT data events found.")
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        print(f"Parsed {len(df)} QPT data points")
        return df
        
    except FileNotFoundError:
        print(f"File not found: {trace_path}")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error parsing QPT data: {e}")
        return pd.DataFrame()


def calculate_rail_power_v2(qpt_df: pd.DataFrame) -> pd.DataFrame:
    """
    Calculate average power (W) using first and last energy points - simplified version
    
    Parameters:
    qpt_df: DataFrame containing QPT data
    
    Returns:
    pd.DataFrame: Power calculation results with columns:
        - channel_name: Channel identifier
        - channel_display_name: Display name
        - channel_hex: Hex ID
        - has_name: Whether channel has real name
        - time_diff_s: Time duration
        - power_w: Calculated power in watts
    """
    if qpt_df.empty:
        return pd.DataFrame()
        
    results = []
    channels = qpt_df['channel_name'].unique()
    
    for channel in channels:
        channel_df = qpt_df[qpt_df['channel_name'] == channel].sort_values('timestamp')
        if len(channel_df) <= 1:
            continue
            
        first = channel_df.iloc[0]
        last = channel_df.iloc[-1]
        
        time_diff = last['timestamp'] - first['timestamp']
        energy_diff = last['energy_uj'] - first['energy_uj']
        
        if time_diff > 0.001:  # Minimum time threshold
            power_w = energy_diff / time_diff / 1e6  # Convert to watts
            # Filter unreasonable values (0 to 100W)
            if 0 <= power_w <= 100:
                results.append({
                    'channel_name': channel,
                    'channel_display_name': channel_df.iloc[0]['channel_display_name'],
                    'channel_hex': channel_df.iloc[0]['channel_hex'],
                    'has_name': channel_df.iloc[0]['has_name'],
                    'time_diff_s': time_diff,
                    'power_w': power_w
                })
                
    if not results:
        return pd.DataFrame()
        
    return pd.DataFrame(results).sort_values('power_w', ascending=False)


def display_results_v2(qpt_results: pd.DataFrame, show_debug: bool = False, show_unknown_channels: bool = False, return_df: bool = False):
    """
    Display the power analysis results with filtering options
    
    Parameters:
    qpt_results: Power calculation results
    show_debug: Show debug channels (default: False - hidden)
    show_unknown_channels: Show unknown hex-only channels (default: False - hidden)
    """
    if qpt_results.empty:
        print("No results to display.")
        return

    # Apply additional filtering if needed
    filtered_results = qpt_results.copy()
    
    if not show_debug:
        filtered_results = filtered_results[
            ~filtered_results['channel_display_name'].str.lower().str.contains('debug', na=False)
        ]
    
    if not show_unknown_channels:
        filtered_results = filtered_results[filtered_results['has_name'] == True]
    
    if filtered_results.empty:
        print("No results to display after filtering.")
        return

    # Convert to mW and format output
    filtered_results["Power (mW)"] = (filtered_results["power_w"] * 1000.0).round(1)
    
    # Create simplified display dataframe with only Channel and Power
    simple_df = pd.DataFrame({
        "Channel": filtered_results["channel_display_name"],
        "Power (mW)": filtered_results["Power (mW)"]
    })
    
    # Sort by power consumption
    simple_df = simple_df.sort_values("Power (mW)", ascending=False).reset_index(drop=True)
    
    # Display summary with clear filtering status
    total_channels = len(filtered_results)
    named_channels = len(filtered_results[filtered_results["has_name"] == True])
    unnamed_channels = total_channels - named_channels
    
    print(f"Power Analysis Results Summary:")
    print(f"Total channels displayed: {total_channels}")
    print(f"Named channels: {named_channels}")
    print(f"Unnamed channels (hex only): {unnamed_channels}")
    print(f"Debug channels: {'Shown' if show_debug else 'Hidden (default)'}")
    print(f"Unknown hex channels: {'Shown' if show_unknown_channels else 'Hidden (default)'}")
    print("-" * 50)
    
    # Print simplified table with just Channel and Power
    table_output = format_table_output(simple_df, "mW")
    print(table_output)
    
    if return_df:
        return filtered_results


def plot_rails_power_bokeh_v2(qpt_df, min_power_w=0, palette_name='viridis', 
                              show_debug=False, show_unknown_channels=False):
    """
    Plot rails' power with filtering options using Bokeh
    
    Parameters:
    qpt_df: QPT data DataFrame
    min_power_w: Minimum power threshold for plotting
    palette_name: Color palette ('viridis', 'plasma', 'turbo')
    show_debug: Show debug channels (default: False - hidden)
    show_unknown_channels: Show unknown hex-only channels (default: False - hidden)
    
    Returns:
    Bokeh plot object or None if no data
    """
    if not BOKEH_AVAILABLE:
        print("Bokeh not available. Cannot create interactive plots.")
        return None
        
    if qpt_df.empty:
        print("No data available for plotting")
        return None
    
    # Apply filtering
    filtered_df = qpt_df.copy()
    
    if not show_debug:
        filtered_df = filtered_df[
            ~filtered_df['channel_display_name'].str.lower().str.contains('debug', na=False)
        ]
    
    if not show_unknown_channels:
        filtered_df = filtered_df[filtered_df['has_name'] == True]
    
    if filtered_df.empty:
        print("No data available after filtering")
        return None
    
    # Calculate the start time for relative time
    start_time = filtered_df['timestamp'].min()
    filtered_df['relative_time'] = filtered_df['timestamp'] - start_time
    
    # Calculate average power for each channel and filter
    channels_power = {}
    for channel in filtered_df['channel_name'].unique():
        channel_data = filtered_df[filtered_df['channel_name'] == channel]
        power_w = channel_data['avg_power_uw'].mean() / 1e3  # Convert to mW
        if power_w >= min_power_w * 1000:  # Convert threshold to mW
            display_name = channel_data.iloc[0]['channel_display_name']
            channels_power[channel] = {'power': power_w, 'display_name': display_name}
    
    # Sort channels by average power (descending)
    sorted_channels = sorted(channels_power.items(), key=lambda x: x[1]['power'], reverse=True)
    filtered_channels = [ch for ch, _ in sorted_channels]
    
    if not filtered_channels:
        print(f"No channels with average power above {min_power_w}W after filtering")
        return None
    
    # Create figure with updated title showing filter status
    debug_status = "shown" if show_debug else "hidden"
    unknown_status = "shown" if show_unknown_channels else "hidden"
    
    p = figure(
        title=f"QPT Power Analysis (debug: {debug_status}, unknown: {unknown_status}, min: {min_power_w}W)",
        x_axis_label="Time (s)",
        y_axis_label="Power (mW)",
        width=1000,
        height=600,
        tools="pan,box_zoom,wheel_zoom,reset,save,hover",
        toolbar_location="above"
    )
    
    # Select color palette
    if palette_name.lower() == 'plasma':
        palette = Plasma256
    elif palette_name.lower() == 'turbo':
        palette = Turbo256
    else:
        palette = Viridis256  # Default
    
    # Get evenly spaced colors
    num_channels = len(filtered_channels)
    step = max(1, len(palette) // num_channels)
    colors = [palette[i * step % len(palette)] for i in range(num_channels)]
    
    # Plot each channel
    for i, channel in enumerate(filtered_channels):
        channel_data = filtered_df[filtered_df['channel_name'] == channel]
        power_mw = channel_data['avg_power_uw'] / 1e3  # Convert to mW
        display_name = channels_power[channel]['display_name']
        avg_power = channels_power[channel]['power']
        
        # Create data source
        source = ColumnDataSource(data={
            'x': channel_data['relative_time'],
            'y': power_mw,
            'channel': [display_name] * len(channel_data),
            'power_uw': channel_data['avg_power_uw'],
            'abs_time': channel_data['timestamp'],
        })
        
        # Plot line
        p.line(
            'x', 'y', 
            source=source, 
            line_width=2, 
            color=colors[i], 
            alpha=0.8,
            legend_label=f"{display_name} ({avg_power:.1f}mW)"
        )
    
    # Configure hover tool
    hover = p.select(dict(type=HoverTool))
    hover.tooltips = [
        ("Channel", "@channel"),
        ("Time", "@x{0.0} s"),
        ("Power", "@y{0.0} mW")
    ]
    hover.mode = 'vline'
    
    # Configure legend
    p.legend.location = "right"
    p.legend.click_policy = "hide"
    p.legend.border_line_color = "navy"
    p.legend.background_fill_alpha = 0.7
    p.legend.label_text_font_size = "8pt"
    p.legend.glyph_width = 15
    p.legend.spacing = 0
    p.legend.padding = 5
    
    # Add grid lines
    p.grid.grid_line_alpha = 0.3
    
    return p


def save_results_to_excel_enhanced(qpt_results: pd.DataFrame, qpt_raw_data: pd.DataFrame, 
                                  trace_path: str, config: Dict[str, Any]) -> str:
    """
    Enhanced Excel export with configurable options
    
    Parameters:
    qpt_results: Power calculation results DataFrame
    qpt_raw_data: Raw QPT data DataFrame
    trace_path: Original trace file path
    config: Configuration dictionary
    
    Returns:
    str: Path to saved Excel file
    """
    if qpt_results.empty:
        if config.get('verbose', True):
            print("No results to save.")
        return None
    
    # Get configuration parameters
    output_dir = config.get('output_dir', None)
    excel_template = config.get('excel_filename_template', "{trace_name}_power_analysis_{timestamp}.xlsx")
    max_raw_rows = config.get('excel_max_raw_rows', 10000)
    power_unit = config.get('power_unit', 'mW')
    decimal_places = config.get('decimal_places', 1)
    verbose = config.get('verbose', True)
    
    # Determine output directory and filename
    if output_dir is None:
        output_dir = os.path.dirname(trace_path)
    
    trace_filename = os.path.splitext(os.path.basename(trace_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = excel_template.format(
        trace_name=trace_filename,
        timestamp=timestamp
    )
    excel_path = os.path.join(output_dir, excel_filename)
    
    try:
        # Power unit conversion
        power_multiplier = {'W': 1.0, 'mW': 1000.0, 'uW': 1000000.0}
        multiplier = power_multiplier.get(power_unit, 1000.0)
        
        # Create Excel writer
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            
            # Sheet 1: All Channels
            all_results = qpt_results.copy()
            if not all_results.empty:
                all_results[f"Power ({power_unit})"] = (all_results["power_w"] * multiplier).round(decimal_places)
                all_summary = pd.DataFrame({
                    "Channel": all_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
                    f"Power ({power_unit})": all_results[f"Power ({power_unit})"],
                    "Duration (s)": all_results["time_diff_s"].round(2),
                    "Samples": all_results["sample_count"],
                    "Energy (uJ)": all_results["energy_diff_uj"]
                })
                all_summary = all_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                all_summary.to_excel(writer, sheet_name='All Channels', index=False)
            
            # Sheet 2: Named Channels Only
            named_results = qpt_results[qpt_results['has_name'] == True].copy()
            if not named_results.empty:
                named_results[f"Power ({power_unit})"] = (named_results["power_w"] * multiplier).round(decimal_places)
                named_summary = pd.DataFrame({
                    "Channel": named_results["channel_display_name"],
                    f"Power ({power_unit})": named_results[f"Power ({power_unit})"],
                    "Duration (s)": named_results["time_diff_s"].round(2),
                    "Samples": named_results["sample_count"]
                })
                named_summary = named_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                named_summary.to_excel(writer, sheet_name='Named Channels', index=False)
            
            # Sheet 3: Non-Debug Channels
            non_debug_results = qpt_results[
                ~qpt_results['channel_display_name'].str.lower().str.contains('debug', na=False)
            ].copy()
            if not non_debug_results.empty:
                non_debug_results[f"Power ({power_unit})"] = (non_debug_results["power_w"] * multiplier).round(decimal_places)
                non_debug_summary = pd.DataFrame({
                    "Channel": non_debug_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
                    f"Power ({power_unit})": non_debug_results[f"Power ({power_unit})"],
                    "Duration (s)": non_debug_results["time_diff_s"].round(2),
                    "Samples": non_debug_results["sample_count"]
                })
                non_debug_summary = non_debug_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                non_debug_summary.to_excel(writer, sheet_name='Non-Debug Channels', index=False)
            
            # Sheet 4: Clean View
            clean_results = qpt_results[
                (qpt_results['has_name'] == True) & 
                (~qpt_results['channel_display_name'].str.lower().str.contains('debug', na=False))
            ].copy()
            if not clean_results.empty:
                clean_results[f"Power ({power_unit})"] = (clean_results["power_w"] * multiplier).round(decimal_places)
                clean_summary = pd.DataFrame({
                    "Channel": clean_results["channel_display_name"],
                    f"Power ({power_unit})": clean_results[f"Power ({power_unit})"],
                    "Duration (s)": clean_results["time_diff_s"].round(2),
                    "Samples": clean_results["sample_count"]
                })
                clean_summary = clean_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                clean_summary.to_excel(writer, sheet_name='Clean View', index=False)
            
            # Sheet 5: Analysis Info
            total_power = (qpt_results["power_w"] * multiplier).sum()
            metadata = {
                "Analysis Information": [
                    "Trace File",
                    "Analysis Date",
                    "Total Channels Found",
                    "Named Channels",
                    "Unnamed Channels (hex only)",
                    "Debug Channels",
                    f"Total Power ({power_unit})",
                    "Power Unit",
                    "Decimal Places",
                    "Min Power Threshold (W)",
                    "Max Power Threshold (W)"
                ],
                "Value": [
                    os.path.basename(trace_path),
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    len(qpt_results),
                    len(qpt_results[qpt_results["has_name"] == True]),
                    len(qpt_results[qpt_results["has_name"] == False]),
                    len(qpt_results[qpt_results['channel_display_name'].str.lower().str.contains('debug', na=False)]),
                    round(total_power, decimal_places),
                    power_unit,
                    decimal_places,
                    config.get('min_power_w', 0.001),
                    config.get('max_power_w', 100.0)
                ]
            }
            metadata_df = pd.DataFrame(metadata)
            metadata_df.to_excel(writer, sheet_name='Analysis Info', index=False)
        
        if verbose:
            print(f"Results saved to Excel file: {excel_path}")
            print(f"Sheets created: All Channels, Named Channels, Non-Debug Channels, Clean View, Analysis Info")
        return excel_path
        
    except Exception as e:
        if verbose:
            print(f"Error saving to Excel: {e}")
        return None


def create_enhanced_plot(qpt_df: pd.DataFrame, config: Dict[str, Any]):
    """
    Create enhanced plot with configurable options
    
    Parameters:
    qpt_df: QPT data DataFrame
    config: Configuration dictionary
    
    Returns:
    Bokeh plot object or None if no data
    """
    if not BOKEH_AVAILABLE:
        if config.get('verbose', True):
            print("Bokeh not available. Cannot create interactive plots.")
        return None
        
    if qpt_df.empty:
        if config.get('verbose', True):
            print("No data available for plotting")
        return None
    
    # Get configuration parameters
    min_power_w = config.get('min_power_w', 0.001)
    palette_name = config.get('palette', 'viridis')
    plot_width = config.get('plot_width', 1000)
    plot_height = config.get('plot_height', 600)
    power_unit = config.get('power_unit', 'mW')
    show_debug = config.get('show_debug', False)
    show_unknown_channels = config.get('show_unknown_channels', False)
    
    # Apply filtering
    filtered_df = qpt_df.copy()
    
    if not show_debug:
        filtered_df = filtered_df[
            ~filtered_df['channel_display_name'].str.lower().str.contains('debug', na=False)
        ]
    
    if not show_unknown_channels:
        filtered_df = filtered_df[filtered_df['has_name'] == True]
    
    if filtered_df.empty:
        if config.get('verbose', True):
            print("No data available after filtering")
        return None
    
    # Calculate the start time for relative time
    start_time = filtered_df['timestamp'].min()
    filtered_df['relative_time'] = filtered_df['timestamp'] - start_time
    
    # Power unit conversion
    power_multiplier = {'W': 1.0, 'mW': 1000.0, 'uW': 1000000.0}
    multiplier = power_multiplier.get(power_unit, 1000.0)
    
    # Calculate average power for each channel and filter
    channels_power = {}
    for channel in filtered_df['channel_name'].unique():
        channel_data = filtered_df[filtered_df['channel_name'] == channel]
        power_converted = channel_data['avg_power_uw'].mean() / 1e6 * multiplier  # Convert to selected unit
        if power_converted >= min_power_w * multiplier:  # Convert threshold to selected unit
            display_name = channel_data.iloc[0]['channel_display_name']
            channels_power[channel] = {'power': power_converted, 'display_name': display_name}
    
    # Sort channels by average power (descending)
    sorted_channels = sorted(channels_power.items(), key=lambda x: x[1]['power'], reverse=True)
    filtered_channels = [ch for ch, _ in sorted_channels]
    
    if not filtered_channels:
        if config.get('verbose', True):
            print(f"No channels with average power above {min_power_w}W after filtering")
        return None
    
    # Create figure with enhanced title
    debug_status = "shown" if show_debug else "hidden"
    unknown_status = "shown" if show_unknown_channels else "hidden"
    
    p = figure(
        title=f"QPT Power Analysis (debug: {debug_status}, unknown: {unknown_status}, min: {min_power_w}W)",
        x_axis_label="Time (s)",
        y_axis_label=f"Power ({power_unit})",
        width=plot_width,
        height=plot_height,
        tools="pan,box_zoom,wheel_zoom,reset,save,hover",
        toolbar_location="above"
    )
    
    # Select color palette
    if palette_name.lower() == 'plasma':
        palette = Plasma256
    elif palette_name.lower() == 'turbo':
        palette = Turbo256
    else:
        palette = Viridis256  # Default
    
    # Get evenly spaced colors
    num_channels = len(filtered_channels)
    step = max(1, len(palette) // num_channels)
    colors = [palette[i * step % len(palette)] for i in range(num_channels)]
    
    # Plot each channel
    for i, channel in enumerate(filtered_channels):
        channel_data = filtered_df[filtered_df['channel_name'] == channel]
        power_converted = channel_data['avg_power_uw'] / 1e6 * multiplier  # Convert to selected unit
        display_name = channels_power[channel]['display_name']
        avg_power = channels_power[channel]['power']
        
        # Create data source
        source = ColumnDataSource(data={
            'x': channel_data['relative_time'],
            'y': power_converted,
            'channel': [display_name] * len(channel_data),
            'power_uw': channel_data['avg_power_uw'],
            'abs_time': channel_data['timestamp'],
        })
        
        # Plot line
        p.line(
            'x', 'y', 
            source=source, 
            line_width=2, 
            color=colors[i], 
            alpha=0.8,
            legend_label=f"{display_name} ({avg_power:.1f}{power_unit})"
        )
    
    # Configure hover tool
    hover = p.select(dict(type=HoverTool))
    hover.tooltips = [
        ("Channel", "@channel"),
        ("Time", "@x{0.0} s"),
        (f"Power ({power_unit})", "@y{0.0}")
    ]
    hover.mode = 'vline'
    
    # Configure legend
    p.legend.location = "right"
    p.legend.click_policy = "hide"
    p.legend.border_line_color = "navy"
    p.legend.background_fill_alpha = 0.7
    p.legend.label_text_font_size = "8pt"
    p.legend.glyph_width = 15
    p.legend.spacing = 0
    p.legend.padding = 5
    
    # Add grid lines
    p.grid.grid_line_alpha = 0.3
    
    return p


def save_plot_to_file(plot, trace_path: str, config: Dict[str, Any]):
    """
    Save plot to file
    
    Parameters:
    plot: Bokeh plot object
    trace_path: Original trace file path
    config: Configuration dictionary
    """
    if not BOKEH_AVAILABLE or plot is None:
        return
    
    try:
        from bokeh.io import export_png, export_svgs
        from bokeh.plotting import output_file, save
        
        output_dir = config.get('output_dir', os.path.dirname(trace_path))
        plot_format = config.get('plot_format', 'html')
        trace_filename = os.path.splitext(os.path.basename(trace_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if plot_format == 'html':
            plot_path = os.path.join(output_dir, f"{trace_filename}_plot_{timestamp}.html")
            output_file(plot_path)
            save(plot)
            if config.get('verbose', True):
                print(f"Plot saved to: {plot_path}")
        elif plot_format == 'png':
            plot_path = os.path.join(output_dir, f"{trace_filename}_plot_{timestamp}.png")
            export_png(plot, filename=plot_path)
            if config.get('verbose', True):
                print(f"Plot saved to: {plot_path}")
        elif plot_format == 'svg':
            plot_path = os.path.join(output_dir, f"{trace_filename}_plot_{timestamp}.svg")
            plot.output_backend = "svg"
            export_svgs(plot, filename=plot_path)
            if config.get('verbose', True):
                print(f"Plot saved to: {plot_path}")
                
    except Exception as e:
        if config.get('verbose', True):
            print(f"Error saving plot: {e}")


def analyze_power_telemetry_v2(trace_path: str, 
                              config: Dict[str, Any] = None,
                              **kwargs):
    """
    Main function to analyze power telemetry data with enhanced configuration options
    
    Parameters:
    trace_path: Path to trace file
    config: Configuration dictionary (optional)
    **kwargs: Override configuration parameters
    
    Returns:
    Tuple[pd.DataFrame, bokeh.plotting.figure, str]: QPT data, plot object, and Excel file path
    """
    
    # Load configuration
    if config is None:
        config = DEFAULT_CONFIG.copy()
    else:
        config = config.copy()
    
    # Override with kwargs
    config.update(kwargs)
    
    # Extract parameters from config
    show_debug = config.get('show_debug', False)
    show_unknown_channels = config.get('show_unknown_channels', False)
    min_power_w = config.get('min_power_w', 0.001)
    max_power_w = config.get('max_power_w', 100.0)
    min_time_diff_s = config.get('min_time_diff_s', 0.001)
    palette = config.get('palette', 'viridis')
    save_excel = config.get('save_excel', True)
    output_dir = config.get('output_dir', None)
    verbose = config.get('verbose', True)
    auto_detect_channels = config.get('auto_detect_channels', True)
    channel_filters = config.get('channel_filters', [])
    power_unit = config.get('power_unit', 'mW')
    decimal_places = config.get('decimal_places', 1)
    
    if verbose:
        print("=" * 60)
        print("Power Telemetry Analysis V2")
        print("=" * 60)
        print(f"Trace file: {trace_path}")
        print(f"Show debug channels: {show_debug}")
        print(f"Show unknown hex channels: {show_unknown_channels}")
        print(f"Power range: {min_power_w}W - {max_power_w}W")
        print(f"Minimum time difference: {min_time_diff_s}s")
        print(f"Power unit: {power_unit}")
        if channel_filters:
            print(f"Channel filters: {channel_filters}")
        print("-" * 60)
    
    # 1. Parse QPT data with all channels to get complete dataset
    qpt_df_all = parse_qpt_data_update_v2(trace_path, show_debug=True, show_unknown_channels=True)
    
    if qpt_df_all.empty:
        if verbose:
            print("No QPT data found.")
        return None, None, None
    
    # Auto-detect channel strategy if enabled
    if auto_detect_channels:
        summary = get_channel_summary_silent(qpt_df_all)
        if summary['named'] == 0 and summary['unnamed'] > 0:
            if verbose:
                print("Auto-detected: All channels are unnamed, enabling unknown channels display")
            show_unknown_channels = True
    
    # Apply channel filters if specified
    if channel_filters:
        qpt_df_all = apply_channel_filters(qpt_df_all, channel_filters, verbose)
    
    # 2. Calculate power for all channels with enhanced parameters
    qpt_results_all = calculate_rail_power_enhanced(qpt_df_all, min_time_diff_s, min_power_w, max_power_w)
    
    if qpt_results_all.empty:
        if verbose:
            print("No power results calculated.")
        return qpt_df_all, None, None
    
    # 3. Parse QPT data with current filter settings for display
    qpt_df = parse_qpt_data_update_v2(trace_path, show_debug, show_unknown_channels)
    if channel_filters:
        qpt_df = apply_channel_filters(qpt_df, channel_filters, verbose)
    qpt_results = calculate_rail_power_enhanced(qpt_df, min_time_diff_s, min_power_w, max_power_w) if not qpt_df.empty else pd.DataFrame()
    
    # 4. Display results with current filter settings
    if not qpt_results.empty:
        display_results_enhanced(qpt_results, show_debug, show_unknown_channels, power_unit, decimal_places, verbose)
    else:
        if verbose:
            print("No results to display with current filter settings.")
    
    # 5. Save to Excel if requested (using all data for comprehensive sheets)
    excel_path = None
    if save_excel:
        excel_path = save_results_to_excel_enhanced(qpt_results_all, qpt_df_all, trace_path, config)
    
    # 6. Create visualization with current filter settings
    plot = None
    if not qpt_df.empty:
        plot = create_enhanced_plot(qpt_df, config)
    
    # 7. Save plot if requested
    if config.get('save_plot', False) and plot is not None:
        save_plot_to_file(plot, trace_path, config)
    
    return qpt_df, plot, excel_path


def get_channel_summary(trace_path: str) -> Dict[str, int]:
    """
    Get a summary of channels in the trace file without filtering (with output)
    
    Parameters:
    trace_path: Path to trace file
    
    Returns:
    Dict with channel counts
    """
    # Parse all data without filtering
    qpt_df = parse_qpt_data_update_v2(trace_path, show_debug=True, show_unknown_channels=True)
    
    summary = get_channel_summary_silent(qpt_df)
    
    print("Channel Summary:")
    print(f"Total unique channels: {summary['total']}")
    print(f"Named channels: {summary['named']}")
    print(f"Unnamed (hex only) channels: {summary['unnamed']}")
    print(f"Debug channels: {summary['debug']}")
    
    return summary


def main():
    """
    Main function for command line interface
    """
    # Suppress warnings
    warnings.filterwarnings('ignore')
    
    parser = create_argument_parser()
    args = parser.parse_args()
    
    # Get trace file
    trace_file = args.trace_file
    if not trace_file:
        parser.print_help()
        return
    
    if not os.path.exists(trace_file):
        print(f"Error: Trace file not found: {trace_file}")
        return
    
    # Determine Excel filename
    excel_filename = args.excel_file
    if excel_filename:
        # Use provided filename
        if not excel_filename.endswith('.xlsx'):
            excel_filename += '.xlsx'
        output_dir = os.path.dirname(excel_filename) or os.path.dirname(trace_file)
        excel_template = os.path.basename(excel_filename)
    else:
        # Use default naming
        output_dir = os.path.dirname(trace_file)
        excel_template = "{trace_name}_power_analysis_{timestamp}.xlsx"
    
    # Create configuration
    config = DEFAULT_CONFIG.copy()
    config.update({
        'save_excel': True,
        'output_dir': output_dir,
        'excel_filename_template': excel_template,
        'verbose': True
    })
    
    # Perform analysis
    try:
        qpt_data, plot, excel_path = analyze_power_telemetry_v2(trace_file, config)
        
        if plot and BOKEH_AVAILABLE:
            show(plot)
        
        if excel_path:
            print(f"\n=== Analysis Complete ===")
            print(f"Excel file: {excel_path}")
                
    except Exception as e:
        print(f"Error during analysis: {e}")
        return


def get_channel_summary_silent(qpt_df: pd.DataFrame) -> Dict[str, int]:
    """
    Get channel summary without printing (for internal use)
    
    Parameters:
    qpt_df: QPT DataFrame
    
    Returns:
    Dict with channel counts
    """
    if qpt_df.empty:
        return {"total": 0, "named": 0, "unnamed": 0, "debug": 0}
    
    total_channels = qpt_df['channel_name'].nunique()
    named_channels = len(qpt_df[qpt_df['has_name'] == True]['channel_name'].unique())
    unnamed_channels = len(qpt_df[qpt_df['has_name'] == False]['channel_name'].unique())
    debug_channels = len([ch for ch in qpt_df['channel_display_name'].unique() 
                         if is_debug_channel(ch)])
    
    return {
        "total": total_channels,
        "named": named_channels,
        "unnamed": unnamed_channels,
        "debug": debug_channels
    }


def apply_channel_filters(qpt_df: pd.DataFrame, filters: List[str], verbose: bool = True) -> pd.DataFrame:
    """
    Apply channel name filters to DataFrame
    
    Parameters:
    qpt_df: QPT DataFrame
    filters: List of filter patterns (supports regex)
    verbose: Print filter results
    
    Returns:
    pd.DataFrame: Filtered DataFrame
    """
    if not filters or qpt_df.empty:
        return qpt_df
    
    original_count = len(qpt_df)
    filtered_df = qpt_df.copy()
    
    for filter_pattern in filters:
        try:
            # Apply regex filter to channel display names
            mask = filtered_df['channel_display_name'].str.contains(filter_pattern, case=False, na=False, regex=True)
            filtered_df = filtered_df[mask]
            if verbose:
                print(f"Applied filter '{filter_pattern}': {len(filtered_df)} channels remaining")
        except Exception as e:
            if verbose:
                print(f"Error applying filter '{filter_pattern}': {e}")
    
    if verbose and len(filters) > 0:
        print(f"Channel filtering: {original_count} -> {len(filtered_df)} channels")
    
    return filtered_df


def calculate_rail_power_enhanced(qpt_df: pd.DataFrame, min_time_diff_s: float = 0.001, 
                                 min_power_w: float = 0.0, max_power_w: float = 100.0) -> pd.DataFrame:
    """
    Enhanced power calculation with configurable thresholds
    
    Parameters:
    qpt_df: DataFrame containing QPT data
    min_time_diff_s: Minimum time difference threshold
    min_power_w: Minimum power threshold
    max_power_w: Maximum power threshold
    
    Returns:
    pd.DataFrame: Power calculation results
    """
    if qpt_df.empty:
        return pd.DataFrame()
        
    results = []
    channels = qpt_df['channel_name'].unique()
    
    for channel in channels:
        channel_df = qpt_df[qpt_df['channel_name'] == channel].sort_values('timestamp')
        if len(channel_df) <= 1:
            continue
            
        first = channel_df.iloc[0]
        last = channel_df.iloc[-1]
        
        time_diff = last['timestamp'] - first['timestamp']
        energy_diff = last['energy_uj'] - first['energy_uj']
        
        if time_diff > min_time_diff_s:  # Configurable time threshold
            power_w = energy_diff / time_diff / 1e6  # Convert to watts
            # Filter with configurable power range
            if min_power_w <= power_w <= max_power_w:
                results.append({
                    'channel_name': channel,
                    'channel_display_name': channel_df.iloc[0]['channel_display_name'],
                    'channel_hex': channel_df.iloc[0]['channel_hex'],
                    'has_name': channel_df.iloc[0]['has_name'],
                    'time_diff_s': time_diff,
                    'power_w': power_w,
                    'energy_start_uj': first['energy_uj'],
                    'energy_end_uj': last['energy_uj'],
                    'energy_diff_uj': energy_diff,
                    'sample_count': len(channel_df)
                })
                
    if not results:
        return pd.DataFrame()
        
    return pd.DataFrame(results).sort_values('power_w', ascending=False)


def display_results_enhanced(qpt_results: pd.DataFrame, show_debug: bool = False, 
                           show_unknown_channels: bool = False, power_unit: str = 'mW',
                           decimal_places: int = 1, verbose: bool = True):
    """
    Enhanced display function with configurable units and precision
    
    Parameters:
    qpt_results: Power calculation results
    show_debug: Show debug channels
    show_unknown_channels: Show unknown hex-only channels
    power_unit: Power unit for display ('mW', 'W', 'uW')
    decimal_places: Number of decimal places
    verbose: Enable verbose output
    """
    if qpt_results.empty:
        if verbose:
            print("No results to display.")
        return

    # Apply additional filtering if needed
    filtered_results = qpt_results.copy()
    
    if not show_debug:
        filtered_results = filtered_results[
            ~filtered_results['channel_display_name'].str.lower().str.contains('debug', na=False)
        ]
    
    if not show_unknown_channels:
        filtered_results = filtered_results[filtered_results['has_name'] == True]
    
    if filtered_results.empty:
        if verbose:
            print("No results to display after filtering.")
        return

    # Convert power to requested unit
    power_multiplier = {'W': 1.0, 'mW': 1000.0, 'uW': 1000000.0}
    multiplier = power_multiplier.get(power_unit, 1000.0)
    
    filtered_results[f"Power ({power_unit})"] = (filtered_results["power_w"] * multiplier).round(decimal_places)
    
    # Create simplified display dataframe with only Channel and Power
    simple_df = pd.DataFrame({
        "Channel": filtered_results["channel_display_name"],
        f"Power ({power_unit})": filtered_results[f"Power ({power_unit})"]
    })
    
    # Sort by power consumption
    simple_df = simple_df.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
    
    if verbose:
        # Display summary with clear filtering status
        total_channels = len(filtered_results)
        named_channels = len(filtered_results[filtered_results["has_name"] == True])
        unnamed_channels = total_channels - named_channels
        total_power = simple_df[f"Power ({power_unit})"].sum()
        
        print(f"Power Analysis Results Summary:")
        print(f"Total channels displayed: {total_channels}")
        print(f"Named channels: {named_channels}")
        print(f"Unnamed channels (hex only): {unnamed_channels}")
        print(f"Debug channels: {'Shown' if show_debug else 'Hidden'}")
        print(f"Unknown hex channels: {'Shown' if show_unknown_channels else 'Hidden'}")
        print(f"Total power consumption: {total_power:.{decimal_places}f} {power_unit}")
        print("-" * 50)
    
    # Print simplified table with just Channel and Power
    table_output = format_table_output(simple_df, power_unit)
    print(table_output)


# Convenience functions for backward compatibility
def quick_analysis_clean(trace_path: str, save_excel: bool = False, output_dir: str = None):
    """Quick analysis - only show named non-debug channels"""
    config = DEFAULT_CONFIG.copy()
    config.update({
        'show_debug': False,
        'show_unknown_channels': False,
        'save_excel': save_excel,
        'output_dir': output_dir
    })
    return analyze_power_telemetry_v2(trace_path, config)


def quick_analysis_with_debug(trace_path: str, save_excel: bool = False, output_dir: str = None):
    """Quick analysis - show all named channels (including debug)"""
    config = DEFAULT_CONFIG.copy()
    config.update({
        'show_debug': True,
        'show_unknown_channels': False,
        'save_excel': save_excel,
        'output_dir': output_dir
    })
    return analyze_power_telemetry_v2(trace_path, config)


def quick_analysis_all(trace_path: str, save_excel: bool = False, output_dir: str = None):
    """Quick analysis - show all channels"""
    config = DEFAULT_CONFIG.copy()
    config.update({
        'show_debug': True,
        'show_unknown_channels': True,
        'save_excel': save_excel,
        'output_dir': output_dir
    })
    return analyze_power_telemetry_v2(trace_path, config)


def quick_analysis_unknown_only(trace_path: str, save_excel: bool = False, output_dir: str = None):
    """Quick analysis - only show unknown channels (for debugging)"""
    config = DEFAULT_CONFIG.copy()
    config.update({
        'show_debug': False,
        'show_unknown_channels': True,
        'save_excel': save_excel,
        'output_dir': output_dir
    })
    return analyze_power_telemetry_v2(trace_path, config)


def batch_analysis_with_excel(trace_paths: List[str], output_dir: str = None, 
                              show_debug: bool = False, show_unknown_channels: bool = False) -> List[str]:
    """
    Batch analysis of multiple trace files with Excel export
    
    Parameters:
    trace_paths: List of trace file paths
    output_dir: Output directory for Excel files (default: same as each trace file)
    show_debug: Show debug channels
    show_unknown_channels: Show unknown hex-only channels
    
    Returns:
    List[str]: List of generated Excel file paths
    """
    config = DEFAULT_CONFIG.copy()
    config.update({
        'show_debug': show_debug,
        'show_unknown_channels': show_unknown_channels,
        'save_excel': True,
        'output_dir': output_dir
    })
    
    excel_files = []
    
    for i, trace_path in enumerate(trace_paths, 1):
        print(f"\n{'='*60}")
        print(f"Processing file {i}/{len(trace_paths)}: {os.path.basename(trace_path)}")
        print(f"{'='*60}")
        
        try:
            qpt_data, plot, excel_path = analyze_power_telemetry_v2(trace_path, config)
            
            if excel_path:
                excel_files.append(excel_path)
                
        except Exception as e:
            print(f"Error processing {trace_path}: {e}")
            continue
    
    print(f"\n{'='*60}")
    print(f"Batch Analysis Complete")
    print(f"{'='*60}")
    print(f"Processed {len(trace_paths)} files")
    print(f"Generated {len(excel_files)} Excel files")
    
    if excel_files:
        print("\nGenerated Excel files:")
        for excel_file in excel_files:
            print(f"  - {excel_file}")
    
    return excel_files


# Example usage and testing
if __name__ == "__main__":
    import sys
    
    # If command line arguments provided, use CLI mode
    if len(sys.argv) > 1:
        main()
    else:
        # Example usage mode
        # Example trace path - update this to your actual file path
        trace_path = r"C:\Project\DoU\CASE01_ftrace.txt"
        # trace_path = r"C:\Project\telemetry\tencentvideo_new_20251210\ftrace\1\Default_ftrace.txt"
        
        # Check if file exists
        if not os.path.exists(trace_path):
            print(f"Example trace file not found: {trace_path}")
            print("Please update the trace_path variable with your actual file path.")
            print("\nOr use command line mode:")
            print("python TelemetryParserQPTV2.py your_trace_file.txt")
            print("python TelemetryParserQPTV2.py --help  # for more options")
        else:
            print("=== Channel Summary ===")
            summary = get_channel_summary(trace_path)
            
            # Decide analysis strategy based on summary
            if summary['named'] == 0 and summary['unnamed'] > 0:
                print("\nNote: All channels are unnamed (hex only) channels")
                print("Will use show_unknown_channels=True for analysis\n")
                show_unknown = True
            else:
                show_unknown = False
            
            print("\n=== Power Analysis with Excel Export ===")
            # Perform analysis and generate Excel file with all sheets
            script_dir = os.path.dirname(os.path.abspath(__file__))
            
            config = DEFAULT_CONFIG.copy()
            config.update({
                'show_debug': False,
                'show_unknown_channels': show_unknown,
                'save_excel': True,
                'output_dir': script_dir
            })
            
            qpt_data, power_plot, excel_path = analyze_power_telemetry_v2(trace_path, config)
            
            if power_plot and BOKEH_AVAILABLE:
                show(power_plot)
            
            if excel_path:
                print(f"\n=== Excel File Generated ===")
                print(f"File path: {excel_path}")
                print("Contains the following sheets:")
                print("- All Channels: All channels (show name if available, show ID if not)")
                print("- Named Channels: Only named channels")
                print("- Non-Debug Channels: Non-debug channels")
                print("- Clean View: Clean view (named non-debug channels)")
                print("- Analysis Info: Analysis information")