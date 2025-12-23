#!/usr/bin/env python3
"""
TelemetryParserQPTV3.py - Simplified QPT Telemetry Parser V3

Features:
- Direct channel name extraction from logs
- Auto-detection of channel types
- Excel export functionality
- Simple command line interface
- Streamlined output showing only Channel and Power

Author: Generated for Power Analysis
Version: 3.0
"""

import os
import re
import warnings
import argparse
from typing import Dict, Any, List
import pandas as pd
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


def save_results_to_excel_custom(qpt_results: pd.DataFrame, excel_path: str, 
                               power_unit: str = 'mW', decimal_places: int = 1) -> str:
    """Save results to a specified Excel file path"""
    if qpt_results.empty:
        print("No results to save.")
        return None
    
    try:
        # Power unit conversion
        power_multiplier = {'W': 1.0, 'mW': 1000.0, 'uW': 1000000.0}
        multiplier = power_multiplier.get(power_unit, 1000.0)
        
        # Create Excel writer
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            
            # Sheet 1: All Channels - Include all columns for comprehensive data (including debug channels)
            all_results = qpt_results.copy()
            if not all_results.empty:
                all_results[f"Power ({power_unit})"] = (all_results["power_w"] * multiplier).round(decimal_places)
                
                # Create a comprehensive dataframe with all relevant information
                all_summary = pd.DataFrame({
                    "Channel": all_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
                    "Channel ID": all_results["channel_hex"],
                    f"Power ({power_unit})": all_results[f"Power ({power_unit})"],
                    "Has Name": all_results["has_name"],
                    "Is Debug": all_results.apply(lambda row: is_debug_channel(row["channel_display_name"]), axis=1),
                    "Duration (s)": all_results["time_diff_s"].round(2),
                    "Energy (uJ)": all_results["energy_diff_uj"],
                    "Samples": all_results["sample_count"]
                })
                
                all_summary = all_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                all_summary.to_excel(writer, sheet_name='All Channels', index=False)
            
            # Sheet 2: Named Channels Only - Simplified view (excluding debug channels)
            named_results = qpt_results[qpt_results['has_name'] == True].copy()
            if not named_results.empty:
                # Filter out debug channels from named channels sheet
                named_results = named_results[~named_results['channel_display_name'].str.lower().str.contains('^debug[_-]', regex=True)]
                
                if not named_results.empty:
                    named_results[f"Power ({power_unit})"] = (named_results["power_w"] * multiplier).round(decimal_places)
                    named_summary = pd.DataFrame({
                        "Channel": named_results["channel_display_name"],
                        f"Power ({power_unit})": named_results[f"Power ({power_unit})"]
                    })
                    named_summary = named_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                    named_summary.to_excel(writer, sheet_name='Named Channels', index=False)
            
            # Sheet 3: Non-Debug Channels - Simplified view
            non_debug_results = qpt_results[
                ~qpt_results['channel_display_name'].str.lower().str.contains('debug', na=False)
            ].copy()
            if not non_debug_results.empty:
                non_debug_results[f"Power ({power_unit})"] = (non_debug_results["power_w"] * multiplier).round(decimal_places)
                non_debug_summary = pd.DataFrame({
                    "Channel": non_debug_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
                    f"Power ({power_unit})": non_debug_results[f"Power ({power_unit})"]
                })
                non_debug_summary = non_debug_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                non_debug_summary.to_excel(writer, sheet_name='Non-Debug Channels', index=False)
        
        return excel_path
        
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return None


def format_table_output(results_df: pd.DataFrame, power_unit: str = 'mW') -> str:
    """Format results as a nice table for CLI output"""
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


def create_argument_parser():
    """Create command line argument parser"""
    parser = argparse.ArgumentParser(
        description="QPT Telemetry Parser V3 - Enhanced power analysis tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python TelemetryParserQPTV3.py trace.txt
  python TelemetryParserQPTV3.py trace.txt output.xlsx
        """
    )
    
    # Required arguments
    parser.add_argument('trace_file', help='Path to trace file')
    parser.add_argument('excel_file', nargs='?', help='Excel output filename (optional)')
    
    return parser


def is_debug_channel(channel_name: str) -> bool:
    """Check if channel is a debug channel"""
    if pd.isna(channel_name):
        return False
    channel_str = str(channel_name).lower()
    return channel_str.startswith('debug-') or channel_str.startswith('debug_')


def parse_qpt_data_update_v3(trace_path: str, show_debug: bool = False, show_unknown_channels: bool = False) -> pd.DataFrame:
    """Parse qpt_data_update events from trace file"""
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
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        return df
        
    except FileNotFoundError:
        print(f"File not found: {trace_path}")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error parsing QPT data: {e}")
        return pd.DataFrame()


def calculate_rail_power_v3(qpt_df: pd.DataFrame, min_time_diff_s: float = 0.001, 
                           min_power_w: float = 0.0, max_power_w: float = 100.0) -> pd.DataFrame:
    """Calculate power with configurable thresholds"""
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


def display_results_v3(qpt_results: pd.DataFrame, show_debug: bool = False, 
                      show_unknown_channels: bool = False, power_unit: str = 'mW',
                      decimal_places: int = 1, verbose: bool = True):
    """Display results in a simple table format with only Channel and Power"""
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
    
    # Create simplified display dataframe with only Channel and Power
    simple_df = pd.DataFrame({
        "Channel": filtered_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
        f"Power ({power_unit})": filtered_results[f"Power ({power_unit})"]
    })
    
    # Sort by power consumption (descending)
    simple_df = simple_df.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
    
    # Display formatted table with just Channel and Power
    table_output = format_table_output(simple_df, power_unit)
    print(table_output)


def save_results_to_excel_v3(qpt_results: pd.DataFrame, trace_path: str, 
                            power_unit: str = 'mW', decimal_places: int = 1) -> str:
    """Save results to Excel file with simplified sheets"""
    if qpt_results.empty:
        print("No results to save.")
        return None
    
    # Determine output directory and filename
    output_dir = os.path.dirname(trace_path)
    trace_filename = os.path.splitext(os.path.basename(trace_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"{trace_filename}_power_analysis_{timestamp}.xlsx"
    excel_path = os.path.join(output_dir, excel_filename)
    
    try:
        # Power unit conversion
        power_multiplier = {'W': 1.0, 'mW': 1000.0, 'uW': 1000000.0}
        multiplier = power_multiplier.get(power_unit, 1000.0)
        
        # Create Excel writer
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            
            # Sheet 1: All Channels - Include all columns for comprehensive data
            all_results = qpt_results.copy()
            if not all_results.empty:
                all_results[f"Power ({power_unit})"] = (all_results["power_w"] * multiplier).round(decimal_places)
                
                # Create a comprehensive dataframe with all relevant information
                all_summary = pd.DataFrame({
                    "Channel": all_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
                    "Channel ID": all_results["channel_hex"],
                    f"Power ({power_unit})": all_results[f"Power ({power_unit})"],
                    "Has Name": all_results["has_name"],
                    "Duration (s)": all_results["time_diff_s"].round(2),
                    "Energy (uJ)": all_results["energy_diff_uj"],
                    "Samples": all_results["sample_count"]
                })
                
                all_summary = all_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                all_summary.to_excel(writer, sheet_name='All Channels', index=False)
            
            # Sheet 2: Named Channels Only - Simplified view (excluding debug channels)
            named_results = qpt_results[qpt_results['has_name'] == True].copy()
            if not named_results.empty:
                # Filter out debug channels from named channels sheet
                named_results = named_results[~named_results['channel_display_name'].str.lower().str.contains('^debug[_-]', regex=True)]
                
                if not named_results.empty:
                    named_results[f"Power ({power_unit})"] = (named_results["power_w"] * multiplier).round(decimal_places)
                    named_summary = pd.DataFrame({
                        "Channel": named_results["channel_display_name"],
                        f"Power ({power_unit})": named_results[f"Power ({power_unit})"]
                    })
                    named_summary = named_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                    named_summary.to_excel(writer, sheet_name='Named Channels', index=False)
            
            # Sheet 3: Non-Debug Channels - Simplified view
            non_debug_results = qpt_results[
                ~qpt_results['channel_display_name'].str.lower().str.contains('debug', na=False)
            ].copy()
            if not non_debug_results.empty:
                non_debug_results[f"Power ({power_unit})"] = (non_debug_results["power_w"] * multiplier).round(decimal_places)
                non_debug_summary = pd.DataFrame({
                    "Channel": non_debug_results.apply(lambda row: row["channel_display_name"] if row["has_name"] else row["channel_hex"], axis=1),
                    f"Power ({power_unit})": non_debug_results[f"Power ({power_unit})"]
                })
                non_debug_summary = non_debug_summary.sort_values(f"Power ({power_unit})", ascending=False).reset_index(drop=True)
                non_debug_summary.to_excel(writer, sheet_name='Non-Debug Channels', index=False)
        
        return excel_path
        
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return None


def get_channel_summary_silent(qpt_df: pd.DataFrame) -> Dict[str, int]:
    """Get channel summary without printing (for internal use)"""
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


def analyze_power_telemetry_v3(trace_path: str, show_debug: bool = False, 
                              show_unknown_channels: bool = False, save_excel: bool = True,
                              power_unit: str = 'mW', decimal_places: int = 1, excel_path: str = None):
    """Main function for power telemetry analysis V3"""
    # 1. Parse QPT data for display (with filtering)
    qpt_df = parse_qpt_data_update_v3(trace_path, show_debug, show_unknown_channels)
    
    if qpt_df.empty:
        print("No QPT data found.")
        return None, None
    
    # 2. Calculate power for display
    qpt_results = calculate_rail_power_v3(qpt_df)
    
    if qpt_results.empty:
        print("No power results calculated.")
        return qpt_df, None
    
    # 3. Display results (with filtering)
    display_results_v3(qpt_results, show_debug, show_unknown_channels, power_unit, decimal_places)
    
    # 4. For Excel export, parse ALL data including debug channels
    if save_excel:
        # Parse all data without filtering for Excel export
        qpt_df_all = parse_qpt_data_update_v3(trace_path, show_debug=True, show_unknown_channels=True)
        qpt_results_all = calculate_rail_power_v3(qpt_df_all)
        
        output_excel_path = None
        if excel_path:
            # Use specified Excel path
            output_excel_path = save_results_to_excel_custom(qpt_results_all, excel_path, power_unit, decimal_places)
        else:
            # Use default naming
            output_excel_path = save_results_to_excel_v3(qpt_results_all, trace_path, power_unit, decimal_places)
        
        return qpt_df, output_excel_path
    
    return qpt_df, None


def main():
    """Main function for command line interface"""
    # Suppress warnings
    warnings.filterwarnings('ignore')
    
    parser = create_argument_parser()
    args = parser.parse_args()
    
    # Get trace file
    trace_file = args.trace_file
    if not os.path.exists(trace_file):
        print(f"Error: Trace file not found: {trace_file}")
        return
    
    # Determine Excel filename
    excel_filename = args.excel_file
    excel_path = None
    if excel_filename:
        # Use provided filename
        if not excel_filename.endswith('.xlsx'):
            excel_filename += '.xlsx'
        output_dir = os.path.dirname(excel_filename) or os.path.dirname(trace_file)
        if not output_dir:
            output_dir = os.getcwd()
        excel_path = os.path.join(output_dir, os.path.basename(excel_filename))
    
    # Perform analysis
    try:
        # First get channel summary to determine whether to show unknown channels
        qpt_df_all = parse_qpt_data_update_v3(trace_file, show_debug=True, show_unknown_channels=True)
        summary = get_channel_summary_silent(qpt_df_all)
        
        # If only unnamed channels, automatically show them
        show_unknown = False
        if summary['named'] == 0 and summary['unnamed'] > 0:
            show_unknown = True
        
        # Parse QPT data with current filter settings for display
        qpt_df = parse_qpt_data_update_v3(trace_file, show_debug=False, show_unknown_channels=show_unknown)
        
        if qpt_df.empty:
            return
        
        # Calculate power for display
        qpt_results = calculate_rail_power_v3(qpt_df)
        
        if qpt_results.empty:
            return
        
        # Display results
        display_results_v3(qpt_results, show_debug=False, show_unknown_channels=show_unknown)
        
        # For Excel export, parse ALL data including debug channels
        qpt_df_all = parse_qpt_data_update_v3(trace_file, show_debug=True, show_unknown_channels=True)
        qpt_results_all = calculate_rail_power_v3(qpt_df_all)
        
        # Save to Excel
        if excel_path:
            # Use specified Excel path
            save_results_to_excel_custom(qpt_results_all, excel_path)
        else:
            # Use default naming
            save_results_to_excel_v3(qpt_results_all, trace_file)
                
    except Exception as e:
        print(f"Error during analysis: {e}")
        return


if __name__ == "__main__":
    import sys
    
    # If command line arguments provided, use CLI mode
    if len(sys.argv) > 1:
        main()
    else:
        # Example usage mode
        # Example trace path - update this to your actual file path
        trace_path = r"C:\Project\DoU\CASE01_ftrace.txt"
        
        # Check if file exists
        if not os.path.exists(trace_path):
            print(f"Example trace file not found: {trace_path}")
            print("Please update the trace_path variable with your actual file path.")
            print("\nOr use command line mode:")
            print("python TelemetryParserQPTV3.py your_trace_file.txt")
            print("python TelemetryParserQPTV3.py your_trace_file.txt output.xlsx")
        else:
            # Decide analysis strategy based on summary
            qpt_df_all = parse_qpt_data_update_v3(trace_path, show_debug=True, show_unknown_channels=True)
            summary = get_channel_summary_silent(qpt_df_all)
            
            show_unknown = False
            if summary['named'] == 0 and summary['unnamed'] > 0:
                show_unknown = True
            
            # Parse QPT data with current filter settings for display
            qpt_df = parse_qpt_data_update_v3(trace_path, show_debug=False, show_unknown_channels=show_unknown)
            
            if not qpt_df.empty:
                # Calculate power for display
                qpt_results = calculate_rail_power_v3(qpt_df)
                
                if not qpt_results.empty:
                    # Display results
                    display_results_v3(qpt_results, show_debug=False, show_unknown_channels=show_unknown)
                    
                    # For Excel export, parse ALL data including debug channels
                    qpt_df_all = parse_qpt_data_update_v3(trace_path, show_debug=True, show_unknown_channels=True)
                    qpt_results_all = calculate_rail_power_v3(qpt_df_all)
                    
                    # Default Excel path in script directory
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    trace_filename = os.path.splitext(os.path.basename(trace_path))[0]
                    default_excel_path = os.path.join(script_dir, f"{trace_filename}_power_analysis_{timestamp}.xlsx")
                    
                    # Save to Excel with ALL channels
                    save_results_to_excel_custom(qpt_results_all, default_excel_path)