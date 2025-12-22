import os
import re
import sys
from typing import Optional, Dict, Any, Tuple, List
import pandas as pd
import numpy as np
from tabulate import tabulate

# ----------------------------------------------------------------------------------------
# Channel Mapping Configuration
# ----------------------------------------------------------------------------------------
channel_mapping = pd.DataFrame({
    'Channel_name': [
        'debug-12', 'debug-11', 'debug-10', 'debug-9', 'debug-8', 'debug-7', 'debug-6', 'debug-5',
        'debug-4', 'debug-3', 'debug-2', 'debug-1', 'debug-0', 'nsp-spare', 'nsp', 'gpu-spare',
        'gpu', 'cpu-l-spare', 'cpu-l', 'cpu-m-spare', 'cpu-m'
    ],
    'Channel_ID': [
        '0x8ad', '0x3b6', '0x9a1', '0x99e', '0x5ad', '0x6ad', '0x9a4', '0x6b0',
        '0x3ad', '0x39b', '0x6a7', '0x8b6', '0x5b3', '0x3a4', '0x39e', '0x5b6',
        '0x59b', '0x8b3', '0x89e', '0x8b0', '0x6b6'
    ],
    'Breakdown_rail': [
        '', '', '', '', '', '', '', '',
        '', '', '', '', '', 'NSP2 CX', 'NSP1 CX', 'GFX_MXC',
        'GFX', 'APC1_MX', 'APC1_CX', 'APC0_MX', 'APC0_CX'
    ]
})

def fill_channel_names_from_mapping(qpt_df: pd.DataFrame, mapping_df: pd.DataFrame) -> pd.DataFrame:
    """Fill missing channel names using the channel ID mapping."""
    if qpt_df.empty or mapping_df.empty:
        return qpt_df
    
    m = mapping_df.copy()
    m["Channel_ID_norm"] = m["Channel_ID"].astype(str).str.strip().str.lower()
    id_to_name = dict(zip(m["Channel_ID_norm"], m["Channel_name"]))
    
    def get_name_from_id(row):
        if pd.notna(row['channel_name']):
            return row['channel_name']
        channel_hex = row['channel_hex']
        return id_to_name.get(channel_hex, f"channel_{channel_hex}")
        
    qpt_df['channel_name'] = qpt_df.apply(get_name_from_id, axis=1)
    return qpt_df

def parse_qpt_data_update(trace_path: str) -> pd.DataFrame:
    """Parse qpt_data_update events from trace file."""
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
                        data.append({
                            'timestamp': float(match.group(2)),
                            'channel_hex': match.group(3).lower(),
                            'channel_name': match.group(4),
                            'energy_uj': int(match.group(6)),
                            'avg_power_uw': int(match.group(7))
                        })
                    else:
                        match = pattern_without_name.search(line)
                        if match:
                            data.append({
                                'timestamp': float(match.group(2)),
                                'channel_hex': match.group(3).lower(),
                                'channel_name': None,
                                'energy_uj': int(match.group(5)),
                                'avg_power_uw': int(match.group(6))
                            })
                            
        if not data:
            print("No QPT data events found.")
            return pd.DataFrame()
            
        df = pd.DataFrame(data)
        if df['channel_name'].isna().any():
            df = fill_channel_names_from_mapping(df, channel_mapping)
        return df
        
    except FileNotFoundError:
        print(f"File not found: {trace_path}")
        return pd.DataFrame()
    
def attach_channel_mapping(qpt_df: pd.DataFrame, mapping_df: pd.DataFrame) -> pd.DataFrame:
    """Attach Channel_ID and Breakdown_rail to the QPT dataframe."""
    if qpt_df.empty or mapping_df.empty:
        return qpt_df
        
    m = mapping_df.copy()
    m["Channel_name_norm"] = m["Channel_name"].astype(str).str.strip().str.lower()
    
    df = qpt_df.copy()
    df["channel_name_norm"] = df["channel_name"].astype(str).str.strip().str.lower()
    
    # Merge on channel name
    joined = df.merge(
        m[["Channel_name_norm", "Channel_ID", "Breakdown_rail"]],
        left_on="channel_name_norm",
        right_on="Channel_name_norm",
        how="left"
    )
    return joined

def calculate_rail_power(qpt_df: pd.DataFrame) -> pd.DataFrame:
    """Calculate average power (W) using first and last energy points."""
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
        
        if time_diff > 0.001:
            power_w = energy_diff / time_diff / 1e6
            # Filter unreasonable values (0 to 100W)
            if 0 <= power_w <= 100:
                results.append({
                    'channel_name': channel,
                    'time_diff_s': time_diff,
                    'power_w': power_w
                })
                
    if not results:
        return pd.DataFrame()
        
    return pd.DataFrame(results).sort_values('power_w', ascending=False)

def display_results(qpt_results: pd.DataFrame, qpt_full: pd.DataFrame, output_file: Optional[str] = None):
    """Build and display the QPT power analysis results."""
    if qpt_results.empty:
        print("No results to display.")
        return

    # Get mapping details for each channel
    meta = qpt_full.groupby('channel_name', as_index=False)[['Channel_ID', 'Breakdown_rail']].first()
    merged = qpt_results.merge(meta, on='channel_name', how='left')
    
    # Convert to mW
    merged["QPT Power (mW)"] = merged["power_w"] * 1000.0
    
    # Format Output
    output_columns = {
        "Channel": merged["channel_name"],
        "Channel ID": merged["Channel_ID"].fillna(""),
        "Breakdown Rail": merged["Breakdown_rail"].fillna(""),
        "QPT Power (mW)": merged["QPT Power (mW)"].round(1)
    }
    
    final_df = pd.DataFrame(output_columns)
    
    # Filter debug channels
    mask = ~final_df["Channel"].astype(str).str.lower().str.startswith("debug")
    final_df = final_df[mask].sort_values("QPT Power (mW)", ascending=False).reset_index(drop=True)
    
    # Print results to console
    print("\n" + "="*80)
    print("QPT POWER ANALYSIS RESULTS")
    print("="*80 + "\n")
    
    # Use tabulate for pretty printing
    try:
        print(tabulate(final_df, headers='keys', tablefmt='grid', showindex=False))
    except:
        # Fallback to pandas default display
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', None)
        print(final_df.to_string(index=False))
    
    # Calculate and display summary statistics
    print("\n" + "="*80)
    print("SUMMARY STATISTICS")
    print("="*80)
    print(f"Total QPT Power: {final_df['QPT Power (mW)'].sum():.1f} mW")
    print(f"Number of Channels: {len(final_df)}")
    print(f"Average Power per Channel: {final_df['QPT Power (mW)'].mean():.1f} mW")
    print(f"Max Power Channel: {final_df.iloc[0]['Channel']} ({final_df.iloc[0]['QPT Power (mW)']:.1f} mW)")
    print("="*80 + "\n")
    
    # Save to CSV if output file specified
    if output_file:
        final_df.to_csv(output_file, index=False)
        print(f"Results saved to: {output_file}\n")
    
    return final_df

def main(trace_path: str, output_file: Optional[str] = None):
    """Main execution function for QPT-only analysis."""
    print("\n" + "="*80)
    print("QPT TELEMETRY PARSER - Power Analysis Tool")
    print("="*80 + "\n")
    
    # 1. Parse QPT
    print("[1/3] Parsing QPT data...")
    qpt_df = parse_qpt_data_update(trace_path)
    if qpt_df.empty:
        print("ERROR: No QPT data found. Exiting.")
        return None
    
    qpt_df = attach_channel_mapping(qpt_df, channel_mapping)
    print(f"      Found {len(qpt_df)} QPT events across {qpt_df['channel_name'].nunique()} channels\n")
    
    # 2. Calculate QPT Power
    print("[2/3] Calculating QPT power...")
    qpt_results = calculate_rail_power(qpt_df)
    if qpt_results.empty:
        print("ERROR: Could not calculate power. Exiting.")
        return None
    print(f"      Calculated power for {len(qpt_results)} channels\n")
    
    # 3. Display Final Table
    print("[3/3] Generating results...\n")
    results_df = display_results(qpt_results, qpt_df, output_file)
    
    return results_df


if __name__ == "__main__":
    # ==========================================
    # INPUT CONFIGURATION
    # ==========================================
    
    # Default paths (can be overridden by command line arguments)
    TRACE_PATH = r"C:\Project\telemetry\tencentvideo_new_20251210\ftrace\1\Default_ftrace.txt"
    OUTPUT_FILE = None  # Set to a path like "results.csv" to save output
    
    # Parse command line arguments if provided
    if len(sys.argv) > 1:
        TRACE_PATH = sys.argv[1]
    if len(sys.argv) > 2:
        OUTPUT_FILE = sys.argv[2]
    
    # Validate trace path
    if not os.path.exists(TRACE_PATH):
        print(f"ERROR: Trace file not found: {TRACE_PATH}")
        print("\nUsage: python TelemetryParserQPT.py <trace_path> [output_csv]")
        sys.exit(1)
    
    # Run analysis
    try:
        results = main(TRACE_PATH, OUTPUT_FILE)
        if results is not None:
            print("QPT Analysis completed successfully!")
        else:
            print("Analysis failed.")
            sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)