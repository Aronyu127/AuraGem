#!/usr/bin/env python3
"""
Sync Google Sheets data and generate JSON metadata files

This script:
1. Downloads CSV from Google Sheets
2. Converts CSV to JSON metadata files

Usage:
    python sync_google_sheet.py
    python sync_google_sheet.py --skip-download
    python sync_google_sheet.py --spreadsheet-id <ID> --gid <GID>
"""
import urllib.request
import urllib.error
import csv
import json
import sys
import argparse
from pathlib import Path

# Default Google Sheets configuration
DEFAULT_SPREADSHEET_ID = "1BsTCewXvYLHTEatDZbPTgQOyu0fQf9BowmycI8dI25k"
DEFAULT_GID = "1432562920"

# Default paths
SCRIPT_DIR = Path(__file__).parent
DEFAULT_CSV_FILE = SCRIPT_DIR / "metadata" / "raw_data" / "gem.csv"
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "metadata" / "json"

# Required CSV columns (filter out extra columns)
REQUIRED_COLUMNS = [
    'ID',
    '名稱',
    'Category',
    'Color',
    'Cut',
    'Carat',
    'Clarity',
    'Rarity',
    'image',
    'animation_url'
]


def download_csv(spreadsheet_id, gid, output_file):
    """
    Download CSV from Google Sheets and save to local file
    
    Args:
        spreadsheet_id: Google Sheets spreadsheet ID
        gid: Google Sheets tab GID
        output_file: Path to save the CSV file
    
    Returns:
        True if successful, False otherwise
    """
    # Construct the export URL
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?"
        f"format=csv&gid={gid}"
    )
    
    print(f"📥 Downloading CSV from Google Sheets...")
    print(f"   URL: {export_url}")
    
    try:
        # Create request with headers to avoid 403 errors
        req = urllib.request.Request(
            export_url,
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
        )
        
        # Download the file
        with urllib.request.urlopen(req) as response:
            data = response.read()
            
            # Ensure output directory exists
            output_file = Path(output_file)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Write to file
            with open(output_file, 'wb') as f:
                f.write(data)
            
            # Get file size
            file_size = len(data)
            print(f"✓ Successfully downloaded CSV file")
            print(f"  Saved to: {output_file}")
            print(f"  File size: {file_size:,} bytes")
            
            # Count lines
            with open(output_file, 'r', encoding='utf-8') as f:
                line_count = sum(1 for _ in f)
            print(f"  Total lines: {line_count}\n")
            
            return True
            
    except urllib.error.HTTPError as e:
        print(f"✗ HTTP Error: {e.code} - {e.reason}")
        if e.code == 403:
            print("  This might be because the sheet is not publicly accessible.")
            print("  Please make sure the Google Sheet is shared with 'Anyone with the link' can view.")
        return False
        
    except urllib.error.URLError as e:
        print(f"✗ URL Error: {e.reason}")
        print("  Please check your internet connection.")
        return False
        
    except Exception as e:
        print(f"✗ Error: {str(e)}")
        return False


def generate_description(cut, color):
    """
    Generate description based on cut and color
    
    Args:
        cut: Cut type (e.g., "Brilliant", "Princess")
        color: Color name (e.g., "Ruby", "Emerald")
    
    Returns:
        Description string
    """
    # Handle special case for "RoundBrilliant" vs "Brilliant"
    if cut == "Brilliant":
        cut_display = "RoundBrilliant"
    else:
        cut_display = cut
    
    return f"A unique {cut_display} cut gemstone in {color} color."


def csv_to_json(csv_file, output_dir):
    """
    Convert CSV data to JSON metadata files
    
    Args:
        csv_file: Path to CSV file
        output_dir: Directory to save JSON files
    
    Returns:
        True if successful, False otherwise
    """
    csv_file = Path(csv_file)
    output_dir = Path(output_dir)
    
    # Ensure output directory exists
    output_dir.mkdir(parents=True, exist_ok=True)
    
    if not csv_file.exists():
        print(f"✗ CSV file not found: {csv_file}")
        return False
    
    print(f"📝 Converting CSV to JSON files...")
    print(f"   Input: {csv_file}")
    print(f"   Output: {output_dir}")
    print(f"   Using only required columns: {', '.join(REQUIRED_COLUMNS)}")
    
    try:
        with open(csv_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            
            # Only process required columns, extra columns will be ignored
            processed_count = 0
            skipped_count = 0
            
            for row in reader:
                # Skip empty rows
                if not row.get('ID') or not row.get('ID').strip():
                    continue
                
                try:
                    # Only extract required columns, ignore extra columns
                    gem_id = row.get('ID', '').strip()
                    name = row.get('名稱', '').strip()
                    category = row.get('Category', '').strip()
                    color_code = row.get('Color', '').strip()
                    cut = row.get('Cut', '').strip()
                    carat = row.get('Carat', '').strip()
                    clarity = row.get('Clarity', '').strip()
                    rarity = row.get('Rarity', '').strip()
                    image = row.get('image', '').strip()
                    animation_url = row.get('animation_url', '').strip()
                    
                    # Use Category as color name (e.g., "Ruby"), Color column is the hex code
                    color_name = category if category else color_code
                    
                    # Validate required fields
                    if not all([name, color_name, cut, carat, clarity, rarity, image, animation_url]):
                        print(f"⚠ Skipping row {gem_id}: Missing required fields")
                        skipped_count += 1
                        continue
                    
                    # Convert carat to integer
                    try:
                        carat_value = int(carat)
                    except ValueError:
                        print(f"⚠ Skipping row {gem_id}: Invalid carat value '{carat}'")
                        skipped_count += 1
                        continue
                    
                    # Generate description using color name (Category)
                    description = generate_description(cut, color_name)
                    
                    # Build JSON structure
                    json_data = {
                        "name": name,
                        "description": description,
                        "attributes": [
                            {
                                "trait_type": "Cut",
                                "value": cut
                            },
                            {
                                "trait_type": "Color",
                                "value": color_name
                            },
                            {
                                "trait_type": "Carat",
                                "value": carat_value
                            },
                            {
                                "trait_type": "Clarity",
                                "value": clarity
                            },
                            {
                                "trait_type": "Rarity",
                                "value": rarity
                            }
                        ],
                        "image": image,
                        "animation_url": animation_url
                    }
                    
                    # Write JSON file
                    output_file = output_dir / f"{gem_id}.json"
                    with open(output_file, 'w', encoding='utf-8') as json_file:
                        json.dump(json_data, json_file, indent=4, ensure_ascii=False)
                    
                    processed_count += 1
                    
                    if processed_count % 50 == 0:
                        print(f"  Processed {processed_count} files...")
                
                except KeyError as e:
                    print(f"⚠ Skipping row: Missing column '{e}'")
                    skipped_count += 1
                    continue
                except Exception as e:
                    print(f"⚠ Error processing row {row.get('ID', 'unknown')}: {str(e)}")
                    skipped_count += 1
                    continue
            
            print(f"\n✓ Successfully generated {processed_count} JSON files")
            if skipped_count > 0:
                print(f"⚠ Skipped {skipped_count} rows")
            print(f"  Output directory: {output_dir}")
            
            return True
            
    except Exception as e:
        print(f"✗ Error reading CSV file: {str(e)}")
        return False


def sync_google_sheet(
    spreadsheet_id,
    gid,
    csv_file,
    json_output_dir,
    skip_download=False
):
    """
    Sync Google Sheets data: download CSV and generate JSON files
    
    Args:
        spreadsheet_id: Google Sheets spreadsheet ID
        gid: Google Sheets tab GID
        csv_file: Path to save/read CSV file
        json_output_dir: Directory to save JSON files
        skip_download: If True, skip downloading and use existing CSV
    
    Returns:
        True if successful, False otherwise
    """
    print("=" * 60)
    print("🔄 Syncing Google Sheets Data")
    print("=" * 60)
    print()
    
    # Step 1: Download CSV (unless skipped)
    if not skip_download:
        success = download_csv(spreadsheet_id, gid, csv_file)
        if not success:
            print("\n✗ Failed to download CSV. Aborting.")
            return False
    else:
        csv_path = Path(csv_file)
        if not csv_path.exists():
            print(f"✗ CSV file not found: {csv_file}")
            print("  Cannot skip download when CSV file doesn't exist.")
            return False
        print(f"⏭ Skipping download, using existing CSV: {csv_file}\n")
    
    # Step 2: Convert CSV to JSON
    success = csv_to_json(csv_file, json_output_dir)
    if not success:
        print("\n✗ Failed to generate JSON files.")
        return False
    
    print()
    print("=" * 60)
    print("✅ Sync completed successfully!")
    print("=" * 60)
    
    return True


def main():
    """Main function with command line argument parsing"""
    parser = argparse.ArgumentParser(
        description="Sync Google Sheets data and generate JSON metadata files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sync_google_sheet.py
  python sync_google_sheet.py --skip-download
  python sync_google_sheet.py --spreadsheet-id <ID> --gid <GID>
  python sync_google_sheet.py --csv ./custom.csv --output ./custom_json
        """
    )
    
    parser.add_argument(
        '--spreadsheet-id',
        default=DEFAULT_SPREADSHEET_ID,
        help=f'Google Sheets spreadsheet ID (default: {DEFAULT_SPREADSHEET_ID})'
    )
    
    parser.add_argument(
        '--gid',
        default=DEFAULT_GID,
        help=f'Google Sheets tab GID (default: {DEFAULT_GID})'
    )
    
    parser.add_argument(
        '--csv',
        default=str(DEFAULT_CSV_FILE),
        help=f'CSV file path (default: {DEFAULT_CSV_FILE})'
    )
    
    parser.add_argument(
        '--output',
        default=str(DEFAULT_OUTPUT_DIR),
        help=f'Output directory for JSON files (default: {DEFAULT_OUTPUT_DIR})'
    )
    
    parser.add_argument(
        '--skip-download',
        action='store_true',
        help='Skip downloading CSV and use existing file'
    )
    
    args = parser.parse_args()
    
    success = sync_google_sheet(
        spreadsheet_id=args.spreadsheet_id,
        gid=args.gid,
        csv_file=args.csv,
        json_output_dir=args.output,
        skip_download=args.skip_download
    )
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

