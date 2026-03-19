#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET Taxonomy Dictionary Auto-Update Script

This script automatically downloads the latest EDINET taxonomy from FSA
and regenerates the edinet_taxonomy_dict.py file.

Features:
    - Automatic remote update detection using ETag/Last-Modified headers
    - Efficient HEAD request to check for updates without downloading
    - Conditional GET (If-None-Match/If-Modified-Since) for 304 optimization
    - SHA256 hash verification for downloaded files
    - Metadata tracking to avoid unnecessary downloads

Update Detection Strategy (Hybrid Approach):
    1. HEAD request to get remote ETag/Last-Modified
    2. Compare with local metadata (.edinet_taxonomy.meta)
    3. Download only if remote has changed
    4. Save old file hash BEFORE download
    5. Compare old hash vs new hash AFTER download
    6. Regenerate dictionary only if hash changed

Usage:
    python3 update_edinet_taxonomy.py [--force]

Options:
    --force    Force update even if the file hasn't changed
               (skips remote check and hash verification)

Exit Codes:
    0 - Success (updated or no update needed)
    1 - Failure (download error, generation error, etc.)
"""

import os
import sys
import urllib.request
import hashlib
import openpyxl
import json
from datetime import datetime

# EDINET Taxonomy URL
EDINET_TAXONOMY_URL = "https://disclosure2dl.edinet-fsa.go.jp/guide/static/disclosure/download/ESE140115.xlsx"
TAXONOMY_FILE = "edinet_taxonomy_elements.xlsx"
OUTPUT_FILE = "edinet_taxonomy_dict.py"
HASH_FILE = ".edinet_taxonomy.hash"
METADATA_FILE = ".edinet_taxonomy.meta"  # Stores ETag and Last-Modified

def check_remote_update():
    """
    Check if remote file has been updated using ETag or Last-Modified headers.
    Returns:
        tuple: (needs_update: bool, metadata: dict)
    """
    print("Checking for remote updates...")

    try:
        # Send HEAD request to get metadata without downloading
        req = urllib.request.Request(EDINET_TAXONOMY_URL, method='HEAD')
        with urllib.request.urlopen(req, timeout=10) as response:
            remote_etag = response.headers.get('ETag')
            remote_last_modified = response.headers.get('Last-Modified')
            remote_content_length = response.headers.get('Content-Length')

            remote_metadata = {
                'etag': remote_etag,
                'last_modified': remote_last_modified,
                'content_length': remote_content_length,
                'checked_at': datetime.now().isoformat()
            }

            print(f"  Remote ETag: {remote_etag or 'N/A'}")
            print(f"  Remote Last-Modified: {remote_last_modified or 'N/A'}")
            print(f"  Remote Size: {remote_content_length or 'N/A'} bytes")

            # Load local metadata if exists
            if os.path.exists(METADATA_FILE):
                try:
                    with open(METADATA_FILE, 'r') as f:
                        local_metadata = json.load(f)

                    # Compare ETag (most reliable)
                    if remote_etag and local_metadata.get('etag'):
                        if remote_etag == local_metadata['etag']:
                            print("  ✓ ETag matches - no remote update")
                            return False, remote_metadata
                        else:
                            print("  ✗ ETag changed - remote update detected")
                            return True, remote_metadata

                    # Fallback to Last-Modified
                    if remote_last_modified and local_metadata.get('last_modified'):
                        if remote_last_modified == local_metadata['last_modified']:
                            print("  ✓ Last-Modified matches - no remote update")
                            return False, remote_metadata
                        else:
                            print("  ✗ Last-Modified changed - remote update detected")
                            return True, remote_metadata

                    # Fallback to Content-Length
                    if remote_content_length and local_metadata.get('content_length'):
                        if remote_content_length != local_metadata['content_length']:
                            print("  ✗ File size changed - remote update detected")
                            return True, remote_metadata
                        else:
                            print("  ⚠ No ETag/Last-Modified, but size unchanged - assuming no update")
                            return False, remote_metadata

                except json.JSONDecodeError:
                    print("  ⚠ Local metadata corrupted, assuming update needed")
                    return True, remote_metadata
            else:
                print("  ⚠ No local metadata found - first run or forced update")
                return True, remote_metadata

            # If we reach here, metadata exists but couldn't determine - assume update needed
            print("  ⚠ Could not determine update status - assuming update needed")
            return True, remote_metadata

    except Exception as e:
        print(f"  ⚠ Remote check failed ({e}), will proceed with download")
        return True, {}

def download_taxonomy(use_conditional_request=False, metadata=None):
    """
    Download EDINET taxonomy file with optional conditional request.

    Args:
        use_conditional_request: If True, use If-None-Match or If-Modified-Since headers
        metadata: Previous metadata dict with ETag/Last-Modified

    Returns:
        tuple: (success: bool, was_modified: bool)
    """
    print(f"Downloading EDINET taxonomy from: {EDINET_TAXONOMY_URL}")

    try:
        req = urllib.request.Request(EDINET_TAXONOMY_URL)

        # Add conditional headers if available
        if use_conditional_request and metadata:
            if metadata.get('etag'):
                req.add_header('If-None-Match', metadata['etag'])
                print(f"  Using If-None-Match: {metadata['etag']}")
            elif metadata.get('last_modified'):
                req.add_header('If-Modified-Since', metadata['last_modified'])
                print(f"  Using If-Modified-Since: {metadata['last_modified']}")

        try:
            with urllib.request.urlopen(req, timeout=30) as response:
                # Download the file
                with open(TAXONOMY_FILE, 'wb') as f:
                    f.write(response.read())

                print(f"✓ Downloaded: {TAXONOMY_FILE}")

                # Save new metadata
                new_metadata = {
                    'etag': response.headers.get('ETag'),
                    'last_modified': response.headers.get('Last-Modified'),
                    'content_length': response.headers.get('Content-Length'),
                    'downloaded_at': datetime.now().isoformat()
                }
                save_metadata(new_metadata)

                return True, True

        except urllib.error.HTTPError as e:
            if e.code == 304:
                # 304 Not Modified - no need to download
                print("✓ Remote file unchanged (304 Not Modified)")
                return True, False
            else:
                raise

    except Exception as e:
        print(f"✗ Download failed: {e}")
        return False, False

def save_metadata(metadata):
    """Save remote file metadata (ETag, Last-Modified, etc.)"""
    try:
        with open(METADATA_FILE, 'w') as f:
            json.dump(metadata, f, indent=2)
        print(f"✓ Saved remote metadata")
    except Exception as e:
        print(f"⚠ Could not save metadata: {e}")

def load_metadata():
    """Load saved remote file metadata"""
    if os.path.exists(METADATA_FILE):
        try:
            with open(METADATA_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            return None
    return None

def calculate_file_hash(filename):
    """Calculate SHA256 hash of file"""
    if not os.path.exists(filename):
        return None

    sha256_hash = hashlib.sha256()
    with open(filename, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def check_if_file_changed_after_download(old_hash):
    """
    Check if downloaded file differs from the previous version.

    Args:
        old_hash: Hash of the file before download (or None if no previous file)

    Returns:
        bool: True if file has changed (or is new), False if unchanged
    """
    if not os.path.exists(TAXONOMY_FILE):
        print("✗ Downloaded file not found (unexpected error)")
        return False

    new_hash = calculate_file_hash(TAXONOMY_FILE)

    if old_hash is None:
        print("✓ New file downloaded (no previous version)")
        return True

    if new_hash == old_hash:
        print("✓ Downloaded file is identical to previous version (hash match)")
        return False

    print(f"✓ File content changed")
    print(f"  Old hash: {old_hash[:16]}...")
    print(f"  New hash: {new_hash[:16]}...")
    return True

def get_current_file_hash():
    """Get hash of current taxonomy file (before download)"""
    if not os.path.exists(TAXONOMY_FILE):
        return None
    return calculate_file_hash(TAXONOMY_FILE)

def save_hash():
    """Save current hash of taxonomy file"""
    current_hash = calculate_file_hash(TAXONOMY_FILE)
    with open(HASH_FILE, 'w') as f:
        f.write(current_hash)
    print(f"✓ Saved taxonomy hash: {current_hash[:16]}...")

def get_column_index_map(ws, header_row=2):
    """
    Create column name to index mapping from header row.
    This makes the code resilient to column reordering.

    Args:
        ws: openpyxl worksheet
        header_row: Row number containing headers (1-indexed)

    Returns:
        dict: {column_name: index}
    """
    headers = [cell.value for cell in ws[header_row]]
    idx_map = {name: i for i, name in enumerate(headers) if name}
    return idx_map

def generate_dictionary():
    """Generate dictionary from EDINET taxonomy (all industry sheets)"""
    print(f"Generating dictionary from: {TAXONOMY_FILE}")

    try:
        wb = openpyxl.load_workbook(TAXONOMY_FILE, data_only=True)

        # Skip metadata sheets (not taxonomy data)
        SKIP_SHEETS = ['目次', '勘定科目リストについて']

        # Extract from ALL industry sheets (not just '一般商工業')
        # This ensures we capture industry-specific elements (banking, insurance, etc.)
        edinet_dict = {}
        sheets_processed = []
        namespace_stats = {}  # Track namespace usage for transparency

        for sheet_name in wb.sheetnames:
            if sheet_name in SKIP_SHEETS:
                continue

            ws = wb[sheet_name]
            sheet_count = 0

            # Build column index map from header (row 2)
            # This makes code resilient to column reordering
            idx_map = get_column_index_map(ws, header_row=2)

            # Validate required columns exist
            required_columns = ['要素名', '名前空間プレフィックス', '標準ラベル（日本語）']
            missing_columns = [col for col in required_columns if col not in idx_map]
            if missing_columns:
                print(f"  ⚠ Skipping sheet '{sheet_name}': Missing columns {missing_columns}")
                continue

            # Namespace filtering: Use BLACKLIST instead of whitelist
            # This allows IFRS, extensions, and future taxonomies
            NAMESPACE_BLACKLIST = {
                '名前空間プレフィックス',  # Header itself (not actual data)
                None,                      # Empty namespace
                '',                        # Empty string
                # Add more here if needed (e.g., internal test namespaces)
            }

            for row in ws.iter_rows(min_row=3, values_only=True):
                # Use header-based indexing instead of hard-coded positions
                element_name = row[idx_map['要素名']]
                namespace = row[idx_map['名前空間プレフィックス']]
                jp_label = row[idx_map['標準ラベル（日本語）']]

                # Apply blacklist filter (more permissive than whitelist)
                # This captures: jppfs_cor, jpigp_cor, ifrs_full, jpcrp_cor, extensions, etc.
                if not element_name or not jp_label or namespace in NAMESPACE_BLACKLIST:
                    continue

                # Skip if already exists (first occurrence wins - usually from '一般商工業')
                if element_name in edinet_dict:
                    continue

                label = str(jp_label)
                # Shorten labels with "or loss" notation
                if '又は' in label and ('損失' in label or '損' in label) and '（△）' in label:
                    parts = label.split('又は')
                    if len(parts) == 2:
                        label = parts[0].strip()
                edinet_dict[element_name] = label
                sheet_count += 1

                # Track namespace usage
                namespace_stats[namespace] = namespace_stats.get(namespace, 0) + 1

            if sheet_count > 0:
                sheets_processed.append(f"{sheet_name}({sheet_count})")

        print(f"✓ Extracted {len(edinet_dict)} items from EDINET taxonomy")
        print(f"  Processed {len(sheets_processed)} sheets:")
        # Show detailed sheet statistics
        for i, sheet_info in enumerate(sheets_processed):
            if i < 10:  # Show first 10 sheets
                print(f"    - {sheet_info}")
        if len(sheets_processed) > 10:
            print(f"    ... and {len(sheets_processed) - 10} more sheets")

        # Show namespace statistics (for transparency)
        if namespace_stats:
            print(f"  Namespaces found:")
            for ns, count in sorted(namespace_stats.items(), key=lambda x: x[1], reverse=True):
                print(f"    - {ns}: {count} elements")

        # Custom mappings (IFRS variants, abbreviations, etc.)
        custom_mappings = {
            # Generic abbreviations
            'Notes': '注記',
            'Inventory': '棚卸資産',

            # Financial statement names
            'ConsolidatedBalanceSheet': '連結貸借対照表',
            'ConsolidatedStatementOfIncome': '連結損益計算書',
            'ConsolidatedStatementOfCashFlows': '連結キャッシュ・フロー計算書',
            'ConsolidatedStatementOfChangesInEquity': '連結株主資本等変動計算書',
            'ConsolidatedStatementOfFinancialPosition': '連結財政状態計算書',
            'ConsolidatedStatementOfProfitOrLoss': '連結損益計算書',
            'ConsolidatedStatementOfFinancialPositionIFRS': '連結財政状態計算書',
            'ConsolidatedStatementOfProfitOrLossIFRS': '連結損益計算書',
            'ConsolidatedStatementOfCashFlowsIFRS': '連結キャッシュ・フロー計算書',
            'ConsolidatedStatementOfChangesInEquityIFRS': '連結株主資本等変動計算書',
            'ConsolidatedStatementOfComprehensiveIncomeIFRS': '連結包括利益計算書',

            # IFRS variants / Abbreviations
            'NetSalesIFRS': '売上収益',
            'RevenueIFRS': '売上収益',
            'CostOfSalesIFRS': '売上原価',
            'GrossProfitIFRS': '売上総利益',
            'SellingGeneralAndAdministrativeExpensesIFRS': '販売費及び一般管理費',
            'OtherOperatingIncome': 'その他の営業収益',
            'OtherOperatingExpenses': 'その他の営業費用',
            'OtherOperatingExpense': 'その他の営業費用',
            'OtherIncomeIFRS': 'その他の収益',
            'OtherExpensesIFRS': 'その他の費用',
            'OtherOperatingIncomeIFRS': 'その他の営業収益',
            'OtherOperatingExpensesIFRS': 'その他の営業費用',
            'ShareOfProfitLossOfInvestmentsAccountedForUsingEquityMethodIFRS': '持分法による投資利益',
            'OperatingProfitLossIFRS': '営業利益',
            'FinanceIncomeIFRS': '金融収益',
            'FinanceCostsIFRS': '金融費用',
            'ProfitLossBeforeTaxIFRS': '税引前当期利益',
            'IncomeTaxExpenseIFRS': '法人所得税費用',
            'ProfitLossIFRS': '当期利益',
            'ProfitLossAttributableToOwnersOfParentIFRS': '親会社の所有者に帰属する当期利益',
            'ProfitLossAttributableToNonControllingInterestsIFRS': '非支配持分',
            'BasicEarningsPerShareIFRS': '基本的１株当たり当期利益（円）',

            # J-GAAP variants / Abbreviations
            'OperatingProfit': '営業利益',
            'FinanceIncome': '金融収益',
            'FinancialIncome': '金融収益',
            'FinanceCosts': '金融費用',
            'FinancialExpenses': '金融費用',
            'FinanceExpenses': '金融費用',
            'ProfitBeforeTax': '税引前利益',
            'IncomeTaxExpense': '法人所得税費用',
            'Profit': '当期利益',
            'NetIncome': '当期利益',
            'ProfitLossAttributableToAbstract': '当期利益の帰属',
            'ProfitAttributableToOwnersOfParent': '親会社株主に帰属する当期純利益',
            'ProfitAttributableToNoncontrollingInterests': '非支配持分',
            'ProfitLossAttributableToNoncontrollingInterests': '非支配持分',
            'BasicEarningsPerShare': '基本的１株当たり当期純利益（円）',
            'BasicEarningsLossPerShare': '基本的１株当たり当期純利益（円）',
            'SellingGeneralAndAdministrativeExpense': '販売費及び一般管理費',
            'ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod': '持分法による投資利益',
        }

        # Merge dictionaries
        final_dict = {}
        final_dict.update(edinet_dict)
        final_dict.update(custom_mappings)

        print(f"✓ Total items: {len(final_dict)} (EDINET: {len(edinet_dict)}, Custom: {len(custom_mappings)})")

        return final_dict, len(edinet_dict), len(custom_mappings)

    except Exception as e:
        print(f"✗ Dictionary generation failed: {e}")
        import traceback
        traceback.print_exc()
        return None, 0, 0

def write_dictionary_file(final_dict, edinet_count, custom_count):
    """Write dictionary to Python file"""
    print(f"Writing dictionary to: {OUTPUT_FILE}")

    try:
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write('#!/usr/bin/env python3\n')
            f.write('# -*- coding: utf-8 -*-\n')
            f.write('"""\n')
            f.write('EDINET Taxonomy Dictionary\n')
            f.write('\n')
            f.write('Auto-generated from EDINET Official Taxonomy:\n')
            f.write(f'{EDINET_TAXONOMY_URL}\n')
            f.write('\n')
            f.write(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write(f'Total: {len(final_dict):,} items\n')
            f.write(f'- EDINET Official Taxonomy: {edinet_count:,} items\n')
            f.write(f'- Custom Mappings (IFRS variants, etc.): {custom_count} items\n')
            f.write('"""\n')
            f.write('\n')
            f.write('# EDINET Taxonomy common dictionary\n')
            f.write('# Maps element names to Japanese account labels\n')
            f.write('common_dict = {\n')

            # Sort and write entries
            for key in sorted(final_dict.keys()):
                value = final_dict[key]
                # Escape single quotes
                value_escaped = value.replace("'", "\\'")
                f.write(f"    '{key}': '{value_escaped}',\n")

            f.write('}\n')

        print(f"✓ Dictionary written to: {OUTPUT_FILE}")
        return True

    except Exception as e:
        print(f"✗ File write failed: {e}")
        return False

def main():
    """Main function"""
    print("=" * 80)
    print("EDINET Taxonomy Dictionary Auto-Update")
    print("=" * 80)
    print()

    force = '--force' in sys.argv

    # Step 1: Check for remote updates (unless file doesn't exist)
    needs_download = False
    remote_metadata = None

    if not os.path.exists(TAXONOMY_FILE):
        print("Local taxonomy file not found - download required")
        needs_download = True
    elif force:
        print("Force update requested - skipping remote check")
        needs_download = True
    else:
        # Check if remote file has been updated
        needs_remote_update, remote_metadata = check_remote_update()
        if needs_remote_update:
            print("Remote update detected - download required")
            needs_download = True
        else:
            print("Remote file unchanged - no download needed")

    # Step 2: Download if needed
    file_changed = False
    if needs_download:
        # CRITICAL: Save hash of old file BEFORE download
        old_hash = get_current_file_hash()
        print(f"Old file hash: {old_hash[:16] + '...' if old_hash else 'N/A (no previous file)'}")

        # Load existing metadata for conditional request (304 optimization)
        existing_metadata = load_metadata()
        success, was_modified = download_taxonomy(
            use_conditional_request=not force,
            metadata=existing_metadata
        )

        if not success:
            print("\n✗ Update failed: Could not download taxonomy file")
            return 1

        if not was_modified and not force:
            print("\n✓ No update needed (304 Not Modified)")
            return 0

        # Step 3: Verify if downloaded file differs from old file (hash comparison)
        if force:
            print("\nForce mode: Skipping hash verification")
            file_changed = True  # Force regeneration
        else:
            print("\nVerifying downloaded file...")
            file_changed = check_if_file_changed_after_download(old_hash)

            if not file_changed:
                print("\n✓ No update needed (downloaded file is identical to previous version)")
                return 0

    # If we didn't download, check if we even have a file
    elif not os.path.exists(TAXONOMY_FILE):
        print("\n✗ No taxonomy file found and no download performed")
        return 1
    else:
        # No download needed, file exists, assume it's already processed
        print("\n✓ No update needed (remote unchanged, local file exists)")
        return 0

    # Step 4: Generate dictionary (only if file changed or force)
    print()
    print("Generating dictionary from updated taxonomy file...")
    final_dict, edinet_count, custom_count = generate_dictionary()
    if final_dict is None:
        print("\n✗ Update failed: Could not generate dictionary")
        return 1

    # Step 5: Write dictionary file
    if not write_dictionary_file(final_dict, edinet_count, custom_count):
        print("\n✗ Update failed: Could not write dictionary file")
        return 1

    # Step 6: Save hash of new file
    save_hash()

    print()
    print("=" * 80)
    print("✓ Dictionary update completed successfully!")
    print("=" * 80)

    return 0

if __name__ == '__main__':
    sys.exit(main())
