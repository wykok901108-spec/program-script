#!/usr/bin/env python3

import os
import re
import sys
import time
import collections
import mam_constants
import mam_mipc
from mam_report import Lot_Report
from parse_psel_blacklist import parse_psel_blacklist_file
import GeneralLib
import pandas as pd
from datetime import datetime

def to_int(x, default=0):
    try:
        return int(str(x).strip())
    except Exception:
        return default


def write_summary_to_excel(tstsi_rows, tstpg_rows):

    timestamp = datetime.now().strftime('%Y%m%d')
    #os.makedirs("excel_reports", exist_ok=True)
    fname = f"/home/weeyap/python/jira_create_issue/excel_reports/blacklist_inventory_report_{timestamp}.xlsx"
    fname_empty = f"/home/weeyap/python/jira_create_issue/excel_reports/blacklist_inventory_report_{timestamp}_empty.xlsx"
    
    if not tstsi_rows and not tstpg_rows:
        df = pd.DataFrame()
        print(f"Created empty Excel file: {fname_empty}")
        with pd.ExcelWriter(fname_empty, engine='openpyxl') as writer:
            df.to_excel(writer, index=False) 
        return
    
    with pd.ExcelWriter(fname, engine="openpyxl") as writer:
        if tstsi_rows:
            df_tstsi = pd.DataFrame(
                tstsi_rows,
                columns=[
                    'Summary Key',
                    'Lots (count)',
                    'Total Qty',
                    'Max Lot Qty',
                    'Table Title',
                    'Design ID',
                    'Reticle Wave ID',
                    'Major Probe Prog Rev',
                    'Number Of Die In Pkg',
                ],
            )
            df_tstsi.to_excel(writer, sheet_name='TSTSI', index=False)
        
        if tstpg_rows:
            df_tstpg = pd.DataFrame(
                tstpg_rows,
                columns=[
                    'Summary Key',
                    'Lots (count)',
                    'Total Qty',
                    'Max Lot Qty',
                    'Table Title',
                    'Design ID',
                    'Reticle Wave ID',
                    'Major Probe Prog Rev',
                    'Number Of Die In Pkg',
                ],
            )
            df_tstpg.to_excel(writer, sheet_name='TSTPG', index=False)

    print(f"Excel file with separate sheets written: {fname}")


def main():
    tstsi_tables_rows = []
    tstpg_tables_rows = []

    search_criteria_array = parse_psel_blacklist_file.return_search_item_list_req_attn()

    # whitelist
    with open('/home/tianyifeng/DATA/PSEL/blacklist_inventory_monitoring/lot_result_whitelist', 'r') as f:
        lot_result_whitelist_file_content = f.read().splitlines()

    for search_criteria_item in search_criteria_array:
        title = search_criteria_item['title']
        if ("CELL" in title and "REV" in title) or "PSPT" in title or "SSD-REBALL" in title:
            continue

        search_display = [
            'CURRENT QTY', 
            'LOT LOCATION', 
            'INVENTORY LOCATION', 
            'RETICLE WAVE ID',
            'MAJOR PROBE PROG REV', 
            'DESIGN ID', 
            'NUMBER OF DIE IN PKG', 
            'HOLD LOT',
            'SPECTEK SOURCE', 
            'CELL REVISION', 
            'CMOS REVISION', 
            'SPTK TST CONTAINMENT', 
            'LEAD COUNT'
        ]

        # Process both reports
        for mam_instance, rows_collection in [('TSTSI', tstsi_tables_rows), ('TSTPG', tstpg_tables_rows)]:
            mam_query_fail = True
            query_time = 1
            print(f"Current querying for {search_criteria_item['title']} on {mam_instance} with search criteria as:")
            print(search_criteria_item['criteria'])
            
            while mam_query_fail:
                try:
                    report = Lot_Report(mam_instance, report_title='Lot Inventory', criteria=search_criteria_item['criteria'], display=search_display).send_via()
                    mam_query_fail = False
                except:
                    time.sleep(10)  #Wait 10s if failed last mam query
                    query_time += 1
                    if (query_time > 5):
                        print(title)
                        query_time += 1
                        break
            
            if (query_time > 6):
                continue

            report_result = report._unwrap()
            for report_result_lot in report_result:
                for attrs_list in report_result_lot.keys():
                    if report_result_lot[attrs_list] is None:
                        report_result_lot[attrs_list] = 'N/A'


            ordered_report_result = sorted(
                report_result,
                key=lambda k: (
                    k['DESIGN ID'], 
                    k['LEAD COUNT'], 
                    k['NUMBER OF DIE IN PKG'],
                    k['SPECTEK SOURCE'], 
                    k['RETICLE WAVE ID'], 
                    k['MAJOR PROBE PROG REV']
                )
            )

            summary_report_list = {}
            for lot in ordered_report_result:
                if lot['RETICLE WAVE ID'] in ['N/A', 'MIXED'] or lot['MAJOR PROBE PROG REV'] in ['N/A', 'MIXED']:
                    continue


                ndp = to_int(lot['NUMBER OF DIE IN PKG'])
                if ndp == 1:
                    num_of_die_in_pkg = "SDP"
                elif ndp == 2:
                    num_of_die_in_pkg = "DDP"
                elif ndp == 4:
                    num_of_die_in_pkg = "QDP"
                elif ndp == 8:
                    num_of_die_in_pkg = "ODP"
                else:
                    num_of_die_in_pkg = f"{lot['NUMBER OF DIE IN PKG']}DP"

                lot_summary_entry = (
                    f"{lot['DESIGN ID']}_{lot['LEAD COUNT']}_{num_of_die_in_pkg}_"
                    f"{lot['SPECTEK SOURCE']}_{lot['RETICLE WAVE ID']}_{lot['MAJOR PROBE PROG REV']}"
                )

                if lot_summary_entry in lot_result_whitelist_file_content:
                    continue

                qty = to_int(lot['CURRENT QTY'])

                if lot_summary_entry not in summary_report_list:
                    summary_report_list[lot_summary_entry] = [1, qty, qty]  # [count, total, max]
                else:
                    summary_report_list[lot_summary_entry][0] += 1
                    summary_report_list[lot_summary_entry][1] += qty
                    if qty > summary_report_list[lot_summary_entry][2]:
                        summary_report_list[lot_summary_entry][2] = qty

            if not len(summary_report_list):
                continue

            table_title = search_criteria_item['title']
            for lot_summary_entry, vals in summary_report_list.items():
                lots_count, total_qty, max_lot_qty = vals
                if total_qty > 5000:  # preserve your threshold
                    # Parse fields from Summary Key (B68S_154/195_QDP_ASSEMBLY_WAVE010_69)
                    parts = lot_summary_entry.split('_', 5)
                    design_id = parts[0] if len(parts) > 0 else ''
                    num_of_die_in_pkg_summary= parts[2] if len(parts)> 0 else''
                    reticle_wave = parts[4] if len(parts) > 4 else ''
                    major_probe_prog_rev = parts[5] if len(parts) > 5 else ''

                    rows_collection.append({
                        'Summary Key': lot_summary_entry,
                        'Lots (count)': lots_count,
                        'Total Qty': total_qty,
                        'Max Lot Qty': max_lot_qty,
                        'Table Title': table_title,
                        'Design ID': design_id,
                        'Reticle Wave ID': reticle_wave,
                        'Major Probe Prog Rev': major_probe_prog_rev,
                        'Number Of Die In Pkg': num_of_die_in_pkg_summary
                    })

    # write both datasets to different sheets
    write_summary_to_excel(tstsi_tables_rows, tstpg_tables_rows)


if __name__ == "__main__":
    main()
