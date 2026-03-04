import os
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import GeneralLib 
import time
from datetime import datetime

# --------------------------
# HTTP Session
# --------------------------
session = requests.Session()
session.auth = HTTPBasicAuth("weeyap@micron.com", "ATATT3xFfGF0kHi-lHGOXNLSShgiSkexu2wiG610kWpPqADZw0bGVHAq3ye6TaxxZ1w-NLzCHHaM_hHIkPfsQhct0D64CyQUmDB_ufIZxmfaeMDMqiGKOWyuSDolbwjuLjoSnr_keDGU4eZK3jPQOVYA4KSnY8tPvqTKE1nenv1RJgUkkTRmUac=5F189EFB")
session.headers.update({
    "Accept": "application/json",
    "Content-Type": "application/json"
})
#session.verify = r'C:\ProgramData\pip\micronCAchain.pem'

def jira_issue_exists(summary_key: str) -> bool:
 
    jql = f'project = "SPTKNAND" AND cf[10264]~ "{summary_key}"'
    url = "https://micron.atlassian.net/rest/api/3/search/jql"

    payload = {
        "jql": jql,
        "maxResults": 1
    }

    resp = session.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    print(f"Search result for {summary_key}: {data}")
    
    # Check if any issues were found
    issues = data.get("issues", [])
    print(issues)
    print(len(issues))
    return len(issues) > 0

def create_jira_issue(summary_key: str,
                      design_id: str,
                      reticle_wave_id: str,
                      major_probe_prog_rev: str,
                      number_of_die_in_pkg:str) -> dict:
    url = "https://micron.atlassian.net/rest/api/3/issue"

    payload = {
        "fields": 
        {
            "project":   {"key": "SPTKNAND"},
            "issuetype": {"id": "10120"},
            "summary":   f"{design_id} {reticle_wave_id} MPPR{major_probe_prog_rev} {number_of_die_in_pkg} New Material",
            "customfield_10264": summary_key,
            "components": [{"name": design_id}, {"name": "Qual"}],
            #"assignee":  {"accountId": "712020:aed26921-96d3-4061-9672-5f9fa8aa9db1"},
            #"priority":  {"id": "3"}
        }
    }
        
    resp = session.post(url, json=payload, headers=session.headers, auth=session.auth, timeout=30)
    if resp.status_code >= 300:
        raise requests.HTTPError(f"JIRA create failed: {resp.status_code} {resp.text}", response=resp)
    return resp.json()


def notify_issue_created(email_to: str,
                         issue_key: str,
                         summary_key: str,
                         design_id: str,
                         reticle_wave_id: str,
                         major_probe_prog_rev: str,
                         number_of_die_in_pkg: str,
                         sheet_name: str):

    site=sheet_name
    
    if sheet_name == "TSTSI":
        try:
            timestamp = datetime.now().strftime('%Y%m%d')
            excel_file = f'z:/python/jira_create_issue/excel_reports/blacklist_inventory_report_{timestamp}.xlsx' 
            df_tstpg = pd.read_excel(excel_file, sheet_name='TSTPG')
            
            # Check if summary_key exists in TSTPG
            if summary_key in df_tstpg['Summary Key'].values:
                print(f"DUPLICATE FOUND: {summary_key} exists in both MSB (TSTSI) and MMP (TSTPG)")
                site="TSTSI and TSTPG"
                
        except Exception as e:
            print(f"Error checking for duplicates: {e}")

    try:
        subject = f"[AUTO] {issue_key} for {design_id} {reticle_wave_id} MPPR{major_probe_prog_rev} {number_of_die_in_pkg} New Material"
        
        html_body = f"""
            <html>
            <body style="font-family: Calibri, Arial, sans-serif; font-size: 16px;">

            <p>Hi MSB Engineers,</p>

            <p>A new Jira issue has been created.</p>

            <table style="border-collapse: collapse;">
            <tr><td><b>Summary Title:</b></td><td>{design_id} {reticle_wave_id} MPPR{major_probe_prog_rev} {number_of_die_in_pkg} New Material</td></tr>
            <tr><td><b>Issue Key:</b></td><td>{issue_key}</td></tr>
            <tr><td><b>Design ID:</b></td><td>{design_id}</td></tr>
            <tr><td><b>Reticle Wave:</b></td><td>{reticle_wave_id}</td></tr>
            <tr><td><b>MPPR:</b></td><td>{major_probe_prog_rev}</td></tr>
            <tr><td><b>Number of Die in Package:</b></td><td>{number_of_die_in_pkg}</td></tr>
            <tr><td><b>Summary Key:</b></td><td>{summary_key}</td></tr>
            <tr><td><b>Site:</b></td><td>{site}</td></tr>
            </table>

            <br>

            <p>
            <b>Direct link:</b> 
            <a href="https://micron.atlassian.net/browse/{issue_key}">
            https://micron.atlassian.net/browse/{issue_key}
            </a>
            </p>

            <br>

            <p><strong>Thanks,<br>
            WeeYap</strong></p>

            </body>
            </html>
            """
        
        if "TSTPG" in site:
            recipients = ["SPECTEK_MSB_TESTENG@micron.com","amohamadtamb@micron.com"]
            GeneralLib.SendMail(
                From="weeyap@micron.com",
                To=recipients,
                Subject=subject,
                Msg=html_body,
            )
            print(f"[MAIL] Notification sent to {recipients} for {issue_key}")
        else:
            GeneralLib.SendMail(
                From="weeyap@micron.com",
                To=[email_to],
                Subject=subject,
                Msg=html_body,
            )
            print(f"[MAIL] Notification sent to {email_to} for {issue_key}")
            
    except Exception as e:
        print(f"[MAIL] Failed to send notification for {issue_key}: {e}")

def process_records(records, sheet_name):
    print(f"\n--- Processing {sheet_name} sheet ---")
    
    for record in records:
        summary_key = record["Summary Key"]
        design_id = record["Design ID"]
        reticle_wave_id = record["Reticle Wave ID"]
        major_probe_prog_rev = record["Major Probe Prog Rev"]
        number_of_die_in_pkg = record["Number Of Die In Pkg"]
        #print(record)
        try:
            if jira_issue_exists(summary_key):
                print(f"SKIP — Jira already created before for summary_key {summary_key} ({sheet_name})")
            else:
                issue = create_jira_issue(summary_key, design_id, reticle_wave_id, major_probe_prog_rev,number_of_die_in_pkg)
                
                issue_key = issue.get('key')
                notify_issue_created(
                    email_to="SPECTEK_MSB_TESTENG@micron.com",
                    issue_key=issue_key,
                    summary_key=summary_key,          
                    design_id=design_id,
                    reticle_wave_id=reticle_wave_id,
                    major_probe_prog_rev=major_probe_prog_rev,
                    number_of_die_in_pkg=number_of_die_in_pkg,
                    sheet_name=sheet_name,
                )

                print(f"CREATED — New Jira for summary_key {summary_key} ({sheet_name})")
                
        except requests.HTTPError as e:
            print(f"HTTP error when checking JIRA for {summary_key} ({sheet_name}): {e} | {getattr(e.response, 'text', '')[:300]}")
        except Exception as e:
            print(f"Error when checking JIRA for {summary_key} ({sheet_name}): {e}")

# --------------------------
# Main
# --------------------------
def main():
    timestamp = datetime.now().strftime('%Y%m%d')
    excel_file = f'z:/python/jira_create_issue/excel_reports/blacklist_inventory_report_{timestamp}.xlsx' 
    
    try:
        df_tstsi = pd.read_excel(excel_file, sheet_name='TSTSI')
        records_si = df_tstsi.to_dict(orient="records")  
        process_records(records_si, "TSTSI")
    except Exception as e:
        print(f"Error processing TSTSI sheet: {e}")
    
    try:
        df_tstpg = pd.read_excel(excel_file, sheet_name='TSTPG')
        records_pg = df_tstpg.to_dict(orient="records")
        process_records(records_pg, "TSTPG")
    except Exception as e:
        print(f"Error processing TSTPG sheet: {e}")

if __name__ == "__main__":
    main()