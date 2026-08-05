# ============================================================
#         EMAIL SENDER - GITHUB ACTIONS VERSION
# ============================================================
# What this script does:
#   1. Reads state file from process_invoice.py
#   2. Determines which email to send (Email 1 or Email 2)
#   3. Builds professional HTML email body
#   4. Attaches CRD Report and/or Not Found Summary
#   5. Sends email via Gmail SMTP
#   6. Updates state file after sending
#
# Email 1 (WITH Not Found):
#   • Attachments : CRD Report + Not Found Summary
#   • Body        : Mentions not found entities exist
#
# Email 1 (WITHOUT Not Found):
#   • Attachments : CRD Report only
#   • Body        : Standard body
#
# Email 2 (Final after correction):
#   • Attachments : Final CRD Report only
#   • Body        : Final report body
# ============================================================

import os
import json
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email                import encoders
from datetime             import datetime


# ============================================================
#                    PATH CONFIGURATION
# ============================================================

def get_repo_root():
    """Returns the root directory of the repository."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def read_state_file():
    """Reads state file and resolves all paths."""
    repo_root  = get_repo_root()
    state_path = os.path.join(
        repo_root, "reports", ".run_state.json"
    )

    if not os.path.exists(state_path):
        raise FileNotFoundError(
            f"❌ State file not found: {state_path}"
        )

    with open(state_path, "r") as f:
        state = json.load(f)

    # Resolve all path fields to absolute paths
    path_keys = [
        "output_dir",
        "not_found_summary_path",
        "not_found_json_path",
        "combined_file_path",
        "sample_invoice_path",
        "final_report_path",
        "final_sample_invoice_path"
    ]

    for key in path_keys:
        if key in state and state[key]:
            state[key] = resolve_path(state[key])

    return state, state_path


def update_state_file(state_path, updates):
    """Updates specific keys in the state JSON file."""
    with open(state_path, "r") as f:
        state = json.load(f)

    state.update(updates)

    with open(state_path, "w") as f:
        json.dump(state, f, indent=4)

    return state


# ============================================================
#                    EMAIL CREDENTIALS
# ============================================================

def get_email_config():
    """
    Reads email configuration from environment variables.
    All values come from GitHub Secrets.

    Required Secrets:
        GMAIL_SENDER        → sender@gmail.com
        GMAIL_APP_PASSWORD  → Gmail App Password
        RECIPIENT_EMAILS    → email1@x.com,email2@x.com
        CC_EMAILS           → cc1@x.com,cc2@x.com
    """
    sender    = os.environ.get("GMAIL_SENDER", "").strip()
    password  = os.environ.get("GMAIL_APP_PASSWORD", "").strip()
    recipients_raw = os.environ.get("RECIPIENT_EMAILS", "").strip()
    cc_raw         = os.environ.get("CC_EMAILS", "").strip()

    if not sender:
        raise ValueError("❌ GMAIL_SENDER secret is not set.")
    if not password:
        raise ValueError("❌ GMAIL_APP_PASSWORD secret is not set.")
    if not recipients_raw:
        raise ValueError("❌ RECIPIENT_EMAILS secret is not set.")

    recipients = [
        e.strip() for e in recipients_raw.split(",")
        if e.strip()
    ]
    cc_list = [
        e.strip() for e in cc_raw.split(",")
        if e.strip()
    ] if cc_raw else []

    return sender, password, recipients, cc_list


# ============================================================
#                    EMAIL BODY BUILDERS
# ============================================================

def format_month_range_display(start_month, end_month):
    """
    Convert month strings to display format.

    Example:
        'APR-2025' , 'JUL-2026'
        → 'APR-25 to JUL-26'
    """
    def shorten(m):
        try:
            dt = datetime.strptime(m, "%b-%Y")
            return dt.strftime("%b-%y").upper()
        except Exception:
            return m.upper()

    return f"{shorten(start_month)} to {shorten(end_month)}"


def build_email1_with_not_found_body(
    start_month, end_month,
    not_found_count, run_date_str
):
    """
    Email 1 body when Not Found entities exist.
    Concise body with professional Action Required highlight.
    """
    month_range = format_month_range_display(
        start_month, end_month
    )

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {{
      margin: 0; padding: 0;
      background-color: #e5e5e5;
      font-family: Segoe UI, Arial, sans-serif;
    }}
    .container {{ width: 875px; margin: 0 auto; }}
  </style>
</head>
<body>
  <div class="container">
    <center>
      <table width="100%" bgcolor="#e5e5e5"
             cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td align="center" style="padding:30px;">
            <table width="100%" cellpadding="0"
                   cellspacing="0" border="0"
                   style="background-color:#ffffff;
                          border-top:6px solid #00085D;
                          border-radius:6px;">

              <!-- ── HEADER ── -->
              <tr>
                <td style="padding:25px 30px 15px 30px;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0">
                    <tr>
                      <td style="font-size:16px;
                                 color:#333333;
                                 font-weight:bold;
                                 padding-bottom:5px;">
                        CRD Client-Wise Booking &amp;
                        Revenue Report: {month_range}
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

              <!-- ── GREETING ── -->
              <tr>
                <td style="padding:0 30px 5px 30px;
                           font-size:14px;
                           color:#333333;">
                  <strong>Dear Sir / Ma'am,</strong>
                </td>
              </tr>

              <!-- ── BODY ── -->
              <tr>
                <td style="padding:10px 30px 15px 30px;
                           font-size:13px;
                           color:#333333;
                           line-height:1.8;">
                  Please find attached the
                  <strong>CRD Client-Wise Booking &amp;
                  Revenue Report</strong> for the period
                  <strong>{month_range}</strong>
                  for your reference.
                  The attached file contains the
                  <strong>Business Report</strong> and a
                  <strong>Not Found Summary</strong>
                  sheet.
                </td>
              </tr>
              <!-- ── ACTION REQUIRED BOX ── -->
              <tr>
                <td style="padding:0 30px 20px 30px;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0"
                         style="border-radius:6px;
                                overflow:hidden;">

                    <!-- Orange Header Bar -->
                    <tr>
                      <td style="background-color:#E67E22;
                                 padding:10px 16px;">
                        <table width="100%" cellpadding="0"
                               cellspacing="0" border="0">
                          <tr>
                            <td style="font-size:13px;
                                       font-weight:bold;
                                       color:#ffffff;
                                       letter-spacing:0.5px;">
                              ACTION REQUIRED
                            </td>
                            <td align="right"
                                style="font-size:12px;
                                       color:#FDEBD0;
                                       font-weight:bold;">
                              {not_found_count}
                              Unmapped Combination(s)
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>

                    <!-- Content -->
                    <tr>
                      <td style="background-color:#FEF9F0;
                                 padding:14px 16px;
                                 border:1px solid #F0A500;
                                 border-top:none;
                                 font-size:13px;
                                 color:#333333;
                                 line-height:1.8;">
                        <strong>{not_found_count}
                        Corporate&ndash;Hub
                        combination(s)</strong> could not
                        be mapped to a Group Entity or
                        Vertical and have been
                        <strong>excluded</strong> from the
                        Business Report.<br><br>
                        Please fill in the
                        <strong>Group Entity</strong> and
                        <strong>Vertical</strong> columns
                        in the <strong>Not Found
                        Summary</strong> sheet of the
                        attached file and share the
                        updated Not Found Summary file.<br><br>
                        The corrected report will then be
                        generated and shared accordingly.
                      </td>
                    </tr>

                  </table>
                </td>
              </tr>
              <!-- ── SIGNATURE ── -->
              <tr>
                <td style="padding:15px 30px 25px 30px;
                           background-color:#f9f9f9;
                           border-top:1px solid #dddddd;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0">
                    <tr>
                      <td style="vertical-align:top;">
                        <table cellpadding="0"
                               cellspacing="0" border="0">
                          <tr>
                            <td style="font-size:12px;
                                       color:#444444;
                                       padding-bottom:5px;">
                              <strong>
                                Thanks &amp; Regards,
                              </strong>
                            </td>
                          </tr>
                          <tr>
                            <td style="font-size:13px;
                                       color:#00085D;
                                       font-weight:bold;
                                       padding-bottom:3px;">
                              Adit Sharma
                            </td>
                          </tr>
                          <tr>
                            <td style="font-size:12px;
                                       color:#666666;">
                              Analytics Team
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td align="right"
                          style="vertical-align:top;">
                        <img src="https://dgdlm6ddvctpd.cloudfront.net/bannerimages/wtiintelligentmobility.png"
                             alt="Company Logo"
                             width="100"
                             style="display:block;
                                    border:0;">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </center>
  </div>
</body>
</html>"""

    return html
def build_email1_no_not_found_body(
    start_month, end_month, run_date_str
):
    """
    Email 1 body when ALL entities are mapped.
    Single attachment — CRD Business Report only.
    """
    month_range = format_month_range_display(
        start_month, end_month
    )

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {{
      margin: 0; padding: 0;
      background-color: #e5e5e5;
      font-family: Segoe UI, Arial, sans-serif;
    }}
    .container {{ width: 875px; margin: 0 auto; }}
  </style>
</head>
<body>
  <div class="container">
    <center>
      <table width="100%" bgcolor="#e5e5e5"
             cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td align="center" style="padding:30px;">
            <table width="100%" cellpadding="0"
                   cellspacing="0" border="0"
                   style="background-color:#ffffff;
                          border-top:6px solid #00085D;
                          border-radius:6px;">

              <!-- ── HEADER ── -->
              <tr>
                <td style="padding:25px 30px 15px 30px;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0">
                    <tr>
                      <td style="font-size:16px;
                                 color:#333333;
                                 font-weight:bold;
                                 padding-bottom:5px;">
                        CRD Client-Wise Booking &amp;
                        Revenue Report: {month_range}
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

              <!-- ── GREETING ── -->
              <tr>
                <td style="padding:0 30px 5px 30px;
                           font-size:14px;
                           color:#333333;">
                  <strong>Dear Sir / Ma'am,</strong>
                </td>
              </tr>

              <!-- ── BODY ── -->
              <tr>
                <td style="padding:10px 30px 20px 30px;
                           font-size:13px;
                           color:#333333;
                           line-height:1.8;">
                  Please find attached the
                  <strong>CRD Client-Wise Booking &amp;
                  Revenue Report</strong> for the period
                  <strong>{month_range}</strong>
                  for your reference.
                </td>
              </tr>

              <!-- ── SIGNATURE ── -->
              <tr>
                <td style="padding:15px 30px 25px 30px;
                           background-color:#f9f9f9;
                           border-top:1px solid #dddddd;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0">
                    <tr>
                      <td style="vertical-align:top;">
                        <table cellpadding="0"
                               cellspacing="0" border="0">
                          <tr>
                            <td style="font-size:12px;
                                       color:#444444;
                                       padding-bottom:5px;">
                              <strong>
                                Thanks &amp; Regards,
                              </strong>
                            </td>
                          </tr>
                          <tr>
                            <td style="font-size:13px;
                                       color:#00085D;
                                       font-weight:bold;
                                       padding-bottom:3px;">
                              Adit Sharma
                            </td>
                          </tr>
                          <tr>
                            <td style="font-size:12px;
                                       color:#666666;">
                              Analytics Team
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td align="right"
                          style="vertical-align:top;">
                        <img src="https://dgdlm6ddvctpd.cloudfront.net/bannerimages/wtiintelligentmobility.png"
                             alt="Company Logo"
                             width="100"
                             style="display:block;
                                    border:0;">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </center>
  </div>
</body>
</html>"""

    return html
def build_email2_final_body(
    start_month, end_month, run_date_str
):
    """
    Email 2 body — Final corrected report.
    """
    month_range = format_month_range_display(
        start_month, end_month
    )

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {{
      margin: 0; padding: 0;
      background-color: #e5e5e5;
      font-family: Segoe UI, Arial, sans-serif;
    }}
    .container {{ width: 875px; margin: 0 auto; }}
  </style>
</head>
<body>
  <div class="container">
    <center>
      <table width="100%" bgcolor="#e5e5e5"
             cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td align="center" style="padding:30px;">
            <table width="100%" cellpadding="0"
                   cellspacing="0" border="0"
                   style="background-color:#ffffff;
                          border-top:6px solid #00085D;
                          border-radius:6px;">

              <!-- ── HEADER ── -->
              <tr>
                <td style="padding:25px 30px 15px 30px;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0">
                    <tr>
                      <td style="font-size:16px;
                                 color:#333333;
                                 font-weight:bold;
                                 padding-bottom:5px;">
                        CRD Client-Wise Booking &amp;
                        Revenue Report: {month_range}
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

              <!-- ── GREETING ── -->
              <tr>
                <td style="padding:0 30px 5px 30px;
                           font-size:14px;
                           color:#333333;">
                  <strong>Dear Sir / Ma'am,</strong>
                </td>
              </tr>

              <!-- ── BODY ── -->
              <tr>
                <td style="padding:10px 30px 20px 30px;
                           font-size:13px;
                           color:#333333;
                           line-height:1.8;">
                  Please find attached the final corrected
                  <strong>CRD Client-Wise Booking &amp;
                  Revenue Report</strong> for the period
                  <strong>{month_range}</strong>
                  for your reference.
                </td>
              </tr>

              <!-- ── SIGNATURE ── -->
              <tr>
                <td style="padding:15px 30px 25px 30px;
                           background-color:#f9f9f9;
                           border-top:1px solid #dddddd;">
                  <table width="100%" cellpadding="0"
                         cellspacing="0" border="0">
                    <tr>
                      <td style="vertical-align:top;">
                        <table cellpadding="0"
                               cellspacing="0" border="0">
                          <tr>
                            <td style="font-size:12px;
                                       color:#444444;
                                       padding-bottom:5px;">
                              <strong>
                                Thanks &amp; Regards,
                              </strong>
                            </td>
                          </tr>
                          <tr>
                            <td style="font-size:13px;
                                       color:#00085D;
                                       font-weight:bold;
                                       padding-bottom:3px;">
                              Adit Sharma
                            </td>
                          </tr>
                          <tr>
                            <td style="font-size:12px;
                                       color:#666666;">
                              Analytics Team
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td align="right"
                          style="vertical-align:top;">
                        <img src="https://dgdlm6ddvctpd.cloudfront.net/bannerimages/wtiintelligentmobility.png"
                             alt="Company Logo"
                             width="100"
                             style="display:block;
                                    border:0;">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </center>
  </div>
</body>
</html>"""

    return html

def resolve_path(relative_or_absolute_path):
    """
    Resolve a path that may be relative or absolute
    from a different runner.
    Always returns absolute path on current runner.
    """
    if not relative_or_absolute_path:
        return relative_or_absolute_path

    repo_root = get_repo_root()

    # If already absolute and exists → use it
    if os.path.isabs(relative_or_absolute_path):
        if os.path.exists(relative_or_absolute_path):
            return relative_or_absolute_path

        # Extract relative part from old absolute path
        for marker in ['reports/', 'reports' + os.sep,
                       'data/', 'data' + os.sep]:
            idx = relative_or_absolute_path.find(marker)
            if idx != -1:
                rel_part = relative_or_absolute_path[idx:]
                resolved = os.path.join(repo_root, rel_part)
                return resolved

    # Relative path → join with repo root
    return os.path.join(
        repo_root, relative_or_absolute_path
    )
# ============================================================
#                    ATTACHMENT HELPER
# ============================================================

def attach_file(msg, file_path, display_name=None):
    """
    Attach a file to the email message.

    Args:
        msg          : MIMEMultipart email object
        file_path    : Full path to the file
        display_name : Name shown in email (optional)
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(
            f"❌ Attachment not found: {file_path}"
        )

    filename = display_name or os.path.basename(file_path)

    with open(file_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={filename}"
    )
    msg.attach(part)
    print(f"      📎 Attached: {filename}")


# ============================================================
#                    EMAIL SENDER CORE
# ============================================================

def send_email(sender, password, recipients,
               cc_list, subject, html_body,
               attachments):
    """
    Send HTML email with attachments via Gmail SMTP.

    Args:
        sender      : Sender Gmail address
        password    : Gmail App Password
        recipients  : List of To email addresses
        cc_list     : List of CC email addresses
        subject     : Email subject line
        html_body   : HTML string for email body
        attachments : List of dicts:
                      [{'path': '...', 'name': '...'}]
    """
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    # Attach HTML body
    msg.attach(MIMEText(html_body, "html"))

    # Switch to mixed for attachments
    msg_mixed = MIMEMultipart("mixed")
    msg_mixed["Subject"] = subject
    msg_mixed["From"]    = sender
    msg_mixed["To"]      = ", ".join(recipients)

    if cc_list:
        msg_mixed["Cc"] = ", ".join(cc_list)

    msg_mixed.attach(MIMEText(html_body, "html"))

    # Attach files
    for attachment in attachments:
        attach_file(
            msg_mixed,
            attachment["path"],
            attachment.get("name")
        )

    # All recipients (To + CC)
    all_recipients = recipients + cc_list

    print(f"\n      📧 Connecting to Gmail SMTP...")
    print(f"      📤 From    : {sender}")
    print(f"      📥 To      : {', '.join(recipients)}")
    if cc_list:
        print(f"      📋 CC      : {', '.join(cc_list)}")
    print(f"      📌 Subject : {subject}")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(
            sender,
            all_recipients,
            msg_mixed.as_string()
        )

    print(f"      ✅ Email sent successfully!")


# ============================================================
#                    MASTER FILE UPDATER
# ============================================================

def update_master_file_with_corrections(
    state, repo_root
):
    """
    After Email 2 flow:
    Read the corrected combined invoice file,
    extract newly mapped Corporate+Hub combinations,
    add them to the master file (skip complete duplicate rows).

    Logic:
        1. Load current master file
        2. Load corrected sample invoice
        3. Extract rows where Group Entity & Vertical are filled
           and NOT already in master
        4. Append new rows to master
        5. Remove complete duplicate rows
        6. Save updated master file
    """
    import pandas as pd

    master_path    = os.path.join(
        repo_root, "data", "Group_Entity_Vertical_Master.xlsx"
    )
    combined_path  = state.get("combined_file_path", "")

    if not os.path.exists(combined_path):
        print(
            f"      ⚠️  Combined file not found: {combined_path}"
        )
        return

    print(f"\n      🔄 Updating master file...")
    print(f"         Master   : {master_path}")
    print(f"         Combined : {combined_path}")

    # Load master
    master_df = pd.read_excel(master_path, sheet_name='Sheet1')

    # Load combined (corrected)
    combined_df = pd.read_excel(combined_path, sheet_name='Sheet1')

    # Filter rows that now have valid Group Entity & Vertical
    valid_mask = (
        combined_df['Group Entity'].notna() &
        combined_df['Vertical'].notna() &
        (~combined_df['Group Entity'].isin(
            ['Not Found', '', 'nan', 'None']
        )) &
        (~combined_df['Vertical'].isin(
            ['Not Found', '', 'nan', 'None']
        ))
    )
    mapped_df = combined_df[valid_mask].copy()

    # Build new master rows from mapped data
    new_rows = mapped_df[[
        'Corporate', 'Hub', 'Group Entity', 'Vertical'
    ]].drop_duplicates().rename(columns={
        'Corporate'    : 'Corporate Name',
        'Hub'          : 'Hub',
        'Group Entity' : 'GroupEntity',
        'Vertical'     : 'Vertical'
    })

    # Combine with existing master
    combined_master = pd.concat(
        [master_df, new_rows], ignore_index=True
    )

    # Remove COMPLETE duplicate rows only
    combined_master.drop_duplicates(inplace=True)
    combined_master.reset_index(drop=True, inplace=True)

    # Save updated master
    combined_master.to_excel(
        master_path, sheet_name='Sheet1', index=False
    )

    added = len(combined_master) - len(master_df)
    print(
        f"         ✅ Master updated: "
        f"{added} new row(s) added. "
        f"Total rows: {len(combined_master):,}"
    )


# ============================================================
#                       MAIN FUNCTION
# ============================================================

def main():
    # Read which email mode from command line argument
    # Usage:
    #   python email_sender.py email1
    #   python email_sender.py email2

    if len(sys.argv) < 2:
        raise ValueError(
            "❌ Please specify email mode: "
            "python email_sender.py email1|email2"
        )

    email_mode = sys.argv[1].lower().strip()

    if email_mode not in ("email1", "email2"):
        raise ValueError(
            f"❌ Invalid email mode: {email_mode}. "
            f"Use 'email1' or 'email2'."
        )

    print("\n" + "=" * 65)
    print(f"     📧 EMAIL SENDER — MODE: {email_mode.upper()}")
    print("=" * 65)

    # ── Step 1: Read State ───────────────────────────────────
    print("\n[1/5] Reading state file...")
    state, state_path = read_state_file()

    has_not_found          = state.get("has_not_found", False)
    not_found_summary_path = state.get("not_found_summary_path", "")
    final_report_path      = state.get("final_report_path", "")
    start_month            = state.get("start_month", "")
    end_month              = state.get("end_month", "")
    run_date_str           = state.get("run_date", "")

    print(f"      Has Not Found   : {has_not_found}")
    print(f"      Start Month     : {start_month}")
    print(f"      End Month       : {end_month}")
    print(f"      Report Path     : {final_report_path}")

    # ── Step 2: Get Email Config ─────────────────────────────
    print("\n[2/5] Loading email configuration...")
    sender, password, recipients, cc_list = get_email_config()

    month_range_display = format_month_range_display(
        start_month, end_month
    )

    # ── Step 3: Build Email Content ──────────────────────────
    print("\n[3/5] Building email content...")

    repo_root = get_repo_root()

    if email_mode == "email1":

        # ── EMAIL 1 ──────────────────────────────────────────
        if has_not_found:
            # Email 1 WITH Not Found
            not_found_count = 0
            if os.path.exists(not_found_summary_path):
                import pandas as pd
                nf_df = pd.read_excel(
                    not_found_summary_path,
                    engine='openpyxl'
                )
                not_found_count = len(nf_df)

            subject = (
                f"CRD Client-Wise Booking & Revenue Report: "
                f"{month_range_display} "
            )
            html_body = build_email1_with_not_found_body(
                start_month, end_month,
                not_found_count, run_date_str
            )
            # Single file attachment — report contains
            # both CRD Report (Sheet 1) and
            # Not Found Summary (Sheet 2)
            attachments = [
                {
                    "path": final_report_path,
                    "name": (
                        f"CRD_Report_"
                        f"{month_range_display.replace(' ', '_')}"
                        f".xlsx"
                    )
                }
            ]
            print(
                f"      Mode: Email 1 WITH Not Found "
                f"({not_found_count} records)"
            )

        else:
            # Email 1 WITHOUT Not Found
            subject = (
                f"CRD Client-Wise Booking & Revenue Report: "
                f"{month_range_display}"
            )
            html_body = build_email1_no_not_found_body(
                start_month, end_month, run_date_str
            )
            attachments = [
                {
                    "path": final_report_path,
                    "name": f"CRD_Report_{month_range_display.replace(' ', '_')}.xlsx"
                }
            ]
            print("      Mode: Email 1 — Complete (No Not Found)")

    else:
        # ── EMAIL 2 ──────────────────────────────────────────
        # Two valid scenarios for Email 2:
        # Scenario A: master_updated = True
        #             (normal followup after Not Found fix)
        # Scenario B: workflow_dispatch manual trigger
        #             (forced send regardless of flags)

        master_updated = state.get("master_updated", False)
        is_manual      = (
            len(sys.argv) > 2 and
            sys.argv[2] == "manual"
        ) or os.environ.get(
            "GITHUB_EVENT_NAME", ""
        ) == "workflow_dispatch"

        print(f"      master_updated  : {master_updated}")
        print(f"      is_manual       : {is_manual}")

        # Send if master was updated OR manual trigger
        if not master_updated and not is_manual:
            print(
                "\n      ℹ️  master_updated = False "
                "and not manual trigger. "
                "Email 2 not required."
            )
            return

        subject = (
            f"CRD Client-Wise Booking & Revenue Report: "
            f"{month_range_display}"
        )
        html_body = build_email2_final_body(
            start_month, end_month, run_date_str
        )
        attachments = [
            {
                "path": final_report_path,
                "name": (
                    f"CRD_Report_Final_"
                    f"{month_range_display.replace(' ', '_')}"
                    f".xlsx"
                )
            }
        ]
        print("      Mode: Email 2 — Final Corrected Report")

    # ── Step 4: Send Email ───────────────────────────────────
    print("\n[4/5] Sending email...")
    send_email(
        sender      = sender,
        password    = password,
        recipients  = recipients,
        cc_list     = cc_list,
        subject     = subject,
        html_body   = html_body,
        attachments = attachments
    )

    # ── Step 5: Post-Send Actions ────────────────────────────
    print("\n[5/5] Running post-send actions...")

    if email_mode == "email1":
        update_state_file(state_path, {"email1_sent": True})
        print("      ✅ State updated: email1_sent = True")

    elif email_mode == "email2":
        update_state_file(state_path, {
            "email2_sent"   : True,
            "master_updated": False
        })
        print("      ✅ State updated: email2_sent = True")
        print("      ✅ master_updated reset to False")

    # ── Final Summary ────────────────────────────────────────
    print("\n" + "=" * 65)
    print("              ✅ EMAIL PROCESS COMPLETE")
    print("=" * 65)
    print(f"  📧 Mode       : {email_mode.upper()}")
    print(f"  📌 Subject    : {subject}")
    print(f"  📥 Recipients : {', '.join(recipients)}")
    if cc_list:
        print(f"  📋 CC         : {', '.join(cc_list)}")
    print(f"  📎 Files      : {len(attachments)} attachment(s)")
    print("=" * 65)


if __name__ == "__main__":
    main()
