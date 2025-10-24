# ========================================
# Gmail Mail Merge Tool - Batch + Resume (Silent Mode)
# ========================================
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge for Phoenixx IT", layout="wide")
st.title("📧 Gmail Mail Merge Tool for Phoenixx IT")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.compose",
]

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

# ========================================
# Constants
# ========================================
DONE_FILE = "/tmp/mailmerge_done.json"
BATCH_SIZE_DEFAULT = 50

# ========================================
# Recovery Logic
# ========================================
if os.path.exists(DONE_FILE) and not st.session_state.get("done", False):
    try:
        with open(DONE_FILE, "r") as f:
            done_info = json.load(f)
        file_path = done_info.get("file")
        if file_path and os.path.exists(file_path):
            st.session_state["file_path"] = file_path  # Store in session_state
            st.success("✅ Previous mail merge completed successfully.")
            with open(st.session_state["file_path"], "rb") as f:
                st.download_button(
                    "⬇️ Download Updated CSV",
                    data=f,
                    file_name=os.path.basename(st.session_state["file_path"]),
                    mime="text/csv",
                )
            if st.button("🔁 Reset for New Run"):
                os.remove(DONE_FILE)
                st.session_state.clear()
                st.experimental_rerun()
            st.stop()
    except Exception:
        pass

# ========================================
# Helpers
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"""
    <html><body style="font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;">
        {text}
    </body></html>
    """

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        created_label = service.users().labels().create(
            userId="me",
            body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
        ).execute()
        return created_label["id"]
    except Exception:
        return None

def send_email_backup(service, csv_path):
    try:
        user_email = service.users().getProfile(userId="me").execute()["emailAddress"]
        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        msg.attach(MIMEText("Attached is the backup CSV for your mail merge run.", "plain"))
        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        st.info(f"📧 Backup CSV emailed to {user_email}")
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

def fetch_message_id_header(service, message_id):
    for _ in range(6):
        try:
            msg_detail = service.users().messages().get(
                userId="me", id=message_id, format="metadata", metadataHeaders=["Message-ID"]
            ).execute()
            headers = msg_detail.get("payload", {}).get("headers", [])
            for h in headers:
                if h.get("name", "").lower() == "message-id":
                    return h.get("value")
        except Exception:
            pass
        time.sleep(random.uniform(1, 2))
    return ""

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
else:
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code[0])
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Session Setup
# ========================================
if "sending" not in st.session_state:
    st.session_state["sending"] = False
if "done" not in st.session_state:
    st.session_state["done"] = False

# ========================================
# MAIN UI
# ========================================
if not st.session_state["sending"]:
    st.header("📤 Upload Recipient List")
    st.info("⚠️ Upload up to **70–80 contacts** for smooth performance.")
    uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.lower().endswith("csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        # Add missing columns
        for col in ["ThreadId", "RfcMessageId", "Status"]:
            if col not in df.columns:
                df[col] = ""

        df.reset_index(drop=True, inplace=True)
        st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

        # Editable data grid
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        # ✅ Sync deletions before sending
        df = edited_df.reset_index(drop=True)

        pending_indices = df.index[df["Status"] != "Sent"].tolist()

        subject_template = st.text_input("Subject", "Hello {Name}")
        body_template = st.text_area(
            "Body",
            """Dear {Name},

Welcome to **Mail Merge App** demo.

Thanks,  
**Your Company**""",
            height=250,
        )
        label_name = st.text_input("Gmail label", "Mail Merge Sent")
        delay = st.slider("Delay (seconds)", 20, 75, 20)
        send_mode = st.radio("Choose mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        # Preview
        if not df.empty:
            preview_row = df.iloc[0]
            try:
                preview_subject = subject_template.format(**preview_row)
                preview_body = convert_bold(body_template.format(**preview_row))
            except Exception as e:
                preview_subject = subject_template
                preview_body = body_template
                st.warning(f"⚠️ Could not render preview: {e}")

            st.markdown("### 👀 Preview (First Row)")
            st.markdown(f"**Subject:** {preview_subject}")
            st.markdown(preview_body, unsafe_allow_html=True)

        if st.button("🚀 Send Emails / Save Drafts"):
            st.session_state.update({
                "sending": True,
                "df": df,
                "pending_indices": pending_indices,
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode
            })
            st.rerun()

# ========================================
# Sending Mode with Batch Labeling
# ========================================
if st.session_state["sending"]:
    df = st.session_state["df"]
    pending_indices = st.session_state["pending_indices"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]

    st.markdown("<h3>📨 Sending emails... please wait.</h3>", unsafe_allow_html=True)
    progress = st.progress(0)
    status_box = st.empty()

    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    total = len(pending_indices)
    sent_count, skipped, errors = 0, [], []
    batch_count = 0
    sent_message_ids = []

    for i, idx in enumerate(pending_indices):
        if send_mode != "💾 Save as Draft" and batch_count >= BATCH_SIZE_DEFAULT:
            break

        row = df.iloc[idx]

        pct = int(((i + 1) / total) * 100)
        progress.progress(min(max(pct, 0), 100))
        status_box.info(f"Processing {i + 1}/{total}")

        to_addr = extract_email(str(row.get("Email", "")).strip())
        if not to_addr:
            skipped.append(row.get("Email"))
            df.loc[idx, "Status"] = "Skipped"
            continue

        try:
            subject = subject_template.format(**row)
            body_html = convert_bold(body_template.format(**row))
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject

            msg_body = {}
            if send_mode == "↩️ Follow-up (Reply)":
                thread_id = str(row.get("ThreadId", "")).strip()
                rfc_id = str(row.get("RfcMessageId", "")).strip()
                if thread_id and rfc_id:
                    message["In-Reply-To"] = rfc_id
                    message["References"] = rfc_id
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    msg_body = {"raw": raw, "threadId": thread_id}
                else:
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    msg_body = {"raw": raw}
            else:
                raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                df.loc[idx, "Status"] = "Draft"
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                msg_id = sent_msg.get("id", "")
                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                df.loc[idx, "RfcMessageId"] = fetch_message_id_header(service, msg_id) or msg_id
                df.loc[idx, "Status"] = "Sent"
                if send_mode == "🆕 New Email" and label_id:
                    sent_message_ids.append(msg_id)
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))

            sent_count += 1
            batch_count += 1
        except Exception as e:
            df.loc[idx, "Status"] = "Error"
            errors.append((to_addr, str(e)))
            st.error(f"Error for {to_addr}: {e}")

    # Save & backup
    if send_mode != "💾 Save as Draft":
        if sent_message_ids and label_id:
            try:
                service.users().messages().batchModify(
                    userId="me",
                    body={"ids": sent_message_ids, "addLabelIds": [label_id]}
                ).execute()
            except Exception as e:
                st.warning(f"Batch labeling failed: {e}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"Updated_{safe_label}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)
        df.to_csv(file_path, index=False)

        # ✅ Store file_path in session_state
        st.session_state["file_path"] = file_path

        try:
            send_email_backup(service, file_path)
        except Exception as e:
            st.warning(f"Backup email failed: {e}")

        with open(DONE_FILE, "w") as f:
            json.dump({"done_time": str(datetime.now()), "file": file_path}, f)

    st.session_state["sending"] = False
    st.session_state["done"] = True
    st.session_state["summary"] = {"sent": sent_count, "errors": errors, "skipped": skipped}
    st.rerun()

# ========================================
# Completion
# ========================================
if st.session_state["done"]:
    summary = st.session_state.get("summary", {})
    st.success(f"✅ Process completed. Sent: {summary.get('sent', 0)}")
    if summary.get("errors"):
        st.error(f"❌ {len(summary['errors'])} errors occurred.")
    if summary.get("skipped"):
        st.warning(f"⚠️ Skipped: {summary['skipped']}")

    # ✅ Show download button using session_state
    file_path = st.session_state.get("file_path")
    if file_path and os.path.exists(file_path):
        with open(file_path, "rb") as f:
            st.download_button(
                "⬇️ Download Updated CSV",
                data=f,
                file_name=os.path.basename(file_path),
                mime="text/csv",
            )

    if st.button("🔁 New Run / Reset"):
        if os.path.exists(DONE_FILE):
            os.remove(DONE_FILE)
        st.session_state.clear()
        st.experimental_rerun()

