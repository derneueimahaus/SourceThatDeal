# SourceThatDeal

A standalone Windows desktop tool for automating e-mail mass outreach for sales activties in an environment where modern tools are not efasible due to compliannce conerns.. Load a contact list, pick an HTML email template, and bulk-create personalized draft emails directly in your local Outlook Drafts folder — ready to review and send.

No cloud APIs. No email credentials. Everything stays on your machine.

---

## Core Features

### 1. Campaigns
The primary workflow. Map `{{placeholder}}` fields in an email template to columns in a contact list (.xlsx), optionally schedule a deferred send time, and generate personalized Outlook drafts for every row in one click. Your Outlook signature is appended automatically.

### 2. Contact Lists
Import and manage `.xlsx` contact lists. View and edit rows directly in the app. Each list is stored locally in the `/data` folder.

### 3. Template Library
A library of HTML email templates organized in named folders. Includes a rich WYSIWYG editor with merge-field insertion (e.g. `{{First Name}}`, `{{Company}}`), font/size controls, and folder management.

---

## Key Design Choices

**OOM-only Outlook integration — no MAPI, no OMG dialogs**

Drafts are created via the Outlook Object Model (COM/`win32com`). Recipients are set as plain display strings (`mail.To = addr`) rather than resolved through the address book, which avoids triggering Outlook's Object Model Guard security dialog. No `Recipients.Add()`, no `GetInspector`.

**Filesystem signature reading**

The default Outlook signature is read directly from `%APPDATA%\Microsoft\Signatures\` (checking the Windows registry for the configured name, falling back to the only `.htm` file if unset). No COM call is made for signature extraction — avoids the `CurrentUser.Address` OMG trigger.

**No cloud APIs**

All email is handled through the local Outlook desktop client. Nothing is sent to external services.

**Mock backend for non-Windows / dev environments**

On Linux or Windows without `win32com`, all Outlook operations silently no-op with a log message so the UI and logic can be developed and tested without Outlook installed.

---

## Requirements

- Windows 10/11
- Microsoft Outlook desktop (logged in, running)
- Python 3.11+

---

## Getting Started

```bash
# 1. Clone the repo
git clone https://github.com/your-username/SourceThatDeal.git
cd SourceThatDeal

# 2. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
python main.py
```

NiceGUI starts a local web server (default port 8080) and opens the app in your browser automatically.

> **Note:** Make sure Outlook is open before launching — the app connects to the running Outlook instance via COM.
