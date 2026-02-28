# PE Deal Sourcing Automator

**Goal:** A standalone Windows tool for automating CEO outreach via local Outlook.

## Tech Stack

- **Frontend:** NiceGUI (Tailwind-based)
- **Logic:** Python 3.11+
- **Integration:** win32com (Local Outlook Desktop)
- **Data:** Local Excel (.xlsx) & HTML (.html) templates

## Project Architecture

- `main.py`: Entry point & NiceGUI Layout.
- `email_engine.py`: Logic for switching between Mock (Linux) and Real (Windows) Outlook.
- `file_manager.py`: CRUD operations for local templates and contact lists.
- `/templates`: Directory for .txt outreach scripts.
- `/data`: Directory for deal sourcing Excel files.

## Constraints

1. **Compliance:** No cloud APIs for email. Must use local Outlook session.
2. **Portability:** Must run as a standalone .exe via PyInstaller.
3. **UI Vibe:** Clean & Corporate (White/Slate/Navy).

## Software core features

1. **Campaings:** This is the core feature which merges e-mail templates with contact lists and prepares the draft e-mails via a workflow.

2. **Contact Lists:** This feature manages the contact lists

3. **Template Libary:** This feature acts as a libary for e-mail templates
