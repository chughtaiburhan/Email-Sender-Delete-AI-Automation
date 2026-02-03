# Email-Sender-Delete-AI-Automation

A Google Apps Script solution that automates personalized email sends, template reuse, and Gmail cleanup.

## Features
- **Profile + Sheet creation**: A form captures name, email, and phone, then creates a connected Google Sheet.
- **Templates**: Use a Google Doc template or save a draft subject/body.
- **Personalization**: Use placeholders like `${name}`, `${email}`, `${about}` mapped from sheet headers.
- **Chunked sending**: Sends emails in batches and splits workloads into two pipelines for large lists.
- **Gmail cleanup**: Move threads older than 30 days to Trash.

## Setup
1. Create a new Apps Script project and add the files from this repo.
2. (Optional) Add a Google Drive folder ID in `CONFIG.templateFolderId` to limit templates.
3. Deploy as a web app (execute as yourself) and open the URL.

## Usage
1. **Connect**: Fill the profile form to create your sheet.
2. **Templates**: Load your Google Docs templates or save a draft subject/body.
3. **Recipients**: Add rows in the `Recipients` sheet with headers like `email`, `name`, `about`, `company`.
4. **Send**: Click **Send Emails**. Use placeholders `${header}` in templates for personalization.
5. **Cleanup**: Click **Clean 30d Old Email** to move old threads to Trash.

## Notes
- The system follows a DRY approach with reusable helpers for chunking, templating, and validation.
- Gmail send limits apply to Apps Script accounts.
