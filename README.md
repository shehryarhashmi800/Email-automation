# Email Automation
Free automated cold email outreach system built with Google Apps Script.

## Features
- Reads client data from Google Sheets
- Reads email templates from Google Docs
- Sends emails automatically at scheduled times
- One email per 70 seconds — looks 100% human
- Tracks email opens with count, first open, last open
- Detects real replies vs auto-replies
- Cancels sequence automatically when client replies
- Runs 24/7 on Google servers — no hosting needed
- Completely free forever

## Requirements
- Google Account (free)
- Google Sheets
- Google Docs
- Google Apps Script (free — built into Google)

## Sheet Structure
Row 1 headers in exact order:
S.No | Brand Name | Whose Email | Email | Send Time (PKT) | Email 1 UID | Email 1 Date | Email 1 Status | Email 1 Sent At | Email 1 Open Count | Email 1 First Open | Email 1 Last Open | ... | Replied

## Doc Structure
[BRAND NAME - EMAIL 1]
Subject: Your subject line
Body:
Your email body here...

[BRAND NAME - EMAIL 2]
Subject: Your subject line
Body:
Your email body here...

## Setup Steps
1. Create Google Sheet with headers above
2. Create Google Doc with email templates
3. Open script.google.com
4. Paste script
5. Fill CONFIG section with your details
6. Deploy as Web App
7. Run setupTriggers
8. Done — runs automatically forever

## CONFIG Setup
Edit these values in the script:
- FROM_EMAIL — email to send from
- FROM_NAME — your name
- ALERT_EMAIL — where to receive error alerts
- SHEET_ID — your Google Sheet ID
- DOC_ID — your Google Doc ID
- TRACKING_URL — your Web App URL after deployment

## Auto-Reply Detection
System automatically ignores:
- Out of office replies
- Support ticket confirmations
- Review/survey requests
- Bounce messages
- Marketing auto-replies

## Open Tracking
- Tracks how many times each email opened
- Records first open timestamp
- Records last open timestamp
- Uses unique UID per email — no false counts

## License
MIT — free to use, modify, share
