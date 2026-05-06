# Organic Commerce Hub — College of Agriculture, Balaghat

A Google Apps Script web app for selling organic inputs from the college unit.

## What it includes
- Public product store
- Cart and online order form
- Admin-only product editing
- Student-only sales entry with access codes
- Google Sheet database
- Dashboard with revenue, expenses, cash-in-hand, and monthly chart
- PDF invoice download
- Email notifications through GmailApp
- Optional WhatsApp / SMS webhook hooks
- Loyalty points
- Hindi / English UI
- Basic AI crop doctor with voice input

## Important notes
- Email notifications work directly inside Google Apps Script.
- WhatsApp and SMS require your own provider API URL and token.
- UPI QR can be displayed from your UPI ID, but automatic UPI verification needs a payment gateway or a webhook.
- The AI doctor in this build is rule-based. It supports photo upload in the UI, but true image analysis needs a computer-vision API.

## Setup
1. Create a new Google Apps Script project.
2. Add `Code.gs` and `Index.html`.
3. Open **Project Settings** and set the timezone to `Asia/Kolkata` if needed.
4. Run `setup()` once from the editor.
5. Deploy → **New deployment** → **Web app**
   - Execute as: Me
   - Who has access: Anyone with the link, or your preferred access level
6. Copy the deployed URL and paste it into the share-link box in the app.
7. Change the admin PIN in Settings after first login.

## Default admin PIN
`2022`

## Default student access codes
`STU-2022-A`, `STU-2022-B`, `STU-2022-C`

## Spreadsheet tabs
- Products
- Orders
- Sales
- Expenses
- Customers
- Staff
- Settings
- Audit