# SMA Fellows Evaluation Tool — Deployment Guide

**Who this is for:** The person setting up the tool for the first time, or the technical contact making future updates.  
**Time required:** ~30 minutes for initial setup.  
**Accounts needed:** Access to the SMA Google Workspace, a Netlify account (free).

---

## Overview

The tool has two parts:
- **The form** (`index.html`) — what fellows see and fill out. Hosted on Netlify.
- **The backend** (`Code.gs`) — a Google Apps Script that receives submissions and writes them to a Google Sheet.

These two parts are connected by a single URL. Most of the setup below is generating that URL and putting it in the right place.

---

## Part 1 — Create the Google Sheet

1. Sign in to Google using the SMA Workspace account (e.g. `fellows-eval@schoolformoralambition.org`).
2. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet.
3. Name it something clear — e.g. **SMA Fellows Evaluation — Responses**.
4. Share it with any team members who need to view or analyze data (use the Share button, top right). Viewers only need **Viewer** access; anyone who will analyze data in the sheet can have **Editor** access.
5. Leave this tab open — you'll come back to it.

---

## Part 2 — Set Up the Apps Script

1. In your Google Sheet, click **Extensions** in the top menu → **Apps Script**. A new tab opens showing a code editor.
2. Delete everything in the editor (it will show a small default function — select all and delete it).
3. Open the file `Code.gs` from the project folder in any text editor (Notepad, TextEdit, VS Code, etc.). Select all the text and copy it.
4. Paste it into the Apps Script editor.
5. Click the **Save** button (floppy disk icon, or Ctrl+S / Cmd+S). Name the project **SMA Eval Backend** when prompted.

---

## Part 3 — Run the Setup Function (once only)

1. In the Apps Script editor, find the dropdown that says **"Select function"** near the top toolbar. Click it and choose **`setup`**.
2. Click the **Run** button (▶ play icon).
3. A pop-up will appear asking you to authorize the script — click **Review permissions**, choose the SMA Workspace account, and click **Allow**. This lets the script write to your Google Sheet.
4. Wait about 10–15 seconds. When it finishes, you'll see **"Execution completed"** at the bottom of the screen.
5. Switch back to your Google Sheet tab and refresh the page. You should now see:
   - A **Responses** tab with column headers across row 1.
   - A **Data Dictionary** tab listing every column and what it means.

> **If you see an error instead:** Take a screenshot and send it to the technical contact. Do not run setup() again until the error is resolved.

---

## Part 4 — Deploy the Script as a Web App

This step generates the URL that connects the form to the sheet.

1. In the Apps Script editor, click **Deploy** (top right) → **New deployment**.
2. Click the gear icon ⚙ next to "Select type" and choose **Web app**.
3. Fill in the settings exactly as follows:
   - **Description:** `SMA Eval v1`
   - **Execute as:** `Me (fellows-eval@schoolformoralambition.org)`
   - **Who has access:** `Anyone`
4. Click **Deploy**.
5. Another authorization prompt may appear — click through and allow.
6. You'll see a screen with a **Web app URL** — it looks like `https://script.google.com/macros/s/ABC123.../exec`. **Copy this URL and save it somewhere safe** (paste it into a note or document). You will need it in the next step.
7. Click **Done**.

---

## Part 5 — Connect the Form to the Backend

1. Open `index.html` in a text editor (VS Code recommended, but Notepad works).
2. Near the top of the file (around line 950), find this line:
   ```
   const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
   ```
3. Replace `YOUR_APPS_SCRIPT_URL_HERE` with the URL you copied in Part 4. Keep the surrounding quote marks. It should look like:
   ```
   const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/ABC123.../exec';
   ```
4. Save the file.

---

## Part 6 — Host the Form on Netlify

1. Go to [netlify.com](https://netlify.com) and sign up for a free account (or log in if you already have one).
2. From the Netlify dashboard, find the section that says **"Deploy manually"** or **"Drag and drop your site folder here"**.
3. Drag the `index.html` file into that area (or click to browse and select it).
4. Netlify will give you a URL — something like `https://sma-fellows-eval.netlify.app`. This is the link you send to fellows.
5. Optional: To use a custom URL (e.g. `eval.schoolformoralambition.org`), click **Domain settings** and follow Netlify's instructions to connect a custom domain.

---

## Part 7 — Test the Full System

Before sharing with fellows, do a quick end-to-end test:

1. Open the Netlify URL in your browser.
2. Click **Begin Assessment** → select **T1 — Pre-orientation**.
3. Enter a real email address you have access to, and fill in a few responses on each page.
4. Submit the form. You should see a spinner, then a "Thank you" confirmation.
5. Go to your Google Sheet and check the **Responses** tab — a new row should appear with your email and the answers you entered.

If the row appears: everything is working. If the form shows an error message instead, check that the URL in `index.html` matches the one from Part 4 exactly, and that you saved the file before uploading to Netlify.

---

## Updating the Tool in the Future

### If the form questions change (index.html is updated):
1. Upload the new `index.html` to Netlify using the same drag-and-drop process. Netlify automatically replaces the old version.

### If the backend script changes (Code.gs is updated):
1. Open the Apps Script editor (Google Sheet → Extensions → Apps Script).
2. Make your changes.
3. Click **Deploy** → **Manage deployments**.
4. Find the existing deployment, click the **pencil / edit icon ✏️**.
5. Change **Version** to **"New version"**.
6. Click **Deploy**.

> ⚠️ **Important:** Always edit the existing deployment — never click "New deployment." Creating a new deployment generates a new URL, which breaks the connection to the form.

### If the script stops accepting submissions:
This is usually a re-authorization issue. Open the Apps Script editor, click **Run** on any function, and follow any authorization prompts that appear. Takes about 2 minutes.

---

## Key Information to Keep on File

| Item | Where to find it |
|---|---|
| Apps Script web app URL | Saved from Part 4 — also visible in Deploy → Manage deployments |
| Google Sheet | Bookmarked or in SMA Workspace Google Drive |
| Netlify site URL | Netlify dashboard |
| Backup files | Google Drive folder named **SMA Eval Backups** — updated automatically every Monday |
| Data dictionary | **Data Dictionary** tab in the Google Sheet |
| Technical contact | *(add name and email here)* |

---

## Common Questions

**Q: A fellow says they submitted but their data isn't in the sheet.**  
A: Check the **Responses** tab and search for their email. If it's missing, ask them to try again — the form will show an error message if something went wrong. If errors persist, check the Apps Script editor under **Executions** for recent error logs.

**Q: There are two rows for the same person.**  
A: This means they used a different email at two different touchpoints. Look for a row where `Link_Status` is flagged in yellow. Copy the data from the incomplete row into the correct row (matching by name), then delete the duplicate. The Data Dictionary tab explains every column if you're unsure what goes where.

**Q: The Link_Status column shows something like "T3_no_T2".**  
A: This means T3 was submitted without a T2 on file for that email. It does not mean data is missing — T1 and T3 data are both on the same row, and the T2 columns are simply blank. No action required unless T2 data is expected and should be tracked down.

**Q: Can I add or remove fellows' rows manually?**  
A: You can add rows manually if needed, but do not delete or rearrange existing rows or columns — the script writes to specific column positions and a structural change could cause data to land in the wrong columns. If you need to correct an error in a row, editing cell values is fine.
