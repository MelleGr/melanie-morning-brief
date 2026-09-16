# Kampkalenderen — backend setup (one-time, ~5 minutes)

This is the free, no-account-for-players backend for the handball tool at
`docs/kampkalender/`. It runs as a Google Apps Script Web App attached to a
Google Sheet you own — nobody but you needs a Google account.

## Deploy

1. Go to [sheets.google.com](https://sheets.google.com) → create a new blank
   spreadsheet. Name it e.g. "Kampkalenderen data".
2. In the sheet, go to **Extensions → Apps Script**.
3. Delete the placeholder code in `Code.gs` and paste in the contents of
   `handball-backend/Code.gs` from this repo.
4. Near the bottom, find `setAdminKey()` and replace `YOUR_SECRET_HERE` with
   a passphrase only you (and any co-manager) will know — this is what lets
   you add/edit/delete matches from the tool. Save (Ctrl/Cmd+S).
5. In the function dropdown at the top (next to the bug icon), select
   `setAdminKey`, then click **Run**. The first time, Google will ask you to
   authorize the script — click through **Advanced → Go to (project) (unsafe)**
   and **Allow**. This is your own script running on your own sheet, so this
   prompt is expected.
6. Click **Deploy → New deployment**. Click the gear next to "Select type" →
   **Web app**.
   - Description: anything, e.g. "Kampkalenderen v1"
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Click **Deploy**, authorize again if asked, then copy the **Web app URL**
   it gives you (ends in `/exec`).
8. Send me that URL and the passphrase you set in step 4 — I'll wire them
   into the page and push it live.

That's it — no billing, no separate account, nothing to install. The sheet
itself is your live database; you can always open it directly to see the
raw data.

## If you ever need to redeploy

Any time you change `Code.gs`, use **Deploy → Manage deployments → edit
(pencil) → New version → Deploy** to publish the change — the URL stays the
same.
