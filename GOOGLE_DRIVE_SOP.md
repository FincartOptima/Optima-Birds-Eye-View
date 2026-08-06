# SOP — Google Drive Auto-Load for Birds Eye View Webpage

## 1. Purpose

Before this feature, whoever wanted to see the dashboard had to manually
upload three files (Tradebook, Client File, Account Statement) every single
time they opened the site. This was slow, easy to forget, and meant the site
could show stale or wrong data if someone uploaded the wrong file by mistake.

This feature removes that step. The three files now live in one shared
Google Drive folder. The webpage reads them directly from Drive every time
it's opened — no upload button, no file picker, nothing to remember.

## 2. End Result

- Opening **https://optima-birds-eye-view.onrender.com/** shows the latest
  data automatically, straight away.
- To update the data, someone just replaces a file in the Drive folder — no
  need to open the website at all.
- A **"🔄 Refresh from Drive"** button on the site lets anyone pull the very
  latest version on demand, instead of waiting for the automatic refresh.
- The old manual upload option is still there as a backup, in case Drive is
  ever unreachable or someone wants to test a one-off file.

## 3. How the Pieces Fit Together

- **Google Drive folder** — holds exactly 3 files, with fixed names:
  `Tradebook`, `Client File`, `Account Statement`.
- **A "robot" Google account (service account)** — a special account created
  in Google Cloud that the webpage logs in as. It only has *read* access to
  the folder, nothing else. Its login key lives on the server (Render), not
  in the file itself.
- **The webpage (Render)** — every time it's opened, it asks the robot
  account to fetch the 3 files, then processes them exactly like a manual
  upload would.

## 4. Setting This Up From Scratch (if it ever needs to be redone)

Use this section if the whole thing needs to be rebuilt — e.g. a new Google
account is being used, or the current setup was lost.

### Step 1 — Create a Google Cloud project
1. Go to **console.cloud.google.com**, signed in with the Google account
   that should own this.
2. Click the project dropdown near the top → **New Project**.
3. Name it anything memorable (e.g. `Birds Eye View`) → **Create**.
4. Make sure this new project is selected in that same dropdown before
   moving to the next step.

### Step 2 — Turn on the Google Drive API
1. Left menu → **APIs & Services → Library**.
2. Search **"Google Drive API"** → click it → **Enable**.

### Step 3 — Create the robot account (service account)
1. Left menu → **APIs & Services → Credentials**.
2. **+ Create Credentials → Service Account**.
3. Give it a name (e.g. `bev-drive-reader`) → **Create and Continue** →
   **Continue** → **Done** (no special permissions needed here).
4. Copy its email address — it looks like
   `bev-drive-reader@your-project.iam.gserviceaccount.com`. You'll need it
   in Step 5.

### Step 4 — Create its login key (JSON key file)
1. Click on the service account you just made.
2. Go to the **Keys** tab → **Add Key → Create new key → JSON → Create**.
3. A file downloads automatically. This is effectively a password —
   see **Precautions** below.

### Step 5 — Share the Drive folder with the robot
1. Create (or reuse) a Drive folder to hold the 3 files.
2. Right-click it → **Share**.
3. Paste the robot's email from Step 3.
4. Set its role to **Viewer** → **Send**.
5. Inside that folder, make sure there are exactly 3 files named (case
   doesn't matter, but spelling does):
   - `Tradebook` (the .xlsx trade allocation master)
   - `Client File` (the .csv client holdings snapshot)
   - `Account Statement` (the .xls account statement)

### Step 6 — Give the key to the website (Render)
1. Open the downloaded JSON key file in a text editor and copy its entire
   contents.
2. Go to the **Render dashboard → this service → Environment**.
3. Add a new environment variable:
   - **Key:** `GDRIVE_SERVICE_ACCOUNT_JSON`
   - **Value:** the entire JSON content you copied.
4. Save. Render redeploys automatically.
5. (Optional, only if the folder ever changes) Also add a variable
   `GDRIVE_FOLDER_ID` with the folder's ID from its Drive URL — otherwise
   the site keeps using the folder it was originally set up with.
6. Open the website and confirm it loads data automatically with no upload
   step.

## 5. How to Update the Data (day-to-day use)

This is all that's needed going forward — no code, no Render, no Google
Cloud:

1. Open the shared Drive folder.
2. **Replace** the relevant file (drag the new file directly onto the old
   one so it keeps the same name — this is safer than deleting first).
3. Wait up to 10 minutes for the site to notice on its own, **or** open the
   site and click **"🔄 Refresh from Drive"** to pull the new version
   immediately.

### Giving someone else permission to update the files
This is separate from everything in Section 4 — no Google Cloud access
needed at all:
1. Right-click the Drive folder → **Share**.
2. Add their Google account email.
3. Give them **Editor** access.

They now just use Drive normally, the same way they'd update any shared
document.

## 6. Precautions

- **Treat the JSON key file like a password.** Anyone who has it can read
  whatever the Drive folder holds. Never email it, never put it in the
  GitHub repository, never paste it anywhere public. It should only ever
  live in Render's environment variables and, for local testing, in this
  project's `credentials/` folder (which is already set up to be ignored by
  git, so it can never be accidentally committed).
- **If the key is ever exposed** (shared by mistake, leaked, etc.), delete
  it immediately in Google Cloud Console (Credentials → the service account
  → Keys tab → delete the compromised key), generate a new one the same
  way as Step 4, and update the `GDRIVE_SERVICE_ACCOUNT_JSON` value on
  Render.
- **Keep the 3 file names exact.** If a file is renamed, the site will
  silently stop finding it — there's no error banner for this, it just
  won't show updated data for that piece. Always double-check the name
  after uploading.
- **Avoid having two files with the same name in the folder at once.** If
  someone deletes-then-reuploads instead of replacing in place, there can
  be a brief moment where two files share a name — the site isn't
  guaranteed to pick the newest one in that case. Replacing in place (drag
  onto the existing file) avoids this entirely.
- **Give data-updaters Editor access to the Drive folder only** — never
  give them access to the Google Cloud project or the service account
  itself. They don't need it, and it's an unnecessary risk.
- **The manual upload option is still there as a fallback.** If Drive is
  ever unreachable, or the key expires, the site will simply show the old
  upload screen again — nothing breaks, it just stops auto-loading until
  the issue is fixed.

## 7. Future Scope of Improvement

- **Prefer "most recently modified" automatically** if two files ever share
  the same name, instead of picking whichever Drive happens to list first —
  a small safety net for the duplicate-file edge case above.
- **Move the 3 files into Google Sheets** instead of raw Excel/CSV files
  (matching how fund NAV data already works). Sheets update live the
  instant someone edits a cell — no re-uploading a file at all, and no risk
  of a corrupted binary file breaking the fetch.
- **Show a "data as of" timestamp on the webpage itself**, so anyone
  viewing it can see exactly how fresh the numbers are, instead of trusting
  the 10-minute refresh silently.
- **Alert on failure** — e.g. an email or Slack message if the Drive fetch
  ever fails (expired key, renamed file, folder unshared), so a silent
  failure doesn't go unnoticed for days.
- **Move from an in-memory cache to a scheduled background refresh** (e.g.
  a small job that re-fetches every few minutes) so page loads are always
  instant, rather than occasionally waiting on a live Drive fetch.
- **Consider Workload Identity Federation** instead of a static key file,
  per Google's own recommendation — removes the need to ever rotate a key
  manually. This mainly matters if the app is ever moved to run on Google
  Cloud infrastructure itself; it doesn't have an equivalent on Render today.
- **Extend the same Drive-based approach to the Allocation Rules history
  table**, which is currently hardcoded in the website's code rather than
  pulled from a shared file.
