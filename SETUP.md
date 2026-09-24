# Setup guide

[Project overview](README.md) | [Analytics guide](ANALYTICS.md)

## Deploy the app

Complete the initial setup in a desktop browser. You'll need a Google account and the `Code.gs` and `Index.html` files from this repository. Node.js is only needed for local tests, not to deploy or use the app.

### 1. Prepare your spreadsheet

1. Sign in to the Google account you want to use for your journal.
2. Create a Google Sheet and name the document **Fitness Journal** (or any name you prefer).
3. Rename the first worksheet tab to **StrengthWorkoutINPUT**, matching the capitalization exactly. The document title and worksheet tab name are separate settings.
4. In row 1, enter these headers in columns A–G, one per cell: `Date`, `Exercise`, `Sets`, `Reps`, `Load`, `Focus`, `Notes`.
5. In the spreadsheet's **File > Settings**, check the time zone and save your changes.

### 2. Add the app code

1. From that spreadsheet, open **Extensions > Apps Script**. This creates a script project bound to your journal spreadsheet.
2. Name the project **Fitness Journal**.
3. Open the default `Code.gs` file in the script editor. Replace its starter code with the complete contents of this repository's `Code.gs`.
4. Next to **Files**, click **+ > HTML**. Enter **Index** as the name; the editor adds the `.html` extension.
5. Replace the new HTML file's starter content with the complete contents of this repository's `Index.html`. When copying from GitHub, use the file's **Raw** view to copy only the source code.
6. Save the project. Open **Project Settings** (the gear icon) and set its time zone to match the spreadsheet.
7. Leave `IMPORT_SOURCE_SHEETS = []` and `DIGEST_RECIPIENT_EMAIL = null` for a new installation. Historical imports and a custom email recipient are optional.

### 3. Deploy your private web app

1. In the Apps Script editor, click **Deploy > New deployment**.
2. Open the **Select type** menu (gear icon) and choose **Web app**.
3. Enter a description, such as **Initial release**.
4. Set **Execute as** to **Me** and **Who has access** to **Only myself**. Available choices can depend on your Google account or organization's policies. This is a single-user journal: everyone given access to one deployment would share its drafts, templates, and workout history.
5. Click **Deploy**. If prompted, choose **Authorize access**, select the account that owns your journal, review the requested permissions, and complete authorization. If Google displays an unverified-app notice, confirm that it identifies your own script project before proceeding. A work or school account may require administrator approval.
6. Copy the **Web app URL** from the deployment dialog and click **Done**. Save this URL in a private bookmark. Use the deployed URL ending in `/exec`; a test deployment ending in `/dev` is intended for development.
7. Open the copied URL while signed in to the same Google account. The Fitness Journal form should load.

See Google's [web app deployment guide](https://developers.google.com/apps-script/guides/web) and [access settings](https://developers.google.com/apps-script/manifest/web-app-api-executable).

### 4. Check your installation

1. In the web app, enter your first workout, including at least one exercise name.
2. Press **Save Session** and wait for the saved confirmation.
3. Return to your spreadsheet and check that the workout appears in `StrengthWorkoutINPUT`.
4. Reload the spreadsheet to show the **Fitness Journal** menu. It provides dashboard rebuilding, historical imports, exercise name merging, and optional weekly digest setup.

The app creates supporting tabs as needed. For weekly email summaries, follow the digest setup below.

## Configuration and data

No API keys are required. Apps Script uses Google authorization to access the bound spreadsheet and send the optional digest. Workout data stays in the spreadsheet; this repository contains source code and synthetic test fixtures only.

Historical imports are disabled by default (`IMPORT_SOURCE_SHEETS = []`). Add your own source tab names locally if needed. The optional digest recipient defaults to the user running the script; configure an override only in your own Apps Script project.

Do not commit spreadsheet exports, personal email addresses, deployment URLs or IDs, or credentials to the public repository. The private deployment still requires Google authorization; its URL is not a substitute for authentication.

Each installation should use its own spreadsheet and Apps Script project.

## Weekly digest setup

Enable weekly email summaries from your spreadsheet:

1. Open your journal spreadsheet and reload it if the **Fitness Journal** menu is missing.
2. Choose **Fitness Journal > Set Up Weekly Digest (Monday mornings)**.
3. Complete any Google authorization prompts. The digest needs permission to send email.
4. The app schedules delivery for Monday mornings, approximately 8–9 a.m. in the script project's time zone.
5. After logging at least one workout, choose **Fitness Journal > Send Weekly Digest Now (test)** to send an immediate email and check delivery. This sends a real digest; it does not create another schedule.

Running setup again from the same account does not create a duplicate digest trigger. The recipient defaults to the account running the script; you can set `DIGEST_RECIPIENT_EMAIL` in your own Apps Script project to use another address. An empty workout log does not generate a digest.

See [the digest contents and reporting period](ANALYTICS.md#weekly-digest) for what the email covers.

## Add the journal to your phone

Send yourself the deployed `/exec` URL or access it through your private bookmarks. Open it directly in Safari or Chrome, sign in to the Google account that deployed the app, and wait for the journal form to load before creating a shortcut. Bookmark the web app, rather than the spreadsheet, script editor, or Google sign-in page.

### iPhone: Safari

1. Open **Safari** and visit your deployed web app URL.
2. Tap **Share** (the square with an upward arrow). Depending on your Safari layout, open the page menu first, then tap **Share**.
3. Scroll through the actions and tap **Add to Home Screen**. If it is missing, look under **Edit Actions**.
4. Name the shortcut **Fitness Journal**. If **Open as Web App** is available, enable it to open in its own window.
5. Tap **Add**, return to your Home Screen, and tap the new icon to check that the journal opens. Sign in again if prompted.

See Apple's [Safari Home Screen instructions](https://support.apple.com/guide/iphone/iphea86e5236/ios).

### Android: Chrome

1. Open **Chrome** and visit your deployed web app URL.
2. Tap the **three-dot menu** beside the address bar.
3. Choose **Install and create shortcut > Create shortcut**. Depending on your Chrome version, the menu may instead show **Add to Home screen**.
4. Name the shortcut **Fitness Journal**, then tap **Add** and confirm placement if your phone asks.
5. Return to your Home Screen and tap the new icon to check that it opens your journal. A shortcut with the Chrome badge opens in Chrome.

See Google's [Chrome shortcut instructions](https://support.google.com/chrome/answer/15085120?co=GENIE.Platform%3DAndroid&hl=en).

### If the shortcut does not open correctly

- **Access denied or repeated sign-in:** Open the original `/exec` URL in a regular browser tab and check that you're using the deploying Google account. If multiple accounts are signed in, retry with only that account signed in. Keep the deployment restricted to yourself.
- **The shortcut opens a login page:** Open the saved deployment URL, finish signing in, and recreate the shortcut once the journal is visible.
- **The iPhone web-app window cannot complete sign-in:** Try the URL in Safari. You can also recreate the Home Screen shortcut with **Open as Web App** disabled, if that option is available, or use a Safari bookmark.
- **Changes to the code are missing:** Follow the deployment update steps below, then reopen or refresh the journal.

The shortcut provides quick access to the hosted app. An internet connection is required to load history, save drafts, and save sessions; it does not add offline support. Wait for the save confirmation before closing a completed session.

## Deployment updates

This repository stores the source code. Pushing to GitHub does not automatically update the Apps Script deployment; copy changes to the script project and update its deployment separately.

1. Open your spreadsheet, then **Extensions > Apps Script**.
2. Update `Code.gs` and `Index.html` with the new source, preserving any settings you customized, and save.
3. Click **Deploy > Manage deployments** and select your existing web app deployment.
4. Click **Edit** (pencil icon). Under **Version**, select **New version**, add a description, and click **Deploy**. Complete authorization if new permissions are requested.
5. Reopen or refresh the web app on your phone. Updating the existing deployment preserves its URL, so your Home Screen shortcut continues to work.

See Google's [deployment update guide](https://developers.google.com/apps-script/concepts/deployments).
