# Fitness Journal

A self-hosted Google Apps Script workout journal backed by your own Google Sheet. Log exercises from a web app, restore unfinished sessions, reuse templates, and review training history through a dashboard and weekly email digest.

## Who is this for?

Fitness Journal is my attempt to make logging my workouts easier while gathering more actionable information from those logs. **"You can't improve what you don't measure."**

I used to log my workouts in my phone's Notes app and then transcribe that information into an Excel sheet. I had the raw data, but gathering it was time-consuming, and what I ended up with was mostly rows and rows of workout information that were difficult to navigate and hard to act upon

I could have paid for any number of fitness apps, but I wanted to retain control of my information and be able to access it whenever I wanted.

In my experience, using the app has made me more involved in my workouts and made the process more enjoyable. The information I've gathered has been valuable to me, even when its value is simply being able to see my progress over time.

Deploying and setting up this app is definitely more technically involved than using a paid fitness app, but what you get in return is:

- Freedom: The app and the code are yours, forever.
- More privacy: Your workout data remains in your own Google Sheet rather than in a separate fitness platform.
- Access: As long as Google Apps Script and Google Sheets remain available, you'll be able to access your app and its underlying data.

If you're starting your own fitness journey and want a way to track your progress indefinitely without paying for a subscription, and retaining control of all the underlying information, maybe this is for you.

## Major capabilities

- Log straight sets or individual sets with different reps and loads.
- Restore unfinished sessions and reuse workout templates.
- Track load PRs and the most reps at your maximum-ever load, including longest isometric holds at that load.
- Review volume, exercise progress, muscle-group gaps, and workout density in Google Sheets.
- Receive an optional weekly training digest by email.
- Open your journal from a shortcut on your phone's Home Screen.

## Documentation

- [Setup guide](SETUP.md): deployment, permissions, mobile shortcuts, troubleshooting, and updates.
- [Analytics guide](ANALYTICS.md): dashboard and digest contents, calculations, and limitations.
- [Changelog](CHANGELOG.md): completed changes and upgrade notes.

## Quick installation

1. Create a Google Sheet with a worksheet named **StrengthWorkoutINPUT**. Add these headers in columns A-G: `Date`, `Exercise`, `Sets`, `Reps`, `Load`, `Focus`, `Notes`.
2. Open **Extensions > Apps Script**. Replace `Code.gs` with this repository's code and create an HTML file named **Index** containing `Index.html`.
3. Save the project and match its time zone to the spreadsheet's time zone.
4. Choose **Deploy > New deployment > Web app**, execute as **Me**, and restrict access to **Only myself**. Complete authorization.
5. Open the deployed `/exec` URL using the same Google account.

For the full walkthrough and phone shortcuts, see [SETUP.md](SETUP.md). This is a single-user journal: each installation should have its own spreadsheet and script project. No API keys are required.

## Basic usage

1. Set the workout date; it defaults to today, but you can change it for past workouts.
2. Fill in at least one exercise card. Use **Add Exercise** for additional exercises.
3. For varying sets, choose **Enter sets individually**. Use **Back to single entry** to return.
4. To reuse the workout, choose **Save Current Exercises as Template** before saving the session. Load a saved template to prefill its exercises later.
5. Press **Save Session** and wait for confirmation. The app announces new records and clears the cards for your next session.

Each card has these fields:

- **Focus:** Choose the muscle group first to filter recent exercise suggestions.
- **Exercise:** The exercise name.
- **Brand:** Optional equipment brand for this logged entry.
- **Sets:** Number of sets.
- **Reps:** Reps per set, or hold duration in seconds for isometrics.
- **Load:** Weight in kilograms.
- **Notes:** Additional information.

Use a name containing `iso` for an isometric exercise. Exercise suggestions use the existing dashboard summary, filtered by Focus and an adjustable 30-day recency cutoff. An empty Focus field can be filled from the exercise's historical focus associations. You can still type any known exercise in full, including older exercises, with its historical PR and 1RM hints. Sets, reps, and load are prefilled by templates, not automatically restored from history.

The app shows muscle-group training gaps and live PR hints. Rep or hold records appear when your entered load meets or exceeds your saved load PR. See [PR definitions](ANALYTICS.md#personal-records) for details.

Open the spreadsheet's **Dashboard** tab for charts. After manual spreadsheet edits, choose **Fitness Journal > Rebuild Dashboard**. Enable the optional weekly email using the [digest setup instructions](SETUP.md#weekly-digest-setup).

## Focus, Brand, and bodyweight

The app adds a Brand header in column H of StrengthWorkoutINPUT and Templates. Existing A:G columns and historical exercise names are preserved; no brand extraction or name migration runs automatically. Brand is saved in drafts and templates and does not split exercise volume, PRs, or progression. See [schema and upgrade notes](SETUP.md#brand-and-focus-upgrade).

Dashboard rebuilds normalize historical Focus formatting (for example, chest, BICEPS becomes Chest,Biceps). The existing exercise summary includes all historical focus associations; no second catalog is created. Muscle Group Gaps tracks each group in a combined focus separately, while volume charts use the first-listed group.

Optional bodyweight is stored in Bodyweight. Assistance uses negative loads, unassisted sets use zero, and added weight uses positive loads. ExerciseSettings stores persistent, editable bodyweight classifications. Reporting uses smoothed bodyweight plus logged load; app PRs retain raw loads. See [analytics](ANALYTICS.md) for details.

## Files

- [Code.gs](Code.gs): Apps Script backend, spreadsheet storage, dashboard, and digest.
- [Index.html](Index.html): Web app interface.
- [tests/pr-records.test.js](tests/pr-records.test.js): Local regression tests for PR calculations and hints.

## Tests

With Node.js installed, run:

```sh
node --test tests/pr-records.test.js
```

Tests use mocked spreadsheet services. Live Apps Script deployment and Google service integration require separate verification.

## Updates

All updates will be committed to the repo, and will have to be deployed manually to your own Google Sheet and Apps Script. The core functionality is unlikely to change in the near future. Follow [Updating your deployment](SETUP.md#deployment-updates) to keep your existing phone shortcut working.
