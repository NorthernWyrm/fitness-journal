# Analytics guide

[Project overview](README.md) | [Setup guide](SETUP.md)

The dashboard and digest summarize the workout rows in your Google Sheet. PRs are derived from existing history; no separate record migration is required.

## Dashboard

The **Dashboard** tab in your spreadsheet contains charts and summary tables built from your workout history. The app rebuilds it after saving a session. To refresh it after editing the spreadsheet directly, or to recalculate the rolling date windows on a day when you haven't logged a workout:

1. Open your journal spreadsheet.
2. Choose **Fitness Journal > Rebuild Dashboard**. Reload the spreadsheet if the menu is missing.
3. Open the **Dashboard** tab and scroll through the tables and charts.

The dashboard includes:

- **Training volume:** Daily totals, with trend charts for the last 7 days and last 3 months.
- **Volume by exercise:** The top 15 exercises by total volume over the last 3 weeks.
- **Muscle-group balance:** Volume grouped by Focus over the last 7 days and last 3 months.
- **Exercise progress:** Number of logged entries, first and latest loads and average reps per set, changes between those entries, first and latest training dates, best estimated 1RM using the Epley and Brzycki formulas, and most common set pattern.
- **Muscle Group Gaps:** The most recent training date for each recorded group and days since it was trained, capped at 15 days.
- **Workout density:** Volume per minute for sessions with recorded timing.

Charts appear when enough relevant data is available; trend lines require at least two data points. The dashboard is regenerated, so keep your own notes and custom charts on a separate tab.

## Weekly digest

The digest summarizes the rolling last 7 days, including the day it runs, and compares them with the preceding 7 days. It is not a fixed Monday-to-Sunday report. See [SETUP.md](SETUP.md#weekly-digest-setup) to enable delivery or send a test email.

The email includes:

- **Weekly activity:** Exercises performed, their logged counts and volume, total volume, and the change from the previous 7 days.
- **Muscle-group coverage:** Volume by group and up to five recorded groups with the longest training gaps.
- **Rest-day patterns:** Days since the last session, average days between training dates, and the most often skipped weekday.
- **Workload ratio:** Last 7 days' volume divided by the average weekly volume over the last 28 days.
- **Exercise trends:** Week-over-week volume and average representative load for exercises trained in both periods. Changes within 10% are labeled plateaued.
- **Workout density:** Average volume per minute across timed sessions in the reporting period.
- **New load PRs:** Loads that exceeded an earlier recorded best during the reporting period. First-ever entries are excluded here; rep and hold-duration PRs currently appear in the app's live hints and save confirmations rather than the digest.
- **Plateau alerts:** Exercises whose maximum load has not been exceeded for at least four subsequent logged entries, subject to the configured history and volume thresholds. Repeating the same maximum does not reset that count.

## Volume

For straight sets, volume is sets * reps * load. For individual sets, sum each set's reps * load. The app labels this weight-based training volume in kilograms. Isometric holds count once per set for volume; a 30-second hold is not treated as 30 repetitions. Exercises without a recorded load contribute no weight-based volume.

A single reps or load value repeats across the corresponding sets. Comma-separated values are paired by position. If a multi-value load list is shorter than the reps list, the current calculation repeats its final load. Use the individual-set fields and complete each pair to avoid ambiguous input.

## Personal records

After you enter an exercise and a valid load, the app shows your saved load personal record (PR), or indicates that you're establishing your first record. When the entered load meets or exceeds your load PR, it also shows the most reps ever recorded specifically at your maximum-ever load, if available. For isometric exercises, this is the longest hold in seconds at the maximum-ever load held. A heavier load establishes a new reps or hold-duration baseline at that weight.

A load PR is the greatest recorded load for an exercise. At that load, the rep record is the highest paired rep count across its sets; lighter sets cannot raise that record. Isometric hold records use the same pairing but retain raw seconds. Unknown reps at the maximum load stay unknown instead of borrowing reps from a lighter set.

The save confirmation announces new load, rep, or hold records. A new exercise with a valid load establishes an initial load PR. The digest's PR section currently tracks load improvements over an existing best only, excluding first-ever entries.

## Estimated 1RM and relative intensity

The app calculates estimated one-rep maximum using Epley and Brzycki and keeps the best estimate from the recorded sets for each formula. The dashboard shows both estimates. Brzycki is omitted for sets of 37 or more reps.

The formulas used are `load * (1 + reps / 30)` for Epley and `load * 36 / (37 - reps)` for Brzycki.

The live relative intensity indicator compares the entered set's load with the average of the historical best Epley and Brzycki estimates. Its intensity label also depends on reps. Isometric holds use one rep for this calculation; their actual seconds are retained separately for hold PRs.

## Workload ratio

The acute-to-chronic workload ratio (ACWR) is the last 7 days' volume divided by the average weekly volume over the last 28 days. The chronic weekly average is the 28-day total divided by four. The periods overlap, and no ratio is shown when the chronic average is zero.

The code does not require 28 complete days of history before displaying a ratio. A new or incomplete log can therefore produce a distorted result.

## Trends and plateau logic

Week-over-week trends compare only exercises logged in both 7-day periods. They compare total volume and average representative load, where each row's representative load is its heaviest weight. Changes within +/-10% are labeled plateaued.

The separate plateau alert counts logged exercise entries since the all-time maximum load was first achieved. Matching that load does not reset the count. The defaults require at least four subsequent entries without exceeding it, at least two entries overall, and at least 300 kg of total volume. These thresholds are configurable in `Code.gs`.

## Workout density

Density is that training date's total volume divided by the timed session duration in minutes. The digest reports the arithmetic average of the qualifying session densities in its reporting period.

Session timing starts when the app opens (or resumes the start time from a restored draft), resets after a successful save, and ends at the next save. Idle time affects density. Sessions shorter than one minute are omitted. If you log multiple sessions on one date, each timing entry currently uses that date's total volume, which can inflate density. Historical imports without timing records do not produce density entries.

## Muscle groups, counts, and edge cases

The exercise summary uses each row's heaviest load and average reps, so those columns are not the precise reps-at-maximum-load record shown in the app. Logged counts represent exercise rows, while rest-day patterns count distinct training dates.

Combined Focus labels use the first listed muscle segment for classification rather than splitting volume across groups. Unrecognized labels remain separate groups. Muscle Group Gaps includes groups found in your history and caps displayed gaps at 15 days; it does not infer groups you have never logged.

Dashboard windows are recalculated when the dashboard is rebuilt, rather than continuously as the date changes. Missing loads, inconsistent exercise names, and incomplete per-set entries affect the resulting summaries.

Estimated 1RM, workload ratios, and plateau labels are summaries of your logs. Short or incomplete histories can distort them; use them to review trends rather than as precise measures of capacity or recovery.
