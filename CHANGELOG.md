# Changelog

## 2026-10-08

- Added optional Brand on each exercise entry, preserved through sessions, drafts, templates, and historical imports. Brand is column H in StrengthWorkoutINPUT and Templates; existing A:G positions remain intact. Setup refuses to overwrite occupied column H.
- Kept Brand separate from exercise identity, volume, PRs, progression, and hints. Historical names and brands are not automatically migrated.
- Moved Focus before Exercise. Autocomplete and typo suggestions use the existing dashboard exercise summary, filtered by selected Focus and an adjustable 30-day cutoff. Full-history canonical matching and hints remain available.
- Added Focus associations to the existing dashboard summary and normalized historical and newly saved Focus formatting. No duplicate catalog or audit tab is created.
- Fixed combined Muscle Group Gaps categories, including Legs recognition. Each listed muscle group receives its own last-trained date; volume balance keeps its first-listed-group rule.
- Added persistent bodyweight exercise settings, optional weigh-ins, smoothed effective-load reporting, signed-load PR handling, settings-aware name merges, and explicit per-set validation. Raw workout loads remain intact.
- Verified 38 regression tests covering storage, compatibility, Focus, Brand, assisted loads, validation, and PR behavior.

Deployment: update both Code.gs and Index.html in Apps Script and deploy a new version. Run Set Up Brand Columns or open the app; rebuild the dashboard after manual history edits. Pushing Git does not update Apps Script. No Deltas tab or automatic brand/name migration is included.
