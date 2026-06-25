# Google Apps Script Peer Evaluation App

This package is intended for a **new Google Sheet / new Apps Script deployment**. It uses a fresh schema and does not include backward-compatibility handling for older column layouts.

Files included:

- `Code.gs` — backend, spreadsheet setup, activity-aware registrations, validation, duplicate blocking, faculty unlock, and structured report data.
- `Index.html` — student registration and evaluation portal with Activity/Class/Group input, contribution display, and 1–10 star ratings.
- `Faculty.html` — password-gated faculty report portal with dropdown filters, print, and CSV export.

## Main features

1. Students register under an `Activity Key`, `Class ID`, and `Group Key`.
2. The same student may register again for a different activity or a different class.
3. Duplicate registration is blocked only for the same `Activity Key + Class ID + Student ID`.
4. Evaluations are isolated by `Activity Key + Class ID + Group Key`.
5. Students evaluate teammates with 1–10 star ratings.
6. When selecting a teammate, the evaluator sees that teammate's role and contribution.
7. Students cannot evaluate themselves.
8. Duplicate peer evaluations are blocked server-side.
9. Faculty codes are hidden until the faculty PIN is accepted.
10. After unlocking, faculty can select existing codes from dropdowns.
11. Faculty reports support these modes:
    - Class + Group across all activities
    - Activity + Class across all groups
    - Exact Activity + Class + Group
    - Whole class across all activities and groups
12. Reports include cohort breakdown, activity breakdown, group breakdown, grade summaries, completion tracking, missing evaluations, detailed logs, roster, print, and CSV export.

## Setup

1. Create a new Google Sheet.
2. Open Extensions > Apps Script.
3. Replace `Code.gs` with the provided `Code.gs`.
4. Add or replace `Index.html` with the provided `Index.html`.
5. Add or replace `Faculty.html` with the provided `Faculty.html`.
6. Save the project.
7. Run `ensureSetup()` once from the Apps Script editor and approve permissions.
8. Deploy as a Web App.
9. Student portal: open the main web app URL.
10. Faculty portal: open the web app URL with `?view=faculty` appended.

## Activity workflow

For each activity, the teacher should provide students with:

- Activity Key / Name, for example `ACTIVITY-1`, `LAB2`, `CAPSTONE`, or `MIDTERM`
- Class ID / Class Number, for example `CS101` or `BSIT-2A`
- Group Key, for example `1A`, `GROUP-3`, or `ALPHA`

A student can register for:

- `ACTIVITY-1 + CS101`
- `ACTIVITY-2 + CS101`
- `ACTIVITY-1 + CS102`

The same Student ID cannot register twice for the exact same Activity + Class.

## Faculty report examples

After entering the PIN, the faculty portal shows dropdowns only for existing codes.

Use **Class + Group: all activities** when you want:

- Class: `CS101`
- Group: `1A`
- Result: all activities for that class/group.

Use **Activity + Class: all groups** when you want:

- Activity: `ACTIVITY-1`
- Class: `CS101`
- Result: all groups in that activity/class.

Use **Activity + Class + Group** for one exact cohort.

Use **Whole class** for all activities and all groups under one class.

## Faculty PIN

Default PIN is currently `Admin2026` in `Code.gs`. For better security, change it by running:

```javascript
changeFacultyPin('Admin2026', 'YourNewSecurePIN')
```

A stronger option is to store `FACULTY_PIN` in Apps Script Project Settings > Script Properties.

## Important security note

This app blocks duplicate evaluations and validates group membership on the server, but Student ID alone is not true identity authentication. For high-stakes grading, deploy within a Google Workspace domain and add email-based checks or per-student tokens.
