# Peer Evaluation Google Apps Script App

This version is optimized for simple setup by less technical users.

## Files to copy into Apps Script

You only need three files:

- `Code.gs` - all backend logic in one file
- `Index.html` - student registration and evaluation portal
- `Faculty.html` - faculty dashboard and reports

No extra `.gs` helper files are required.

## Default-first behavior

Students may enter an Activity Key, Class ID, and Group Key directly. If the Activity/Class/Group does not exist yet, the backend automatically creates it with default settings:

- Status: Active
- Student registration: Allowed
- Registration deadline: none
- Evaluation deadline: none
- Rubric: default weighted rubric

The teacher only needs to change settings in the Faculty Dashboard when they want to override those defaults.

## Added features

- Student-entered activities with automatic activity creation
- Activities sheet
- Rubrics sheet
- Configurable rubric criteria per Activity + Class
- Weighted grading
- Registration and evaluation open/close windows
- Faculty activity/deadline dashboard
- Faculty correction tools:
  - move student to another group
  - update contribution
  - deactivate/reactivate registration
- Dashboard cards for incomplete evaluators, single-member groups, cohorts with no submissions, and totals
- Dynamic student star ratings from the active rubric
- Accessible star buttons with aria labels, radio roles, and keyboard arrow support
- Improved scoring storage using raw score, max raw score, weighted score, weighted percent, and Scores JSON

## Setup

1. Create a new Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the default `Code.gs` with this repo's `Code.gs`.
4. Add an HTML file named `Index` and paste `Index.html`.
5. Add an HTML file named `Faculty` and paste `Faculty.html`.
6. Run `ensureSetup()` once from the Apps Script editor and approve permissions.
7. Deploy as a web app.
8. Student portal: open the main web app URL.
9. Faculty portal: append `?view=faculty` to the web app URL.

## Important

This version changes the schema. Use a new Google Sheet or clear old tabs before installing.

Default Faculty PIN is currently `Admin2026`. For production, run `changeFacultyPin('Admin2026', 'YourNewSecurePIN')` or set `FACULTY_PIN` in Apps Script Properties.
