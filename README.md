# Peer Evaluation Google Apps Script App

This version adds the requested dashboard-driven defaults and teacher overrides.

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

## Files

- `Code.gs`
- `Index.html`
- `Faculty.html`

## Setup

1. Create a new Google Sheet.
2. Open Extensions > Apps Script.
3. Replace `Code.gs` with this package's `Code.gs`.
4. Add `Index.html`.
5. Add `Faculty.html`.
6. Run `ensureSetup()` once.
7. Deploy as a web app.
8. Faculty portal: append `?view=faculty` to the web app URL.

## Important

This version changes the schema. Use a new Google Sheet or clear old tabs before installing.
