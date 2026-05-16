# Term IV Class Schedule → Google Calendar

A single static page that lets each MBA student (IIM Udaipur) subscribe to their
personalised set of Google Calendars for the Term IV timetable, with minimum
friction: **search your name → confirm your courses → one button → guided sequence.**

No backend. No login. Everything runs in the browser. Hostable on GitHub Pages.

> **Note:** These calendars are **not dynamic**. If MBASecondYear reschedules
> classes in the Excel timetable, that change will not propagate to the calendars
> automatically — they reflect the schedule as published.

```
index.html              the whole app (HTML + CSS + vanilla JS, single file)
calendars.json          38 calendars: key -> { name, calId }   (generated)
enrollments.json        367 students: roll -> { name, sections[] } (generated)
calendars.template.csv  course_code,display_name,calendar_id  (source of names)
build-data.py           regenerates the two .json files
```

`calendars.json` and `enrollments.json` are committed so the site works as pure
static files. Regenerate them with `build-data.py` whenever data changes.

---

## How the data is wired

- Each course-section has a stable **key** derived from its course code:
  `B2B M (A)` → `B2B_M_SecA`, `HRM(IR)` → `HRM_IR`, `CV` → `CV`, `Rev Mgmt` → `Rev_Mgmt`.
- `calendars.json` maps **key → { name, calId }** (the 38 Google Calendar IDs).
- `enrollments.json` maps **roll number → { name, sections: [key, …] }**.
- The two files share the same keys, so they always stay consistent because the
  same script produces both.

### The subscribe link

The calendars are Google Workspace group calendars
(`c_…@group.calendar.google.com`). The page builds the add-link as:

```
https://calendar.google.com/calendar/u/0/r?cid=<base64(calId)>
```

base64-encoding the calendar id is the form that subscribes reliably for this
type of calendar. If you ever need the raw form instead, set
`CID_BASE64 = false` near the top of the `<script>` in `index.html`.

---

## Updating the Calendar IDs

The 38 IDs live in **one place**: the `CALENDAR_IDS` dict at the top of
`build-data.py`. To change one:

1. In Google Calendar → the calendar’s **Settings** → **Integrate calendar** →
   copy the **Calendar ID** (looks like `c_abc123@group.calendar.google.com`).
2. Paste it next to the right key in `CALENDAR_IDS`.
3. `python3 build-data.py "/path/to/Enrollment details"`
4. Commit the updated `calendars.json`.

---

## Regenerating enrollments next term (for ACAD)

Each term ACAD already produces one spreadsheet per course-section. Keep that
workflow:

1. Put the 38 `.xlsx` files in a folder. **Each file name must equal the
   `course_code`** in `calendars.template.csv` (e.g. `B2B M (A).xlsx`,
   `CV.xlsx`, `HRM(IR).xlsx`). Each sheet needs a header row containing
   **`Roll No.`** and **`Name`** columns.
2. If course offerings changed, edit `calendars.template.csv`
   (`course_code,display_name,calendar_id`) and the `CALENDAR_IDS` dict.
3. Run:

   ```bash
   pip install openpyxl
   python3 build-data.py "/path/to/Enrollment details"
   ```

   It prints a summary and rewrites `calendars.json` + `enrollments.json`.
   It **stops with an error** if any course is missing a calendar ID or
   `.xlsx` file, so you can’t deploy a half-built dataset.
4. Commit and push the two regenerated `.json` files.

The script reads only `Roll No.` and `Name`; email/other columns are ignored.

---

## Deploy to GitHub Pages

1. Create a repo (e.g. `term4-calendar`) and add these files at the **root**:
   `index.html`, `calendars.json`, `enrollments.json`
   (`build-data.py`, `calendars.template.csv`, `README.md` are optional but
   recommended to keep in the repo).
2. Push to `main`.
3. Repo → **Settings → Pages** → Source: **Deploy from a branch** →
   Branch: `main` / folder: `/ (root)` → **Save**.
4. The site goes live at `https://<org-or-user>.github.io/<repo>/` in ~1 min.

The page fetches `calendars.json` / `enrollments.json` over HTTP from the same
folder, which GitHub Pages serves automatically. (Opening `index.html` directly
from disk with `file://` will fail the fetch — use Pages, or a local server:
`python3 -m http.server` then open `http://localhost:8000`.)

---

## What a student sees

1. **Search** — type roll number or name; suggest-as-you-type. A name shared by
   several students shows all of them; an unknown roll shows *“No match found.”*
2. **Confirm** — greeting + a clean dash-bulleted list of their courses
   (display only, no per-course buttons).
3. **One button — “Add all to Google Calendar.”** It opens the first calendar in
   a new tab. The student taps **Add** in Google, returns to the page, and the
   next one opens automatically; each added course gets ticked off.
   - Returning focus is detected automatically. If a browser swallows that
     (some mobile browsers) or blocks the pop-up, a reliable **“open next →”**
     button is always shown as the manual fallback.
   - Re-adding an already-added calendar is harmless — Google de-dupes it.
   - A tip reminds phone users to be signed in to Google Calendar first.
4. **Done** — *“All N calendars added ✓”* plus a reminder that the calendars are
   not dynamic and won’t reflect later reschedules automatically.

Collapsible footer: a “How does this work?” explainer (including the
not-dynamic note), a **browse-all-38** list with individual Add buttons
(audits / friends), and the privacy note.

> Note: the current dataset has **367 students** across **38 calendars**
> (5–9 courses each). Anyone with the link can look up anyone — accepted by
> design; no authentication.
