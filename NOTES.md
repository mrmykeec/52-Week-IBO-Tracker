# 52-Week IBO Tracker: Decision Log

A working notes file for future-Mike. Covers the *why* behind decisions that aren't obvious from the code.

## What this is
A single-file HTML tracker for Jolene & Jake Carrithers' 52-week IBO training program. Originally built as a personal gift for Jackie, then expanded for sharing with the WWG team.

## Core architecture decisions

### Why a single HTML file
- Zero install. Zero accounts. Zero hosting costs.
- Works offline. Survives anywhere a browser exists.
- Trivial to share by email or upload to Drive.
- The trade-off: no cross-device sync, but Export/Import bridges that.

### Why localStorage instead of a backend
- Privacy: no data ever leaves the user's device.
- No user accounts to manage.
- No infrastructure to maintain over time.
- Trade-off: clearing browser data wipes progress. Mitigated by Export Backup.

### Why CSS variables for theming
- Lets the user swap palettes instantly without re-rendering anything.
- 6 themes × 2-mode = 12 visual combinations from one CSS file.

### Why no analytics
- Privacy-first. Plus IBO compliance is sensitive about data on customers/prospects.
- Trade-off: no feedback loop on actual usage. Have to ask Jackie directly.

## Key UX decisions

### Today vs. Month view
- **Today** = "what do I do today?" Focused, low-overwhelm.
- **Month** = "where am I in the program?" For review and planning.
- Default opens to Today on first launch. Persisted across reloads.

### Why Today doesn't auto-advance to the next week
- Earlier versions skipped Week 1 → Week 2 the moment Week 1 hit 100%, robbing the user of the celebration moment.
- Now: explicit "Start Week N+1" button. User decides when to advance.
- `state.todayWeek` anchors the user's current week so completion celebrations actually fire.

### Why JOT (Just One Thing) picks a random task with smart bias
- Combats decision paralysis on overwhelmed days.
- Smart bias: prefers Learning > Habit > Networking > Event (lightest first).
- One pick per day persists; reroll button if she doesn't want what was picked.

### Why notes are collapsed-with-preview in Today, always-visible in Month
- Today is meant to be quick. Notes are a thoughtful activity.
- Month is for reflection, so notes belong always-visible there.
- Same `state.weekNotes[N]` data backs both. Editing in one view shows in the other.

### Why ranks use IBO pin tiers (Eagle / Double Eagle / Platinum / Emerald / Diamond)
- Cultural alignment with the program's actual milestone language.
- Framed as "training milestones reached" not the actual business pins (legitimacy).
- 5 tiers at 20/40/60/80/100% gives more frequent celebrations than the original 25/50/75/100.

### Why streak counts calendar days, not 24-hour windows
- Industry-standard for habit trackers (Duolingo, Habitica, Streaks app).
- Simpler mental model: "did I do anything today?"
- `state.bestStreak` preserves the lifetime record so a missed day doesn't feel like everything's lost.

### Why counter tasks (5 MAs + 2 DMs etc.)
- Some training tasks are quantitative. Binary checkbox doesn't capture progress.
- `+/-` buttons let her track partial progress mid-week.
- Tapping the checkbox auto-fills counters to target (shortcut to "I did all 5").

### Why Reset Progress also clears program start date
- A fresh start should feel fresh. Half-resets are confusing.
- Theme/mode preferences are kept (visual settings shouldn't be data).

## Design constraints (don't change without thinking)

- **Carrithers attribution stays** in the footer. They designed the curriculum.
- **Mike's dedication stays.** Two places: footer ("Made with 💜 for Jackie by Mike Carrizales") is the neutral team-safe version. The final tour card carries the personal version ("Made with 💜 for Jackie, who deserves every dream she's chasing.") so it appears only on first launch and replays. Don't merge them.
- **The 4 task categories are fixed**: Learning / Daily Habit / Networking / Event.
- **Storage keys bump with major versions**: `iboTracker_v6` → `iboTracker_v7` → etc. Forces a clean state for new users; prevents schema drift bugs.
- **No external dependencies in v8+**. All CSS, fonts, scripts are inlined.
- **No em dashes** in user-facing text (Mike's writing voice rule). Applies to in-app copy and to this notes file.

## Version history

| Ver | Theme | Highlights |
|-----|-------|-----------|
| v1  | Initial | Simple HTML checklist, one-tab view, basic checkbox + state |
| v2  | Spreadsheet | xlsx generator (later abandoned) + html with month grid |
| v3  | Themes & polish | 6 themes, light/dark/auto, accessibility pass, mobile improvements |
| v4  | Tracking depth | Program start date, week dates, pace status, counters, per-week notes, export/import |
| v5  | Today/Month split | Today view + JOT button + monthly notes + rank tiers + streak |
| v6  | Polish | Various bug fixes, accessibility, "Set program start date" date picker, Pro tip from Jolene |
| v7  | Bug review | Counter cleanup on import, validate todayWeek, keyboard accessibility for headers, aria-live celebrations |
| v8  | Self-contained | Removed Tailwind & Google Fonts CDN, inlined utilities, system font stack, fully offline-capable for sharing |
| v9  | Celebration polish | Month-complete celebration upgrade (last week + all goals), "Set as Today" jump-back button on Month view, calendar mode no longer skips past unfinished earlier weeks, fixed wrong-month goals card refresh in Today view, default mode set to Auto, week celebration moved below task list |
| v10 | Search & a11y polish | Search feature (🔍 icon top-right of hero) with debounced filter, highlighted matches via `<mark>`, results grouped by week, "Set as Today's week" jump-to-week button. Linkify helper extended to monthly goals (URLs in goal text now clickable). Accessibility improvements: aria-live on search status, focus restored to trigger on close, larger 40×40 tap target on search-close, contextual aria-labels on "Set as Today" buttons, iOS Safari zoom fix on search input (font-size 16px). Removed wwdb.com from linkify map (web search returns no results for audio numbers, app is the real source of truth). |
| v11 | Onboarding tour + search-result jump | Internal dev snapshot. Build-up version that introduced the welcome tour and search-result jump but never went out to users. Superseded by v12. |
| v12 | First broad team release | Onboarding consolidation, simplification pass, and the version actually shipped to the WWG team. Welcome banner removed; its content (set start date, Pro tip from Jolene, "data lives here" reminder) folded into the welcome tour so onboarding is one consolidated experience. Tour: 8 steps + final dedication card (Welcome → First, set your start date → Your week → JOT → Notes → Streak/Rank → From Jolene → Add to Home Screen → final). "From Jolene" sits late so the tour pivots from feature tour to philosophy → install → goodbye, and so the three centered steps cluster cleanly at the end. Tour skips the start-date step on replay if the date is already set. Skip button, Esc/arrow-key support, dot indicators, focus trap (Tab cycles Skip/Back/Next), Enter-on-button passes to native click, replayable from data-tools row ("❓ Show tour", now first button in the row). Spotlight uses CSS clip-path with reduced-motion fallback; `card.focus({preventScroll:true})` keeps the viewport from jumping. Reset Progress also clears the tour-seen flag and replays the tour. Search-result jump: each search hit gets a ↗ button that closes search, jumps to the matched week in Today view, scrolls the task into center, and pulses it briefly. Removed Pause feature: JOT + lifetime-best streak already cover the "hard week" case Pause was built for. Data-tools order: Show tour → Export Backup → Import Backup → Reset Progress. "Track the journey" tour step lists every rank with emoji + name + percentage. "Start Week N+1" now scrolls back to top of page after advancing. Curriculum fix: Week 3 Howie Danzik Wrap-Up search number corrected to `1531`. Install instructions added to tour with icon-replacement guidance for Android users. |

## Curriculum changes

These edits diverge from the original PDF curriculum. Both were sourced from Jolene's posts in the team Topics thread about the 52-week training. Worth re-checking against future updates from her.

| Week | Change | Source |
|------|--------|--------|
| Week 1 | Removed "Watch: Basics of the Business" task entirely | Jolene's Topics post in the team training thread |
| Week 3 | Updated "Howie Danzik Wrap-Up" search number from `1376` → `1531` | Jolene's Topics post in the team training thread (corrected 2026-05-03 from an earlier typo of `151`) |

If Jolene posts further curriculum updates, add the change here with the source thread/date so future-Mike has a clean trail back to authority.

## Open questions / future ideas

- **Voice memo dedication** from Mike on tap-the-heart Easter egg (would add a bit of personality)
- **Weekly summary email** (would require backend)
- **Cloud sync** (would require backend; explicitly avoided to date)
- **Multi-user "team view"** (way out of scope for single-file)
- **Auto-prompt to export backup** every N days (concern: nagging)
- **More easter eggs.** Five total feels like the right ceiling.
- **PWA manifest + service worker** for true install/offline (sketched in v12 ideas)
- **Activity heatmap, three-state tasks, per-task notes** also queued for v12+

## Files in this folder

- **`52-week-tracker.html`**: **the file to share with the team.** User-friendly filename, identical contents to the latest dev version.
- `tracker-v12.html`: current dev version (source of truth for `52-week-tracker.html`)
- `tracker.html`: v1 baseline (legacy)
- `tracker-v2.html` through `tracker-v11.html`: version history (legacy, kept for diffing/reference)
- `_test_linkify.js`, `_test_linkify.py`: link regex smoke tests (run when curriculum URLs change)
- `tracker.xlsx`: abandoned spreadsheet output
- `52-Week Training.pdf`: original source curriculum (private, do not redistribute)
- `build_tracker.py`: abandoned xlsx generator
- `NOTES.md`: this file

### Filename convention

Going forward:
- Dev work happens on `tracker-vN.html` (versioned, used for diffing and rollback).
- When a version is ready to ship, copy it over `52-week-tracker.html`:
  ```
  cp tracker-v13.html 52-week-tracker.html
  ```
- Always share `52-week-tracker.html` with the team. The version-numbered filenames are internal.
- Internal version markers (`BACKUP_VERSION`, storage keys) keep bumping with each `vN` so backups can warn about mismatches and new users get a clean state.

## Permissions / sharing notes

The curriculum content embedded in this file belongs to Jake & Jolene Carrithers. Always check with them before:
- Sharing publicly (GitHub, blog, social media)
- Sharing outside the immediate team
- Posting in any searchable archive

Sharing within the WWG team (downline, sponsors, sister teams) is generally OK but a heads-up to the Carrithers is appropriate.

## Bug-report channel

If a user finds a bug, they message Mike directly. No issue tracker, no support form. Keep it human.
