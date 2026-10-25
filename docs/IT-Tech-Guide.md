# Branthaven IT Help Desk – Technician Guide

Site: https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2
Business day for SLA purposes: the hourly clock (HD6, and the BusinessHoursOpen/SLAHours it writes) runs **8am–9pm**, every day. Response-time credit ("Response (min)") is more generous — it counts from 7am (support opens) regardless of when the ticket was actually raised; see the note under "Service levels" below.

## How tickets arrive

| Source | What happens |
|---|---|
| Email to the help desk mailbox | HD1 creates a ticket. A reply containing an [IT-123] tag updates that ticket. Branthaven senders always get a ticket; outside senders only if their address or domain is on the **External Senders** list (see below). Everything else lands in the mailbox's **Review** folder and is ignored. |
| IT Support form (Microsoft Forms) | HD2 creates a ticket. |
| **Report an IT issue** Teams app | HD14 creates a ticket (source: "Teams form"). |

Every new ticket then runs HD3: it sets permissions (the requester sees only their own ticket), creates the **Planner task**, emails and messages the requester a confirmation (as IT Support), alerts IT in the Help Desk channel, and sends Martin and Joseph a direct Teams message from IT Support with the ticket number, requester, priority and link.

Every HD4 run (any note, status change, or resolution) also rewrites the Planner task's own **Notes** to the ticket title, a link, and the ticket's conversation history (capped at ~1800 characters) — so a tech reading the task in Planner sees the same notes and resolution text as the ticket itself, without opening SharePoint. It's a full rebuild each time, not an append, so nothing you type directly into the Planner task's notes will survive the next ticket update — put any notes on the ticket (Edit ticket / note box) instead.

**Attaching a file to a ticket:** attaching it doesn't notify the requester by itself — attaching a file doesn't change Status, IT Assigned, or Note to requester, and only those trigger HD4. Put a line in **Note to requester** (even just "see attached file") so HD4 fires. When it does, HD4 checks the ticket's attachments fresh each run and, if any haven't been linked before, adds a direct download link for each one right in the email and Teams message — the requester doesn't need to open the ticket just to find the file.

Only *new* attachments get linked, not every file ever added to the ticket — a **NotifiedAttachments0** field (pipe-separated filenames, kept off every form and view) tracks what's already gone out, so the requester's own originally-submitted file doesn't get re-linked on every later update. It only advances when a note/status notification actually sends — a reassignment-only edit with no note won't mark a pending file as "seen." (It's `NotifiedAttachments0`, not `NotifiedAttachments` — see `CLAUDE.md` for why.)

The one exception: the status email/Teams message sent when a ticket reaches **Resolved Pending Confirmation** or **Closed** lists *every* attachment on the ticket — whatever the requester originally sent plus whatever a tech added along the way — not just the ones not yet linked, since that's the message most likely to get kept or forwarded. Every other note or status change still only links new files.

## Your daily workflow

1. Watch the **Help Desk channel** and the **Planner board** (buckets: New, In progress, Waiting, Done).
2. **Claim a ticket**: on the card, **Start work** if you are starting now (assigns it to you and sets *In Progress* in one update — one message to the requester), or **Assign to me** if you will get to it later (status *Assigned*; use Start work when you begin).
3. **Work it**: keep Status accurate – it drives the SLA clock, the Planner bucket and the messages the requester receives.
4. **Change priority, category or who it's assigned to**: use **Edit ticket** on the card. Pick only what you want to change and leave the rest on *(keep current)*. Reassigning notifies the new technician and the requester; priority/category changes are noted in the card's thread.
5. **Fixed it on the spot?** Use **Quick close** on the card: type the resolution, Send. The ticket is assigned to you and closed in one step and the requester gets one message ("Resolved" with your resolution) instead of three or four. Use it for walk-ups, phone fixes and password resets; for anything the user needs to try first, use Resolved Pending Confirmation instead. The same button exists in the **SharePoint list** (the ✓ icon next to the ticket number, hidden once a ticket is Closed): SharePoint asks for the resolution, then HD18 assigns the ticket to you (unless it already has a tech) and closes it in one update — the requester gets the same single "Resolved" message. The other icons in that column are Edit, Assign to me and Start work; their definition is `dashboard/tickets-id-column.json` (change it there, then paste into the ID column's Format this column, never only in SharePoint). HD18 must list every tech under **Run-only users** (flow details page), with both connections set to *Use this connection (IT_Support)* — otherwise the button shows "cannot be used to activate this flow". Add new techs there when they join. SharePoint always shows a confirm panel with a **Run flow** button before a column-button flow actually runs (same as Microsoft's own native "Flow" command — there's no way to skip it or make it submit on Enter); HD18's flow **Description** (Power Automate → flow → Details → Edit) is set to a one-line reminder to click Run flow, which SharePoint shows under the owner line in that panel — change it there (never only on a copy of the flow) if the wording needs updating.
   If you resolve or close a ticket by editing it directly and forget to assign it, HD4 assigns it to you (the editor) before sending anything, so the requester's email never says "Not yet assigned" for a ticket a tech just fixed.
   Quick close (Teams or SharePoint) assigns and closes the ticket in one update, so HD4 would normally both assign you in Planner (a Planner "assigned to you" notification) and mark the task complete in that same run; since Oct 2026 HD4 skips the Planner assignment step whenever the ticket is landing on Resolved/Closed in the same update, so Quick close no longer generates an assignment notification for a task that's already done — the Planner task still moves to the Done bucket and 100%. A tech assigned to a ticket that stays open (not via Quick close) still gets the normal Planner assignment notification.
6. **Ask the requester a question**: use the card/reply action on the ticket thread in Teams. The question reaches them as an IT Support message and their answer (by Teams or email) is posted back on the ticket.
7. **Resolve**: set Status to *Resolved Pending Confirmation* (or complete the Planner task), with a resolution written in the message box. The requester is notified, including your resolution text.
8. The system **closes it automatically after 48 hours**, with a reminder to the requester at 24 hours; the automatic close itself is silent (the resolved email already told them it would happen). Set *Closed* yourself if they confirm earlier — **a resolution is required to close a ticket**: if none was captured yet (e.g. closing directly without going through Resolved Pending Confirmation), enter it in the message box when you set Status to Closed, or the flow will bounce it back to you and leave the ticket open. This is enforced three ways: the Teams card blocks it outright, if a ticket is set to Closed by any other means (e.g. editing the list item directly) with no resolution on file HD4 reverts it to its prior status and messages you to add one, and completing the ticket's **Planner task** only moves the ticket to Resolved Pending Confirmation if a resolution is already on file — otherwise HD5 reopens the task and messages you to add the resolution first.

## Logging a ticket for someone else

In the **Report an IT issue** app, fill in **On behalf of** with the person's Branthaven email address (e.g. jane@branthaven.com). The ticket then belongs to that person: they get the confirmation and every update, and the ticket shows *Logged by <you>*. It is **not** assigned to you. It lands unassigned like any other ticket, so the team can pick it up. Leave the box empty to log a ticket for yourself. Only @branthaven.com addresses are accepted.

## External senders (vendors, trades, consultants)

Outside addresses can raise and work tickets **by email only**. Add them to the **External Senders** list on the help desk site: one row per entry, **Title** = either a whole domain (`acmeplumbing.ca`) or one address (`bob.smith@gmail.com`); case doesn't matter. Remove the row to cut them off. Changes take effect on the next email — no flow change needed.

What's different for an external ticket:

- The Conversation log starts with "Ticket created from Email (external sender)". The Requester person field is empty; **Requester Email** holds their address.
- They get the same confirmation, status and reminder **emails**, without the View-ticket / portal links — they have no access to the site. They never get Teams messages.
- They reply by email (keeping IT-<n> in the subject) and that's posted to the ticket exactly like an internal reply. Techs work the ticket the same way.
- Anything from an address that isn't on the list stays in the **Review** folder. Glance at it now and then; a legitimate request from an unknown vendor means "add them to the list and forward the email to the mailbox", not a manual ticket.

## Status guide

| Status | Use when | SLA clock |
|---|---|---|
| Open | New, unassigned | Running |
| Assigned | Given to a tech | Running |
| In Progress | You are working it | Running |
| Escalated or Waiting Internally | Waiting on another internal team | Running |
| Waiting for End User | You need info from the requester | Paused |
| To Be Scheduled | Booking a time with the requester | Paused |
| Resolved Pending Confirmation | Fix delivered, awaiting confirmation | Stopped |
| Closed | Finished (requesters see this status as **Resolved** in their emails and Teams messages) | Stopped |

Status changes automatically notify the requester (email + Teams) and move the Planner card.

## Per-tech service levels

The **By IT Support (30 days)** view of the Tickets list groups the last 30 days of tickets by technician. Each ticket shows **Response (min)** (from the ticket's effective start — 7am for anything raised outside 7am–9pm — to the first IT action), **Response met** (within 4 h) and **Resolved in 24 h** (24 SLA-clock hours). The totals row under each group gives that tech's averages; the bottom row is the team. The same numbers, laid out properly, are on the **IT Dashboard → Service levels by technician** panel (HD15, refreshed every 5 minutes). HD4 writes the three figures once per ticket; they are never recalculated, so a reassignment does not move the credit.

## Clock, reports and reminders

- **HD6 (every day, 8am–9pm):** counts business hours open, handles the 24h reminder and 48h auto-close for resolved tickets. A ticket unassigned for 4 hours is posted in the Help Desk channel; at 5 hours (once only) Martin and Joseph (first-line) both get a direct Teams ping. Every full business day it stays open after that (24h, 48h, ...) sends a "still open" reminder — while it's still unassigned, that goes to Martin and Joseph (same audience as the 5-hour ping, just repeating daily); once a tech is assigned, it goes to that tech alone.
- **HD8 (9pm daily):** daily report of overdue tickets (open longer than a business day) and unassigned tickets (4+ hours), emailed to Martin.
- **HD7:** a requester can ask for "My open tickets" and gets a list.

## Monitoring (is the system healthy?)

**Look first at the IT Dashboard → "System health" panel** (top of https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2/SitePages/IT-Dashboard.aspx). Seven cards, each green OK / amber CHECK / red PROBLEM / grey – (could not check), refreshed every hour at :30, around the clock, by **HD16 Health check v2**. Under the cards, one line per problem (ticket numbers only; detail is in the daily email). If HD16 itself stops, the whole panel fades and the headline pill turns grey **"STALE — health check has not run since …"** (the view compares `Slot22`, the time of the check as epoch milliseconds, with now; 90 min grace).

| Card | What it checks | Red means |
|---|---|---|
| Connections | A read-only call through every connector each hour: Outlook (mailbox), Teams (channels), Teams as IT Support (joined teams), Planner (buckets), Office 365 Users, OneDrive, Power Automate Management; SharePoint is exercised by the ticket queries. Microsoft Forms has no read action and is **not** probed. | Names the broken connection — Power Automate › Connections › Reconnect, or recreate it (see "Emails stop but Teams messages still arrive" below) |
| Flows | Power Automate Management "List My Flows": every expected flow (HD1–HD9, HD11, HD12, HD15, HD16 v2) is **on, exactly once**. HD13/HD17 (run-once) are not expected on; HD14 is owned by Martin's account so it is not visible to this check. HD19/HD20/HD21 (IT Projects) are checked separately — see "IT Project Systems" below — and deliberately **not** folded into this card, to avoid double-reporting and to keep this card's "14 flows on, one copy each" meaning unchanged. | "HDx is off or missing" / "HDx has 2 copies on" — turn the right one on, delete the duplicate |
| New tickets set up | Tickets created in the last 24 h have a Planner task, a channel post and the "Ticket created" line | HD3 is not finishing — open its run history |
| Updates notified (HD4) | No ticket whose Status differs from LastNotifiedStatus for more than 30 min (the requester was never told) | HD4 is not running — check its run history and the SharePoint connection |
| Email intake (HD1) | No mail older than 20 min still sitting in the IT Support **Inbox** (HD1 moves everything to Processed or Review; Microsoft/Power Automate system mail ignored); amber if mail landed in **Review** in the last 24 h (system mail ignored) | HD1 is not ticketing email |
| Clock & auto-close (HD6) | An open ticket has been touched in the last 3 h (checked 11:00–21:59 only, since HD6 runs 8–21); nothing resolved more than 72 h ago is still "Resolved Pending Confirmation" | HD6 stopped — SLA clock, reminders and auto-close are all off |
| Dashboard data (HD15) | The stats rows behind the charts were refreshed in the last 30 min (checked 7am–9pm) | HD15 stopped — the charts below are stale |

Problems reported by the HD3/HD4 self-checks in the last 24 h are added to the band under the cards (the step name is in the alert email).

Behind the panel, three layers of alerts all go to Martin (mpersaud@branthaven.com), each **by email and as a Teams message from IT Support**:

1. **Per-ticket, immediately.** HD3 (ticket created) and HD4 (ticket updated) check their own work at the end of every run — the named notification steps directly (confirmation email, confirmation chat, Planner task, channel post; status email, status chat, note email/chat, new-tech chat, Planner move), not just the scope summaries, so a failure nested inside a condition is caught. If any did not succeed, Martin gets an email and a Teams message naming the ticket, the failed step, the error text and a link to the flow run. The ticket itself is still created/updated — only the notification is reported.
2. **HD16 health check, hourly, 24/7.** Updates the panel every run; emails Martin a report once a day at 9:30 am (subject "[Help desk health] All clear" or "[Help desk health] PROBLEMS: n"); emails and Teams-messages him whenever the issue count changes, **between 7 am and 9 pm only** — an overnight problem waits on the panel and is in the 7:30 am run.
3. **Hard flow failures.** Power Automate emails the flow owner (the IT Support mailbox) whenever a run fails outright, when a connection breaks, and in a weekly digest. HD1 recognises those (sender @microsoft.com, subject mentioning a flow, a failure or a connection that isn't working), relays each one to Martin by email ("[Help desk health] <original subject>") and Teams, and parks the original in the Review folder.

If an alert names a flow run, open the link: the failed step is highlighted in red with the full error.

The panel's rendering is the JSON view formatting on the **Live overview** list's *Chart - Health* view (filtered to item 14). The source of truth is `dashboard/health-view.json` in the repo: paste it into the view via *Format current view → Advanced mode* whenever it changes. Column map: card n uses `Slot n` (name), `Slot n+14` (description) and `Share n` (0 OK / 50 amber / 100 red / −1 grey); `Total` = issue count, `Refreshed` = time shown, `Slot22` = time of the check in epoch ms, `Grad` = the problem band, `Slot10`/`Slot11` = last HD3/HD4 self-reported problem and its time.

### IT Project Systems panel

Below the ticket-systems health panel, the IT Dashboard page has a second, smaller panel — **"IT Project Systems"** — reporting just HD19/HD20/HD21. Same mechanism as the main panel (HD16 writes it, hourly), different DASH row (**16**, not 14), different view (**Chart - ProjHealth**, not Chart - Health), and no staleness dimming (it shares HD16's own health, so if HD16 stops the main panel's STALE pill already covers that). Three cards only, one per flow: badge **OK** / **PROBLEM** / grey **"Not checked yet"** (no amber "CHECK" state — a project flow is either on exactly once or it isn't).

HD16 computes this from the same `FlowStr` string it already builds for the main Flows card (see `EXPECTED_PROJECT_FLOWS`/`PROJECT_FLOW_LABELS`/`DASH_PROJHEALTH` in `generator/gen.py`) — no extra API calls. Column map on row 16: `Slot1`–`Slot3` = card names (HD19/HD20/HD21 labels), `Slot4`–`Slot6` = card descriptions ("on" / "off or missing" / "N copies on" / "could not list flows"), `Share1`–`Share3` = 0 OK / 100 red / −1 grey, `Total` = issue count (−1 if flows couldn't be listed), `Refreshed` = time shown, `Grad` = the problem band (semicolon-joined names of the flows that are off/duplicated). The source of truth is `dashboard/project-health-view.json`: paste it into the *Chart - ProjHealth* view the same way as the main panel.

## IT Projects (project tracking)

A separate list, **IT Projects** (same site, nav link next to Tickets), tracks longer-running work that isn't a single ticket — rollouts, infrastructure work, process changes. It has its own Status (`Not Started` / `In Progress` / `On Hold` / `Waiting on Vendor/Approval` / `Completed` / `Cancelled`) and Category (`Infrastructure` / `Security` / `Software Rollout` / `Network` / `Hardware Refresh` / `Process Improvement` / `Other`) choice columns — keep both lists in sync with `PRJ_STATUSES`/`PRJ_CATEGORIES` in `generator/gen.py` if either ever changes. "Active" for monitoring/staleness purposes means every status except Completed/Cancelled.

Project stats live on the same **IT Dashboard** page as ticket stats, styled the same way: a "Projects" section (Project by Status / Priority / Category donuts, DASH rows **11/12/13**, same `chart()` builder as the ticket donuts but reading from the IT Projects list and titled `ProjStatus`/`ProjPriority`/`ProjCategory` so they don't collide with the ticket rows' `Status`/`Priority`/`Category` titles) and a `Projects` tile in the KPI row (count of active projects), both refreshed by HD15 alongside the ticket stats. The "IT Project Systems" health bucket described above sits with the ticket-systems health panel, not here — it's about whether HD19/HD20/HD21 are running, not about the project data itself.

Three flows run it, the same pattern as HD3/HD4/HD6 for tickets:

- **HD19 Project created** — breaks permission inheritance on the new item (mirrors HD3): Owners get full control, and if a Sponsor is set they get Read. Emails and Teams-messages the creator, Martin, and the sponsor (if any). Stamps `LastNotifiedStatus`/`LastNotifiedSponsorLen`/`LastNotifiedInternalLen` so HD20 can tell what's already been told.
- **HD20 Project updated** — a project has two note logs: **SponsorUpdates** (sponsor-facing — a status change or a new sponsor note) and **InternalNotes** (techs-only). A status change or a new SponsorUpdates note notifies the creator, Martin, *and* the sponsor; a new InternalNotes-only note notifies Martin and the project Owner alone — the sponsor never sees internal notes. Each note field is a prepended log, same convention as the ticket Conversation field; HD20 diffs the field length against `LastNotifiedSponsorLen`/`LastNotifiedInternalLen` to find just the newly-added text (not the whole log) for the notification, then restamps both lengths and `LastNotifiedStatus` at the end of every run. Like HD4, the trigger itself carries a condition (status/sponsor/internal "changed" check) so a run that wouldn't notify anyone never starts — without it, the restamp at the end of every run would bump Modified and re-trigger itself forever (see the troubleshooting table).
- **HD21 No-update check** — runs Fridays at 9am; finds every *active* project not modified in 7+ days and emails/Teams-messages the creator and Martin a nudge. No self-check (same as HD6 — it's a scheduled sweep, not a per-item flow).

There's no sponsor-facing input yet — sponsors only *receive* email/Teams notifications; they have no form or portal to add their own update. That's planned as Phase 2 (a sponsor-reply flow and/or a Power Apps form, tentatively HD22/HD23) and hasn't been built.

## Reply to a closed ticket

If someone replies — by email or in the Teams chat — to a ticket that's already Closed, the system doesn't reopen it. It starts a **new ticket** instead, tagged "Follow-up to IT-<old number>" in Details, and (via Teams) titled "Follow-up: <old title>" with the same Category/Priority carried over. If the old ticket had a resolution on file, it's copied into the new ticket's Details too ("Previous resolution (IT-<old number>): ..."), so you can see what was already tried without flipping back to the old ticket. The old ticket stays closed as the record of what was done; work the new one like any other.

## Rules of the road

- Only the list owner edits the Tickets list structure; techs work tickets through the ticket form/Planner.
- Do not delete tickets. Close them.
- Don't email the requester separately for status changes; the system already does.
- Ticket numbers are IT-<ID>. Put them in any email subject about the ticket.

### Who can see a ticket (permissions)

- The **IT Support Tickets** list has **unique permissions** (it does not inherit from the site): only the site Owners group (IT Support). The site's Visitors group (BH Head Office) and Members must *not* be on the list, or every staff member can read every ticket.
- HD3 then breaks inheritance on each new ticket and grants Owners full control and the requester **Read** on their own item. Requesters update a ticket by replying (email "Reply to IT" / Teams chat), never by editing the list.
- The flows address the list by GUID (`TICKETS` in `gen.py`), never by title — renaming the list must not break them. (Until Oct 2026 the permission steps used the old title `Tickets` and failed silently after the list was renamed; run **HD17 Fix ticket permissions** once (button flow) to stop the list inheriting, drop Visitors/Members, and backfill every existing ticket. Safe to re-run.)
- Check: `…/_api/web/lists(guid'22c09b0b-…')?$select=HasUniqueRoleAssignments` should be `true` for the list and for every item.

### Other lists locked down from the site (2026-10-23)

Three more lists on the help desk site hold data end users shouldn't see, but — unlike Tickets — nothing re-locks them automatically, so if one is ever recreated it needs to be redone by hand:

- **External Senders**, **Live overview** (the dashboard's backing data), and **Pending replies** (raw requester reply text, see below) all have **unique permissions**: the site Owners group gets Full Control, Martin/Talal/Joseph get **Edit**, and the IT Support service account gets **Read** (flows read/write these as whichever SharePoint connection the flow runs under — currently Martin's). The site's Members/Visitors groups are **not** on any of the three.
- Why it mattered: all three previously inherited site permissions, so any of the ~75 site Members/Visitors could open them from **Site Contents** — External Senders exposed the allow-list itself, Live overview exposed SLA/health numbers, and Pending replies exposed other requesters' raw reply text keyed by email address.
- Check: `…/_api/web/lists(guid'<list guid>')?$select=HasUniqueRoleAssignments` should be `true` for each of the three (guids: External Senders `37b12028-370c-4b21-b31c-11362129c7fe`, Live overview `24631152-bdde-4654-82a4-d64bb0bb9b3b`, Pending replies `189e9634-988d-40ae-afce-0cb6eef5314c`).
- Left alone: Documents, Events, Form Templates, Holidays, Site Assets, Site Pages, Style Library — all empty or non-sensitive (Holidays is just date rows used by the business-hours check).

## When something looks wrong

| Symptom | Check |
|---|---|
| No confirmation sent | HD3 run history (Power Automate), requester email field on the ticket |
| Ticket not created from email | HD1 run history; was it sent to the right mailbox? Outside sender: is their domain/address on the External Senders list? Check the Review folder. |
| App says error on Send | HD14 run history; the Category chosen must match a SharePoint Category choice |
| Reminder/close didn't fire | HD6 run history (runs 8am–9pm) |
| Emails stop but Teams messages still arrive | Power Automate › Connections › Office 365 Outlook. If a flow's email step fails with `Unexpected connection parameter set name: 'oauthDelegated'`, the connection was recreated as the wrong type: make a new Office 365 Outlook connection (default sign-in), delete the old one, re-import `HelpDesk-all-flows.zip` with **Update** on every row. Never import with *Create as new* — that makes duplicate flows. |
| Every HD3 / HD4 run shows **Failed** but the emails and Teams posts still go out | The run died in its last step, `Health_problems` (the self-check). Open a failed run: if the message is `substring … start index must be less than the length of the string`, a leaf action with an empty error text hit `substring('',0,0)`. Fixed 2026-10-22 in `generator/gen.py` (`leaf_item` pads the text before `substring`); re-import `HD3.zip` and `HD4.zip` from `flows/` with **Update**. Microsoft's weekly "N of your flow(s) have failed" email counts these runs even though nothing user-facing was lost. |
| SharePoint Quick close button opens a **Run flow** panel with no resolution box, just "Review connections and actions" | The trigger's `schema` wasn't wrapped the way Power Automate's SharePoint button connector expects — it needs a top-level `{"type":"object","properties":{"rows":{...}},"required":["rows"]}` with an `entity` (the row ID) alongside `text` under `items.properties`, not the flatter shape. Found 2026-10-22 by touching the trigger's description in the flow designer and letting Power Automate re-save/canonicalize it. Fixed in `generator/gen.py` (`hd18`'s `trig` schema); re-import `HD18.zip` from `flows/` with **Update**. |
| **HD3 fails on `Card`** (channel post never goes out) with `substring … start index must be … less than the length of the string`, or **HD6's 5‑day reminder email** fails the same way, or **HD4's `Sync_task_notes`** (Planner description) fails the same way | Same `substring('',0,0)` gotcha as the `Health_problems` row above, just in a different spot: HD3's `Card` substrings the ticket's `Details` field to build the Teams card, HD6's reminder email substrings `Conversation`, and HD4's `Sync_task_notes` substrings `varConversation` — all without padding first, so a ticket with no Details yet (common right after creation) or no Conversation history yet (a ticket whose only update so far is an attachment, with no note/status/assignment change) throws instead of posting. Found 2026-10-25 testing a ticket created with no Details field. Fixed in `generator/gen.py` (all three now `substring(concat(x,' '),0,min(length(x),N))`); re-import `HD3.zip`, `HD4.zip` and `HD6.zip` from `flows/` with **Update**. |
| **HD20's `InternalLine`** fails with `InvalidTemplate … the provided value is of type 'Float'` | `LastNotifiedSponsorLen`/`LastNotifiedInternalLen` are SharePoint **Number** columns — the REST API always returns them as a Float, even for whole numbers, so any `coalesce`/`sub`/`max`/`min` expression touching them silently becomes a Float, and `substring()`'s length parameter rejects a Float. Only the Internal Notes path showed it directly; the Sponsor Updates path had the identical bug but it stayed hidden because `NoteLine`'s enclosing `if()` short-circuits the untaken branch. Found 2026-10-25 full-circle-testing HD20's Internal Notes path. Fixed in `generator/gen.py` (both length expressions now wrapped in `int(...)`); re-import `HD20.zip` from `flows/` with **Update**. |
| **HD20 keeps running on its own**, roughly once a minute, long after a project was last actually touched (no notifications go out, but the run history fills up) | HD20's `Update_item` step patches `LastNotifiedStatus`/`LastNotifiedSponsorLen`/`LastNotifiedInternalLen` **unconditionally** on every run, to keep them in sync — but that PATCH itself bumps the item's Modified date like any other edit, which re-fires the "item created or modified" trigger, which finds nothing new to notify about but still runs `Update_item` again, forever. HD4 avoids this with a **trigger-level condition** (only start a run if Status/AssignedTo differ from LastNotified* or NoteToUser is non-empty) — HD20 was missing the equivalent. Found 2026-10-25 full-circle-testing: a test project kept re-triggering HD20 every ~30–40s for 10+ minutes after a single note was saved. Fixed in `generator/gen.py` (`hd20`'s trigger now carries the same `status_changed`/`sponsor_changed`/`internal_changed` check as a `conditions` expression, so a run that wouldn't notify never starts); re-import `HD20.zip` from `flows/` with **Update**. |

## Flow reference

HD1 Email intake · HD2 Form intake · HD3 Ticket created · HD4 Ticket updated · HD5 Planner task completed · HD6 Clock · HD7 My open tickets · HD8 Daily report · HD9 Card actions · HD11 Tech thread replies · HD12 Requester replies in Teams · HD13 One-time setup of the pending-replies list · HD14 Teams app intake · HD15 Dashboard stats · HD16 Health check · HD17 Fix ticket permissions (run once) · HD18 Quick close (SharePoint) — the list button's flow; its ID goes into `dashboard/tickets-id-column.json` (`HD18_FLOW_ID`) · HD19 Project created · HD20 Project updated · HD21 No-update check (weekly, Fridays 9am) — see "IT Projects" above.

*Flows are generated from `generator/gen.py`. Never edit the live flows directly without updating the generator.*

Live flow IDs (environment Default-26e851a0…, rebuilt 2026-10-21 after the Outlook connection was replaced). When importing a zip with **Update**, pick the flow with this ID — never *Create as new*:

| Flow | ID (first 8) | Flow | ID (first 8) |
|---|---|---|---|
| HD1 Email intake | 3975d310 | HD9 Card actions | 90dd749d (original — Teams card responses are bound to it) |
| HD2 Form intake | e4d7a201 | HD11 Tech thread replies | bfd6591a |
| HD3 Ticket created | c646c647 | HD12 Requester replies in Teams | 2eb059df |
| HD4 Ticket updated | 5f9a1671 | HD13 Setup pending-replies (off) | b747215c |
| HD5 Planner task completed | 0cd62cf6 | HD14 Teams app intake | 4bede0e5 (under *Shared with me*; the app calls it by ID) |
| HD6 Hourly clock | 964cba3d | HD15 Dashboard stats | bd04cabb |
| HD7 My open tickets | 7172c1c0 | HD16 Health check v2 | da2f43b4 (old HD16 482dc407 kept off as rollback) |
| HD8 Daily report | 3c9907f7 | HD17 Fix ticket permissions (off) | 1a62b6b6 |
| HD18 Quick close (SharePoint) | 0d643975 (its full ID is in `dashboard/tickets-id-column.json`) | HD20 Project updated | e5d3bcee |
| HD19 Project created | b4d90357 | HD21 No-update check | acc884cc |

**Why HD14 is not in "My flows":** HD14 was created on 2026-10-10 under Martin's own account, with IT Support added as co-owner. "My flows" only lists flows whose *primary owner* is the signed-in account, so when working as IT_Support, HD14 appears under **Shared with me** instead. It is the only flow like this; every other flow was imported as IT_Support. The Teams app calls HD14 by its ID (4bede0e5…), so a copy of HD14 imported as IT_Support is *not* used by the app — delete any such spare, or re-point the app's Send button at it in Power Apps Studio if you want HD14 owned by IT_Support like the rest.
