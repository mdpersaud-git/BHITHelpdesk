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
| Flows | Power Automate Management "List My Flows": every expected flow (HD1–HD9, HD11, HD12, HD15, HD16 v2) is **on, exactly once**. HD13/HD17 (run-once) are not expected on; HD14 is owned by Martin's account so it is not visible to this check. | "HDx is off or missing" / "HDx has 2 copies on" — turn the right one on, delete the duplicate |
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

## Flow reference

HD1 Email intake · HD2 Form intake · HD3 Ticket created · HD4 Ticket updated · HD5 Planner task completed · HD6 Clock · HD7 My open tickets · HD8 Daily report · HD9 Card actions · HD11 Tech thread replies · HD12 Requester replies in Teams · HD13 One-time setup of the pending-replies list · HD14 Teams app intake · HD15 Dashboard stats · HD16 Health check · HD17 Fix ticket permissions (run once) · HD18 Quick close (SharePoint) — the list button's flow; its ID goes into `dashboard/tickets-id-column.json` (`HD18_FLOW_ID`).

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
| HD18 Quick close (SharePoint) | 0d643975 (its full ID is in `dashboard/tickets-id-column.json`) | | |

**Why HD14 is not in "My flows":** HD14 was created on 2026-10-10 under Martin's own account, with IT Support added as co-owner. "My flows" only lists flows whose *primary owner* is the signed-in account, so when working as IT_Support, HD14 appears under **Shared with me** instead. It is the only flow like this; every other flow was imported as IT_Support. The Teams app calls HD14 by its ID (4bede0e5…), so a copy of HD14 imported as IT_Support is *not* used by the app — delete any such spare, or re-point the app's Send button at it in Power Apps Studio if you want HD14 owned by IT_Support like the rest.
