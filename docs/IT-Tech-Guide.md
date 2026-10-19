# Branthaven IT Help Desk – Technician Guide

Site: https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2
Business day for SLA purposes: **7am–9pm**, every day.

## How tickets arrive

| Source | What happens |
|---|---|
| Email to the help desk mailbox | HD1 creates a ticket. A reply containing an [IT-123] tag updates that ticket. Branthaven senders always get a ticket; outside senders only if their address or domain is on the **External Senders** list (see below). Everything else lands in the mailbox's **Review** folder and is ignored. |
| IT Support form (Microsoft Forms) | HD2 creates a ticket. |
| **Report an IT issue** Teams app | HD14 creates a ticket (source: "Teams form"). |

Every new ticket then runs HD3: it sets permissions (the requester sees only their own ticket), creates the **Planner task**, emails and messages the requester a confirmation (as IT Support), alerts IT in the Help Desk channel, and sends Martin and Joseph a direct Teams message from IT Support with the ticket number, requester, priority and link.

## Your daily workflow

1. Watch the **Help Desk channel** and the **Planner board** (buckets: New, In progress, Waiting, Done).
2. **Claim a ticket**: open it from the alert/Planner link, set **Assigned To** and move Status to *Assigned* then *In Progress*.
3. **Work it**: keep Status accurate – it drives the SLA clock, the Planner bucket and the messages the requester receives.
4. **Ask the requester a question**: use the card/reply action on the ticket thread in Teams. The question reaches them as an IT Support message and their answer (by Teams or email) is posted back on the ticket.
5. **Resolve**: set Status to *Resolved Pending Confirmation* (or complete the Planner task), with a resolution written in the message box. The requester is notified, including your resolution text.
6. The system **closes it automatically after 48 hours**, with a reminder to the requester at 24 hours. Set *Closed* yourself if they confirm earlier — **a resolution is required to close a ticket**: if none was captured yet (e.g. closing directly without going through Resolved Pending Confirmation), enter it in the message box when you set Status to Closed, or the flow will bounce it back to you and leave the ticket open. This is enforced three ways: the Teams card blocks it outright, if a ticket is set to Closed by any other means (e.g. editing the list item directly) with no resolution on file HD4 reverts it to its prior status and messages you to add one, and completing the ticket's **Planner task** only moves the ticket to Resolved Pending Confirmation if a resolution is already on file — otherwise HD5 reopens the task and messages you to add the resolution first.

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
| Closed | Finished | Stopped |

Status changes automatically notify the requester (email + Teams) and move the Planner card.

## Clock, reports and reminders

- **HD6 (every day, 7am–9pm):** counts business hours open, handles the 24h reminder and 48h auto-close for resolved tickets. A ticket unassigned for 4 hours is posted in the Help Desk channel; at 5 hours Martin and Joseph (first-line) both get a direct Teams ping. Once it's assigned — to either tech or Martin — the normal per-day open reminder takes over and goes to whoever it's assigned to.
- **HD8 (9pm daily):** daily report of overdue tickets (open longer than a business day) and unassigned tickets (4+ hours), emailed to Martin.
- **HD7:** a requester can ask for "My open tickets" and gets a list.

## Monitoring (is the system healthy?)

**Look first at the IT Dashboard → "System health" panel** (top of https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2/SitePages/IT-Dashboard.aspx). Five cards, each green OK / amber CHECK / red PROBLEM, refreshed every hour 7am–9pm by HD16:

| Card | What it checks | Red means |
|---|---|---|
| New tickets set up | Tickets created in the last 24 h have a Planner task, a channel post and the "Ticket created" line | HD3 is not finishing — open its run history |
| Hourly clock (HD6) | At least one open ticket has been touched in the last 3 h | HD6 stopped — SLA clock, reminders and auto-close are all off |
| Auto-close (HD6) | No ticket resolved more than 72 h ago is still "Resolved Pending Confirmation" | Auto-close isn't firing |
| Emails not ticketed | Mail in the mailbox's **Review** folder in the last 24 h (amber, not red — it's a "go look" not a failure) | — |
| Notifications (HD3/HD4) | HD3/HD4 reported a failed requester/tech notification in the last 24 h | A requester or tech did not get a message — the card says which ticket and step |

Behind the panel, three layers of alerts all go to Martin:

1. **Per-ticket, immediately.** HD3 (ticket created) and HD4 (ticket updated) check their own work at the end of every run: requester email, requester Teams message, Planner task, channel post, tech ping. If any did not succeed, Martin gets an email and a Teams message from IT Support naming the ticket, the failed step, the error text and a link to the flow run, and the Notifications card turns red for 24 h. The ticket itself is still created/updated — only the notification is reported.
2. **HD16 health check, hourly.** Updates the panel every run; emails Martin a report once a day at 9:30 am (subject "[Help desk health] All clear" or "[Help desk health] PROBLEMS: n"); Teams-messages him whenever the issue count changes, so a new problem pings once rather than every hour.
3. **Hard flow failures.** Power Automate emails the flow owner (IT Support) whenever a run fails outright, plus a weekly digest. To receive these as yourself, open each flow → ⋯ → *Flow failure alerts*.

If an alert names a flow run, open the link: the failed step is highlighted in red with the full error.

## Reply to a closed ticket

If someone replies — by email or in the Teams chat — to a ticket that's already Closed, the system doesn't reopen it. It starts a **new ticket** instead, tagged "Follow-up to IT-<old number>" in Details, and (via Teams) titled "Follow-up: <old title>" with the same Category/Priority carried over. If the old ticket had a resolution on file, it's copied into the new ticket's Details too ("Previous resolution (IT-<old number>): ..."), so you can see what was already tried without flipping back to the old ticket. The old ticket stays closed as the record of what was done; work the new one like any other.

## Rules of the road

- Only the list owner edits the Tickets list structure; techs work tickets through the ticket form/Planner.
- Do not delete tickets. Close them.
- Don't email the requester separately for status changes; the system already does.
- Ticket numbers are IT-<ID>. Put them in any email subject about the ticket.

## When something looks wrong

| Symptom | Check |
|---|---|
| No confirmation sent | HD3 run history (Power Automate), requester email field on the ticket |
| Ticket not created from email | HD1 run history; was it sent to the right mailbox? Outside sender: is their domain/address on the External Senders list? Check the Review folder. |
| App says error on Send | HD14 run history; the Category chosen must match a SharePoint Category choice |
| Reminder/close didn't fire | HD6 run history (runs 7am–9pm) |

## Flow reference

HD1 Email intake · HD2 Form intake · HD3 Ticket created · HD4 Ticket updated · HD5 Planner task completed · HD6 Clock · HD7 My open tickets · HD8 Daily report · HD9 Card actions · HD11 Tech thread replies · HD12 Requester replies in Teams · HD13 One-time setup of the pending-replies list · HD14 Teams app intake · HD15 Dashboard stats · HD16 Health check.

*Flows are generated from `generator/gen.py`. Never edit the live flows directly without updating the generator.*
