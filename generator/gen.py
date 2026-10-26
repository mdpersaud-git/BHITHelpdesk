import json, uuid, os, zipfile, shutil, copy

OUT = os.path.join(os.path.dirname(__file__), 'out')
SITE = "https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2"
ENV = "Default-26e851a0-5d3b-4687-80a3-7d758c6fc2b2"   # Power Automate environment (for links in alerts)
TICKETS = "22c09b0b-bbcd-42a9-9aeb-844d81089636"
PROJECTS = "3de1b401-4cac-4994-9c43-0d7cdd464b47"   # "IT Projects" list, added 2026-10-25
PRJ_LIST_URL = SITE + "/Lists/IT%20Projects/Full%20Detail.aspx?FilterField1=ID&FilterValue1="   # tech view
PRJ_ITEM_URL = SITE + "/Lists/IT%20Projects/DispForm.aspx?ID="   # read-only item page (also used by the sponsor for now)
PRJ_STATUSES = ["Not Started", "In Progress", "On Hold", "Waiting on Vendor/Approval", "Completed", "Cancelled"]
PRJ_ACTIVE = ["Not Started", "In Progress", "On Hold", "Waiting on Vendor/Approval"]   # not Completed/Cancelled
PRJ_CATEGORIES = ["Infrastructure", "Security", "Software Rollout", "Network", "Hardware Refresh", "Process Improvement", "Other"]
# keep PRJ_STATUSES/PRJ_CATEGORIES in sync with: the IT Projects list's Status/Category choice columns
HOLIDAYS = "430cd9b6-9cdc-45a2-8e9b-66a88cd19b74"
EXTERNAL_SENDERS = "37b12028-370c-4b21-b31c-11362129c7fe"   # "External Senders" allow-list: Title = domain or full address
MBX = "IT_Support@branthaven.com"
INTERNAL_DOMAIN = "@branthaven.com"
ME = "mpersaud@branthaven.com"
TECH1 = "tsiddiqui@branthaven.com"   # Talal Siddiqui
TECH2 = "jelkadi@branthaven.com"     # Joseph El-Kadi, first-line support
UNASSIGNED_PING_TEAM = [ME, TECH2]   # 5-hour unassigned-ticket ping: Martin + first-line support only
NEW_TICKET_PING_TEAM = [ME, TECH2]   # 1:1 Teams message the moment a ticket is created (HD3)
TECHS = [(ME, "Martin Persaud"), (TECH1, "Talal Siddiqui"), (TECH2, "Joseph El-Kadi")]   # "Edit ticket" card choices
PRIORITIES = ["High", "Medium", "Low"]
CATEGORIES = ["Access Request", "Application", "Bookings", "BuilderGo/OnLocation Warranty", "Email", "Horizon", "InFormXL",
              "Landline", "Laptop / Desktop", "Mobile (iPad / Cell)", "Mouse / Keyboard / Monitor", "New User Setup", "NewStar",
              "OneDrive / SharePoint", "Other", "Password or Sign-In", "Printer", "SaleFish", "SaleXpress", "Teams", "VPN", "Uncategorized"]
# keep in sync with: the Category choice column, ddCategory.Items in the canvas app, and the Chart - Category view (22 slots)
GROUP = "f6e29c64-7081-428a-9713-f9f4b3373f67"
CHANNEL = "19:a5b109f2452d44a9a1a80e1d72075232@thread.tacv2"
PLAN = "AP29gK5xm06UPxOU3IjkQWUAFqRt"
B_NEW = "exgIyM6b1UWrsrntUDiTn2UADctB"
B_PROG = "XvpcjPJjIU2T_WUM0hPQ6mUAD0Lw"
B_WAIT = "QwPq1DcpE0aT8r3VUmrCNmUAH-XC"
B_DONE = "j9sW6pWSLUm0T-6dkkkJ92UAAs1h"
F_REVIEW = "Id::AAMkADQ1OWM4ZWViLWNjMDItNGZkOC04MTk3LWU5YzU3NTVlOTc2YQAuAAAAAADEPNF_h2PyQJliA7jqv5DlAQCWlQEoNgX0Sa8FZZmjc_9SAAAAAAE-AAA="
F_PROCESSED = "Id::AAMkADQ1OWM4ZWViLWNjMDItNGZkOC04MTk3LWU5YzU3NTVlOTc2YQAuAAAAAADEPNF_h2PyQJliA7jqv5DlAQCWlQEoNgX0Sa8FZZmjc_9SAAAAAAE_AAA="
FORM1 = "oFHoJjtdh0aAo311jG_CsmSc4vaBcIpClxP59LM3P2dURTQ4QUJGSFVWSTFLN1VMTUY1WTRWT05VUiQlQCN0PWcu"
FORM2 = "oFHoJjtdh0aAo311jG_CsmSc4vaBcIpClxP59LM3P2dUMVdBWUU2OFEyS1FOQjZEUEg5TFBERUlKTyQlQCN0PWcu"
Q_WHAT = "r7b9c158532f549f0a2ebace3feef1795"
Q_CAT = "r6ca7f9430ba64026825ac19e7d8fd71f"
Q_URG = "ra6e75a77b45e4077b8b29217c325cec9"
Q_DET = "r1058e21c2b0e4f079fba1c54924d5302"
Q_SHOT = "rcbdab21fdac14fd6bde6f7009de24d9e"
TZ = "Eastern Standard Time"
MARKER = "## Reply above this line ##"
LIST_URL = SITE + "/Lists/Tickets/AllItems.aspx?FilterField1=ID&FilterValue1="  # tech view: formatted list showing one ticket
REQ_URL = SITE + "/Lists/Tickets/DispForm.aspx?ID="  # read-only item page for requesters

APIS = {
    'shared_office365users': ('office365users', 'Office 365 Users', '30b9b2951c534266941f276deb55839b'),
    'shared_office365': ('office365', 'Office 365 Outlook', '1332e4ac698b4bb983fe43adc1c74483'),
    'shared_sharepointonline': ('sharepointonline', 'SharePoint', '907790570e5b4663bd70952e83b2dbe1'),
    'shared_planner': ('planner', 'Planner', '72155b29abdd4c5d9cbebc824d4c1502'),
    'shared_teams': ('teams', 'Microsoft Teams', 'shared-teams-placeholder'),
    'shared_teams_1': ('teams', 'Microsoft Teams (sign in as IT_Support@branthaven.com)', 'shared-teams-itsupport-placeholder'),
    'shared_microsoftforms': ('microsoftforms', 'Microsoft Forms', 'shared-forms-placeholder'),
    'shared_conversionservice': ('conversionservice', 'Content Conversion', 'shared-conversion-placeholder'),
    'shared_onedriveforbusiness': ('onedriveforbusiness', 'OneDrive for Business', 'shared-onedrive-placeholder'),
    'shared_flowmanagement': ('flowmanagement', 'Power Automate Management', 'shared-flowmgmt-placeholder'),
}

# ---------- statuses ----------
S_OPEN, S_ASSIGNED, S_PROG, S_ESC = "Open", "Assigned", "In Progress", "Escalated or Waiting Internally"
S_WAIT, S_SCHED, S_RES, S_CLOSED = "Waiting for End User", "To Be Scheduled", "Resolved Pending Confirmation", "Closed"
STATUSES = [S_OPEN, S_ASSIGNED, S_PROG, S_ESC, S_WAIT, S_SCHED, S_RES, S_CLOSED]
WORKING = [S_OPEN, S_ASSIGNED, S_PROG, S_ESC]          # SLA clock runs
MEANING = {
    S_OPEN: "Status on newly created Ticket",
    S_ASSIGNED: "Ticket has been re-assigned to an IT Team Member",
    S_PROG: "IT Support has begun work on the Ticket",
    S_ESC: "IT Support has reached out for assistance from an internal team",
    S_WAIT: "IT Support has requested further information from you",
    S_SCHED: "IT Support has requested to schedule a meeting with you",
    S_RES: "IT Support has provided resolution and awaiting confirmation that issue has been resolved",
    S_CLOSED: "Ticket has been resolved and closed"}
def arr(lst):
    return "createArray(" + ",".join("'" + x + "'" for x in lst) + ")"
def bucket_expr(S):
    return (f"if(equals({S},'{S_OPEN}'),'{B_NEW}',if(contains({arr([S_ASSIGNED,S_PROG,S_ESC])},{S}),'{B_PROG}',"
            f"if(contains({arr([S_WAIT,S_SCHED])},{S}),'{B_WAIT}','{B_DONE}')))")

# ---------- helpers ----------
def ra(*names, status=("Succeeded",)):
    return {n: list(status) for n in names}

def op(api, opid, params, after=None, **extra):
    a = {"runAfter": after or {}, "type": "OpenApiConnection",
         "inputs": {"parameters": params,
                    "host": {"apiId": f"/providers/Microsoft.PowerApps/apis/shared_{APIS[api][0]}", "connectionName": api, "operationId": opid},
                    "authentication": "@parameters('$authentication')"}}
    a.update(extra)
    return a

def compose(expr, after=None):
    return {"runAfter": after or {}, "type": "Compose", "inputs": expr}

def cond(expr, yes, no=None, after=None):
    return {"runAfter": after or {}, "type": "If",
            "expression": {"and": [{"equals": [expr, "@true"]}]},
            "actions": yes, "else": {"actions": no or {}}}

def seq(actions):
    """chain an ordered list of (name, action) so each runs after the previous"""
    out = {}
    prev = None
    for name, a in actions:
        if prev and not a.get("runAfter"):
            a["runAfter"] = ra(prev)
        out[name] = a
        prev = name
    return out

def sp(opid, params, after=None):
    p = {"dataset": SITE}
    if opid not in ("HttpRequest", "CreateFile", "GetFileContentByPath"):
        p["table"] = TICKETS
    p.update(params)
    return op("shared_sharepointonline", opid, p, after)

def sp_http(method, uri, body=None, after=None, headers=None):
    p = {"dataset": SITE, "parameters/method": method, "parameters/uri": uri,
         "parameters/headers": {"Accept": "application/json;odata=nometadata",
                                "Content-Type": "application/json;odata=nometadata", **(headers or {})}}
    if body is not None:
        p["parameters/body"] = body
    return op("shared_sharepointonline", "HttpRequest", p, after)

LOGO_URL = "https://www.branthaven.com/wp-content/uploads/2026/08/cropped-cropped-BH19-NewLogo_button_black-square-270x270.jpg"
EMAIL_FOOTER = ('<div style="margin-top:24px"><img src="' + LOGO_URL + '" width="80" height="80" alt="Branthaven" '
                'style="display:block;border:0;width:80px;height:80px"></div>')

def mail(to, subject, body, after=None, attachments=None):
    p = {"emailMessage/MailboxAddress": MBX, "emailMessage/To": to,
         "emailMessage/Subject": subject, "emailMessage/Body": body + EMAIL_FOOTER, "emailMessage/Importance": "Normal"}
    if attachments:
        p["emailMessage/Attachments"] = attachments
    return op("shared_office365", "SharedMailboxSendEmailV2", p, after)

def chat(to, msg, after=None):
    return op("shared_teams", "PostMessageToConversation",
              {"poster": "Flow bot", "location": "Chat with Flow bot", "body/recipient": to, "body/messageBody": msg}, after)

SUPPORT = "IT_Support@branthaven.com"

def scope_msg(prefix, to, html):
    return {"type": "Scope", "runAfter": {}, "actions": support_msg(prefix, to, html)}
def team_msgs(prefix, html, team):
    """1:1-message each person in `team` in parallel. Returns {name: action} for a branch.
    The outer Scope name must differ from the inner message-action name (support_msg names
    its post action literally `prefix`), or Logic Apps rejects the flow for duplicate names."""
    return {f"Tell_{prefix}_{i}": scope_msg(f"{prefix}_{i}", who, html) for i, who in enumerate(team)}

# ---------- multi-value Person-field helpers (IT Projects Owner / Project Sponsor) ----------
# Both columns now allow multiple people. The SharePoint connector flattens a multi-value Person field
# into an array of objects shaped like its single-value form (Email, DisplayName, Claims, ...) - the same
# Email/DisplayName keys the existing singular code already reads (pt('Owner/Email') etc.), just
# array-wrapped - confirmed live 2026-10-26 against a real trigger body (Owner.0.Email, Owner.1.Email, ...).
#
# Power Automate's expression language has NO select()/filter() lambda function - "Select" and "Filter
# array" (Query) only exist as dedicated ACTIONS, never as something callable inline inside another
# expression (confirmed live 2026-10-26: HD19 failed with "InvalidTemplate ... the template function
# 'select' is not defined or not valid" the first time this file used an inline select(...) call). So
# extracting values (emails/names, for joining into a recipient list) needs an actual Select action
# first - people_select() below - with the join done afterwards by a Compose reading that action's body.
# A membership *check* (does this array contain a given email) doesn't need a new action, since Query's
# own "where" clause is itself just an expression with no action step of its own to attach a nested Select
# to (this bites HD12's Sponsor_projects Query specifically, whose where clause already runs per outer
# item()) - people_contains() instead lower-cases the whole array's JSON serialization and substring-
# matches '"email":"<address>"', which is safe because a SharePoint Person field's Email can't contain a
# literal '"'.
def people_select(name, expr, key, after=None):
    """A 'Select' action (named `name`) extracting `key` ('Email' or 'DisplayName') from each item of a
    multi-value Person array expression (`expr`, no leading @). Email is lower-cased on the way out.
    Returns {name: action}; its body (an array of plain strings) is read as body('{name}')."""
    sel = f"@toLower(coalesce(item()?['{key}'],''))" if key == "Email" else f"@coalesce(item()?['{key}'],'')"
    return {name: {"type": "Select", "runAfter": after or {},
        "inputs": {"from": f"@coalesce({expr},json('[]'))", "select": sel}}}

def join_emails(select_name, fallback_single=None):
    """';'-joined emails from a people_select(..., 'Email') action's body (`select_name`, the action's
    name, no leading @ needed). `fallback_single`, if given, is a plain-text email expression (no leading
    @) used only when that Select produced an empty array (legacy single-email field)."""
    joined = f"join(body('{select_name}'),';')"
    if fallback_single:
        return f"if(empty(body('{select_name}')),toLower(coalesce({fallback_single},'')),{joined})"
    return joined

def join_names(select_name, empty_text):
    """Comma-joined display names from a people_select(..., 'DisplayName') action's body (`select_name`,
    no leading @ needed); `empty_text` (e.g. 'Not yet assigned') when that Select produced an empty array."""
    return f"if(empty(body('{select_name}')),'{empty_text}',join(body('{select_name}'),', '))"

def people_contains(expr, email_expr):
    """True if the lower-cased `email_expr` (no leading @) is one of the Emails in a multi-value Person
    array expression (`expr`, no leading @) - see the module comment above for why this can't just be
    contains(select(...)) despite looking like it should be able to."""
    return (f"contains(toLower(string(coalesce({expr},json('[]')))),"
            f"concat('\"email\":\"',{email_expr},'\"'))")

def notify_each(prefix, csv_expr, html, exclude=()):
    """1:1-message every non-empty, internal, non-excluded address in a `;`-joined email string
    (`csv_expr`, no leading @ - typically a join_emails(...) Compose's outputs('...')). `exclude` is a list of expressions
    (no leading @, already lower-cased, e.g. outputs('CreatorEmail')) to skip, so someone who is both the
    creator and a sponsor isn't messaged twice. Returns {name: action} with one Scope named `prefix`
    wrapping a Foreach, so self_check's `result('{prefix}')` sees pass/fail the same way it already does
    for a plain Scope (Foreach does not add to CLAUDE.md's If/Switch/Foreach nesting cap - a Scope never
    does, and this Foreach sits at just one level)."""
    who = f"{prefix}_who"
    skip = "".join(f",not(equals(outputs('{who}'),{e}))" for e in exclude)
    return {prefix: {"type": "Scope", "runAfter": {}, "actions": {
        f"{prefix}_loop": {"type": "Foreach", "runAfter": {}, "foreach": f"@split({csv_expr},';')",
            "actions": seq([
                (who, compose("@trim(item())")),
                (f"{prefix}_check", cond(
                    f"@and(not(empty(outputs('{who}'))),not(equals(outputs('{who}'),'{ME}')){skip}," +
                    is_internal(f"outputs('{who}')") + ")",
                    # the If's branch key (f"{prefix}_notify") must differ from scope_msg's own prefix
                    # (f"{prefix}_msg", which names its inner post action literally that), per team_msgs'
                    # duplicate-name warning above.
                    {f"{prefix}_notify": scope_msg(f"{prefix}_msg", f"@outputs('{who}')", html)})),
            ])},
    }}}
RUN_LINK = ("https://make.powerautomate.com/environments/@{workflow()?['tags']?['environmentName']}"
            "/flows/@{workflow()?['name']}/runs/@{workflow()?['run']?['name']}")

def leaf_item(x):
    """A {name,status,error:{message}} record for one named action, built as a JSON string so it can be
    unioned with result() output. Quotes, backslashes and newlines in the error text are flattened so the
    JSON stays valid; the text is cut at 160 characters. The message is padded with one space before
    substring(): a succeeded action has an empty message and substring('',0,0) is an error in Power Automate
    ("start index must be less than the length of the string"), which made every HD3/HD4 run end Failed."""
    msg = "coalesce(actions('%s')?['error']?['message'],'')" % x
    safe = ("trim(replace(replace(replace(substring(concat(%s,' '),0,min(length(%s),160)),'\"',''''),'\\',' '),decodeUriComponent('%%0A'),' '))"
            % (msg, msg))
    return ("json(concat('{\"name\":\"%s\",\"status\":\"',coalesce(actions('%s')?['status'],'Skipped'),"
            "'\",\"error\":{\"message\":\"',%s,'\"}}'))" % (x, x, safe))

def self_check(prefix, scoped, subject, intro, ticket_ref, extra=None, after=None, leaf=(), dash_id=None, dash_title=None):
    """Monitoring: look at the results of the listed scoped actions (Scope / If / Foreach) and, if any top-level
    action in them did not succeed, email + Teams-message Martin with the action names, error text and a link to
    the run. `extra` is an optional expression returning additional {name,status} items (synthetic checks).
    `dash_id`/`dash_title` pick which dashboard row's Slot10/Slot11 "last problem" stamp this call writes to -
    DASH_HEALTH/"Health" (ticket flows HD3/HD4) by default, or DASH_PROJHEALTH/"ProjectHealth" (project flows
    HD19/HD20) when passed explicitly. Keeping these separate per flow family matters: HD16 reads DASH_HEALTH's
    Slot10/Slot11 to build the ticket-system "Notifications (HD3/HD4 self-checks)" digest line, and before this
    was split every self_check call (ticket AND project) stamped the same DASH_HEALTH row - a PRJ-n problem from
    HD19/HD20 would silently overwrite a ticket problem (or vice versa), show up mislabeled under the ticket
    digest, inflate the ticket IssueCount/Total and trip Alert_if_changed for a problem that was actually on the
    project side, while having nowhere of its own to surface on the Project Health panel. Found 2026-10-26 when
    a stale HD20 self-check failure on PRJ-3 (the already-fixed `select()` bug, see the troubleshooting table)
    showed up in the "Help desk health" (ticket) digest as "Notifications (HD3/HD4 self-checks)".
    Returns {name: action}; the first action takes `after`."""
    src = "union(" + ",".join(f"result('{s}')" for s in scoped) + ")"
    if extra:
        src = f"union({src},{extra})"
    # result() only sees the top level of a scope; named leaf actions (an email inside an If inside a Scope) are
    # read directly, so a nested failure is reported too. A leaf that never ran reads as Skipped.
    if leaf:
        src = f"union({src},createArray(" + ",".join(leaf_item(x) for x in leaf) + "))"
    lines = f"@{{join(body('{prefix}_lines'),'')}}"
    html = (f"<p>{intro}</p><ul>{lines}</ul>"
            f"<p><a href=\"{RUN_LINK}\">Open this flow run</a></p>")
    return seq([
        (f"{prefix}_problems", {"type": "Query", "runAfter": after or {},
            "inputs": {"from": f"@{src}", "where": "@not(contains(createArray('Succeeded','Skipped'),item()?['status']))"}}),
        (f"{prefix}_lines", {"type": "Select", "inputs": {"from": f"@body('{prefix}_problems')",
            "select": "@concat('<li><b>',item()?['name'],'</b>: ',coalesce(item()?['error']?['message'],item()?['status']),'</li>')"}}),
        (f"{prefix}_names", {"type": "Select", "inputs": {"from": f"@body('{prefix}_problems')", "select": "@item()?['name']"}}),
        (f"{prefix}_alert", cond(f"@greater(length(body('{prefix}_problems')),0)", seq([
            (f"{prefix}_email", mail(ME, subject, html)),
            (f"{prefix}_teams", {"type": "Scope", "runAfter": ra(f"{prefix}_email", status=("Succeeded", "Failed")),
                                 "actions": support_msg(f"{prefix}_dm", ME, html)}),
            # stamp the dashboard (HD16/HD16's project half shows it in the problem band for 24h) - see dash_id
            # in the docstring above for why this isn't always DASH_HEALTH.
            (f"{prefix}_dashboard", sp("PatchItem", {"table": DASH,
                "id": dash_id if dash_id is not None else DASH_HEALTH,
                "item/Title": dash_title if dash_title is not None else "Health",
                "item/Slot11": "@utcNow('yyyy-MM-ddTHH:mm:ssZ')",
                "item/Slot10": f"@concat('Last problem: ',{ticket_ref},' — ',string(length(body('{prefix}_problems'))),' step(s) failed, ',{nowET('MMM d h:mm tt')},' (details emailed)')"},
                after=ra(f"{prefix}_teams", status=("Succeeded", "Failed")))),
        ]))),
    ])

def support_msg(prefix, to, html):
    """Message a requester 1:1 from the IT Support account (connection shared_teams_1).
    Returns {name: action} for a branch, or use .items() in a seq list."""
    return seq([
        (f"{prefix}_open_chat", op("shared_teams_1", "CreateChat", {"item/members": to})),
        (prefix, op("shared_teams_1", "PostMessageToConversation",
            {"poster": "User", "location": "Group chat",
             "body/recipient": f"@body('{prefix}_open_chat')?['id']", "body/messageBody": html})),
    ])

NEWISSUE_URL = "https://forms.office.com/Pages/ResponsePage.aspx?id=" + FORM1

def support_links(ID, title_expr):
    mailto = f"mailto:{MBX}?subject=@{{encodeUriComponent(concat('[IT-',string({ID}),'] Reply: ',coalesce({title_expr},'')))}}"
    return (f"<p><a href=\"{REQ_URL}@{{{ID}}}\">View ticket</a> &nbsp;·&nbsp; <a href=\"{PORTAL_URL}\">My tickets</a>"
            f" &nbsp;·&nbsp; <a href=\"{mailto}\">Reply by email</a></p>"
            "<p style=\"color:#888;font-size:12px\">To add details or reply to IT, reply in this chat or use Reply by email.</p>")

def channel(msg, after=None):
    return op("shared_teams", "PostMessageToConversation",
              {"poster": "Flow bot", "location": "Channel", "body/recipient/groupId": GROUP,
               "body/recipient/channelId": CHANNEL, "body/messageBody": msg}, after)

def channel_reply(parent, msg, after=None):
    return op("shared_teams", "ReplyWithMessageToConversation",
              {"poster": "Flow bot", "location": "Channel", "body/recipient/groupId": GROUP,
               "body/recipient/channelId": CHANNEL, "body/parentMessageId": parent, "body/messageBody": msg}, after)


def ticket_card(ID, title, S, it, req, cat, pri, src, hist):
    """Adaptive card for the channel post. All args are expressions without a leading @."""
    return {"type": "AdaptiveCard", "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": f"IT-@{{{ID}}}: @{{{title}}}", "weight": "Bolder", "size": "Medium", "wrap": True},
            {"type": "TextBlock", "text": f"Status: @{{{S}}}", "weight": "Bolder", "wrap": True,
             "color": f"@{{if(contains(createArray('Resolved','Closed'),{S}),'Good',if(equals({S},'New'),'Accent','Warning'))}}"},
            {"type": "FactSet", "facts": [
                {"title": "IT", "value": f"@{{coalesce({it},'Not assigned')}}"},
                {"title": "Requester", "value": f"@{{{req}}}"},
                {"title": "Category", "value": f"@{{{cat}}}"},
                {"title": "Priority", "value": f"@{{{pri}}}"},
                {"title": "Source", "value": f"@{{{src}}}"}]},
            {"type": "TextBlock", "text": "History", "weight": "Bolder", "separator": True, "spacing": "Medium"},
            {"type": "TextBlock", "text": f"@{{{hist}}}", "wrap": True, "size": "Small"}],
        "actions": [{"type": "Action.OpenUrl", "title": "View ticket", "url": f"{LIST_URL}@{{{ID}}}"}]}

def channel_card(card_expr, after=None, card_type=None):
    p = {"poster": "Flow bot", "location": "Channel", "body/recipient/groupId": GROUP,
         "body/recipient/channelId": CHANNEL, "body/messageBody": card_expr}
    if card_type:
        p["body/cardTypeId"] = card_type
    return op("shared_teams", "PostCardToConversation", p, after)

def channel_card_update(msgid, card_expr, after=None):
    return op("shared_teams", "UpdateCardInConversation",
              {"poster": "Flow bot", "location": "Channel", "body/messageId": msgid, "body/recipient/groupId": GROUP,
               "body/recipient/channelId": CHANNEL, "body/messageBody": card_expr}, after)

CARD_TYPE = "ITHelpDeskTicketCard"
BLANK_GIF = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"

def action_card(ID, title, req, cat, pri, src, details, pic=None, photo=None, req_name=None):
    """Static ticket card with action buttons (Adaptive Card 1.2 for Teams mobile). Args are expressions without a leading @.
    photo: expression giving a data: URI of the requester's photo ('' when none); req_name: display name expression."""
    tid = f"@{{{ID}}}"
    def text_card(inputs, act):
        return {"type": "AdaptiveCard", "body": inputs,
                "actions": [{"type": "Action.Submit", "title": "Send", "data": {"act": act, "ticket": tid}}]}
    status_inputs = [
        {"type": "TextBlock", "text": "New status", "weight": "Bolder"},
        {"type": "Input.ChoiceSet", "id": "status_choice", "style": "compact", "value": S_PROG,
         "choices": [{"title": x, "value": x} for x in STATUSES if x != S_OPEN]},
        {"type": "TextBlock", "text": "Message (sent to the user for Waiting for End User and To Be Scheduled; "
                                      "the resolution for Resolved; an internal note for Escalated)", "wrap": True, "size": "Small", "isSubtle": True},
        {"type": "Input.Text", "id": "status_text", "isMultiline": True, "placeholder": "Optional message"}]
    keep = {"title": "(keep current)", "value": ""}
    edit_inputs = [
        {"type": "TextBlock", "text": "Assign to", "weight": "Bolder"},
        {"type": "Input.ChoiceSet", "id": "edit_tech", "style": "compact", "value": "",
         "choices": [keep, {"title": "Unassigned", "value": "none"}] + [{"title": n, "value": e} for e, n in TECHS]},
        {"type": "TextBlock", "text": "Priority", "weight": "Bolder"},
        {"type": "Input.ChoiceSet", "id": "edit_priority", "style": "compact", "value": "",
         "choices": [keep] + [{"title": p, "value": p} for p in PRIORITIES]},
        {"type": "TextBlock", "text": "Category", "weight": "Bolder"},
        {"type": "Input.ChoiceSet", "id": "edit_category", "style": "compact", "value": "",
         "choices": [keep] + [{"title": c, "value": c} for c in CATEGORIES]},
        {"type": "TextBlock", "text": "Only the fields you change are updated. Reassigning notifies the new technician and the requester.",
         "wrap": True, "size": "Small", "isSubtle": True}]
    return {"type": "AdaptiveCard", "version": "1.2",
        "body": [
            {"type": "TextBlock", "text": "IT Support [Branthaven] - New Ticket", "weight": "Bolder", "size": "Small", "color": "Accent", "wrap": True},
            {"type": "TextBlock", "text": f"IT-@{{{ID}}}: @{{{title}}}", "weight": "Bolder", "size": "Medium", "wrap": True},
            *([{"type": "ColumnSet", "spacing": "Small", "columns": [
                {"type": "Column", "width": "auto", "verticalContentAlignment": "Center", "isVisible": f"@not(empty({photo}))", "items": [
                    {"type": "Image", "url": f"@if(empty({photo}),'{BLANK_GIF}',{photo})", "size": "Small", "style": "Person"}]},
                {"type": "Column", "width": "stretch", "verticalContentAlignment": "Center", "items": [
                    {"type": "TextBlock", "text": f"@{{{req_name}}}", "weight": "Bolder", "wrap": True},
                    {"type": "TextBlock", "text": f"@{{{req}}}", "size": "Small", "isSubtle": True, "wrap": True, "spacing": "None"}]}]}] if photo else []),
            {"type": "FactSet", "facts": [
                *([] if photo else [{"title": "Requester", "value": f"@{{{req}}}"}]),
                {"title": "Category", "value": f"@{{{cat}}}"},
                {"title": "Priority", "value": f"@{{{pri}}}"},
                {"title": "Source", "value": f"@{{{src}}}"}]},
            {"type": "TextBlock", "text": f"@{{{details}}}", "wrap": True, "size": "Small", "maxLines": 6},
            *([{"type": "TextBlock", "text": "Image attached. Open the ticket to view it.", "weight": "Bolder", "color": "Accent", "wrap": True, "isVisible": pic}] if pic else []),
            {"type": "TextBlock", "text": "Status changes and history appear in the replies below.", "isSubtle": True, "size": "Small", "wrap": True}],
        "actions": [
            {"type": "Action.Submit", "title": "Assign to me", "data": {"act": "assign", "ticket": tid}},
            {"type": "Action.Submit", "title": "Start work", "data": {"act": "start", "ticket": tid}},
            {"type": "Action.ShowCard", "title": "Update status", "card": text_card(status_inputs, "status")},
            {"type": "Action.ShowCard", "title": "Note to user", "card": text_card(
                [{"type": "Input.Text", "id": "note_text", "isMultiline": True, "placeholder": "Message the user sees by email and Teams"}], "note")},
            {"type": "Action.ShowCard", "title": "Edit ticket", "card": text_card(edit_inputs, "edit")},
            {"type": "Action.ShowCard", "title": "Quick close", "card": text_card(
                [{"type": "TextBlock", "text": "Fixed already? Enter the resolution. The ticket is assigned to you and closed in one step; the requester gets a single message.", "wrap": True, "size": "Small", "isSubtle": True},
                 {"type": "Input.Text", "id": "quick_text", "isMultiline": True, "placeholder": "Resolution (required)"}], "quickclose")},
            {"type": "Action.OpenUrl", "title": "View ticket", "url": f"{LIST_URL}@{{{ID}}}"},
            *([{"type": "Action.OpenUrl", "title": "View image", "url": "@{outputs('PicUrl')}", "isVisible": pic}] if pic else [])]}

def move(folder, after=None):
    return op("shared_office365", "MoveV2", {"messageId": "@triggerOutputs()?['body/id']", "folderPath": folder,
                                             "mailboxAddress": MBX}, after)

def planner_update(bucket, done=False, after=None, task="@triggerOutputs()?['body/PlannerTaskId']"):
    p = {"id": task, "body/bucketId": bucket}
    if done:
        p["body/percentComplete"] = 100
    return op("shared_planner", "UpdateTask_V3", p, after)

def link(idexpr):
    return f"<a href=\"{LIST_URL}@{{{idexpr}}}\">View ticket</a>"

def req_link(idexpr):
    return f"<a href=\"{REQ_URL}@{{{idexpr}}}\">View ticket</a>"

def is_internal(email_expr):
    """expression (no leading @): true when the address belongs to a Branthaven user"""
    return f"endsWith({email_expr},'{INTERNAL_DOMAIN}')"

def req_link_if_internal(idexpr, email_expr):
    """View-ticket paragraph for internal requesters only; external senders have no SharePoint access"""
    return f"@{{if({is_internal(email_expr)},concat('<p><a href=\"{REQ_URL}',{idexpr},'\">View ticket</a></p>'),'')}}"

def if_internal(cond_expr, yes, no=None, after=None):
    """Wrap a requester-facing Teams step so it only runs for internal requesters. `yes` is {name: action}."""
    return cond(f"@{cond_expr}", yes, no, after)

def nowET(fmt):
    return f"convertTimeZone(utcNow(),'UTC','{TZ}','{fmt}')"

SUPPORT_OPEN_H, SUPPORT_CLOSE_H = 7, 21
def sla_start_ticks(created_expr):
    """Ticks (UTC) of when the response timer starts for a ticket: its creation time, moved to 7:00 the same day
    if raised before 7am, or 7:00 the next day if raised at/after 9pm (support hours 7am-9pm Eastern)."""
    c = created_expr
    h = f"int(convertTimeZone({c},'UTC','{TZ}','HH'))"
    start_et = (f"if(less({h},{SUPPORT_OPEN_H}),concat(convertTimeZone({c},'UTC','{TZ}','yyyy-MM-dd'),'T07:00:00'),"
                f"if(greaterOrEquals({h},{SUPPORT_CLOSE_H}),concat(convertTimeZone(addDays({c},1),'UTC','{TZ}','yyyy-MM-dd'),'T07:00:00'),"
                f"convertTimeZone({c},'UTC','{TZ}','yyyy-MM-ddTHH:mm:ss')))")
    return f"ticks(convertTimeZone({start_et},'{TZ}','UTC'))"

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ALLDAYS = ["Sunday"] + WEEKDAYS + ["Saturday"]
DAY_H = 14   # HD6's hourly-clock day = 8am-9pm = 14 hours (distinct from SUPPORT_OPEN_H/CLOSE_H's 7am-9pm,
             # which is only used for the Response(min) "effective start" calc, not this counter)

def recurrence(hours, days=WEEKDAYS):
    return {"type": "Recurrence", "recurrence": {
        "frequency": "Week", "interval": 1, "timeZone": TZ,
        "schedule": {"weekDays": days,
                     "hours": [str(h) for h in hours], "minutes": [0]}}}

def sp_trigger(opid, table=TICKETS):
    return {"recurrence": {"frequency": "Minute", "interval": 1},
            "splitOn": "@triggerOutputs()?['body/value']", "type": "OpenApiConnection",
            "inputs": {"parameters": {"dataset": SITE, "table": table},
                       "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                                "connectionName": "shared_sharepointonline", "operationId": opid},
                       "authentication": "@parameters('$authentication')"}}

def spp(opid, params, after=None):
    """Like sp(), but against the IT Projects list instead of Tickets."""
    p = dict(params)
    p["table"] = PROJECTS
    return sp(opid, p, after)

def forms_trigger(form):
    return {"splitOn": "@triggerOutputs()?['body/value']", "type": "OpenApiConnectionWebhook",
            "inputs": {"parameters": {"form_id": form},
                       "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_microsoftforms",
                                "connectionName": "shared_microsoftforms", "operationId": "CreateFormWebhook"},
                       "authentication": "@parameters('$authentication')"}}

def holiday_check(name_prefix=""):
    return [
        ("Get_items_Holidays", sp("GetItems", {"table": HOLIDAYS,
            "$filter": "Title eq '@{" + nowET('yyyy-MM-dd') + "}'", "$top": 1})),
        ("Is_it_a_holiday", cond("@greater(length(body('Get_items_Holidays')?['value']),0)",
            {"Terminate": {"runAfter": {}, "type": "Terminate", "inputs": {"runStatus": "Succeeded"}}})),
    ]

PORTAL_URL = SITE + "/Lists/Tickets/My%20tickets.aspx"
_TD = 'style="border:1px solid #000;padding:6px 8px;font-family:Segoe UI,Arial,sans-serif;font-size:14px"'
_TH = 'style="border:1px solid #000;padding:6px 8px;font-family:Segoe UI,Arial,sans-serif;font-size:14px;font-weight:bold;text-align:center"'
_ROWS = [("Open", "Status on newly created Ticket"),
         ("Assigned", "Ticket has been re-assigned to an IT Team Member"),
         ("In Progress", "IT Support has begun work on the Ticket"),
         ("Escalated or Waiting Internally", "IT Support has reached out for assistance from an internal team"),
         ("Waiting for End User", "IT Support has requested further information from you"),
         ("To Be Scheduled", "IT Support has requested to schedule a meeting with you"),
         ("Resolved Pending Confirmation", "IT Support has provided resolution and awaiting confirmation that issue has been resolved"),
         ("Resolved", "Your issue has been resolved and the ticket is complete")]   # shown to requesters as Resolved; stored as Closed
CONFIRM_BODY = (
    '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333">'
    "<p>Hi @{first(split(trim(coalesce(§T§Requester/DisplayName'],first(split(outputs('ReqEmail'),'@')))),' '))},</p>"
    "<p><b>SUBJECT:</b> @{§T§Title']}<br><b>DESCRIPTION:</b> @{coalesce(§T§Details'],'')}</p>"
    "@{if(and(outputs('HasPic'),outputs('IsInternal')),concat('<p><b>IMAGE:</b> <a href=\"',outputs('PicUrl'),'\">View the image attached to your ticket</a></p>'),'')}"
    "<p>This email was system generated for the purpose of communicating that the referenced Ticket Number has been created to record your contact with the IT Team. "
    "To keep you informed on the progress of your Ticket, you will receive ongoing automated emails@{if(outputs('IsInternal'),' and Teams messages','')} when your IT Ticket Status changes as follows:</p>"
    '<table style="border-collapse:collapse;margin:8px 0 12px 0">'
    f"<tr><th {_TH}>TICKET STATUS</th><th {_TH}>MEANING</th></tr>"
    + "".join(f"<tr><td {_TD}>{a}</td><td {_TD}>{b}</td></tr>" for a, b in _ROWS) +
    "</table>"
    "@{if(outputs('IsInternal'),"
    f"'<p>Please feel free to use our <a href=\"{PORTAL_URL}\">Ticket Portal</a> anytime to view the Status of your Ticket. To add comments or attachments, reply to this email or reply in the IT Support chat in Teams.</p>"
    "<p>You will also receive Teams messages from <b>IT Support</b>. You can reply to IT in that chat or by replying to this email.</p>',"
    "'<p>To add comments or attachments, or to reply to IT, reply to this email and keep the ticket number in the subject.</p>')}"
    "<p><b>TICKET NUMBER:</b> IT-@{§ID§}<br>"
    "<b>TICKET STATUS:</b> @{if(empty(coalesce(§T§AssignedTo/Email'],'')),'Open','Assigned')}<br>"
    "<b>TICKET CONTACT:</b> @{coalesce(§T§Requester/DisplayName'],outputs('ReqEmail'))} (@{outputs('ReqEmail')})<br>"
    "<b>TICKET CREATED:</b> @{convertTimeZone(§T§Created'],'UTC','Eastern Standard Time','MMMM d, yyyy h:mm tt')}<br>"
    "<b>TICKET URGENCY:</b> @{§T§Priority/Value']}<br>"
    "<b>TICKET ASSIGNEE:</b> @{coalesce(§T§AssignedTo/DisplayName'],'Not yet assigned')}</p>"
    "</div>")

SIG_MARKERS = ["Regards", "Thks", "Thanks,", "Thank you,", "Cheers,", "Sent from my", "Get Outlook for", "[cid:", "[https://", "T: ", "________________"]
HARD_SIG_MARKERS = ["Regards", "Sent from my", "Get Outlook for", "[cid:", "[https://", "T: ", "________________"]
def strip_signature(expr, markers=None):
    """Cut text at the first signature marker. expr is an expression without a leading @."""
    out = expr
    for m in (markers or SIG_MARKERS):
        out = f"first(split({out},'{m}'))"
    return f"trim({out})"

SAFETY_TIPS = ["Learn why this is important", "Learn why this could be a risk"]
def strip_banners(expr):
    """Actions that remove Exchange's injected notices from the top of a message's text: a leading
    '[CAUTION: External E-Mail ...]' disclaimer and the first-contact safety tip, which ends in a
    'Learn why ...' link (the URL may follow as '<https://aka.ms/...>'). Done in steps because a single
    nested expression would exceed the 8 KB expression limit. The result is outputs('BodyText')."""
    steps = [("Body0", compose(f"@trim({expr})"))]
    p = "outputs('Body0')"
    steps.append(("Body1", compose(f"@trim(if(startsWith({p},'[CAUTION'),substring({p},add(indexOf({p},']'),1)),{p}))")))
    p = "outputs('Body1')"
    for i, tip in enumerate(SAFETY_TIPS):
        n = f"Body{i+2}"
        steps.append((n, compose(f"@if(contains({p},'{tip}'),trim(substring({p},add(indexOf({p},'{tip}'),{len(tip)}))),{p})")))
        p = f"outputs('{n}')"
    steps.append(("BodyText", compose(f"@trim(if(startsWith({p},'<http'),substring({p},add(indexOf({p},'>'),1)),{p}))")))
    return steps

def clean_text(expr):
    """Strip the signature. If that leaves nothing (the message itself starts with 'Thanks' etc.),
    cut only at the hard signature markers, and as a last resort keep the first 2000 characters."""
    return (f"if(not(empty({strip_signature(expr)})),{strip_signature(expr)},"
            f"if(not(empty({strip_signature(expr, HARD_SIG_MARKERS)})),{strip_signature(expr, HARD_SIG_MARKERS)},"
            f"substring({expr},0,min(length({expr}),2000))))")

USER_CARD_TYPE = "ITHelpDeskUserCard"
CHAT_NOTICE = "To reply to IT, reply in this chat or use Reply by email."

def user_card(ID, blocks, title_expr="triggerOutputs()?['body/Title']"):
    """Card sent to the requester in their Workflows chat. Replies go by email (HD1), because Teams
    does not pass card clicks from the Workflows chat to a flow."""
    body = [{"type": "TextBlock", "text": txt, "wrap": True, **({"weight": "Bolder", "size": "Medium"} if bold else {})}
            for txt, bold in blocks]
    body.append({"type": "TextBlock", "text": CHAT_NOTICE, "wrap": True, "size": "Small", "isSubtle": True})
    mailto = (f"mailto:{MBX}?subject=@{{encodeUriComponent(concat('[IT-',string({ID}),'] Reply: ',coalesce({title_expr},'')))}}"
              )
    return {"type": "AdaptiveCard", "version": "1.2", "body": body,
            "actions": [{"type": "Action.OpenUrl", "title": "Reply to IT", "url": mailto},
                        {"type": "Action.OpenUrl", "title": "View ticket", "url": f"{REQ_URL}@{{{ID}}}"},
                        {"type": "Action.OpenUrl", "title": "My tickets", "url": PORTAL_URL}]}

def user_card_post(card_compose_name, after=None):
    return op("shared_teams", "PostCardToConversation",
              {"poster": "Flow bot", "location": "Chat with Flow bot", "body/recipient": "@outputs('ReqEmail')",
               "body/messageBody": f"@{{string(outputs('{card_compose_name}'))}}"}, after)

# ---------- HD1 Email intake ----------
def hd1():
    T = "triggerOutputs()?['body/subject']"
    title_expr = f"@if(empty(coalesce({T},'')),'(no subject)',substring({T},0,min(length({T}),250)))"
    def attach_loop(item_id, after, suffix):
        return {"runAfter": after, "type": "Foreach", "foreach": "@body('Real_attachments')",
                "actions": {f"Add_attachment{suffix}": sp("CreateAttachment", {
                    "itemId": item_id, "displayName": f"@items('Apply_to_each{suffix}')?['name']",
                    "body": f"@base64ToBinary(items('Apply_to_each{suffix}')?['contentBytes'])"})}}

    new_item = {
        "item/Title": title_expr,
        "item/Details": "@outputs('CleanBody')",
        # an external address cannot be resolved to a person; keep RequesterEmail as the only identity for them
        "item/Requester/Claims": "@if(outputs('IsInternal'),outputs('Sender'),null)",
        "item/RequesterEmail": "@outputs('Sender')",
        "item/Source/Value": "Email",
        "item/Conversation": f"@concat('[',convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm'),'] Ticket created from ',if(outputs('IsInternal'),'Email','Email (external sender)'),' by ',outputs('Sender'),decodeUriComponent('%0A%0A'))",
        "item/Category/Value": "Uncategorized",
        "item/Priority/Value": f"@if(contains(toLower(coalesce({T},'')),'urgent'),'High','Medium')",
        "item/Status/Value": S_OPEN,
    }
    followup = dict(new_item)
    followup["item/Details"] = ("@if(equals(body('Get_item')?['Status']?['Value'],'Closed'),"
                                "concat('Follow-up to IT-',string(outputs('TicketId')),'. ',outputs('ReplyText'),"
                                "if(empty(coalesce(body('Get_item')?['Resolution'],'')),'',concat(decodeUriComponent('%0A%0A'),"
                                "'Previous resolution (IT-',string(outputs('TicketId')),'): ',body('Get_item')?['Resolution']))),outputs('CleanBody'))")

    gi = "body('Get_item')"
    LT = link("outputs('TicketId')")
    reply_yes = seq([
        ("NewStatus", compose(f"@if(contains({arr([S_WAIT, S_SCHED, S_RES])},{gi}?['Status']?['Value']),'{S_PROG}',{gi}?['Status']?['Value'])")),
        ("Update_item", sp("PatchItem", {
            "id": "@outputs('TicketId')",
            "item/Title": f"@{gi}?['Title']",
            "item/Status/Value": "@outputs('NewStatus')",
            "item/LastNotifiedStatus": "@outputs('NewStatus')",
            "item/WaitingHours": 0, "item/ResolvedHours": 0,
            "item/Conversation": ("@concat('[',convertTimeZone(utcNow(),'UTC','" + TZ + "','yyyy-MM-dd HH:mm'),' | ',triggerOutputs()?['body/from'],'] ',"
                                  "outputs('ReplyText'),decodeUriComponent('%0A%0A'),coalesce(" + gi + "?['Conversation'],''))")})),
        ("Apply_to_each", attach_loop("@outputs('TicketId')", None, "")),
        ("Has_a_tech", cond(f"@not(empty(coalesce({gi}?['AssignedTo']?['Email'],'')))", {
            "Tell_the_tech": scope_msg("Tell_tech_msg", f"@{{{gi}?['AssignedTo']?['Email']}}",
                f"<p>@{{triggerOutputs()?['body/from']}} replied on IT-@{{outputs('TicketId')}}:</p><p>@{{outputs('ReplyText')}}</p><p>{LT}</p>")})),
        ("Has_channel_post", cond(f"@not(empty(coalesce({gi}?['ChannelMessageId'],'')))", {
            "Reply_in_channel": channel_reply(f"@{gi}?['ChannelMessageId']",
                f"<p>@{{triggerOutputs()?['body/from']}} replied on IT-@{{outputs('TicketId')}}:</p><p>@{{outputs('ReplyText')}}</p>")})),
        ("Was_resolved", cond(f"@and(equals({gi}?['Status']?['Value'],'{S_RES}'),outputs('IsInternal'))", {
            "Tell_requester_reopened": {"type": "Scope", "runAfter": {}, "actions": support_msg("Post_reopened", "@outputs('Sender')",
                f"<p><b>IT-@{{outputs('TicketId')}} has been reopened</b></p><p>@{{{gi}?['Title']}}</p>" + support_links("outputs('TicketId')", f"{gi}?['Title']"))}})),
    ])
    reply_no = seq([
        ("Create_followup_item", sp("PostItem", followup)),
        ("Apply_to_each_2", attach_loop("@body('Create_followup_item')?['ID']", None, "_2")),
    ])
    reply_branch = seq([
        ("TicketId", compose(f"@if(contains(coalesce({T},''),'[IT-'),int(first(split(last(split({T},'[IT-')),']'))),outputs('RefNum'))")),
        ("ReplyText", compose("@" + clean_text(f"first(split(first(split(outputs('BodyText'),'{MARKER}')),'From: '))"))),
        ("Get_item", sp("GetItem", {"id": "@outputs('TicketId')"})),
        ("Update_existing_ticket", cond(
            f"@and(not(empty(string(coalesce({gi}?['ID'],'')))),not(equals({gi}?['Status']?['Value'],'Closed')),"
            f"or(equals(outputs('Sender'),toLower(coalesce({gi}?['RequesterEmail'],''))),equals(outputs('Sender'),'{ME}')))",
            reply_yes, reply_no)),
    ])
    reply_branch["Update_existing_ticket"]["runAfter"] = ra("Get_item", status=("Succeeded", "Failed"))
    new_branch = seq([
        ("Create_item", sp("PostItem", new_item)),
        ("Apply_to_each_3", attach_loop("@body('Create_item')?['ID']", None, "_3")),
    ])

    # A reply to a [PRJ-n] project notification (HD19/HD20's "Reply to this email to add a note") lands here
    # too, since both lists share the one shared-mailbox trigger. Only the project's Sponsor, Owner, creator or
    # Martin may add a Sponsor Updates note this way; anyone else (or a PRJ-n that doesn't exist) alerts Martin
    # instead of silently creating a bogus ticket titled after the project email.
    gp = "body('Get_project')"
    sponsor_arr, owner_arr, sender = gp + "?['ProjectSponsor']", gp + "?['Owner']", "outputs('Sender')"
    # Owner and Project Sponsor both allow multiple people - the SharePoint "Get item" action exposes a
    # multi-value Person field as an array of {Email,DisplayName,...} items (see hd19()'s docstring), so
    # membership is an exact-match array check (people_contains), not a single ?['Email'] lookup. SponsorEmail
    # is the legacy plain-text fallback, checked only when no Project Sponsor person is set.
    prj_auth = (f"or({people_contains(sponsor_arr, sender)},"
                f"equals({sender},toLower(coalesce({gp}?['SponsorEmail'],''))),"
                f"{people_contains(owner_arr, sender)},"
                f"equals({sender},toLower(coalesce({gp}?['Author']?['Email'],''))),"
                f"equals({sender},'{ME}'))")
    prj_stamp = f"convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm')"
    project_reply_yes = seq([
        ("Update_project", sp("PatchItem", {"table": PROJECTS, "id": "@outputs('PrjId')",
            "item/Title": f"@{gp}?['Title']",
            "item/SponsorUpdates": (f"@concat('[',{prj_stamp},' | ',outputs('Sender'),'] ',outputs('PrjReplyText'),"
                                     f"decodeUriComponent('%0A%0A'),coalesce({gp}?['SponsorUpdates'],''))")})),
        # Confirm back to the sender, the email-side mirror of HD12's Teams "Thank_sponsor" reply - until this was
        # added, a sponsor replying by email got no acknowledgment at all that their note was received (only the
        # Teams-reply path did), unlike every other reply path in the system. Found 2026-10-25 when Martin asked
        # for it after confirming the Teams side worked.
        ("Thank_sponsor_email", mail("@outputs('Sender')",
            f"[PRJ-@{{outputs('PrjId')}}] Update added: @{{{gp}?['Title']}}",
            f"<p>Thanks, your update was added to <b>PRJ-@{{outputs('PrjId')}}: @{{{gp}?['Title']}}</b>.</p>")),
    ])
    project_reply_notfound = {
        "Alert_bad_project_reply": {**mail(ME, "[Help desk health] Unmatched project reply",
            f"<p>A reply naming PRJ-@{{outputs('PrjId')}} from @{{outputs('Sender')}} couldn't be matched to a "
            f"project they're allowed to update (sponsor, owner, creator, or you). Subject: @{{{T}}}</p>"
            f"<p>@{{outputs('PrjReplyText')}}</p>"), "runAfter": {}},
        "Alert_bad_project_reply_teams": {"type": "Scope", "runAfter": ra("Alert_bad_project_reply", status=("Succeeded", "Failed")),
            "actions": support_msg("Alert_bad_project_reply_dm", ME,
                f"<p>A reply naming PRJ-@{{outputs('PrjId')}} from @{{outputs('Sender')}} couldn't be auto-matched. Check email for details.</p>")},
    }
    project_reply_branch = seq([
        ("PrjId", compose(f"@if(contains(coalesce({T},''),'[PRJ-'),int(first(split(last(split({T},'[PRJ-')),']'))),outputs('PrjRefNum'))")),
        ("PrjReplyText", compose("@" + clean_text(f"first(split(first(split(outputs('BodyText'),'{MARKER}')),'From: '))"))),
        ("Get_project", sp("GetItem", {"table": PROJECTS, "id": "@outputs('PrjId')"})),
        ("Project_found_and_authorized", cond(
            f"@and(not(empty(string(coalesce({gp}?['ID'],'')))),{prj_auth})",
            project_reply_yes, project_reply_notfound)),
    ])
    project_reply_branch["Project_found_and_authorized"]["runAfter"] = ra("Get_project", status=("Succeeded", "Failed"))

    actions = seq([
        ("Sender", compose("@toLower(trim(coalesce(triggerOutputs()?['body/from'],'')))")),
        ("SenderDomain", compose("@last(split(outputs('Sender'),'@'))")),
        ("IsInternal", compose("@" + is_internal("outputs('Sender')"))),
        # External senders: allowed only when their address or domain is on the "External Senders" list.
        # Read the whole (small) list and compare in an expression: no OData quoting, case-insensitive.
        ("Get_external_senders", sp("GetItems", {"table": EXTERNAL_SENDERS, "$top": 500})),
        ("Allowed_senders", {"type": "Select", "runAfter": ra("Get_external_senders", status=("Succeeded", "Failed")),
                             "inputs": {"from": "@coalesce(body('Get_external_senders')?['value'],json('[]'))",
                                        "select": "@toLower(trim(coalesce(item()?['Title'],'')))"}}),
        ("IsAllowedExternal", compose("@or(contains(body('Allowed_senders'),outputs('Sender')),contains(body('Allowed_senders'),outputs('SenderDomain')))")),
        # skip inline images (signature logos); only real attachments go on the ticket
        ("Real_attachments", {"type": "Query", "inputs": {"from": "@coalesce(triggerOutputs()?['body/attachments'],json('[]'))",
                                                          "where": "@not(equals(item()?['isInline'],true))"}}),
        # Power Automate's own alerts ("Your flow 'HD3' failed", "connection isn't working", weekly digests) arrive
        # in the IT Support mailbox because IT Support owns the flows. Relay them to Martin by email and Teams.
        ("Is_a_Power_Automate_alert", cond(
            f"@and(endsWith(outputs('Sender'),'@microsoft.com'),or(contains(toLower(coalesce({T},'')),'flow'),contains(toLower(coalesce({T},'')),'isn''t working'),contains(toLower(coalesce({T},'')),'failed')))",
            seq([
                ("Relay_alert_email", mail(ME, f"[Help desk health] @{{{T}}}",
                    "<p style=\"color:#888;font-size:12px\">Relayed from the IT Support mailbox (Power Automate alert).</p>@{triggerOutputs()?['body/body']}")),
                ("Relay_alert_teams", {"type": "Scope", "runAfter": ra("Relay_alert_email", status=("Succeeded", "Failed")),
                    "actions": support_msg("Relay_alert_dm", ME,
                        f"<p><b>Power Automate alert:</b> @{{{T}}}</p><p>Details are in your email. <a href=\"https://make.powerautomate.com/environments/{ENV}/flows\">Open the flows</a></p>")}),
                ("Move_alert_to_Review", {**move(F_REVIEW), "runAfter": ra("Relay_alert_teams", status=("Succeeded", "Failed"))}),
                ("Terminate_alert", {"type": "Terminate", "runAfter": ra("Move_alert_to_Review", status=("Succeeded", "Failed")), "inputs": {"runStatus": "Succeeded"}}),
            ]))),
        # Microsoft Teams "missed activity" relay: when a Teams chat sent to IT_Support's own Teams presence goes
        # unread, Teams emails a notification here (sender always contains 'teams.mail.microsoft', subject "<name>
        # sent a message" / "<name> is trying to reach you in Microsoft Teams"). That sender is never Branthaven
        # and is never going to be on the External Senders list (it's Microsoft's relay address, not the real
        # person's), so without this check it fell straight into the unknown-sender branch below: silently moved
        # to Review and terminated, with no alert at all. Found 2026-10-26: Andrea Royers' "Hi Martin, I'm still
        # not able to access Andre's email" (sent via Teams chat, Fri Oct 23) sat unread and unactioned in Review
        # for 3 days — nobody knew it existed until the Review folder was checked by hand, and by then it had also
        # aged out of HD16's 24h "mail landed in Review" window, so even the dashboard never flagged it. Mirrors
        # the Is_a_Power_Automate_alert relay just above: can't become a real ticket (no reliable sender identity
        # for whoever actually sent the Teams message), so it still parks in Review for the audit trail, but Martin
        # now gets told immediately instead of finding out by accident.
        ("Is_a_Teams_missed_activity", cond(
            "@contains(outputs('Sender'),'teams.mail.microsoft')",
            seq([
                ("Relay_teams_email", mail(ME, f"[Help desk health] Missed Teams message: @{{{T}}}",
                    "<p style=\"color:#888;font-size:12px\">Relayed from the IT Support mailbox — someone messaged "
                    "IT Support in Teams and it went unread.</p>@{triggerOutputs()?['body/body']}")),
                ("Relay_teams_teams", {"type": "Scope", "runAfter": ra("Relay_teams_email", status=("Succeeded", "Failed")),
                    "actions": support_msg("Relay_teams_dm", ME,
                        f"<p><b>Missed Teams message:</b> @{{{T}}}</p><p>Someone messaged IT Support in Teams and it "
                        "wasn't picked up. Full details are in your email.</p>")}),
                ("Move_teams_to_Review", {**move(F_REVIEW), "runAfter": ra("Relay_teams_teams", status=("Succeeded", "Failed"))}),
                ("Terminate_teams", {"type": "Terminate", "runAfter": ra("Move_teams_to_Review", status=("Succeeded", "Failed")), "inputs": {"runStatus": "Succeeded"}}),
            ]))),
        ("Is_it_a_real_email_from_a_known_sender", cond(
            f"@and(or(outputs('IsInternal'),outputs('IsAllowedExternal')),not(equals(outputs('Sender'),toLower('{MBX}'))),"
            f"not(startsWith(coalesce({T},''),'Automatic reply')),not(startsWith(coalesce({T},''),'Undeliverable')))",
            {}, seq([("Move_to_Review", move(F_REVIEW)),
                     ("Terminate", {"type": "Terminate", "inputs": {"runStatus": "Succeeded"}})]))),
        ("Html_to_text", op("shared_conversionservice", "HtmlToText", {"Content": "@triggerOutputs()?['body/body']"})),
        # drop the banners Exchange injects at the top of outside mail: the external-mail CAUTION notice and the
        # "first contact" safety tip ("X appears similar to someone... Learn why this could be a risk")
        *strip_banners("coalesce(body('Html_to_text'),'')"),
        ("CleanBody", compose("@" + clean_text("outputs('BodyText')"))),
        # safety net: no [IT-n] in the subject, but the subject or body names IT-n (first occurrence; digits only)
        ("RefPart", compose(f"@concat(coalesce(split(toUpper(concat(coalesce({T},''),' ',coalesce(body('Html_to_text'),''))),'IT-')?[1],''),'     ')")),
        ("RefNum", compose("@if(not(contains('0123456789',substring(outputs('RefPart'),0,1))),0,"
                           "int(substring(outputs('RefPart'),0,if(not(contains('0123456789',substring(outputs('RefPart'),1,1))),1,"
                           "if(not(contains('0123456789',substring(outputs('RefPart'),2,1))),2,"
                           "if(not(contains('0123456789',substring(outputs('RefPart'),3,1))),3,"
                           "if(not(contains('0123456789',substring(outputs('RefPart'),4,1))),4,5)))))))")),
        # same safety net, for a reply that names a project instead of a ticket
        ("PrjRefPart", compose(f"@concat(coalesce(split(toUpper(concat(coalesce({T},''),' ',coalesce(body('Html_to_text'),''))),'PRJ-')?[1],''),'     ')")),
        ("PrjRefNum", compose("@if(not(contains('0123456789',substring(outputs('PrjRefPart'),0,1))),0,"
                           "int(substring(outputs('PrjRefPart'),0,if(not(contains('0123456789',substring(outputs('PrjRefPart'),1,1))),1,"
                           "if(not(contains('0123456789',substring(outputs('PrjRefPart'),2,1))),2,"
                           "if(not(contains('0123456789',substring(outputs('PrjRefPart'),3,1))),3,"
                           "if(not(contains('0123456789',substring(outputs('PrjRefPart'),4,1))),4,5)))))))")),
        ("Is_project_or_ticket_reply", cond(
            f"@or(contains(coalesce({T},''),'[PRJ-'),greater(outputs('PrjRefNum'),0))", project_reply_branch,
            {"Is_it_a_reply": cond(f"@or(contains(coalesce({T},''),'[IT-'),greater(outputs('RefNum'),0))", reply_branch, new_branch)})),
        ("Move_to_Processed", move(F_PROCESSED)),
    ])
    trigger = {"When_a_new_email_arrives_in_a_shared_mailbox_(V2)": {
        "recurrence": {"frequency": "Minute", "interval": 1},
        "splitOn": "@triggerOutputs()?['body/value']", "type": "OpenApiConnection",
        "inputs": {"parameters": {"mailboxAddress": MBX, "includeAttachments": True},
                   "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                            "connectionName": "shared_office365", "operationId": "SharedMailboxOnNewEmailV2"},
                   "authentication": "@parameters('$authentication')"}}}
    return "HD1 Email intake", trigger, actions

# ---------- HD2 Form intake ----------
def hd2():
    R = "body('Get_response_details')"
    urg = f"{R}?['{Q_URG}']"
    actions = seq([
        ("Get_response_details", op("shared_microsoftforms", "GetFormResponseById",
            {"form_id": FORM1, "response_id": "@triggerOutputs()?['body/resourceData/responseId']"})),
        ("Create_item", sp("PostItem", {
            "item/Title": f"@substring({R}?['{Q_WHAT}'],0,min(length({R}?['{Q_WHAT}']),250))",
            "item/Category/Value": f"@{R}?['{Q_CAT}']",
            "item/Priority/Value": f"@if(startsWith(coalesce({urg},''),'I can'),'High',if(equals({urg},'No rush'),'Low','Medium'))",
            "item/Details": f"@{R}?['{Q_DET}']",
            "item/Requester/Claims": f"@{R}?['responder']",
            "item/RequesterEmail": f"@toLower({R}?['responder'])",
            "item/Source/Value": "Teams form",
            "item/Conversation": f"@concat('[',convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm'),'] Ticket created from Teams form by ',toLower({R}?['responder']),decodeUriComponent('%0A%0A'))",
            "item/Status/Value": S_OPEN})),
        ("Has_a_screenshot", cond(f"@not(empty(coalesce({R}?['{Q_SHOT}'],'')))", {
            "Apply_to_each": {"runAfter": {}, "type": "Foreach", "foreach": f"@json({R}?['{Q_SHOT}'])",
                "actions": seq([
                    ("Get_file_content_using_path", op("shared_sharepointonline", "GetFileContentByPath", {
                        "dataset": "@first(split(items('Apply_to_each')?['link'],'/Shared%20Documents'))",
                        "path": "@concat('/Shared Documents',decodeUriComponent(first(split(last(split(items('Apply_to_each')?['link'],'/Shared%20Documents')),'?'))))",
                        "inferContentType": True})),
                    ("Add_attachment", sp("CreateAttachment", {
                        "itemId": "@body('Create_item')?['ID']",
                        "displayName": "@items('Apply_to_each')?['name']",
                        "body": "@body('Get_file_content_using_path')"})),
                ])}})),
    ])
    return "HD2 Form intake", {"When_a_new_response_is_submitted": forms_trigger(FORM1)}, actions

# ---------- HD3 Ticket created ----------
def hd3():
    t = lambda k: f"triggerOutputs()?['body/{k}']"
    ID = t('ID')
    actions = seq([
        ("ReqEmail", compose(f"@toLower(coalesce({t('RequesterEmail')},{t('Requester/Email')},''))")),
        ("IsInternal", compose("@" + is_internal("outputs('ReqEmail')"))),
        ("Tag", compose(f"@concat('IT-',{ID})")),
        ("Get_owner_group", sp_http("GET", "_api/web/associatedownergroup")),
        ("Break_inheritance", sp_http("POST", f"_api/web/lists(guid'{TICKETS}')/items(@{{{ID}}})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")),
        ("Give_owners_full_control", sp_http("POST", f"_api/web/lists(guid'{TICKETS}')/items(@{{{ID}}})/roleassignments/addroleassignment(principalid=@{{body('Get_owner_group')?['Id']}},roledefid=1073741829)")),
        # external senders have no SharePoint identity: nothing to grant, they work the ticket by email
        ("Has_requester", cond("@and(not(empty(outputs('ReqEmail'))),outputs('IsInternal'))", seq([
            ("Ensure_requester", sp_http("POST", "_api/web/ensureuser",
                body="{\"logonName\":\"i:0#.f|membership|@{outputs('ReqEmail')}\"}")),
            ("Give_requester_read", sp_http("POST", f"_api/web/lists(guid'{TICKETS}')/items(@{{{ID}}})/roleassignments/addroleassignment(principalid=@{{body('Ensure_requester')?['Id']}},roledefid=1073741826)")),
        ]))),
        ("Create_a_task", op("shared_planner", "CreateTask_V3", {
            "body/groupId": GROUP, "body/planId": PLAN,
            "body/title": f"@{{outputs('Tag')}} @{{{t('Title')}}}",
            "body/bucketId": f"@if(empty(coalesce({t('AssignedTo/Email')},'')),'{B_NEW}','{B_PROG}')"})),
        ("Has_tech_at_creation", cond(f"@not(empty(coalesce({t('AssignedTo/Email')},'')))", {
            "Add_assignees_to_a_task": op("shared_planner", "AssignUsers", {
                "id": "@body('Create_a_task')?['id']", "body/assignments": f"@{t('AssignedTo/Email')}"})})),
        ("Update_task_details", op("shared_planner", "UpdateTaskDetails_V2", {
            "id": "@body('Create_a_task')?['id']",
            "body/description": (f"Requester: @{{outputs('ReqEmail')}}\nCategory: @{{{t('Category/Value')}}}\nPriority: @{{{t('Priority/Value')}}}\n"
                                 f"Source: @{{{t('Source/Value')}}}\n\n@{{{t('Details')}}}\n\n@{{if(outputs('HasPic'),concat('Image: ',outputs('PicUrl'),decodeUriComponent('%0A%0A')),'')}}{LIST_URL}@{{{ID}}}")})),
        ("Wait_for_attachment", {"type": "Wait", "inputs": {"interval": {"count": 10, "unit": "Second"}}}),
        ("Get_ticket", sp("GetItem", {"id": f"@{ID}"})),
        ("HasPic", compose("@equals(body('Get_ticket')?['{HasAttachments}'],true)")),
        ("Get_attachments", sp_http("GET", f"_api/web/lists(guid'{TICKETS}')/items(@{{{ID}}})/AttachmentFiles")),
        ("PicUrl", compose("@if(outputs('HasPic'),concat('https://branthavenhomes.sharepoint.com',coalesce(first(body('Get_attachments')?['value'])?['ServerRelativeUrl'],'')),'')")),
        ("IsOpenNow", compose(
            f"@and(greaterOrEquals(int(convertTimeZone(utcNow(),'UTC','{TZ}','HH')),7),less(int(convertTimeZone(utcNow(),'UTC','{TZ}','HH')),21))")),
        ("HoursLine", compose("@if(outputs('IsOpenNow'),'A technician will pick it up shortly.','Support hours are 7:00 am to 9:00 pm, every day. Your ticket will be actioned when support reopens at 7:00 am.')")),
        ("Tell_requester", cond("@not(empty(outputs('ReqEmail')))", seq([
            ("Send_confirmation_email", mail("@outputs('ReqEmail')",
                f"[IT-@{{{ID}}}] We've got your request: @{{{t('Title')}}}",
                CONFIRM_BODY.replace("§ID§", ID).replace("§T§", "triggerOutputs()?['body/"))),
            ("Confirm_in_Teams_if_internal", {**if_internal("outputs('IsInternal')", {
                "Confirm_in_Teams": {"type": "Scope", "runAfter": {},
                    "actions": support_msg("Send_confirmation_chat", "@outputs('ReqEmail')",
                        f"<p><b>IT-@{{{ID}}} created:</b> @{{{t('Title')}}}</p>"
                        f"<p>Status: @{{if(empty(coalesce({t('AssignedTo/Email')},'')),'Open','Assigned')}} · Urgency: @{{{t('Priority/Value')}}}</p>"
                        "@{if(outputs('HasPic'),concat('<p><a href=\"',outputs('PicUrl'),'\">View image</a></p>'),'')}"
                        "<p>You'll get a message here and an email each time the status changes.</p>" + support_links(ID, t('Title')))}}),
                "runAfter": ra("Send_confirmation_email", status=("Succeeded", "Failed"))}),
        ]))),
        ("CreatedLine", compose(f"@concat('[',convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm'),'] Ticket created from ',coalesce({t('Source/Value')},'IT'),' by ',outputs('ReqEmail'),decodeUriComponent('%0A%0A'),coalesce({t('Conversation')},''))")),
        # requester photo for the card (Branthaven users only): 48x48 from Microsoft 365, embedded as a data URI.
        # Best effort - no photo, external sender or a Graph refusal just leaves the card without a picture.
        ("Init_varPhoto", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varPhoto", "type": "string", "value": ""}]}}),
        ("Requester_photo", cond("@outputs('IsInternal')", seq([
            ("Get_photo", op("shared_office365users", "HttpRequest", {"Method": "GET",
                "Uri": "https://graph.microsoft.com/v1.0/users/@{outputs('ReqEmail')}/photos/48x48/$value"})),
            ("Set_varPhoto", {"type": "SetVariable", "inputs": {"name": "varPhoto",
                "value": "@concat('data:image/jpeg;base64,',base64(body('Get_photo')))"}}),
        ]))),
        ("Card", {**compose(action_card(ID, t('Title'), "outputs('ReqEmail')", t('Category/Value'), t('Priority/Value'), t('Source/Value'),
            # pad before substring: a ticket with no Details (e.g. one created without that field) leaves
            # coalesce(Details,'') empty, and substring('',0,0) throws - see CLAUDE.md's substring gotcha.
            f"substring(concat(coalesce({t('Details')},''),' '),0,min(length(coalesce({t('Details')},'')),400))", pic="@outputs('HasPic')",
            photo="variables('varPhoto')", req_name=f"coalesce({t('Requester/DisplayName')},outputs('ReqEmail'))")),
            "runAfter": ra("Requester_photo", status=("Succeeded", "Failed", "Skipped", "TimedOut"))}),
        ("Post_to_Help_Desk_channel", channel_card("@{string(outputs('Card'))}", card_type=CARD_TYPE)),
        ("Update_item", sp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{t('Title')}",
            "item/RequesterEmail": "@outputs('ReqEmail')",
            "item/PlannerTaskId": "@body('Create_a_task')?['id']",
            "item/ChannelMessageId": "@body('Post_to_Help_Desk_channel')?['id']",
            "item/Conversation": f"@if(empty(coalesce({t('Conversation')},'')),outputs('CreatedLine'),{t('Conversation')})"})),
    ])
    # The requester's confirmation must not depend on Planner or the channel: the Planner steps sit in their own
    # scope, after the confirmation, and everything after them runs whether they worked or not.
    ANY3 = ("Succeeded", "Failed", "Skipped")
    planner = {n: actions.pop(n) for n in ("Create_a_task", "Has_tech_at_creation", "Update_task_details")}
    planner["Create_a_task"]["runAfter"] = {}
    planner["Has_tech_at_creation"]["runAfter"] = ra("Create_a_task", status=ANY3)
    planner["Update_task_details"]["runAfter"] = ra("Has_tech_at_creation", status=ANY3)
    actions["Wait_for_attachment"]["runAfter"] = ra("Has_requester", status=ANY3)
    actions["Tell_requester"]["runAfter"] = ra("HoursLine", status=ANY3)
    actions["Planner_task"] = {"type": "Scope", "runAfter": ra("Tell_requester", status=ANY3), "actions": planner}
    actions["CreatedLine"]["runAfter"] = ra("Planner_task", status=ANY3)
    actions["Update_item"]["runAfter"] = ra("Post_to_Help_Desk_channel", status=("Succeeded", "Failed"))
    # Direct 1:1 Teams message to the on-duty people, after the record is complete; never blocks the ticket.
    new_ping = (f"<p><b>New ticket IT-@{{{ID}}}</b> — @{{{t('Title')}}}</p>"
                f"<p>From: @{{outputs('ReqEmail')}} &nbsp;·&nbsp; Priority: @{{coalesce({t('Priority/Value')},'Medium')}}"
                f" &nbsp;·&nbsp; Category: @{{coalesce({t('Category/Value')},'')}}</p>"
                f"<p>{link(ID)}</p>")
    actions["Tell_IT_new_ticket"] = {"type": "Scope", "runAfter": ra("Update_item", status=ANY3),
                                     "actions": team_msgs("NewTicket_msg", new_ping, NEW_TICKET_PING_TEAM)}
    # Monitoring: did the requester confirmation, Planner task, channel post and tech ping all go out?
    no_channel = ("if(empty(body('Post_to_Help_Desk_channel')?['id']),"
                  "createArray(json('{\"name\":\"Post_to_Help_Desk_channel\",\"status\":\"Failed\"}')),json('[]'))")
    actions.update(self_check("Health", ["Tell_requester", "Planner_task", "Tell_IT_new_ticket"],
        f"[Help desk health] Notification problem on new ticket IT-@{{{ID}}}",
        f"Ticket <b>IT-@{{{ID}}}</b> (@{{{t('Title')}}}, requester @{{outputs('ReqEmail')}}) was created, but these steps did not complete:",
        f"concat('IT-',{ID},' (created)')", extra=no_channel, leaf=("Send_confirmation_email", "Send_confirmation_chat", "Create_a_task", "Post_to_Help_Desk_channel"), after=ra("Tell_IT_new_ticket", status=ANY3)))
    return "HD3 Ticket created", {"When_an_item_is_created": sp_trigger("GetOnNewItems")}, actions

# ---------- HD4 Ticket updated ----------
def hd4():
    t = lambda k: f"triggerOutputs()?['body/{k}']"
    ID = t('ID')
    hasTask = f"@not(empty(coalesce({t('PlannerTaskId')},'')))"
    lnk = link(ID)
    V = "variables('varStatus')"
    stamp = f"convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm')"
    # Changes made through the Teams card are saved by the IT Support account; credit the assigned tech instead
    who = (f"if(equals(toLower(coalesce({t('Editor/Email')},'')),'{MBX.lower()}'),"
           f"coalesce({t('AssignedTo/DisplayName')},'IT'),coalesce({t('Editor/DisplayName')},'IT'))")
    first = "outputs('FirstName')"
    # portal mentions only make sense for Branthaven users; external senders work the ticket by email
    portal_note = f"@{{if(outputs('IsInternal'),' You can also view your ticket in the <a href=\"{PORTAL_URL}\">Ticket Portal</a>.','')}}"
    portal_line = f"@{{if(outputs('IsInternal'),'<p>You can view your ticket anytime in the <a href=\"{PORTAL_URL}\">Ticket Portal</a>.</p>','')}}"
    # requesters never see the word "Closed": the final status is shown to them as "Resolved"
    UV = f"if(equals({V},'{S_CLOSED}'),'Resolved',{V})"
    details = ("<p><b>TICKET NUMBER:</b> IT-@{" + ID + "}<br>"
               "<b>TICKET STATUS:</b> @{" + UV + "}<br>"
               "<b>TICKET ASSIGNEE:</b> @{coalesce(" + t('AssignedTo/DisplayName') + ",'Not yet assigned')}<br>"
               "<b>TICKET URGENCY:</b> @{" + t('Priority/Value') + "}</p>")
    _hist = ("replace(replace(replace(substring(coalesce(triggerOutputs()?['body/Conversation'],''),0,min(length(coalesce(triggerOutputs()?['body/Conversation'],'')),3000)),"
             "decodeUriComponent('%0A%0A'),'<br>'),'Status: Closed','Status: Resolved'),'Closed automatically','Resolved automatically')")
    earlier = ("@{if(empty(coalesce(triggerOutputs()?['body/Conversation'],'')),'',concat('<p style=\"margin-top:18px\"><b>Ticket history</b></p>"
               "<div style=\"font-size:13px;color:#555;line-height:1.5\">'," + _hist + ",'</div>'))}")
    style = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333">'
    # Direct download links for whatever's attached to the ticket right now (fetched fresh each run via
    # Get_attachments below, not from the trigger body, since an attachment can be added without touching
    # Status/AssignedTo/NoteToUser). Folded into every note/status notification so the requester doesn't have
    # to open the ticket just to find a file a tech attached.
    files_html = f"@{{if(outputs('HasFiles'),concat('<p><b>Attached file(s):</b> ',join(body('FileLinks'),' &nbsp;&middot;&nbsp; '),'</p>'),'')}}"
    # On the final status (Resolved Pending Confirmation or Closed) — the one message a requester is most
    # likely to keep or forward — list every attachment currently on the ticket, whatever the requester
    # originally sent plus whatever a tech added, not just the ones not yet linked. Everywhere else (a note
    # on its own, or an earlier status change) keep linking only new files via files_html above, so the same
    # file isn't repeated on every single update. AllFileLinks (built below from every item in Get_attachments,
    # unlike FileLinks which only covers New_attachments) backs the Resolved/Closed branch.
    status_files_html = (f"@{{if(contains({arr([S_RES, S_CLOSED])},{V}),"
                          f"if(greater(length(coalesce(body('Get_attachments')?['value'],json('[]'))),0),"
                          f"concat('<p><b>Attachment(s):</b> ',join(body('AllFileLinks'),' &nbsp;&middot;&nbsp; '),'</p>'),''),"
                          f"if(outputs('HasFiles'),concat('<p><b>Attached file(s):</b> ',join(body('FileLinks'),' &nbsp;&middot;&nbsp; '),'</p>'),''))}}")
    # True exactly when a note/status notification with files_html in it is actually going out this run — the
    # same two conditions that gate Email_the_note (note, same status) and Email_status (status changed, has a
    # requester email). Only then does Update_item roll NotifiedAttachments forward; otherwise a run with no
    # outgoing notification (e.g. a pure reassignment) would mark files "seen" without ever showing them.
    files_shown = (f"or(and(not(empty(coalesce({t('NoteToUser')},''))),equals({V},coalesce({t('LastNotifiedStatus')},''))),"
                   f"and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),not(empty(outputs('ReqEmail')))))")

    # Part B: a tech was assigned (runs first so Open -> Assigned feeds the status email)
    partB = seq([
        ("Has_task_B", cond(hasTask, seq([
            ("Had_previous_tech", cond(f"@not(empty(coalesce({t('LastNotifiedTech')},'')))", {
                "Remove_assignees_from_a_task": op("shared_planner", "UnassignUsers", {
                    "id": f"@{t('PlannerTaskId')}", "body/assignments": f"@{t('LastNotifiedTech')}"})})),
            ("Add_assignees_to_a_task", {**op("shared_planner", "AssignUsers", {
                "id": f"@{t('PlannerTaskId')}", "body/assignments": "@outputs('Tech')"}),
                "runAfter": ra("Had_previous_tech", status=("Succeeded", "Failed"))}),
        ]))),
        ("Is_open", {**cond(f"@equals({V},'{S_OPEN}')", {
            "Set_Assigned": {"type": "SetVariable", "inputs": {"name": "varStatus", "value": S_ASSIGNED}}}),
            "runAfter": ra("Has_task_B", status=("Succeeded", "Failed", "Skipped"))}),
        ("First_response", cond("@empty(variables('varFirstResponse'))", {
            "Set_varFirstResponse": {"type": "SetVariable", "inputs": {"name": "varFirstResponse", "value": "@{utcNow()}"}}})),
        ("Reassigned_same_status", cond(f"@and(not(empty(coalesce({t('LastNotifiedTech')},''))),equals({V},coalesce({t('LastNotifiedStatus')},'')),outputs('IsInternal'))", {
            "Tell_requester_new_tech": {"type": "Scope", "runAfter": {}, "actions": support_msg("Post_new_tech", "@outputs('ReqEmail')",
                f"<p><b>IT-@{{{ID}}} is now being handled by @{{{t('AssignedTo/DisplayName')}}}</b></p><p>@{{{t('Title')}}}</p>" + support_links(ID, t('Title')))}})),
        ("AssignedLine", {**compose(f"@concat('[',{stamp},' | ',{who},'] Assigned to ',{t('AssignedTo/DisplayName')},decodeUriComponent('%0A%0A'),variables('varConversation'))"),
                          "runAfter": ra("Reassigned_same_status", status=("Succeeded", "Failed", "Skipped"))}),
        ("Set_varConversation_assigned", {"type": "SetVariable", "inputs": {"name": "varConversation", "value": "@outputs('AssignedLine')"}}),
    ])

    # Part A: a note for the user. Logged always; sent on its own only when the status is not changing
    partA = seq([
        ("NoteLine", compose(f"@concat('[',{stamp},' | ',{who},' (IT)] ',{t('NoteToUser')},decodeUriComponent('%0A%0A'),variables('varConversation'))")),
        ("Set_varConversation_note", {"type": "SetVariable", "inputs": {"name": "varConversation", "value": "@outputs('NoteLine')"}}),
        ("Send_note_on_its_own", cond(f"@equals({V},coalesce({t('LastNotifiedStatus')},''))", seq([
            ("Email_the_note", mail("@outputs('ReqEmail')", f"[IT-@{{{ID}}}] Update on your ticket: @{{{t('Title')}}}",
                style + f"<p style=\"color:#999;font-size:12px\">{MARKER}</p><p>Hi @{{{first}}},</p><p>@{{{who}}} from IT wrote:</p><p>@{{{t('NoteToUser')}}}</p>"
                + files_html +
                f"<p>Reply to this email to answer.{portal_note}</p>" + details + earlier + "</div>")),
            ("Chat_the_note_if_internal", {**if_internal("outputs('IsInternal')", {
                "Chat_the_note": {"type": "Scope", "runAfter": {},
                    "actions": support_msg("Post_note", "@outputs('ReqEmail')",
                        f"<p><b>IT-@{{{ID}}}: message from @{{{who}}}</b></p><p>@{{{t('NoteToUser')}}}</p>" + files_html
                        + f"<p style=\"color:#666\">@{{{t('Title')}}}</p>" + support_links(ID, t('Title')))}}),
                "runAfter": ra("Email_the_note", status=("Succeeded", "Failed"))}),
            ("Note_in_channel", {**cond(f"@not(empty(coalesce({t('ChannelMessageId')},'')))", {
                "Reply_note_in_channel": channel_reply(f"@{t('ChannelMessageId')}",
                    f"<p>Note to user from @{{{who}}}: @{{{t('NoteToUser')}}}</p>")}),
                "runAfter": ra("Chat_the_note_if_internal", status=("Succeeded", "Failed"))}),
        ]))),
    ])

    # Part C: the status changed
    note_html = f"@{{if(empty(coalesce({t('NoteToUser')},'')),'',concat('<p><b>Message from IT:</b> ',{t('NoteToUser')},'</p>'))}}"
    resolution_html = f"@{{if(and(contains({arr([S_RES, S_CLOSED])},{V}),not(empty(coalesce({t('Resolution')},'')))),concat('<p><b>Resolution:</b> ',{t('Resolution')},'</p>'),'')}}"
    extra_html = (f"@{{if(equals({V},'{S_RES}'),'<p>If this is not fixed, reply to this email within 2 days and your ticket will be reopened.</p>',"
                  f"if(contains({arr([S_WAIT, S_SCHED])},{V}),'<p>Please reply to this email to respond to IT.</p>',"
                  f"if(equals({V},'{S_CLOSED}'),'<p>To raise this again, reply to this email or log a new ticket.</p>','')))}}")
    partC = seq([
        ("Has_task_C", cond(hasTask, {"Move_Planner_task": op("shared_planner", "UpdateTask_V3", {
            "id": f"@{t('PlannerTaskId')}", "body/bucketId": "@" + bucket_expr(V),
            "body/percentComplete": f"@if(contains({arr([S_RES, S_CLOSED])},{V}),100,if(equals({V},'{S_OPEN}'),0,50))"})})),
        ("Meanings", {**compose(MEANING), "runAfter": ra("Has_task_C", status=("Succeeded", "Failed", "Skipped"))}),
        ("StatusLine", compose(f"@concat('[',{stamp},' | ',{who},'] Status: ',{V},if(contains({arr([S_RES, S_CLOSED])},{V}),concat(' · Resolution: ',coalesce({t('Resolution')},'')),''),decodeUriComponent('%0A%0A'),variables('varConversation'))")),
        ("Set_varConversation_status", {"type": "SetVariable", "inputs": {"name": "varConversation", "value": "@outputs('StatusLine')"}}),
        ("Has_requester", cond("@not(empty(outputs('ReqEmail')))", seq([
            ("Email_status", mail("@outputs('ReqEmail')", f"[IT-@{{{ID}}}] Ticket status: @{{{UV}}}: @{{{t('Title')}}}",
                style + f"<p style=\"color:#999;font-size:12px\">{MARKER}</p><p>Hi @{{{first}}},</p>"
                f"<p>The status of your ticket IT-@{{{ID}}} has changed to <b>@{{{UV}}}</b>@{{if(equals({V},'{S_CLOSED}'),'',concat(': ',outputs('Meanings')?[{V}]))}}.</p>"
                + resolution_html + note_html + status_files_html + extra_html +
                portal_line + details + earlier + "</div>")),
            ("Chat_status_if_internal", {**if_internal("outputs('IsInternal')", {
                "Chat_status": {"type": "Scope", "runAfter": {},
                    "actions": support_msg("Post_status", "@outputs('ReqEmail')",
                        f"<p><b>IT-@{{{ID}}} is now @{{{UV}}}</b></p>@{{if(equals({V},'{S_CLOSED}'),'',concat('<p>',outputs('Meanings')?[{V}],'.</p>'))}}<p style=\"color:#666\">@{{{t('Title')}}}</p>"
                        f"@{{if(empty(coalesce({t('NoteToUser')},'')),'',concat('<p><b>Message from IT:</b> ',{t('NoteToUser')},'</p>'))}}"
                        + status_files_html +
                        resolution_html +
                        f"@{{if(equals({V},'{S_RES}'),'<p>If this is not fixed, use Reply by email within 2 days and your ticket will be reopened.</p>',if(contains({arr([S_WAIT, S_SCHED])},{V}),'<p>Please use Reply by email to respond to IT.</p>',''))}}"
                        + support_links(ID, t('Title')))}}),
                "runAfter": ra("Email_status", status=("Succeeded", "Failed"))}),
        ]))),
        ("Question_missing", {**cond(f"@and(contains({arr([S_WAIT, S_SCHED])},{V}),empty(coalesce({t('NoteToUser')},'')))", {
            "Nudge_the_tech": scope_msg("Nudge_msg", f"@{{coalesce({t('Editor/Email')},'{ME}')}}",
                f"<p>IT-@{{{ID}}} is set to @{{{V}}}, but no message was sent to the user. Add one with Note to user.</p><p>{lnk}</p>")}),
            "runAfter": ra("Has_requester", status=("Succeeded", "Failed", "Skipped"))}),
        ("Has_channel_post", {**cond(f"@not(empty(coalesce({t('ChannelMessageId')},'')))", {
            "Reply_status_in_channel": channel_reply(f"@{t('ChannelMessageId')}",
                f"<p>IT-@{{{ID}}} is now <b>@{{{V}}}</b>"
                f"@{{if(empty(coalesce({t('AssignedTo/DisplayName')},'')),'',concat(' · IT: ',{t('AssignedTo/DisplayName')}))}}"
                f" · updated by @{{{who}}}</p>"
                f"@{{if(empty(coalesce({t('NoteToUser')},'')),'',concat('<p>Message to user: ',{t('NoteToUser')},'</p>'))}}"
                f"@{{if(contains({arr([S_RES, S_CLOSED])},{V}),concat('<p>Resolution: ',coalesce({t('Resolution')},''),'</p>'),'')}}")}),
            "runAfter": ra("Question_missing", status=("Succeeded", "Failed", "Skipped"))}),
    ])

    def var(name, typ, val):
        return {"type": "InitializeVariable", "inputs": {"variables": [{"name": name, "type": typ, "value": val}]}}
    ALL = ("Succeeded", "Failed", "Skipped", "TimedOut")

    # Safety net: HD9 blocks closing a ticket with no resolution from the Teams card, but a direct
    # edit to the list item (or anything else that bypasses the card) skips that gate entirely.
    # Catch it here too: if Status lands on Closed with no Resolution on file, put it back and tell
    # whoever made the edit, instead of letting a bare close go out to the requester.
    close_guard = seq([
        ("Needs_resolution_HD4", compose(f"@and(equals({V},'{S_CLOSED}'),empty(coalesce({t('Resolution')},'')))")),
        ("Fallback_status", compose(f"@coalesce({t('LastNotifiedStatus')},'{S_OPEN}')")),
    ])
    close_guard["Revert_bad_close"] = {**cond("@outputs('Needs_resolution_HD4')", seq([
        ("Patch_revert_status", sp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{t('Title')}",
            "item/Status/Value": "@outputs('Fallback_status')"})),
        ("Set_varStatus_reverted", {"type": "SetVariable", "inputs": {"name": "varStatus", "value": "@outputs('Fallback_status')"}}),
        ("Tell_editor_resolution_needed", scope_msg("Revert_msg", f"@{{coalesce({t('Editor/Email')},'{ME}')}}",
            f"<p>IT-@{{{ID}}} was set to Closed without a resolution, so it was put back to <b>@{{outputs('Fallback_status')}}</b>. "
            f"Add a resolution and set it to Closed again — use the ticket's <b>Update status</b> card in Teams, or fill in the Resolution field before closing.</p><p>{lnk}</p>")),
    ])), "runAfter": ra("Fallback_status")}

    actions = seq([
        ("Init_varStatus", var("varStatus", "string", f"@{t('Status/Value')}")),
        ("Init_varConversation", var("varConversation", "string", f"@coalesce({t('Conversation')},'')")),
        ("Init_varFirstResponse", var("varFirstResponse", "string", f"@coalesce({t('FirstResponseOn')},'')")),
        ("Tech", compose(f"@toLower(coalesce({t('AssignedTo/Email')},''))")),
        ("ReqEmail", compose(f"@toLower(coalesce({t('RequesterEmail')},{t('Requester/Email')},''))")),
        ("IsInternal", compose("@" + is_internal("outputs('ReqEmail')"))),
        ("FirstName", compose(f"@first(split(trim(coalesce({t('Requester/DisplayName')},first(split(outputs('ReqEmail'),'@')),'there')),' '))")),
        # Fetched fresh (not from the trigger body) so a file attached separately from the edit that fired this
        # run still shows up: attaching a file doesn't itself touch Status/AssignedTo/NoteToUser, so by the time
        # HD4 runs for some other reason, any attachment already on the item is picked up. Only files not already
        # in NotifiedAttachments0 (pipe-separated filenames, not shown on any form) get linked — otherwise the
        # requester's own originally-submitted file would get re-linked on every single later update. Update_item
        # below rolls the baseline forward to every attachment currently on the item whenever a note/status
        # notification actually goes out.
        # Internal name is NotifiedAttachments0, not NotifiedAttachments: the first column was created with
        # Hidden="TRUE" in its schema XML, which SharePoint will neither un-hide nor delete afterward (REST
        # errors "Cannot change Hidden attribute"/"You cannot delete a hidden column") - so that GUID is stranded
        # on the list forever, unused. Recreated as Hidden="FALSE" with ShowInDisplayForm/EditForm/NewForm/
        # ViewForms all FALSE instead (keeps it off every form and view without tripping that restriction), which
        # SharePoint auto-suffixed to NotifiedAttachments0 since the name was taken. Never create a SharePoint
        # column as Hidden="TRUE" for anything a flow needs to reference - use the Show-In-*="FALSE" properties.
        ("Get_attachments", sp_http("GET", f"_api/web/lists(guid'{TICKETS}')/items(@{{{ID}}})/AttachmentFiles")),
        ("NotifiedList", {**compose(f"@split(coalesce({t('NotifiedAttachments0')},''),'|')"),
                          "runAfter": ra("Get_attachments", status=ALL)}),
        ("New_attachments", {"type": "Query", "runAfter": ra("NotifiedList", status=ALL), "inputs": {
            "from": "@coalesce(body('Get_attachments')?['value'],json('[]'))",
            "where": "@not(contains(outputs('NotifiedList'),item()?['FileName']))"}}),
        ("HasFiles", {**compose("@greater(length(coalesce(body('New_attachments'),json('[]'))),0)"),
                      "runAfter": ra("New_attachments", status=ALL)}),
        ("FileLinks", {"type": "Select", "runAfter": ra("HasFiles", status=ALL), "inputs": {"from": "@coalesce(body('New_attachments'),json('[]'))",
            "select": "@concat('<a href=\"https://branthavenhomes.sharepoint.com',item()?['ServerRelativeUrl'],'\">',item()?['FileName'],'</a>')"}}),
        ("AllFileNames", {"type": "Select", "runAfter": ra("FileLinks", status=ALL), "inputs": {
            "from": "@coalesce(body('Get_attachments')?['value'],json('[]'))", "select": "@item()?['FileName']"}}),
        # Same link format as FileLinks, but over every attachment currently on the ticket (not just
        # New_attachments) — backs status_files_html's Resolved/Closed branch above.
        ("AllFileLinks", {"type": "Select", "runAfter": ra("AllFileNames", status=ALL), "inputs": {
            "from": "@coalesce(body('Get_attachments')?['value'],json('[]'))",
            "select": "@concat('<a href=\"https://branthavenhomes.sharepoint.com',item()?['ServerRelativeUrl'],'\">',item()?['FileName'],'</a>')"}}),
        # a failure fetching/attachments never blocks notifications or anything else below it
        ("Close_guard", {"type": "Scope", "runAfter": ra("AllFileLinks", status=ALL), "actions": close_guard}),
        # A tech who resolves/closes a ticket straight from the SharePoint form usually forgets to assign it, so the
        # requester's email says "Not yet assigned". Teams Quick close assigns the responder; do the same here: if
        # the status just moved to Resolved/Closed with no assignee and the editor is a person (not the flow
        # account), write the editor into IT Assigned and stop — that edit re-triggers HD4, which then sends the
        # notifications with the assignee filled in (LastNotifiedStatus is still the old value, so nothing is lost).
        ("Self_assign_on_close", cond(
            f"@and(not(outputs('Needs_resolution_HD4')),empty(coalesce({t('AssignedTo/Email')},'')),"
            f"contains({arr([S_RES, S_CLOSED])},{V}),not(equals({V},coalesce({t('LastNotifiedStatus')},''))),"
            f"not(empty(coalesce({t('Editor/Email')},''))),not(equals(toLower(coalesce({t('Editor/Email')},'')),'{MBX.lower()}')))",
            seq([
                ("Assign_to_editor", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{t('Title')}",
                                                      "item/AssignedTo/Claims": f"@{t('Editor/Email')}"})),
                ("Stop_until_rerun", {"type": "Terminate", "inputs": {"runStatus": "Succeeded"}}),
            ]))),
        # Skip the Planner assignment (and its "assigned to you" notification) when this same update is also
        # landing on Resolved/Closed — e.g. Quick close, which sets AssignedTo and Status together. Part C
        # below still marks the Planner task done; there's no point notifying someone of an assignment to a
        # task that's complete in the same breath. A tech assigned while the ticket stays open still gets
        # notified as before.
        ("Part_B_tech_assigned", cond(f"@and(not(empty(outputs('Tech'))),not(equals(outputs('Tech'),toLower(coalesce({t('LastNotifiedTech')},'')))),not(contains({arr([S_RES, S_CLOSED])},{V})))", partB)),
        ("Part_A_note_for_user", {**cond(f"@not(empty(coalesce({t('NoteToUser')},'')))", partA),
                                  "runAfter": ra("Part_B_tech_assigned", status=ALL)}),
        ("Part_C_status_changed", {**cond(f"@not(equals({V},coalesce({t('LastNotifiedStatus')},'')))", partC),
                                   "runAfter": ra("Part_A_note_for_user", status=ALL)}),
        # Keep the Planner task's own Notes in sync with the ticket: a note added, a status change, and the
        # resolution text on Resolved/Closed (StatusLine already folds "· Resolution: ..." into the conversation
        # line, so it shows up here too) all land in the task without anyone opening SharePoint. Rebuilt in full
        # each run — simpler and more robust than trying to append — and capped well under Planner's description
        # limit.
        ("Sync_task_notes", {**cond(hasTask, {
            "Update_task_notes": op("shared_planner", "UpdateTaskDetails_V2", {
                "id": f"@{t('PlannerTaskId')}",
                "body/description": f"@concat('IT-',{ID},': ',{t('Title')},decodeUriComponent('%0A'),'{LIST_URL}',"
                    # padded before substring - varConversation stays at Init_varConversation's coalesce(Conversation,'')
                    # when this run touched neither a note nor a status (e.g. just an attachment added), which can be empty
                    f"string({ID}),decodeUriComponent('%0A%0A'),substring(concat(variables('varConversation'),' '),0,"
                    f"min(length(variables('varConversation')),1800)))"})}),
            "runAfter": ra("Part_C_status_changed", status=ALL)}),
        # ticks between the SLA start (7am rule) and the first response; 0 when there is no first response yet
        ("ResponseTicks", {**compose(f"@if(empty(variables('varFirstResponse')),0,sub(ticks(variables('varFirstResponse')),{sla_start_ticks(t('Created'))}))"),
                           "runAfter": ra("Sync_task_notes", status=ALL)}),
        ("Update_item", {**sp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{t('Title')}",
            "item/Status/Value": f"@{V}",
            "item/LastNotifiedStatus": f"@{V}",
            "item/LastNotifiedTech": "@outputs('Tech')",
            "item/NoteToUser": "@null",
            "item/Conversation": "@variables('varConversation')",
            "item/FirstResponseOn": "@if(empty(variables('varFirstResponse')),null,variables('varFirstResponse'))",
            "item/NotifiedAttachments0": f"@if({files_shown},join(body('AllFileNames'),'|'),coalesce({t('NotifiedAttachments0')},''))",
            # ResolvedOn is also stamped when a ticket goes straight to Closed (Quick close) so it counts as resolved
            "item/ResolvedOn": f"@if(and(empty({t('ResolvedOn')}),not(equals({V},coalesce({t('LastNotifiedStatus')},''))),contains({arr([S_RES, S_CLOSED])},{V})),utcNow(),{t('ResolvedOn')})",
            "item/ClosedOn": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_CLOSED}')),utcNow(),{t('ClosedOn')})",
            "item/WaitingHours": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_WAIT}')),0,{t('WaitingHours')})",
            "item/ResolvedHours": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_RES}')),0,{t('ResolvedHours')})"}),
            "runAfter": ra("ResponseTicks", status=ALL)}),
        # Per-ticket SLA figures (feed the per-tech averages in the By IT Support view), written once via REST so the
        # flow does not depend on the SharePoint connector having refreshed its column schema.
        ("Write_SLA_figures", {**cond(
            f"@or(and(empty({t('FirstResponseOn')}),not(empty(variables('varFirstResponse')))),"
            f"and(empty({t('ResolvedOn')}),not(equals({V},coalesce({t('LastNotifiedStatus')},''))),contains({arr([S_RES, S_CLOSED])},{V})))",
            {"Merge_SLA_figures": sp_http("POST", f"_api/web/lists(guid'{TICKETS}')/items(@{{{ID}}})",
                body={
                    "ResponseMinutes": f"@if(empty({t('FirstResponseOn')}),if(empty(variables('varFirstResponse')),null,if(less(outputs('ResponseTicks'),0),0,div(outputs('ResponseTicks'),600000000))),{t('ResponseMinutes')})",
                    "ResponseMet": f"@if(empty({t('FirstResponseOn')}),if(empty(variables('varFirstResponse')),null,if(lessOrEquals(outputs('ResponseTicks'),144000000000),1,0)),{t('ResponseMet')})",
                    "ResolvedMet": f"@if(and(empty({t('ResolvedOn')}),not(equals({V},coalesce({t('LastNotifiedStatus')},''))),contains({arr([S_RES, S_CLOSED])},{V})),if(lessOrEquals(float(coalesce({t('SLAHours')},0)),24),1,0),{t('ResolvedMet')})"},
                headers={"IF-MATCH": "*", "X-HTTP-Method": "MERGE"})}),
            "runAfter": ra("Update_item", status=("Succeeded", "Failed"))}),
    ])
    # Monitoring: did every requester/tech notification for this update go out?
    actions.update(self_check("Health", ["Close_guard", "Part_B_tech_assigned", "Part_A_note_for_user", "Part_C_status_changed"],
        f"[Help desk health] Notification problem on ticket update IT-@{{{ID}}}",
        f"Ticket <b>IT-@{{{ID}}}</b> (@{{{t('Title')}}}) was updated to <b>@{{{V}}}</b>"
        f" (assignee @{{coalesce({t('AssignedTo/DisplayName')},'none')}}, requester @{{outputs('ReqEmail')}}), but these steps did not complete:",
        f"concat('IT-',{ID},' (updated)')", after=ra("Write_SLA_figures", status=ALL),
        leaf=("Email_status", "Post_status", "Email_the_note", "Post_note", "Post_new_tech", "Move_Planner_task", "Update_task_notes")))
    trig = sp_trigger("GetOnUpdatedItems")
    trig["conditions"] = [{"expression":
        "@or(not(equals(triggerOutputs()?['body/Status/Value'],triggerOutputs()?['body/LastNotifiedStatus'])),"
        "not(equals(toLower(coalesce(triggerOutputs()?['body/AssignedTo/Email'],'')),toLower(coalesce(triggerOutputs()?['body/LastNotifiedTech'],'')))),"
        "greater(length(coalesce(triggerOutputs()?['body/NoteToUser'],'')),0))"}]
    return "HD4 Ticket updated", {"When_an_item_is_created_or_modified": trig}, actions

# ---------- HD5 Planner task completed ----------
def hd5():
    item = lambda k: f"first(body('Get_items')?['value'])?['{k}']"
    ID = item('ID')
    lnk = link(ID)
    has_res = f"not(empty(coalesce({item('Resolution')},'')))"
    # Same resolution requirement as HD9/HD4, applied to the Planner path: completing the task
    # alone used to flip the ticket to Resolved Pending Confirmation with no resolution captured.
    # Now it only does that when a resolution is already on file; otherwise the task is reopened
    # (back to In progress) and the tech is told to add one before completing it again.
    actions = seq([
        ("Get_items", sp("GetItems", {"$filter": "PlannerTaskId eq '@{triggerOutputs()?['body/id']}'", "$top": 1})),
        ("Condition", cond(f"@and(greater(length(body('Get_items')?['value']),0),not(contains({arr([S_RES, S_CLOSED])},{item('Status')}?['Value'])))", {
            "Has_resolution_HD5": cond(f"@{has_res}", {
                "Update_item": sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{item('Title')}",
                                                "item/Status/Value": S_RES})},
                seq([
                    ("Reopen_task", op("shared_planner", "UpdateTask_V3", {
                        "id": "@triggerOutputs()?['body/id']", "body/percentComplete": 50})),
                    ("Tell_tech_need_resolution_HD5", scope_msg("Need_res_msg_HD5",
                        f"@coalesce({item('AssignedTo')}?['Email'],'{ME}')",
                        f"<p>IT-@{{{ID}}} wasn't resolved — completing the Planner task needs a resolution first. "
                        f"Use the ticket's <b>Update status</b> card in Teams (choose Resolved Pending Confirmation or Closed "
                        f"and enter the resolution), or fill in the Resolution field directly, then complete the task again.</p><p>{lnk}</p>")),
                ])),
        })),
    ])
    trig = {"When_a_task_is_completed": {"recurrence": {"frequency": "Minute", "interval": 1},
        "splitOn": "@triggerOutputs()?['body/value']", "type": "OpenApiConnection",
        "inputs": {"parameters": {"groupId": GROUP, "id": PLAN},
                   "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_planner", "connectionName": "shared_planner", "operationId": "OnCompleteTask_V3"},
                   "authentication": "@parameters('$authentication')"}}}
    return "HD5 Planner task completed", trig, actions

# ---------- HD6 Hourly clock ----------
def hd6():
    it = lambda k: f"items('Apply_to_each')?['{k}']"
    ID = it('ID')
    lnk = link(ID)
    req = f"toLower(coalesce({it('RequesterEmail')},''))"
    rlnk = req_link_if_internal(ID, req)       # already a <p>…</p>, empty for external senders
    internal = is_internal(req)
    biz = "outputs('BizDay')"
    new_case = seq([
        ("H", compose(f"@add(int(coalesce({it('BusinessHoursOpen')},0)),1)")),
        ("S", compose(f"@add(int(coalesce({it('SLAHours')},0)),if({biz},1,0))")),
        ("Update_H", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/BusinessHoursOpen": "@outputs('H')", "item/SLAHours": "@outputs('S')"})),
        ("Unassigned_at_4_hours", cond(f"@and({biz},empty(coalesce({it('AssignedTo')}?['Email'],'')),equals(outputs('S'),4))", {
            "Post_unassigned_4h": channel(f"<p>IT-@{{{ID}}} has had no technician for 4 hours: @{{{it('Title')}}}</p><p>{lnk}</p>")})),
        ("Unassigned_at_5_hours", cond(f"@and({biz},empty(coalesce({it('AssignedTo')}?['Email'],'')),equals(outputs('S'),5))",
            team_msgs("Unassigned5h_msg", f"<p>IT-@{{{ID}}} is still unassigned after 5 hours: @{{{it('Title')}}}</p><p>{lnk}</p>", UNASSIGNED_PING_TEAM))),
        # Once it's assigned, this reminder goes to that tech alone. Only while it's still unassigned does
        # Joseph get it too (alongside Martin, who's the fallback recipient in that case) — same audience as
        # the 5-hour unassigned ping above, just repeating daily instead of firing once.
        ("Open_a_full_business_day", cond(f"@and({biz},greaterOrEquals(outputs('S'),{DAY_H}),equals(mod(outputs('S'),{DAY_H}),0))", seq([
            ("Tell_tech_open_days", scope_msg("Open_days_msg", f"@{{coalesce({it('AssignedTo')}?['Email'],'{ME}')}}",
                f"<p>IT-@{{{ID}}} has been open @{{div(outputs('S'),{DAY_H})}} day(s): @{{{it('Title')}}}, requested by @{{{req}}}</p><p>{lnk}</p>")),
            ("Tell_Joseph_if_still_unassigned", cond(f"@empty(coalesce({it('AssignedTo')}?['Email'],''))",
                {"Open_days_msg_joseph": scope_msg("Open_days_joseph_msg", TECH2,
                    f"<p>IT-@{{{ID}}} has been open @{{div(outputs('S'),{DAY_H})}} day(s): @{{{it('Title')}}}, requested by @{{{req}}}</p><p>{lnk}</p>")})),
        ]))),
    ])
    for k in ("Unassigned_at_5_hours", "Open_a_full_business_day"):
        new_case[k]["runAfter"] = ra(list(new_case.keys())[list(new_case.keys()).index(k) - 1], status=("Succeeded", "Failed", "Skipped"))
    waiting_case = seq([
        ("W", compose(f"@add(int(coalesce({it('WaitingHours')},0)),1)")),
        ("Five_days_waiting", cond(f"@greaterOrEquals(outputs('W'),{5*DAY_H})", {
            "Close_no_reply": sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/Status/Value": S_CLOSED,
                "item/WaitingHours": "@outputs('W')",
                "item/Conversation": f"@concat('Closed automatically: no reply in 5 days.',decodeUriComponent('%0A%0A'),coalesce({it('Conversation')},''))"})},
            seq([
                ("Update_W", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/WaitingHours": "@outputs('W')"})),
                ("Daily_reminder", cond(f"@and(equals(mod(outputs('W'),{DAY_H}),0),not(empty({req})))", seq([
                    ("Email_reminder", mail(f"@{req}", f"[IT-@{{{ID}}}] Reminder: IT is waiting for your reply",
                        f"<p>{MARKER}</p><p>IT is waiting for your reply on IT-@{{{ID}}}: @{{{it('Title')}}}</p>"
                        # padded before substring - an empty Conversation (no history yet) would otherwise throw
                        f"<pre>@{{substring(concat(coalesce({it('Conversation')},''),' '),0,min(length(coalesce({it('Conversation')},'')),1000))}}</pre>"
                        f"<p>Reply to this email to answer. The ticket will be marked Resolved after 5 days without a reply.</p>{rlnk}")),
                    ("Chat_reminder_if_internal", {**if_internal(internal, {
                        "Chat_reminder": {"type": "Scope", "runAfter": {},
                            "actions": support_msg("Post_reminder", f"@{req}",
                                f"<p><b>Reminder: IT is waiting for your reply on IT-@{{{ID}}}</b></p><p>@{{{it('Title')}}}</p>"
                                "<p>The ticket will be marked Resolved after 5 days without a reply.</p>" + support_links(ID, it('Title')))}}),
                        "runAfter": ra("Email_reminder", status=("Succeeded", "Failed"))}),
                ]))),
            ]))),
    ])
    resolved_case = seq([
        ("ResolvedAge", compose(f"@div(sub(ticks(utcNow()),ticks(coalesce({it('ResolvedOn')},utcNow()))),36000000000)")),
        ("Window_48h", cond("@greaterOrEquals(outputs('ResolvedAge'),48)", {
            # Silent close: LastNotifiedStatus is set with the status so HD4 does not run (the resolved email already
            # said the ticket would close itself). Planner, the record line and the channel thread are handled here.
            **seq([
                ("Close_resolved", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}",
                    "item/Status/Value": S_CLOSED, "item/LastNotifiedStatus": S_CLOSED, "item/ClosedOn": "@utcNow()",
                    "item/Conversation": f"@concat('[',{nowET('yyyy-MM-dd HH:mm')},' | System] Status: Closed · closed automatically 48 hours after resolution',decodeUriComponent('%0A%0A'),coalesce({it('Conversation')},''))"})),
                ("Done_in_Planner", cond(f"@not(empty(coalesce({it('PlannerTaskId')},'')))", {
                    "Complete_task": planner_update(B_DONE, done=True, task=f"@{it('PlannerTaskId')}")})),
                ("Note_in_thread", {**cond(f"@not(empty(coalesce({it('ChannelMessageId')},'')))", {
                    "Post_auto_close": channel_reply(f"@{it('ChannelMessageId')}",
                        f"<p>IT-@{{{ID}}} is now <b>Closed</b> · closed automatically 48 hours after resolution (no reply from the requester)</p>")}),
                    "runAfter": ra("Done_in_Planner", status=("Succeeded", "Failed", "Skipped"))}),
            ])},
            {"Reminder_24h": cond(f"@and(greaterOrEquals(outputs('ResolvedAge'),24),lessOrEquals(int(coalesce({it('ResolvedHours')},0)),0),not(empty({req})))", seq([
                ("Mark_reminded", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/ResolvedHours": 1})),
                ("Email_resolved_reminder", mail(f"@{req}", f"[IT-@{{{ID}}}] Reminder: is your issue fixed? @{{{it('Title')}}}",
                    f"<p>{MARKER}</p><p>IT marked ticket IT-@{{{ID}}} as resolved: @{{{it('Title')}}}</p>"
                    f"<p><b>Resolution:</b> @{{coalesce({it('Resolution')},'')}}</p>"
                    "<p>If it is fixed, you don't need to do anything. If it is not fixed, reply to this email"
                    f"@{{if({internal},' or reply in the IT Support chat in Teams','')}} within the next day and your ticket will be reopened. Otherwise it will be marked Resolved automatically 48 hours after the fix was delivered.</p>"
                    f"{rlnk}")),
                ("Chat_resolved_reminder_if_internal", {**if_internal(internal, {
                    "Chat_resolved_reminder": {"type": "Scope", "runAfter": {},
                        "actions": support_msg("Post_resolved_reminder", f"@{req}",
                            f"<p><b>Reminder: is IT-@{{{ID}}} fixed?</b></p><p>@{{{it('Title')}}}</p>"
                            "<p>If it is fixed, no action is needed. If not, reply in this chat or use Reply by email within the next day to reopen it. Otherwise it will be marked Resolved automatically after 48 hours.</p>" + support_links(ID, it('Title')))}}),
                    "runAfter": ra("Email_resolved_reminder", status=("Succeeded", "Failed"))}),
            ]))})),
    ])
    loop = {"type": "Foreach", "foreach": "@body('Get_items_Tickets')?['value']",
            "actions": {"Switch": {"runAfter": {}, "type": "Switch",
                "expression": f"@if(contains({arr(WORKING)},{it('Status')}?['Value']),'Working',if(outputs('BizDay'),{it('Status')}?['Value'],'Skip'))",
                "cases": {"Case_Working": {"case": "Working", "actions": new_case},
                          "Case_Waiting": {"case": S_WAIT, "actions": waiting_case},
                          "Case_Resolved": {"case": S_RES, "actions": resolved_case}},
                "default": {"actions": {}}}}}
    actions = seq([
        # 7 am to 9 pm, every day, holidays included: every run counts
        ("BizDay", compose("@true")),
        ("Get_items_Tickets", sp("GetItems", {"$filter": "Status ne 'Closed' and Created lt '@{addMinutes(utcNow(),-30)}'", "$top": 5000})),
        ("Apply_to_each", loop),
    ])
    return "HD6 Hourly clock", {"Recurrence": recurrence(range(8, 22), ALLDAYS)}, actions

# ---------- HD7 My open tickets ----------
def hd7():
    R = "body('Get_response_details')"
    resp = f"toLower({R}?['responder'])"
    actions = seq([
        ("Get_response_details", op("shared_microsoftforms", "GetFormResponseById",
            {"form_id": FORM2, "response_id": "@triggerOutputs()?['body/resourceData/responseId']"})),
        ("Filter", compose(
            f"@if(contains(string({R}),'Open tickets only'),"
            f"concat('RequesterEmail eq ''',{resp},''' and Status ne ''Closed'''),"
            f"concat('RequesterEmail eq ''',{resp},''' and (Status ne ''Closed'' or Created ge ''',addDays(utcNow(),-30),''')'))")),
        ("Get_items", sp("GetItems", {"$filter": "@outputs('Filter')", "$orderby": "Created desc", "$top": 500})),
        ("Select", {"type": "Select", "inputs": {"from": "@body('Get_items')?['value']", "select": {
            "Ticket": "@concat('IT-',item()?['ID'])", "Summary": "@item()?['Title']",
            "Status": "@item()?['Status']?['Value']", "Technician": "@coalesce(item()?['AssignedTo']?['DisplayName'],'Not assigned yet')",
            "Opened": "@formatDateTime(item()?['Created'],'yyyy-MM-dd')",
            "Link": f"@concat('{REQ_URL}',item()?['ID'])"}}}),
        ("Create_HTML_table", {"type": "Table", "inputs": {"from": "@body('Select')", "format": "HTML"}}),
        ("Any_tickets", cond("@equals(length(body('Get_items')?['value']),0)",
            {"Email_none": mail(f"@{resp}", "Your IT tickets", "<p>You have no open tickets.</p>")},
            {"Email_list": mail(f"@{resp}", "Your IT tickets",
                "<p>Here are your IT tickets. To add information to one, reply to any email about it.</p>@{body('Create_HTML_table')}")})),
        ("Chat_responder", {"type": "Scope", "runAfter": ra("Any_tickets", status=("Succeeded", "Failed")),
            "actions": support_msg("Post_list_sent", f"@{resp}",
                f"<p>Your ticket list has been emailed to you (@{{length(body('Get_items')?['value'])}} tickets).</p><p><a href=\"{PORTAL_URL}\">Open My tickets</a></p>")}),
    ])
    return "HD7 My open tickets", {"When_a_new_response_is_submitted": forms_trigger(FORM2)}, actions

# ---------- HD8 Daily report ----------
def hd8():
    tbl = lambda name: f"@{{body('{name}')}}"
    def section(prefix, where, select):
        return [
            (f"Filter_{prefix}", {"type": "Query", "inputs": {"from": "@body('OpenNow')?['value']", "where": where}}),
            (f"Select_{prefix}", {"type": "Select", "inputs": {"from": f"@body('Filter_{prefix}')", "select": select}}),
            (f"Table_{prefix}", {"type": "Table", "inputs": {"from": f"@body('Select_{prefix}')", "format": "HTML"}}),
        ]
    def or_none(prefix, msg):
        return f"@{{if(equals(length(body('Filter_{prefix}')),0),'<p>{msg}</p>',body('Table_{prefix}'))}}"
    by_tech_loop = {"type": "Foreach", "foreach": "@outputs('Techs')", "actions": seq([
        ("Filter_this_tech", {"type": "Query", "inputs": {"from": "@body('OpenNow')?['value']",
            "where": "@equals(coalesce(item()?['AssignedTo']?['Email'],'Unassigned'),items('Apply_to_each_tech'))"}}),
        ("Append_tech", {"type": "AppendToArrayVariable", "inputs": {"name": "varByTech", "value": {
            "Technician": "@items('Apply_to_each_tech')", "Open": "@length(body('Filter_this_tech'))",
            "Overdue": "@length(filter(body('Filter_this_tech'),'x'))"}}}),
    ])}
    # filter() isn't a workflow function; compute overdue with a second Query instead
    by_tech_loop["actions"] = seq([
        ("Filter_this_tech", {"type": "Query", "inputs": {"from": "@body('OpenNow')?['value']",
            "where": "@equals(coalesce(item()?['AssignedTo']?['Email'],'Unassigned'),items('Apply_to_each_tech'))"}}),
        ("Filter_this_tech_overdue", {"type": "Query", "inputs": {"from": "@body('Filter_this_tech')",
            "where": f"@greaterOrEquals(int(coalesce(item()?['SLAHours'],0)),{DAY_H})"}}),
        ("Append_tech", {"type": "AppendToArrayVariable", "inputs": {"name": "varByTech", "value": {
            "Technician": "@items('Apply_to_each_tech')", "Open": "@length(body('Filter_this_tech'))",
            "Overdue": "@length(body('Filter_this_tech_overdue'))",
            "Pct": "@div(mul(length(body('Filter_this_tech')),100),max(length(body('OpenNow')?['value']),1))"}}}),
    ])
    by_cat_loop = {"type": "Foreach", "foreach": "@outputs('Categories')", "actions": seq([
        ("Filter_this_category", {"type": "Query", "inputs": {"from": "@body('Opened')?['value']",
            "where": "@equals(item()?['Category']?['Value'],items('Apply_to_each_category'))"}}),
        ("Append_category", {"type": "AppendToArrayVariable", "inputs": {"name": "varByCategory", "value": {
            "Category": "@items('Apply_to_each_category')", "Tickets": "@length(body('Filter_this_category'))",
            "Pct": "@div(mul(length(body('Filter_this_category')),100),max(length(body('Opened')?['value']),1))"}}}),
    ])}
    fname = "IT Support Help Desk Report - @{" + nowET('MMMM d, yyyy') + "}"
    ESC = lambda x: f"replace(replace(replace({x},'&','&amp;'),'<','&lt;'),'>','&gt;')"
    LASTTECH = ("if(and(empty(coalesce(item()?['AssignedTo']?['Email'],'')),equals(item()?['Status']?['Value'],'Open')),'',if(not(empty(coalesce(item()?['AssignedTo']?['Email'],''))),coalesce(item()?['AssignedTo']?['DisplayName'],item()?['AssignedTo']?['Email']),"
                "if(not(empty(coalesce(item()?['LastNotifiedTech'],''))),item()?['LastNotifiedTech'],"
                "if(and(not(empty(coalesce(item()?['Editor']?['Email'],''))),not(contains(toLower(coalesce(item()?['Editor']?['Email'],'')),'it_support'))),coalesce(item()?['Editor']?['DisplayName'],item()?['Editor']?['Email']),''))))")
    REQ = "if(empty(item()?['RequesterEmail']),coalesce(item()?['Author']?['DisplayName'],'Unknown'),item()?['RequesterEmail'])"
    def td(expr, extra=""):
        return f"'<td style=\"padding:8px 10px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#0f172a;{extra}\">',{expr},'</td>'"
    def rowx(*cells):
        return "@concat('<tr>'," + ",".join(cells) + ",'</tr>')"
    def th(t, w=""):
        return f"<td{w} style=\"padding:7px 10px;background:#f1f5f9;font-size:11px;font-weight:bold;color:#475569;text-transform:uppercase\">{t}</td>"
    def bar(color):
        return ("concat('<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"><tr><td width=\"',string(max(createArray(int(item()?['Pct']),2))),'%\" bgcolor=\"" + color + "\" style=\"background:" + color + ";border-top:10px solid " + color + ";height:0;font-size:1px;line-height:1px\"></td><td style=\"font-size:1px\">&nbsp;</td></tr></table>')")
    def card(title, accent, count_expr, body_expr):
        return ("<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"margin-top:16px;border:1px solid #e2e8f0;border-collapse:collapse;page-break-inside:avoid;break-inside:avoid\">"
                "<tr style=\"page-break-after:avoid\"><td style=\"background:#f8fafc;color:" + accent + ";border-top:5px solid " + accent + ";padding:9px 12px;font-size:15px;font-weight:bold;page-break-after:avoid\">" + title + count_expr + "</td></tr>"
                "<tr><td style=\"background:#ffffff;padding:0\">" + body_expr + "</td></tr></table>")
    def listbody(filt, headers, tbl, empty):
        return ("@{if(equals(length(body('" + filt + "')),0),'<div style=\"padding:12px;font-size:13px;color:#64748b\">" + empty + "</div>',"
                "concat('<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse:collapse\"><tr>" + headers + "</tr>',outputs('" + tbl + "'),'</table>'))}")
    def tile(label, value_expr, color, bg="#f8fafc"):
        return ("<td width=\"33%\" style=\"padding:4px\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background:" + bg + ";border:1px solid #e2e8f0;border-top:4px solid " + color + "\">"
                "<tr><td style=\"padding:12px 14px\"><div style=\"font-size:30px;font-weight:bold;color:#0f172a;line-height:34px\">" + value_expr + "</div>"
                "<div style=\"font-size:12px;color:#64748b;text-transform:uppercase;margin-top:2px\">" + label + "</div></td></tr></table></td>")
    n_over = "@{length(body('Filter_overdue'))}"
    n_un = "@{length(body('Filter_unassigned'))}"
    report = (
        "<html><body style=\"margin:0;padding:0;background:#eef2f7;font-family:Arial,sans-serif\">"
        "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background:#eef2f7\"><tr><td align=\"center\" style=\"padding:16px\">"
        "<table width=\"680\" cellpadding=\"0\" cellspacing=\"0\" style=\"width:680px;max-width:100%\">"
        "<tr><td style=\"background:#ffffff;color:#0f2a4a;padding:18px 20px;border-top:6px solid #0f2a4a;border-bottom:1px solid #e2e8f0\">"
        "<div style=\"font-size:20px;font-weight:bold\">IT Support [Branthaven]</div>"
        "<div style=\"font-size:13px;color:#475569;margin-top:3px\">Daily report &middot; @{" + nowET('dddd MMMM d, yyyy') + "}</div></td></tr>"
        "<tr><td style=\"padding:10px 0 0 0\"><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\">"
        "<tr>" + tile("Opened today", "@{length(body('Opened')?['value'])}", "#2563eb")
               + tile("Closed today", "@{length(body('ResolvedToday')?['value'])}", "#16a34a")
               + tile("Open now", "@{length(body('OpenNow')?['value'])}", "#0f2a4a") + "</tr>"
        "<tr>" + tile("Overdue", n_over, "#dc2626", "@{if(equals(length(body('Filter_overdue')),0),'#f8fafc','#fef2f2')}")
               + tile("Unassigned 4h+", n_un, "#f59e0b", "@{if(equals(length(body('Filter_unassigned')),0),'#f8fafc','#fffbeb')}")
               + tile("Avg hours to close", "@{formatNumber(outputs('Average'),'0.0')}", "#7c3aed") + "</tr>"
        "</table></td></tr>"
        "<tr><td>"
        + card("Overdue", "#dc2626", " (" + n_over + ")",
               listbody("Filter_overdue", th("Ticket", " width=\"60\"") + th("Requester") + th("Summary") + th("Assigned / last touched by") + th("Days open", " width=\"60\""), "Table_overdue", "Nothing overdue"))
        + card("Unassigned", "#d97706", " (" + n_un + ")",
               listbody("Filter_unassigned", th("Ticket", " width=\"60\"") + th("Requester") + th("Summary") + th("Hours", " width=\"50\""), "Table_unassigned", "Nothing unassigned"))
        + card("Waiting on users", "#7c3aed", " (@{length(body('Filter_waiting'))})",
               listbody("Filter_waiting", th("Ticket", " width=\"60\"") + th("Requester") + th("Summary") + th("Status") + th("Assigned / last touched by") + th("Days waiting", " width=\"70\""), "Table_waiting", "Nobody is waiting"))
        + card("Open IT Tickets", "#0f2a4a", "",
               "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse:collapse\"><tr>" + th("IT Support") + th("Workload", " width=\"220\"") + th("Open", " width=\"45\"") + th("Overdue", " width=\"55\"") + "</tr>@{join(body('Table_by_tech'),'')}</table>")
        + card("Today's tickets by category", "#2563eb", "",
               "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse:collapse\"><tr>" + th("Category") + th("Share", " width=\"220\"") + th("Tickets", " width=\"55\"") + "</tr>@{join(body('Table_by_category'),'')}</table>")
        + "</td></tr>"
        "<tr><td style=\"padding:14px 4px;font-size:11px;color:#94a3b8\">Overdue = open more than one business day. Business hours run 7:00 AM to 9:00 PM daily. Sent automatically by IT Support [Branthaven].</td></tr>"
        "</table></td></tr></table></body></html>")
    actions = seq([
        ("DayStart", compose(
            f"@convertTimeZone(addHours(addDays(startOfDay(convertTimeZone(utcNow(),'UTC','{TZ}')),-1),21),'{TZ}','UTC','yyyy-MM-ddTHH:mm:ssZ')")),
        ("Opened", sp("GetItems", {"$filter": "Created ge '@{outputs('DayStart')}'", "$top": 5000})),
        ("ResolvedToday", sp("GetItems", {"$filter": "ResolvedOn ge '@{outputs('DayStart')}' or (ClosedOn ge '@{outputs('DayStart')}' and ResolvedOn eq null)", "$top": 5000})),
        ("OpenNow", sp("GetItems", {"$filter": f"Status ne '{S_CLOSED}' and Status ne '{S_RES}'", "$top": 5000})),
        ("Init_varTotal", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varTotal", "type": "float", "value": 0}]}}),
        ("Init_varByTech", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varByTech", "type": "array", "value": []}]}}),
        ("Init_varByCategory", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varByCategory", "type": "array", "value": []}]}}),
        ("Apply_to_each", {"type": "Foreach", "foreach": "@body('ResolvedToday')?['value']", "actions": {
            "Increment_varTotal": {"runAfter": {}, "type": "IncrementVariable", "inputs": {"name": "varTotal",
                "value": "@float(coalesce(items('Apply_to_each')?['SLAHours'],0))"}}}}),
        ("Average", compose("@if(equals(length(body('ResolvedToday')?['value']),0),0,div(variables('varTotal'),float(length(body('ResolvedToday')?['value']))))")),
    ] + section("overdue", f"@and(contains({arr(WORKING)},item()?['Status']?['Value']),greaterOrEquals(int(coalesce(item()?['SLAHours'],0)),{DAY_H}))",
                {"Ticket": "@concat('IT-',item()?['ID'])", "Business days open": f"@div(int(coalesce(item()?['SLAHours'],0)),{DAY_H})",
                 "Summary": "@item()?['Title']", "Technician": "@coalesce(item()?['AssignedTo']?['DisplayName'],'Unassigned')"})
      + section("unassigned", "@and(empty(coalesce(item()?['AssignedTo']?['Email'],'')),greaterOrEquals(int(coalesce(item()?['SLAHours'],0)),4))",
                {"Ticket": "@concat('IT-',item()?['ID'])", "Hours unassigned": "@item()?['SLAHours']", "Summary": "@item()?['Title']"})
      + [("Filter_waiting", {"type": "Query", "inputs": {"from": "@body('OpenNow')?['value']", "where": f"@contains({arr([S_WAIT, S_SCHED])},item()?['Status']?['Value'])"}}),
         ("Select_waiting", {"type": "Select", "inputs": {"from": "@body('Filter_waiting')", "select": {
             "Ticket": "@concat('IT-',item()?['ID'])", "Requester": "@item()?['RequesterEmail']",
             "Business days waiting": f"@div(int(coalesce(item()?['WaitingHours'],0)),{DAY_H})"}}}),
         ("Table_waiting", {"type": "Table", "inputs": {"from": "@body('Select_waiting')", "format": "HTML"}}),
         ("Select_techs", {"type": "Select", "inputs": {"from": "@body('OpenNow')?['value']", "select": "@coalesce(item()?['AssignedTo']?['Email'],'Unassigned')"}}),
         ("Techs", compose("@union(body('Select_techs'),body('Select_techs'))")),
         ("Apply_to_each_tech", by_tech_loop),
         ("Table_by_tech", {"type": "Table", "inputs": {"from": "@variables('varByTech')", "format": "HTML"}}),
         ("Select_categories", {"type": "Select", "inputs": {"from": "@body('Opened')?['value']", "select": "@item()?['Category']?['Value']"}}),
         ("Categories", compose("@union(body('Select_categories'),body('Select_categories'))")),
         ("Apply_to_each_category", by_cat_loop),
         ("Table_by_category", {"type": "Table", "inputs": {"from": "@variables('varByCategory')", "format": "HTML"}}),
         ("ReportHtml", compose(report)),
         ("Create_file_OneDrive", op("shared_onedriveforbusiness", "CreateFile", {
             "folderPath": "/HelpDeskReports", "name": fname + ".html", "body": "@outputs('ReportHtml')"})),
         ("Delay", {"type": "Wait", "inputs": {"interval": {"count": 30, "unit": "Second"}}}),
         ("Convert_file", op("shared_onedriveforbusiness", "ConvertFile", {"id": "@body('Create_file_OneDrive')?['Id']", "type": "PDF"})),
         ("Create_file_Reports", op("shared_sharepointonline", "CreateFile", {
             "dataset": SITE, "folderPath": "/Reports", "name": fname + ".pdf", "body": "@body('Convert_file')"})),
         ("Email_me_the_report", {**mail(ME, "IT Support Help Desk Report - @{" + nowET('MMMM d, yyyy') + "}", "@{outputs('ReportHtml')}",
             attachments=[{"Name": fname + ".pdf", "ContentBytes": "@body('Convert_file')"}]),
             "runAfter": ra("Create_file_Reports", status=("Succeeded", "Failed"))}),
         ])
    def setsel(name, expr):
        actions[name]["inputs"]["select"] = expr
    setsel("Select_overdue", rowx(
        td("concat('IT-',item()?['ID'])", "font-weight:bold;"), td(ESC(REQ)), td(ESC("coalesce(item()?['Title'],'')")),
        td(LASTTECH), td(f"div(int(coalesce(item()?['SLAHours'],0)),{DAY_H})", "font-weight:bold;color:#dc2626;")))
    setsel("Select_unassigned", rowx(
        td("concat('IT-',item()?['ID'])", "font-weight:bold;"), td(ESC(REQ)), td(ESC("coalesce(item()?['Title'],'')")),
        td("string(item()?['SLAHours'])", "font-weight:bold;color:#d97706;")))
    setsel("Select_waiting", rowx(
        td("concat('IT-',item()?['ID'])", "font-weight:bold;"), td(ESC(REQ)), td(ESC("coalesce(item()?['Title'],'')")),
        td("coalesce(item()?['Status']?['Value'],'')"), td(LASTTECH), td(f"div(int(coalesce(item()?['WaitingHours'],0)),{DAY_H})")))
    for nm in ("overdue", "unassigned", "waiting"):
        actions["Table_" + nm]["type"] = "Compose"
        actions["Table_" + nm]["inputs"] = f"@join(body('Select_{nm}'),'')"
    actions["Table_by_tech"]["type"] = "Select"
    actions["Table_by_tech"]["inputs"] = {"from": "@variables('varByTech')", "select": rowx(
        td("if(or(equals(item()?['Technician'],'Unassigned'),contains(toLower(item()?['Technician']),'it_support')),'',item()?['Technician'])", "font-weight:bold;"),
        td(bar("#2563eb")), td("string(item()?['Open'])", "font-weight:bold;"),
        td("string(item()?['Overdue'])", "color:#dc2626;font-weight:bold;"))}
    actions["Table_by_category"]["type"] = "Select"
    actions["Table_by_category"]["inputs"] = {"from": "@variables('varByCategory')", "select": rowx(
        td("coalesce(item()?['Category'],'Uncategorized')", "font-weight:bold;"),
        td(bar("#7c3aed")), td("string(item()?['Tickets'])", "font-weight:bold;"))}
    return "HD8 Daily report", {"Recurrence": recurrence([21], ALLDAYS)}, actions

# ---------- HD9 Card actions ----------
def hd9():
    TB = "triggerBody()"
    data = f"{TB}?['entity']?['cardOutputs']"
    actions = seq([
        ("Act", compose(f"@toLower(coalesce({data}?['act'],''))")),
        ("TicketId", compose(f"@int(coalesce({data}?['ticket'],'0'))")),
        ("Get_clicker_profile", op("shared_office365users", "UserProfile_V2",
            {"id": f"@{TB}?['entity']?['teamsFlowRunContext']?['MessagePayload']?['From']?['User']?['Id']"})),
        ("Responder", {**compose("@toLower(coalesce(body('Get_clicker_profile')?['mail'],body('Get_clicker_profile')?['userPrincipalName'],'" + ME + "'))"),
                       "runAfter": ra("Get_clicker_profile", status=("Succeeded", "Failed"))}),
        ("StatusChoice", compose(f"@coalesce({data}?['status_choice'],'')")),
        ("Text", compose(f"@trim(coalesce({data}?['note_text'],{data}?['status_text'],''))")),
        ("Get_item", sp("GetItem", {"id": "@outputs('TicketId')"})),
        ("Switch", {"type": "Switch", "expression": "@outputs('Act')", "cases": {
            "Case_assign": {"case": "assign", "actions": {"Assign": sp("PatchItem", {
                "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                "item/AssignedTo/Claims": "@outputs('Responder')"})}},
            "Case_start": {"case": "start", "actions": {"Start": sp("PatchItem", {
                "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                "item/AssignedTo/Claims": "@coalesce(body('Get_item')?['AssignedTo']?['Email'],outputs('Responder'))",
                "item/Status/Value": S_PROG})}},
            "Case_note": {"case": "note", "actions": {"Note": sp("PatchItem", {
                "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                "item/NoteToUser": "@outputs('Text')"})}},
            "Case_edit": {"case": "edit", "actions": seq([
                ("NewTech", compose(f"@toLower(trim(coalesce({data}?['edit_tech'],'')))")),
                ("NewPriority", compose(f"@trim(coalesce({data}?['edit_priority'],''))")),
                ("NewCategory", compose(f"@trim(coalesce({data}?['edit_category'],''))")),
                ("OldTech", compose("@toLower(coalesce(body('Get_item')?['AssignedTo']?['Email'],''))")),
                ("Edit", sp("PatchItem", {
                    "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                    "item/AssignedTo/Claims": "@if(equals(outputs('NewTech'),''),if(empty(outputs('OldTech')),null,outputs('OldTech')),if(equals(outputs('NewTech'),'none'),null,outputs('NewTech')))",
                    "item/Priority/Value": "@if(empty(outputs('NewPriority')),body('Get_item')?['Priority']?['Value'],outputs('NewPriority'))",
                    "item/Category/Value": "@if(empty(outputs('NewCategory')),body('Get_item')?['Category']?['Value'],outputs('NewCategory'))"})),
                ("Changes", compose(
                    "@trim(concat("
                    "if(or(equals(outputs('NewTech'),''),equals(outputs('NewTech'),outputs('OldTech'))),'',"
                    "concat('Assigned to <b>',if(equals(outputs('NewTech'),'none'),'Unassigned',outputs('NewTech')),'</b> · ')),"
                    "if(or(empty(outputs('NewPriority')),equals(outputs('NewPriority'),body('Get_item')?['Priority']?['Value'])),'',concat('Priority <b>',outputs('NewPriority'),'</b> · ')),"
                    "if(or(empty(outputs('NewCategory')),equals(outputs('NewCategory'),body('Get_item')?['Category']?['Value'])),'',concat('Category <b>',outputs('NewCategory'),'</b> · '))))")),
                ("Note_changes", cond("@and(not(empty(outputs('Changes'))),not(empty(coalesce(body('Get_item')?['ChannelMessageId'],''))))", {
                    "Post_changes": channel_reply("@body('Get_item')?['ChannelMessageId']",
                        "<p>IT-@{outputs('TicketId')} updated by @{coalesce(body('Get_clicker_profile')?['displayName'],outputs('Responder'))}: "
                        "@{substring(outputs('Changes'),0,sub(length(outputs('Changes')),1))}</p>")})),
            ])},
            "Case_quickclose": {"case": "quickclose", "actions": seq([
                ("QuickText", compose(f"@trim(coalesce({data}?['quick_text'],''))")),
                ("Quick_close", cond("@empty(outputs('QuickText'))", {
                    "Say_need_resolution_quick": scope_msg("Need_res_quick_msg", "@outputs('Responder')",
                        "<p>IT-@{outputs('TicketId')} wasn't closed — Quick close needs the resolution typed in the box.</p>")},
                    {"Close_now": sp("PatchItem", {
                        "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                        "item/Status/Value": S_CLOSED,
                        "item/AssignedTo/Claims": "@coalesce(body('Get_item')?['AssignedTo']?['Email'],outputs('Responder'))",
                        "item/Resolution": "@outputs('QuickText')"})})),
            ])},
            "Case_status": {"case": "status", "actions": seq([
                ("NewResolution", compose(f"@if(and(contains({arr([S_RES, S_CLOSED])},outputs('StatusChoice')),not(empty(outputs('Text')))),outputs('Text'),coalesce(body('Get_item')?['Resolution'],''))")),
                ("Needs_resolution", compose(f"@and(equals(outputs('StatusChoice'),'{S_CLOSED}'),empty(outputs('NewResolution')))")),
                ("Status_update", cond("@outputs('Needs_resolution')", {
                    "Say_need_resolution": scope_msg("Need_res_msg", "@outputs('Responder')",
                        "<p>IT-@{outputs('TicketId')} wasn't closed — a resolution is required. Use <b>Update status</b>, choose Closed again, and enter the resolution in the message box.</p>")},
                    seq([
                        ("Set_status", sp("PatchItem", {
                            "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                            "item/Status/Value": "@outputs('StatusChoice')",
                            "item/AssignedTo/Claims": "@coalesce(body('Get_item')?['AssignedTo']?['Email'],outputs('Responder'))",
                            "item/NoteToUser": f"@if(and(contains({arr([S_WAIT, S_SCHED])},outputs('StatusChoice')),not(empty(outputs('Text')))),outputs('Text'),null)",
                            "item/Resolution": "@outputs('NewResolution')"})),
                        ("Internal_note", cond(f"@and(equals(outputs('StatusChoice'),'{S_ESC}'),not(empty(outputs('Text'))),not(empty(coalesce(body('Get_item')?['ChannelMessageId'],''))))", {
                            "Post_internal_note": channel_reply("@body('Get_item')?['ChannelMessageId']",
                                "<p><b>Internal note</b> (not sent to the user): @{outputs('Text')}</p>")})),
                    ]))),
            ])}},
            "default": {"actions": {}}}),
    ])
    import re as _re
    sample = json.loads(_re.sub(r"@\{[^}]*\}", "0", json.dumps(action_card("x", "x", "x", "x", "x", "x", "x"))))
    trig = {"When_someone_responds_to_an_adaptive_card": {"type": "Request", "kind": "ApiConnection",
        "splitOn": "@triggerBody()['rows']",
        "inputs": {"schema": {"type": "object", "properties": {}, "required": []},
                   "host": {"connection": {"name": "@parameters('$connections')['shared_teams']['connectionId']"}},
                   "operationId": "TeamsCardTrigger",
                   "parameters": {"inputsAdaptiveCard": json.dumps(sample, separators=(",", ":")), "CardTypeId": CARD_TYPE}}}}
    return "HD9 Card actions", trig, actions

# ---------- HD10 Ask the requester in Teams ----------
def hd10():
    t = lambda k: f"triggerOutputs()?['body/{k}']"
    ID = t('ID')
    gi = "body('Get_item')"
    stamp = f"convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm')"
    lnk = link(ID)
    card = {"type": "AdaptiveCard", "version": "1.2", "body": [
        {"type": "TextBlock", "text": f"IT has a question about IT-@{{{ID}}}", "weight": "Bolder", "size": "Medium", "wrap": True},
        {"type": "TextBlock", "text": f"**Subject:** @{{{t('Title')}}}", "wrap": True},
        {"type": "TextBlock", "text": f"**From IT:** @{{{t('TeamsQuestion')}}}", "wrap": True},
        {"type": "Input.Text", "id": "answer", "isMultiline": True, "placeholder": "Type your answer"},
        {"type": "TextBlock", "text": "Answer in the box above and select Send answer, or reply in this chat.", "wrap": True, "size": "Small", "isSubtle": True}],
        "actions": [{"type": "Action.Submit", "title": "Send answer"},
                    {"type": "Action.OpenUrl", "title": "View ticket", "url": f"{REQ_URL}@{{{ID}}}"}]}
    who = "coalesce(body('Get_responder')?['displayName'],outputs('ReqEmail'))"
    yes = seq([
        ("NewStatus", compose(f"@if(contains({arr([S_WAIT, S_SCHED, S_RES])},{gi}?['Status']?['Value']),'{S_PROG}',{gi}?['Status']?['Value'])")),
        ("Update_item", sp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{gi}?['Title']",
            "item/Status/Value": "@outputs('NewStatus')", "item/LastNotifiedStatus": "@outputs('NewStatus')",
            "item/WaitingHours": 0,
            "item/Conversation": f"@concat('[',{stamp},' | ',{who},' (Teams)] ',outputs('Answer'),decodeUriComponent('%0A%0A'),coalesce({gi}?['Conversation'],''))"})),
        ("Has_a_tech", cond(f"@not(empty(coalesce({gi}?['AssignedTo']?['Email'],'')))", {
            "Tell_the_tech": chat(f"@{{{gi}?['AssignedTo']?['Email']}}",
                f"<p>@{{{who}}} answered on IT-@{{{ID}}} in Teams:</p><p>@{{outputs('Answer')}}</p><p>{lnk}</p>")})),
        ("Has_channel_post", {**cond(f"@not(empty(coalesce({gi}?['ChannelMessageId'],'')))", {
            "Reply_in_channel": channel_reply(f"@{gi}?['ChannelMessageId']",
                f"<p>@{{{who}}} answered in Teams: @{{outputs('Answer')}}</p>"
                f"@{{if(equals(outputs('NewStatus'),{gi}?['Status']?['Value']),'',concat('<p>IT-',string({ID}),' is now <b>',outputs('NewStatus'),'</b></p>'))}}")}),
            "runAfter": ra("Has_a_tech", status=("Succeeded", "Failed", "Skipped"))}),
    ])
    actions = seq([
        ("Clear_question", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{t('Title')}", "item/TeamsQuestion": "@null"})),
        ("ReqEmail", compose(f"@toLower(coalesce({t('RequesterEmail')},{t('Requester/Email')},''))")),
        ("Card", compose(card)),
        ("Ask_in_Teams", {**op("shared_teams", "PostCardAndWaitForResponse", {
            "poster": "Flow bot", "location": "Chat with Flow bot",
            "body/body/messageBody": "@{string(outputs('Card'))}",
            "body/body/updateMessage": f"Thanks, your answer was added to IT-@{{{ID}}}.",
            "body/body/recipient/to": "@outputs('ReqEmail')"}),
            "type": "OpenApiConnectionWebhook", "limit": {"timeout": "P5D"}}),
        ("Answer", compose("@trim(coalesce(body('Ask_in_Teams')?['data']?['answer'],''))")),
        ("Get_responder", {**op("shared_office365users", "UserProfile_V2",
            {"id": "@{coalesce(body('Ask_in_Teams')?['responder']?['objectId'],body('Ask_in_Teams')?['responder']?['userPrincipalName'],outputs('ReqEmail'))}"}),
            "runAfter": ra("Answer")}),
        ("Get_item", {**sp("GetItem", {"id": f"@{ID}"}), "runAfter": ra("Get_responder", status=("Succeeded", "Failed"))}),
        ("Has_answer", cond(f"@and(not(empty(outputs('Answer'))),not(equals({gi}?['Status']?['Value'],'{S_CLOSED}')))", yes)),
    ])
    trig = sp_trigger("GetOnUpdatedItems")
    trig["conditions"] = [{"expression": "@greater(length(coalesce(triggerOutputs()?['body/TeamsQuestion'],'')),0)"}]
    return "HD10 Ask requester in Teams", {"When_a_question_is_ready": trig}, actions

# ---------- HD11 Tech thread replies ----------
def hd11():
    M = "outputs('Msg')"
    gi = "first(body('Find_ticket')?['value'])"
    actions = seq([
        ("Msg", compose("@coalesce(triggerBody()?['value']?[0],triggerBody())")),
        ("Is_a_person_replying", cond(
            f"@and(not(empty(coalesce({M}?['replyToId'],''))),not(empty(coalesce({M}?['from']?['user']?['id'],''))))", seq([
            ("Reply_to_text", op("shared_conversionservice", "HtmlToText", {"Content": f"@{{coalesce({M}?['body']?['content'],'')}}"})),
            ("ReplyText", compose("@trim(body('Reply_to_text'))")),
            ("Not_internal", cond("@and(not(empty(outputs('ReplyText'))),not(startsWith(toLower(outputs('ReplyText')),'#internal')))", seq([
                ("Find_ticket", sp("GetItems", {"$filter": f"ChannelMessageId eq '@{{{M}?['replyToId']}}'", "$top": 1})),
                ("Ticket_found_and_open", cond(
                    f"@and(greater(length(body('Find_ticket')?['value']),0),not(equals({gi}?['Status']?['Value'],'{S_CLOSED}')))", {
                    "Send_to_user": sp("PatchItem", {
                        "id": f"@{gi}?['ID']", "item/Title": f"@{gi}?['Title']",
                        "item/NoteToUser": "@outputs('ReplyText')"})})),
            ]))),
        ]))),
    ])
    trig = {"When_a_tech_replies_in_the_Help_Desk_channel": {"type": "OpenApiConnectionWebhook",
        "inputs": {"parameters": {"threadType": "channel", "requestBody/channels": [CHANNEL], "requestBody/groupId": GROUP},
                   "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_teams", "connectionName": "shared_teams",
                            "operationId": "WebhookNewMessageTrigger"},
                   "authentication": "@parameters('$authentication')"}}}
    return "HD11 Tech thread replies", trig, actions

# ---------- HD12 Requester replies in Teams (chat with the IT Support account) ----------
def hd12():
    M = "outputs('Msg')"
    # The WebhookChatMessageTrigger notification only carries IDs (subscriptionId, conversationId, messageId,
    # resourceData.id/@odata.id) - no 'from' or 'body'. Microsoft Graph only sends that lite shape for chat
    # messages unless the subscription is set up for encrypted rich notifications; the flow has to fetch the
    # actual message itself. Found 2026-10-25 full-circle-testing the new sponsor-reply path: every recorded run
    # in HD12's 28-day history short-circuited at "From a person" (always False) because M?['from'] was never
    # present on the raw notification, so no inbound Teams reply - ticket or (new) project - had ever actually
    # been processed. Fixed with a "Get message details" call (Teams connector GetMessageDetails, run as
    # IT_Support so it can read any chat IT_Support is a member of) right after the trigger, keyed off the IDs
    # the notification does carry; MD is that resolved message, used everywhere M's 'from'/'body' were assumed.
    MD = "body('Get_message_details')"
    stamp = f"convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm')"
    tk = "outputs('Ticket')"
    TID = f"{tk}?['ID']"
    sender = "toLower(coalesce(body('Get_sender')?['mail'],body('Get_sender')?['userPrincipalName'],''))"
    who = "coalesce(body('Get_sender')?['displayName'],outputs('SenderEmail'))"
    chat_id = f"{M}?['chatId']"
    def reply_here(name, html):
        return {"type": "Scope", "runAfter": {}, "actions": support_msg(name + "_msg", "@outputs('SenderEmail')", html)}
    found = seq([
        ("NewStatus", compose(f"@if(contains({arr([S_WAIT, S_SCHED, S_RES])},{tk}?['Status']?['Value']),'{S_PROG}',{tk}?['Status']?['Value'])")),
        ("Update_ticket", sp("PatchItem", {
            "id": f"@{TID}", "item/Title": f"@{tk}?['Title']",
            "item/Status/Value": "@outputs('NewStatus')", "item/LastNotifiedStatus": "@outputs('NewStatus')",
            "item/WaitingHours": 0, "item/ResolvedHours": 0,
            "item/Conversation": f"@concat('[',{stamp},' | ',{who},' (Teams)] ',outputs('ReplyText'),decodeUriComponent('%0A%0A'),coalesce({tk}?['Conversation'],''))"})),
        ("Thank_requester", reply_here("Thank_requester",
            f"<p>Thanks, your message was added to <b>IT-@{{{TID}}}</b>"
            f"@{{if(equals(outputs('NewStatus'),{tk}?['Status']?['Value']),'',concat(' and the ticket is back with IT (',outputs('NewStatus'),')'))}}.</p>")),
        ("Has_a_tech", {**cond(f"@not(empty(coalesce({tk}?['AssignedTo']?['Email'],'')))", {
            "Tell_the_tech": scope_msg("Tell_tech_msg", f"@{{{tk}?['AssignedTo']?['Email']}}",
                f"<p>@{{{who}}} replied on IT-@{{{TID}}} in Teams:</p><p>@{{outputs('ReplyText')}}</p><p>{link(TID)}</p>")}),
            "runAfter": ra("Thank_requester", status=("Succeeded", "Failed"))}),
        ("Has_channel_post", {**cond(f"@not(empty(coalesce({tk}?['ChannelMessageId'],'')))", {
            "Reply_in_channel": channel_reply(f"@{tk}?['ChannelMessageId']",
                f"<p>@{{{who}}} replied in Teams: @{{outputs('ReplyText')}}</p>"
                f"@{{if(equals(outputs('NewStatus'),{tk}?['Status']?['Value']),'',concat('<p>IT-',string({TID}),' is now <b>',outputs('NewStatus'),'</b></p>'))}}")}),
            "runAfter": ra("Has_a_tech", status=("Succeeded", "Failed", "Skipped"))}),
    ])
    ask_list = {"type": "Select", "inputs": {"from": "@body('Open_tickets')?['value']",
        "select": "@concat('IT-',string(item()?['ID']),' - ',coalesce(item()?['Title'],''))"}}
    import json as _j
    found = _j.loads(_j.dumps(found).replace("outputs('ReplyText')", "outputs('FinalText')"))
    found["Delete_pending"] = {**cond("@outputs('IsBare')", {"Delete_pending_item": sp_http("POST",
        "_api/web/lists/getbytitle('Pending replies')/items(@{outputs('PendingFresh')?['Id']})", headers={"IF-MATCH": "*", "X-HTTP-Method": "DELETE"})}),
        "runAfter": ra("Has_channel_post", status=("Succeeded", "Failed", "Skipped"))}
    not_found = {"Any_open_tickets": {**cond("@greater(length(body('Open_tickets')?['value']),0)",
        seq([("Ticket_list", ask_list),
             ("Save_pending", sp_http("POST", "_api/web/lists/getbytitle('Pending replies')/items",
                body="{\"Title\":\"@{outputs('SenderEmail')}\",\"ReplyText\":@{substring(string(createArray(outputs('ReplyText'))),1,sub(length(string(createArray(outputs('ReplyText')))),2))}}")),
             ("Ask_which", {**reply_here("Ask_which",
                "<p>Which ticket is this for? Just reply with the ticket number (e.g. <b>IT-15</b>) and I'll add your message to it.</p>"
                "<p>Your open tickets:<br>@{join(body('Ticket_list'),'<br>')}</p>"),
                "runAfter": ra("Save_pending", status=("Succeeded", "Failed"))})]),
        {"Say_no_ticket": reply_here("Say_no_ticket",
            f"<p>I couldn't find an open ticket for you. To log a new request, email <a href=\"mailto:{MBX}\">{MBX}</a> or use the IT Support form in Teams.</p>")}), "runAfter": {}}}
    # No open ticket matched, but there's a recently-closed one for this sender: start a follow-up
    # ticket tagged to it, the same way HD1 does for an email reply to a closed ticket.
    ck = "outputs('ClosedTicket')"
    closed_followup = seq([
        ("Create_followup_item", sp("PostItem", {
            "item/Title": f"@concat('Follow-up: ',{ck}?['Title'])",
            "item/Details": (f"@concat('Follow-up to IT-',string({ck}?['ID']),'. ',outputs('ReplyText'),"
                              f"if(empty(coalesce({ck}?['Resolution'],'')),'',concat(decodeUriComponent('%0A%0A'),"
                              f"'Previous resolution (IT-',string({ck}?['ID']),'): ',{ck}?['Resolution'])))"),
            "item/Requester/Claims": "@outputs('SenderEmail')",
            "item/RequesterEmail": "@outputs('SenderEmail')",
            "item/Source/Value": "Teams message",
            "item/Category/Value": f"@coalesce({ck}?['Category']?['Value'],'Uncategorized')",
            "item/Priority/Value": f"@coalesce({ck}?['Priority']?['Value'],'Medium')",
            "item/Conversation": f"@concat('[',{stamp},'] Ticket created from Teams message by ',outputs('SenderEmail'),' — follow-up to IT-',string({ck}?['ID']),decodeUriComponent('%0A%0A'))",
            "item/Status/Value": S_OPEN})),
        ("Tell_requester_followup", reply_here("Tell_requester_followup",
            f"<p>IT-@{{{ck}?['ID']}} was already resolved, so I started a new ticket for this: <b>IT-@{{body('Create_followup_item')?['ID']}}</b>.</p><p>@{{outputs('ReplyText')}}</p>")),
    ])
    not_found_or_followup = {"Check_closed_ticket": cond(f"@not(empty({ck}))", closed_followup, not_found)}
    # No open ticket matched: before falling back to "I couldn't find an open ticket for you", check whether
    # this sender is a Sponsor, Owner or creator on an IT Project - if so, treat their message as a Sponsor
    # Updates note on that project instead (the Teams-side mirror of HD1's [PRJ-n] email-reply handling).
    #
    # Found 2026-10-25 full-circle-testing the sponsor-reply feature: a bare "[PRJ-2] ..." Teams reply got
    # attached to the sender's one open ticket instead, because the old Ticket_found check always ran the
    # existing ticket-matching logic first - and that logic's "exactly one open ticket" convenience fallback
    # (see "Ticket" below) matched regardless of what the message actually said, silently overriding an explicit
    # [PRJ-n] tag. Fixed the same way HD1 already handles it: an explicit PRJ- reference is checked before any
    # ticket matching at all (PrjRefPart/PrjRefNum, mirroring HD1's own safety-net digit scan), and when present
    # it routes straight to a lookup of that *specific* project (not just "any project this sender can update"),
    # with its own not-found/not-authorized alert - never silently falls through to ticket matching.
    pr = "outputs('SponsorProject')"
    PRID = f"{pr}?['ID']"
    project_found = seq([
        ("Update_sponsor_project", sp("PatchItem", {"table": PROJECTS, "id": f"@{PRID}", "item/Title": f"@{pr}?['Title']",
            "item/SponsorUpdates": f"@concat('[',{stamp},' | ',{who},' (Teams)] ',outputs('ReplyText'),decodeUriComponent('%0A%0A'),coalesce({pr}?['SponsorUpdates'],''))"})),
        ("Thank_sponsor", reply_here("Thank_sponsor",
            f"<p>Thanks, your update was added to <b>PRJ-@{{{PRID}}}: @{{{pr}?['Title']}}</b>.</p>")),
    ])
    project_tag_notfound = seq([
        ("Tell_sender_project_tag_notfound", reply_here("Tell_sender_project_tag_notfound",
            "<p>I couldn't match that to a project you're the sponsor, owner, or creator of. If you meant a ticket instead, "
            "just reply with the ticket number (e.g. <b>IT-15</b>).</p>")),
        ("Alert_bad_project_reply_teams", {**scope_msg("Alert_bad_prj_reply_dm", ME,
            f"<p>A Teams reply naming PRJ-@{{outputs('PrjRefNum')}} from @{{{who}}} couldn't be matched to a project they're allowed to update.</p>"),
            "runAfter": ra("Tell_sender_project_tag_notfound", status=("Succeeded", "Failed"))}),
    ])
    person = seq([
        ("Get_sender", op("shared_office365users", "UserProfile_V2", {"id": f"@{{{MD}?['from']?['user']?['id']}}"})),
        ("SenderEmail", compose(f"@{sender}")),
        ("Not_IT_Support", cond(f"@not(equals(outputs('SenderEmail'),toLower('{SUPPORT}')))", seq([
            ("Msg_to_text", op("shared_conversionservice", "HtmlToText", {"Content": f"@{{coalesce({MD}?['body']?['content'],'')}}"})),
            ("ReplyText", compose("@trim(body('Msg_to_text'))")),
            ("Has_text", cond("@not(empty(outputs('ReplyText')))", seq([
                ("Open_tickets", sp("GetItems", {"$filter": f"RequesterEmail eq '@{{outputs('SenderEmail')}}' and Status ne '{S_CLOSED}'",
                                                 "$orderby": "Modified desc", "$top": 50})),
                # a ticket number in the message wins; otherwise the newest ticket waiting on them; otherwise the newest open one
                ("Get_pending", sp_http("GET", "_api/web/lists/getbytitle('Pending replies')/items?$filter=Title eq '@{outputs('SenderEmail')}'&$orderby=Created desc&$top=1&$select=Id,Title,ReplyText,Created")),
                ("Named", {"runAfter": ra("Get_pending", status=("Succeeded", "Failed", "TimedOut")), "type": "Query", "inputs": {"from": "@body('Open_tickets')?['value']",
                    "where": "@contains(concat(' ',toUpper(replace(replace(replace(replace(replace(replace(outputs('ReplyText'),'#',' '),'.',' '),',',' '),'?',' '),'!',' '),':',' ')),' '),concat('IT-',string(item()?['ID']),' '))"}}),
                ("Waiting", {"type": "Query", "inputs": {"from": "@body('Open_tickets')?['value']",
                    "where": f"@contains({arr([S_WAIT, S_SCHED, S_RES])},item()?['Status']?['Value'])"}}),
                ("Ticket", compose("@if(not(empty(body('Named'))),first(body('Named')),if(equals(length(body('Open_tickets')?['value']),1),first(body('Open_tickets')?['value']),null))")),
                ("Closed_tickets", sp("GetItems", {"$filter": f"RequesterEmail eq '@{{outputs('SenderEmail')}}' and Status eq '{S_CLOSED}' and Modified ge '@{{formatDateTime(addDays(utcNow(),-14),'yyyy-MM-ddTHH:mm:ssZ')}}'",
                                                    "$orderby": "Modified desc", "$top": 20})),
                ("Named_closed", {"type": "Query", "inputs": {"from": "@body('Closed_tickets')?['value']",
                    "where": "@contains(concat(' ',toUpper(replace(replace(replace(replace(replace(replace(outputs('ReplyText'),'#',' '),'.',' '),',',' '),'?',' '),'!',' '),':',' ')),' '),concat('IT-',string(item()?['ID']),' '))"}}),
                ("ClosedTicket", compose("@if(not(empty(body('Named_closed'))),first(body('Named_closed')),if(equals(length(body('Closed_tickets')?['value']),1),first(body('Closed_tickets')?['value']),null))")),
                ("PendingFresh", compose("@if(and(not(empty(body('Get_pending')?['value'])),less(sub(ticks(utcNow()),ticks(first(body('Get_pending')?['value'])?['Created'])),864000000000)),first(body('Get_pending')?['value']),null)")),
                ("IsBare", compose("@and(not(empty(outputs('Ticket'))),not(empty(outputs('PendingFresh'))),equals(replace(replace(replace(toUpper(outputs('ReplyText')),' ',''),'#',''),'.',''),concat('IT-',string(outputs('Ticket')?['ID']))))")),
                ("FinalText", compose("@if(outputs('IsBare'),outputs('PendingFresh')?['ReplyText'],outputs('ReplyText'))")),
                # safety net: does the message explicitly name a project, e.g. "[PRJ-2] ..." - mirrors HD1's
                # own PrjRefPart/PrjRefNum digit scan so a bracketed tag or a bare mention both work.
                ("PrjRefPart", compose("@concat(coalesce(split(toUpper(outputs('ReplyText')),'PRJ-')?[1],''),'     ')")),
                ("PrjRefNum", compose("@if(not(contains('0123456789',substring(outputs('PrjRefPart'),0,1))),0,"
                                      "int(substring(outputs('PrjRefPart'),0,if(not(contains('0123456789',substring(outputs('PrjRefPart'),1,1))),1,"
                                      "if(not(contains('0123456789',substring(outputs('PrjRefPart'),2,1))),2,"
                                      "if(not(contains('0123456789',substring(outputs('PrjRefPart'),3,1))),3,"
                                      "if(not(contains('0123456789',substring(outputs('PrjRefPart'),4,1))),4,5)))))))")),
                ("Sponsor_projects_raw", sp("GetItems", {"table": PROJECTS, "$orderby": "Modified desc", "$top": 200})),
                # Owner and Project Sponsor both allow multiple people (see hd19()'s docstring for the
                # multi-value Person field shape), so membership is an exact-match check against each
                # project row's own Owner/ProjectSponsor array - via people_contains()'s JSON-substring
                # match (see that function's docstring for why: this Query's own "where" is itself just an
                # expression with no action of its own to attach a nested Select/Filter-array action to,
                # and Power Automate has no inline select()/filter() expression function to fall back on).
                # item() in each people_contains() call still refers to this Query's own outer loop
                # variable (each project row) - it's inside concat()/string(), not a lambda of its own.
                ("Sponsor_projects", {"type": "Query", "inputs": {"from": "@body('Sponsor_projects_raw')?['value']",
                    "where": ("@or(" + people_contains("item()?['ProjectSponsor']", "outputs('SenderEmail')") + ","
                              "equals(toLower(coalesce(item()?['SponsorEmail'],'')),outputs('SenderEmail')),"
                              + people_contains("item()?['Owner']", "outputs('SenderEmail')") + ","
                              "equals(toLower(coalesce(item()?['Author']?['Email'],'')),outputs('SenderEmail')))")}}),
                # when a PRJ-n tag is present, match ONLY that specific project among the ones this sender is
                # authorized on (never a different one); with no tag, keep the old "first project they can update" fallback
                ("Tagged_sponsor_project", {"type": "Query", "inputs": {"from": "@body('Sponsor_projects')",
                    "where": "@equals(item()?['ID'],outputs('PrjRefNum'))"}}),
                ("SponsorProject", compose(
                    "@if(greater(outputs('PrjRefNum'),0),"
                    "if(greater(length(body('Tagged_sponsor_project')),0),first(body('Tagged_sponsor_project')),null),"
                    "if(greater(length(body('Sponsor_projects')),0),first(body('Sponsor_projects')),null))")),
                # an explicit PRJ-n tag always wins over ticket matching (same precedence as HD1's email side) -
                # the old code ran ticket matching unconditionally first, and its "exactly one open ticket" fallback
                # (see "Ticket" above) would silently hijack a tagged project reply; see the comment above project_found.
                #
                # Structured as exactly ONE new If level here, not two: Project_or_tag_check's True branch picks
                # between project_found/project_tag_notfound (shallow either way), and its False branch drops
                # straight into not_found_or_followup at the SAME depth it sat at before this fix. An earlier
                # version nested a second If (Has_explicit_project_tag) in front of not_found_or_followup, which
                # pushed the pre-existing not_found/Any_open_tickets chain's deepest actions (Ask_which_msg*,
                # Say_no_ticket_msg*) from level 8 to level 9 and failed Power Automate's import-time nesting-depth
                # check ("...nested at level '9' which exceeds the maximum nesting limit of '8'") - found
                # 2026-10-25 when Martin tried to import that build.
                ("Ticket_found", cond(f"@and(equals(outputs('PrjRefNum'),0),not(empty({tk})))", found,
                    {"Project_or_tag_check": cond(f"@or(not(empty({pr})),greater(outputs('PrjRefNum'),0))",
                        {"Tag_project_resolved": cond(f"@not(empty({pr}))", project_found, project_tag_notfound)},
                        not_found_or_followup)})),
            ]))),
        ]))),
    ])
    actions = seq([
        ("Msg", compose("@coalesce(triggerBody()?['value']?[0],triggerBody())")),
        ("Get_message_details", op("shared_teams_1", "GetMessageDetails", {
            "messageId": f"@{{{M}?['messageId']}}", "threadType": "groupchat", "body/recipient": f"@{{{M}?['conversationId']}}"})),
        ("From_a_person", {**cond(f"@not(empty(coalesce({MD}?['from']?['user']?['id'],'')))", person),
            "runAfter": ra("Get_message_details", status=("Succeeded", "Failed"))}),
    ])
    trig = {"When_someone_messages_IT_Support": {"type": "OpenApiConnectionWebhook",
        "inputs": {"parameters": {},
                   "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_teams", "connectionName": "shared_teams_1",
                            "operationId": "WebhookChatMessageTrigger"},
                   "authentication": "@parameters('$authentication')"}}}
    return "HD12 Requester replies in Teams", trig, actions


def hd13():
    H = {"Accept": "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose"}
    xml = "<Field Type='Note' DisplayName='ReplyText' Name='ReplyText' StaticName='ReplyText' NumLines='8' RichText='FALSE' />"
    trig = {"manual": {"type": "Request", "kind": "Button", "inputs": {"schema": {"type": "object", "properties": {}, "required": []}}}}
    acts = seq([
        ("Create_list", sp_http("POST", "_api/web/lists", headers=H,
            body='{"__metadata":{"type":"SP.List"},"Title":"Pending replies","BaseTemplate":100,"Description":"System use: holds a requester reply until they say which ticket it is for.","OnQuickLaunch":false}')),
        ("Add_ReplyText_field", sp_http("POST", "_api/web/lists/getbytitle('Pending replies')/fields/createfieldasxml", headers=H,
            body='{"parameters":{"__metadata":{"type":"SP.XmlSchemaFieldCreationInformation"},"SchemaXml":"' + xml.replace('"', '\\"') + '"}}')),
    ])
    return "HD13 Setup pending-replies list (run once)", trig, acts

# ---------- HD14 Teams app intake (Power Apps V2 trigger) ----------
def hd14():
    def prop(title, desc):
        return {"title": title, "type": "string", "x-ms-dynamically-added": True, "description": desc, "x-ms-content-hint": "TEXT"}
    trig = {"manual": {"type": "Request", "kind": "PowerAppV2", "inputs": {"schema": {"type": "object",
        "properties": {"text": prop("What", "One-line summary"), "text_1": prop("Category", "Category choice"),
                       "text_2": prop("Urgency", "I can't work at all / It's slowing me down / No rush"),
                       "text_3": prop("Details", "Details"),
                       "text_4": prop("OnBehalf", "Optional: email of the person this ticket is for"),
                       "text_5": prop("Screenshot", "Optional: picture as a data URI from the app")},
        "required": ["text", "text_1", "text_2", "text_3"]}}}}
    T = lambda k: f"triggerBody()?['{k}']"
    acts = seq([
        ("Submitter", compose("@toLower(base64ToString(triggerOutputs()?['headers']?['x-ms-user-email-encoded']))")),
        ("OnBehalf", compose("@toLower(trim(coalesce(triggerBody()?['text_4'],'')))")),
        ("Requester", compose("@if(and(endsWith(outputs('OnBehalf'),'@branthaven.com'),not(contains(outputs('OnBehalf'),' ')),not(equals(outputs('OnBehalf'),'@branthaven.com'))),outputs('OnBehalf'),outputs('Submitter'))")),
        ("Create_item", sp("PostItem", {
            "item/Title": f"@substring(coalesce({T('text')},'(no summary)'),0,min(length(coalesce({T('text')},'(no summary)')),250))",
            "item/Category/Value": f"@if(empty(coalesce({T('text_1')},'')),'Uncategorized',{T('text_1')})",
            "item/Priority/Value": f"@if(startsWith(coalesce({T('text_2')},''),'I can'),'High',if(equals({T('text_2')},'No rush'),'Low','Medium'))",
            "item/Details": f"@if(equals(outputs('Requester'),outputs('Submitter')),{T('text_3')},concat('[Logged by ',outputs('Submitter'),' on behalf of ',outputs('Requester'),']',decodeUriComponent('%0A%0A'),{T('text_3')}))",
            "item/Requester/Claims": "@outputs('Requester')",
            "item/RequesterEmail": "@outputs('Requester')",
            "item/Source/Value": "Teams form",
            "item/Conversation": f"@concat('[',convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm'),'] Ticket created from Teams form by ',outputs('Requester'),decodeUriComponent('%0A%0A'))",
            "item/Status/Value": S_OPEN})),
        ("Has_a_screenshot", cond("@contains(coalesce(triggerBody()?['text_5'],''),'base64,')", {
            "Shot_clean": compose("@replace(triggerBody()?['text_5'],'\"','')"),
            "Add_screenshot": sp("CreateAttachment", {
                "itemId": "@body('Create_item')?['ID']",
                "displayName": "@concat('screenshot.',replace(first(split(last(split(outputs('Shot_clean'),'data:image/')),';')),'jpeg','jpg'))",
                "body": "@base64ToBinary(last(split(outputs('Shot_clean'),'base64,')))"}, after={"Shot_clean": ["Succeeded"]})})),
        ("Respond_to_the_app", {"type": "Response", "kind": "PowerApp", "runAfter": ra("Has_a_screenshot"),
            "inputs": {"statusCode": 200, "body": {"ticketid": "@{body('Create_item')?['ID']}"},
                       "schema": {"type": "object", "properties": {"ticketid": {"title": "ticketid", "x-ms-dynamically-added": True, "type": "string"}}}}}),
    ])
    return "HD14 Teams app intake", trig, acts


# ---------- HD15 Dashboard stats (feeds the DashStats list behind the SharePoint dashboard) ----------
DASH = "24631152-bdde-4654-82a4-d64bb0bb9b3b"
DASH_HEALTH = 14   # "Health" row on the Live overview list, rendered by the "Chart - Health" view on the IT dashboard
def hd15():
    ST7 = [S_OPEN, S_ASSIGNED, S_PROG, S_ESC, S_WAIT, S_SCHED, S_RES]
    PR3 = ["High", "Medium", "Low"]
    CAT13 = CATEGORIES
    conc1 = {"runtimeConfiguration": {"concurrency": {"repetitions": 1}}}
    def chart(p, src, field, names, item_id, sub_total=None, title=None):
        key = f"item()?['{field}']?['Value']"
        if field == "Category":
            where = f"@or(equals({key},items('Foreach_{p}')),and(equals(items('Foreach_{p}'),'Uncategorized'),or(empty({key}),equals({key},'Uncategorised'))))"
        else:
            where = f"@equals({key},items('Foreach_{p}'))"
        loop = {"type": "Foreach", "foreach": "@" + arr(names), "actions": seq([
            (f"Count_{p}", {"type": "Query", "inputs": {"from": f"@body('{src}')", "where": where}}),
            (f"Label_{p}", {"type": "AppendToArrayVariable", "inputs": {"name": "varLabels", "value":
                f"@if(equals(length(body('Count_{p}')),0),'',concat(items('Foreach_{p}'),' [',string(length(body('Count_{p}'))),']'))"}}),
            (f"Share_{p}", {"type": "AppendToArrayVariable", "inputs": {"name": "varShares", "value":
                f"@float(formatNumber(div(mul(float(length(body('Count_{p}'))),100),float(max(length(body('{src}')),1))),'0.00'))"}}),
        ]), **conc1}
        item = {"id": item_id, "item/Title": title or {"Status": "Status", "Priority": "Priority", "Category": "Category"}[field],
                "item/Total": f"@length(body('{src}'))"}
        for i in range(max(14, len(names))):
            item[f"item/Slot{i+1}"] = f"@coalesce(variables('varLabels')?[{i}],'')"
            item[f"item/Share{i+1}"] = f"@coalesce(variables('varShares')?[{i}],0)"
        return [
            (f"Reset_{p}_labels", {"type": "SetVariable", "inputs": {"name": "varLabels", "value": []}}),
            (f"Reset_{p}_shares", {"type": "SetVariable", "inputs": {"name": "varShares", "value": []}}),
            (f"Foreach_{p}", loop),
            (f"Update_{p}", sp("PatchItem", {"table": DASH, **item})),
        ]
    d30 = "ticks(addDays(utcNow(),-30))"
    def ranked(p, src, keyexpr, item_id, title, sub, n, total_src):
        loop = {"type": "Foreach", "foreach": f"@outputs('Distinct_{p}')", "actions": seq([
            (f"Count_{p}", {"type": "Query", "inputs": {"from": f"@body('Keys_{p}')", "where": f"@equals(item(),items('Foreach_{p}'))"}}),
            (f"Add_{p}", {"type": "AppendToArrayVariable", "inputs": {"name": "varRanks", "value":
                {"n": f"@length(body('Count_{p}'))", "l": f"@items('Foreach_{p}')"}}}),
        ]), **conc1}
        srt = f"outputs('Sorted_{p}')"
        item = {"id": item_id, "item/Title": title, "item/Sub": sub, "item/Total": f"@length(body('{total_src}'))"}
        for i in range(n):
            item[f"item/Slot{i+1}"] = f"@if(greater(length({srt}),{i}),concat({srt}?[{i}]?['l'],'  ',string({srt}?[{i}]?['n'])),'')"
            item[f"item/Share{i+1}"] = (f"@if(greater(length({srt}),{i}),float(formatNumber(div(mul(float({srt}?[{i}]?['n']),100),"
                                        f"float({srt}?[0]?['n'])),'0.00')),0)")
        return [
            (f"Keys_{p}", {"type": "Select", "inputs": {"from": f"@body('{src}')", "select": "@" + keyexpr}}),
            (f"Distinct_{p}", {"type": "Compose", "inputs": f"@union(body('Keys_{p}'),body('Keys_{p}'))"}),
            (f"Reset_{p}", {"type": "SetVariable", "inputs": {"name": "varRanks", "value": []}}),
            (f"Foreach_{p}", loop),
            (f"Sorted_{p}", {"type": "Compose", "inputs": f"@take(reverse(sort(variables('varRanks'),'n')),{n})"}),
            (f"Update_{p}", sp("PatchItem", {"table": DASH, **item})),
        ]
    def trend():
        day = "formatDateTime(addDays(convertTimeZone(utcNow(),'UTC','" + TZ + "'),sub(items('Foreach_trend'),29)),'yyyy-MM-dd')"
        lab = ("if(equals(items('Foreach_trend'),29),'Today',if(equals(mod(sub(29,items('Foreach_trend')),5),0),"
               "formatDateTime(addDays(convertTimeZone(utcNow(),'UTC','" + TZ + "'),sub(items('Foreach_trend'),29)),'MMM d'),''))")
        loop = {"type": "Foreach", "foreach": "@range(0,30)", "actions": seq([
            ("Count_trend", {"type": "Query", "inputs": {"from": "@body('Days_trend')", "where": f"@equals(item(),{day})"}}),
            ("Label_trend", {"type": "AppendToArrayVariable", "inputs": {"name": "varLabels", "value": "@" + lab}}),
            ("Share_trend", {"type": "AppendToArrayVariable", "inputs": {"name": "varShares", "value": "@length(body('Count_trend'))"}}),
        ]), **conc1}
        item = {"id": 6, "item/Title": "Trend", "item/Sub": "Tickets opened per day, last 30 days",
                "item/Total": "@length(body('Days_trend'))", "item/OpenCount": "@max(variables('varShares'))"}
        for i in range(30):
            item[f"item/Slot{i+1}"] = f"@coalesce(variables('varLabels')?[{i}],'')"
            item[f"item/Share{i+1}"] = f"@coalesce(variables('varShares')?[{i}],0)"
        return [
            ("Days_trend", {"type": "Select", "inputs": {"from": "@body('Created_30')",
                "select": "@formatDateTime(convertTimeZone(item()?['Created'],'UTC','" + TZ + "'),'yyyy-MM-dd')"}}),
            ("Reset_trend_labels", {"type": "SetVariable", "inputs": {"name": "varLabels", "value": []}}),
            ("Reset_trend_shares", {"type": "SetVariable", "inputs": {"name": "varShares", "value": []}}),
            ("Foreach_trend", loop),
            ("Update_trend", sp("PatchItem", {"table": DASH, **item})),
        ]
    def service():
        pct = lambda num, den: (f"@if(equals(length(body('{den}')),0),-1,div(add(mul(length(body('{num}')),100),"
                                f"div(length(body('{den}')),2)),length(body('{den}'))))")
        item = {"id": 7, "item/Title": "Service", "item/Sub": "Service levels, last 30 days",
                "item/Share1": pct("Resp_ok", "Resp_elig"), "item/Slot1": "Responded within 4 h",
                "item/Slot4": "@concat(string(length(body('Resp_ok'))),' of ',string(length(body('Resp_elig'))),' tickets')",
                "item/Share2": pct("Res_ok", "Resolved_30"), "item/Slot2": "Resolved within 24 h",
                "item/Slot5": "@if(equals(length(body('Resolved_30')),0),'no resolved tickets yet',concat(string(length(body('Res_ok'))),' of ',string(length(body('Resolved_30'))),' tickets'))",
                "item/Share3": "@if(equals(length(body('Responded')),0),-1,div(variables('varSumMin'),length(body('Responded'))))",
                "item/Slot3": "Average first response",
                "item/Slot6": "@concat(string(length(body('Responded'))),' tickets answered')"}
        # Response timer starts when support is open: a ticket raised before 7am starts at 7:00 that day, one raised
        # at/after 9pm starts at 7:00 the next day. (Ticks of the effective start, in UTC.)
        start = sla_start_ticks("item()?['Created']")
        fr = "ticks(item()?['FirstResponseOn'])"
        wait = f"max(0,sub({fr},{start}))"
        return [
            ("Resp_elig", {"type": "Query", "inputs": {"from": "@body('Created_30')", "where":
                f"@if(empty(item()?['FirstResponseOn']),less({start},ticks(addHours(utcNow(),-4))),true)"}}),
            ("Resp_ok", {"type": "Query", "inputs": {"from": "@body('Resp_elig')", "where":
                f"@if(empty(item()?['FirstResponseOn']),false,lessOrEquals({wait},144000000000))"}}),
            ("Responded", {"type": "Query", "inputs": {"from": "@body('Created_30')", "where": "@not(empty(item()?['FirstResponseOn']))"}}),
            ("Resp_minutes", {"type": "Select", "inputs": {"from": "@body('Responded')",
                "select": f"@div({wait},600000000)"}}),
            ("Reset_sum", {"type": "SetVariable", "inputs": {"name": "varSumMin", "value": 0}}),
            ("Foreach_sum", {"type": "Foreach", "foreach": "@body('Resp_minutes')", "actions": seq([
                ("Add_sum", {"type": "IncrementVariable", "inputs": {"name": "varSumMin", "value": "@items('Foreach_sum')"}})]), **conc1}),
            ("Resolved_30", {"type": "Query", "inputs": {"from": "@body('All_tickets')?['value']", "where":
                f"@if(empty(item()?['ResolvedOn']),false,greaterOrEquals(ticks(item()?['ResolvedOn']),{d30}))"}}),
            ("Res_ok", {"type": "Query", "inputs": {"from": "@body('Resolved_30')", "where":
                "@lessOrEquals(float(coalesce(item()?['SLAHours'],0)),24)"}}),   # SLA clock hours (HD6's 8am-9pm counter, pauses while waiting on the user)
            ("Update_service", sp("PatchItem", {"table": DASH, **item})),
        ]
    catkey = "if(or(empty(item()?['Category']?['Value']),equals(item()?['Category']?['Value'],'Uncategorised')),'Uncategorized',item()?['Category']?['Value'])"
    acts = seq([
        ("Init_varLabels", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varLabels", "type": "array", "value": []}]}}),
        ("Init_varShares", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varShares", "type": "array", "value": []}]}}),
        ("Init_varRanks", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varRanks", "type": "array", "value": []}]}}),
        ("Init_varSumMin", {"type": "InitializeVariable", "inputs": {"variables": [{"name": "varSumMin", "type": "integer", "value": 0}]}}),
        ("All_tickets", sp("GetItems", {"$top": 5000})),
        ("Open_all", {"type": "Query", "inputs": {"from": "@body('All_tickets')?['value']",
            "where": f"@not(equals(item()?['Status']?['Value'],'{S_CLOSED}'))"}}),
        ("Open_working", {"type": "Query", "inputs": {"from": "@body('Open_all')",
            "where": f"@not(equals(item()?['Status']?['Value'],'{S_RES}'))"}}),
        ("Unassigned", {"type": "Query", "inputs": {"from": "@body('Open_working')",
            "where": "@empty(coalesce(item()?['AssignedTo']?['Email'],''))"}}),
        ("Overdue", {"type": "Query", "inputs": {"from": "@body('Open_working')",
            "where": f"@greaterOrEquals(int(coalesce(item()?['SLAHours'],0)),{DAY_H})"}}),
        ("Waiting", {"type": "Query", "inputs": {"from": "@body('Open_all')",
            "where": f"@equals(item()?['Status']?['Value'],'{S_WAIT}')"}}),
        ("Pending", {"type": "Query", "inputs": {"from": "@body('Open_all')",
            "where": f"@equals(item()?['Status']?['Value'],'{S_RES}')"}}),
        ("Closed_30", {"type": "Query", "inputs": {"from": "@body('All_tickets')?['value']",
            "where": f"@and(equals(item()?['Status']?['Value'],'{S_CLOSED}'),greaterOrEquals(ticks(coalesce(item()?['ClosedOn'],'1970-01-01T00:00:00Z')),{d30}))"}}),
        ("Created_30", {"type": "Query", "inputs": {"from": "@body('All_tickets')?['value']",
            "where": f"@greaterOrEquals(ticks(item()?['Created']),{d30})"}}),
        ("All_projects", spp("GetItems", {"$top": 5000})),
        ("Projects_list", {"type": "Select", "inputs": {"from": "@body('All_projects')?['value']", "select": "@item()"}}),
        ("Active_projects", {"type": "Query", "inputs": {"from": "@body('Projects_list')",
            "where": "@or(" + ",".join(f"equals(item()?['Status']?['Value'],'{s}')" for s in PRJ_ACTIVE) + ")"}}),
        ("Update_KPIs", sp("PatchItem", {"table": DASH, "id": 5, "item/Title": "KPIs",
            "item/OpenCount": "@length(body('Open_working'))", "item/Unassigned": "@length(body('Unassigned'))",
            "item/Overdue": "@length(body('Overdue'))", "item/Waiting": "@length(body('Waiting'))",
            "item/Pending": "@length(body('Pending'))", "item/Closed30": "@length(body('Closed_30'))",
            "item/Projects": "@length(body('Active_projects'))",
            "item/Refreshed": f"@{nowET('h:mm tt')}"})),
        *chart("status", "Open_all", "Status", ST7, 2),
        *chart("priority", "Open_all", "Priority", PR3, 3),
        *chart("category", "Created_30", "Category", CAT13, 4),
        *chart("pstatus", "Projects_list", "Status", PRJ_STATUSES, 11, title="ProjStatus"),
        *chart("ppriority", "Projects_list", "Priority", PRIORITIES, 12, title="ProjPriority"),
        *chart("pcategory", "Projects_list", "Category", PRJ_CATEGORIES, 13, title="ProjCategory"),
        ("Created_7", {"type": "Query", "inputs": {"from": "@body('All_tickets')?['value']",
            "where": "@greaterOrEquals(ticks(item()?['Created']),ticks(addDays(utcNow(),-7)))"}}),
        *trend(),
        *service(),
        *ranked("week", "Created_7", catkey, 8, "Week", "Busiest categories, last 7 days", 5, "Created_7"),
        *ranked("workload", "Open_working", "coalesce(item()?['AssignedTo']?['DisplayName'],'Unassigned')", 9, "Workload", "Open tickets by IT", 8, "Open_working"),
        *ranked("requesters", "Open_working", "coalesce(item()?['Requester']?['DisplayName'],'Unknown')", 10, "Requesters", "Open tickets by requester", 8, "Open_working"),
        *tech_sla(),
    ])
    trig = {"Recurrence": {"type": "Recurrence", "recurrence": {"frequency": "Week", "interval": 1, "timeZone": TZ,
            "schedule": {"weekDays": ALLDAYS, "hours": [str(h) for h in range(7, 21)], "minutes": list(range(0, 60, 5))}}}}
    return "HD15 Dashboard stats", trig, acts

DASH_TECHSLA = 15   # "TechSLA" row on the Live overview list, rendered by the "Chart - Tech SLA" view
DASH_PROJHEALTH = 16   # "IT Project Systems" row on the Live overview list, rendered by the "Chart - ProjHealth" view
def tech_sla():
    """Per-technician service levels over the last 30 days (tickets in Created_30, by assignee): ticket count, average
    first response (ResponseMinutes), % responded within 4h (ResponseMet) and % resolved within 24h (ResolvedMet).
    Rows are written in rank order - fastest average first response first, techs with no data last - so the
    dashboard shows who is in first place. Row i uses Slot(4i+1..4i+4) and Share(3i+1..3i+3)."""
    acts = []
    rows = []
    for i, (email, name) in enumerate(TECHS):
        p = f"T{i}"
        acts += [
            (f"{p}_tickets", {"type": "Query", "inputs": {"from": "@body('Created_30')",
                "where": f"@equals(toLower(coalesce(item()?['AssignedTo']?['Email'],'')),'{email.lower()}')"}}),
            (f"{p}_responded", {"type": "Query", "inputs": {"from": f"@body('{p}_tickets')", "where": "@not(equals(item()?['ResponseMinutes'],null))"}}),
            (f"{p}_minutes", {"type": "Select", "inputs": {"from": f"@body('{p}_responded')", "select": "@float(item()?['ResponseMinutes'])"}}),
            (f"{p}_met", {"type": "Query", "inputs": {"from": f"@body('{p}_responded')", "where": "@equals(float(coalesce(item()?['ResponseMet'],0)),1)"}}),
            (f"{p}_resolved", {"type": "Query", "inputs": {"from": f"@body('{p}_tickets')", "where": "@not(equals(item()?['ResolvedMet'],null))"}}),
            (f"{p}_resmet", {"type": "Query", "inputs": {"from": f"@body('{p}_resolved')", "where": "@equals(float(coalesce(item()?['ResolvedMet'],0)),1)"}}),
            # average minutes: sum via xpath over the JSON array (no loop needed)
            (f"{p}_avg", compose(f"@if(equals(length(body('{p}_minutes')),0),-1,div(float(xpath(xml(json(concat('{{\"r\":{{\"i\":',string(body('{p}_minutes')),'}}}}'))),'sum(/r/i)')),float(length(body('{p}_minutes')))))")),
        ]
        avg = f"outputs('{p}_avg')"
        nresp, nmet, nres, nresmet, ntick = (f"length(body('{p}_responded'))", f"length(body('{p}_met'))",
                                             f"length(body('{p}_resolved'))", f"length(body('{p}_resmet'))", f"length(body('{p}_tickets'))")
        # still-open tickets for this tech (any age), and those past a business day on the SLA clock
        acts += [
            (f"{p}_open", {"type": "Query", "inputs": {"from": "@body('All_tickets')?['value']",
                "where": f"@and(equals(toLower(coalesce(item()?['AssignedTo']?['Email'],'')),'{email.lower()}'),not(contains({arr([S_RES, S_CLOSED])},item()?['Status']?['Value'])))"}}),
            (f"{p}_overdue", {"type": "Query", "inputs": {"from": f"@body('{p}_open')",
                "where": f"@greaterOrEquals(float(coalesce(item()?['SLAHours'],0)),{DAY_H})"}}),
        ]
        nopen, nover = f"length(body('{p}_open'))", f"length(body('{p}_overdue'))"
        # Fairness score 0-100: 45% responded-within-4h rate, 25% resolved-within-24h rate (falls back to the response
        # rate when nothing has been resolved yet), 15% speed (100 at 0 min, 0 at 4 h or slower), 15% backlog
        # (100 with no overdue open tickets, -25 per open ticket older than a business day, floor 0).
        resp_pct = f"if(equals({nresp},0),0,div(mul(float({nmet}),100),float({nresp})))"
        res_pct = f"if(equals({nres},0),{resp_pct},div(mul(float({nresmet}),100),float({nres})))"
        speed = f"if(less({avg},0),0,if(greater({avg},240),0,sub(100,div({avg},2.4))))"
        backlog = f"if(greaterOrEquals({nover},4),0,sub(100,mul(25,{nover})))"
        score = f"add(add(add(mul(0.45,{resp_pct}),mul(0.25,{res_pct})),mul(0.15,{speed})),mul(0.15,{backlog}))"
        rows.append({
            "key": f"@float(if(equals({nresp},0),999999,sub(1000,{score})))",      # sort key (always float: sort() needs one type): best first, no data last
            "score": f"@if(equals({nresp},0),-1,int(formatNumber({score},'0')))",
            "openText": f"@if(equals({nopen},0),'none',concat(string({nopen}),' open',if(equals({nover},0),'',concat(' (',string({nover}),' overdue)'))))",
            "name": name,
            "tickets": f"@concat(string({ntick}),' ticket',if(equals({ntick},1),'','s'))",
            "avgText": (f"@if(less({avg},0),'-',if(less({avg},60),concat(string(int(formatNumber({avg},'0'))),' min'),"
                        f"concat(string(div(int(formatNumber({avg},'0')),60)),' h ',string(mod(int(formatNumber({avg},'0')),60)),' min')))"),
            "metText": f"@if(equals({nresp},0),'none yet',concat(string({nmet}),' of ',string({nresp})))",
            "resp": f"@if(equals({nresp},0),-1,int(formatNumber(div(mul(float({nmet}),100),float({nresp})),'0')))",
            "res": f"@if(equals({nres},0),-1,int(formatNumber(div(mul(float({nresmet}),100),float({nres})),'0')))",
            "avg": f"@{avg}"})
    acts.append(("Tech_rows", compose(rows)))
    acts.append(("Tech_ranked", compose("@sort(outputs('Tech_rows'),'key')")))
    item = {"id": DASH_TECHSLA, "item/Title": "TechSLA", "item/Sub": "Service levels by technician, last 30 days",
            "item/Refreshed": f"@{nowET('h:mm tt')}"}
    for i in range(len(TECHS)):
        r = f"outputs('Tech_ranked')?[{i}]"
        item[f"item/Slot{4*i+1}"] = f"@{r}?['name']"
        item[f"item/Slot{4*i+2}"] = f"@{r}?['tickets']"
        item[f"item/Slot{4*i+3}"] = f"@{r}?['avgText']"
        item[f"item/Slot{4*i+4}"] = f"@{r}?['metText']"
        item[f"item/Share{3*i+1}"] = f"@{r}?['resp']"
        item[f"item/Share{3*i+2}"] = f"@{r}?['res']"
        item[f"item/Share{3*i+3}"] = f"@{r}?['score']"      # -1 = no data (hides the rank badge)
        item[f"item/Slot{13+i}"] = f"@{r}?['openText']"     # still-open tickets, e.g. "2 open (1 overdue)"
    acts.append(("Update_tech_sla", sp("PatchItem", {"table": DASH, **item})))
    return acts

# ---------- HD16 Health check v2 ----------
# Expected live flows (display name as listed by Power Automate Management). HD13/HD17 are run-once (off);
# HD14 is owned by Martin's account so "List My Flows" (as IT_Support) does not return it.
EXPECTED_FLOWS = ["HD1 Email intake", "HD2 Form intake", "HD3 Ticket created", "HD4 Ticket updated",
                  "HD5 Planner task completed", "HD6 Hourly clock", "HD7 My open tickets", "HD8 Daily report",
                  "HD9 Card actions", "HD11 Tech thread replies", "HD12 Requester replies in Teams",
                  "HD15 Dashboard stats", "HD16 Health check v2", "HD18 Quick close (SharePoint)"]
# IT Projects flows, checked separately from EXPECTED_FLOWS above - they get their own dashboard health bucket
# (DASH_PROJHEALTH, "IT Project Systems") rather than folding into the aggregate "Flows" card/count.
EXPECTED_PROJECT_FLOWS = ["HD19 Project created", "HD20 Project updated", "HD21 No-update check"]
PROJECT_FLOW_LABELS = {"HD19 Project created": "New project emails (HD19)",
                        "HD20 Project updated": "Project update emails (HD20)",
                        "HD21 No-update check": "Weekly check-in (HD21)"}
PROBE_LABELS = {"Probe_Outlook": "Outlook", "Probe_Teams_IT": "Teams (IT Support account)", "Probe_Teams": "Teams",
                "Probe_Planner": "Planner", "Probe_Users": "Office 365 Users", "Probe_OneDrive": "OneDrive",
                "Probe_Flows": "Power Automate Management"}

# Flow-level "Description" (Power Automate flow Details > Edit). SharePoint's button-flow "Run flow" confirm
# panel shows this under the owner line, so it's the only place to put on-panel instructions for a flow with
# no adaptive-card UI of its own (see HD18). Keyed by flow display name so a re-import with Update keeps it
# instead of silently dropping back to blank.
FLOW_DESCRIPTIONS = {
    "HD18 Quick close (SharePoint)": "Type the resolution below, then click Run flow to assign this ticket to "
                                      "you, close it, and email the requester.",
}

def hd16():
    """Runs hourly, 24/7. Seven cards on the IT dashboard "System health" panel, a one-line-per-problem band under
    them, an ISO timestamp the dashboard uses to grey the panel out when this flow stops, a daily 9:30 email and a
    Teams message whenever the issue count changes (alerts only 7am-9pm; overnight problems wait on the panel).
      1 Connections   read-only probe through every connector (Forms has no read action, so it is not probed)
      2 Flows         every expected flow is on exactly once (Power Automate Management, List My Flows)
      3 New tickets   tickets created in the last 24h that HD3 did not finish
      4 Updates       tickets whose status changed >30 min ago but the requester was never told (HD4 dead)
      5 Intake        mail sitting in the Inbox >20 min (HD1 dead; red) / mail parked in Review in 24h (amber)
      6 Clock         HD6 heartbeat (nothing touched in 3h) and resolved tickets not auto-closed after 72h
      7 Dashboard     DashStats row refreshed in the last 30 min (HD15 dead), checked 7am-9pm only"""
    it = lambda k: f"item()?['{k}']"
    li = lambda body: f"concat('<li>',{body},'</li>')"
    working = " or ".join(f"Status eq '{s}'" for s in WORKING)
    ANY = ("Succeeded", "Failed", "Skipped", "TimedOut")
    prev = "body('Previous')"
    prev_proj = "body('Previous_proj')"
    NL = "decodeUriComponent('%0A')"
    hourET = f"int({nowET('HH')})"
    office = f"and(greaterOrEquals({hourET},7),less({hourET},21))"

    # --- 1. connection probes: all read-only, run in parallel inside one scope ---
    probes = {
        "Probe_Outlook": op("shared_office365", "HttpRequest", {"Method": "GET",
            "Uri": f"https://graph.microsoft.com/v1.0/users/{MBX}/mailFolders/inbox?$select=id"}),
        "Probe_Teams": op("shared_teams", "GetChannelsForGroup", {"groupId": GROUP}),
        "Probe_Teams_IT": op("shared_teams_1", "GetAllTeams", {}),
        "Probe_Planner": op("shared_planner", "ListBuckets_V3", {"groupId": GROUP, "id": PLAN}),
        "Probe_Users": op("shared_office365users", "UserProfile_V2", {"id": ME}),
        "Probe_OneDrive": op("shared_onedriveforbusiness", "ListRootFolder", {}),
        "Probe_Flows": op("shared_flowmanagement", "ListMyFlows", {"environmentName": "@workflow()?['tags']?['environmentName']"}),
    }
    label_expr = "item()?['name']"
    for k, v in PROBE_LABELS.items():
        label_expr = f"replace({label_expr},'{k}','{v}')"

    # --- 2. flows: "<displayName>=<state>;..." then count how many times each expected flow is Started ---
    flow_terms = []
    for name in EXPECTED_FLOWS:
        c = f"sub(length(split(outputs('FlowStr'),'{name}=Started')),1)"
        flow_terms.append(f"if(equals({c},1),'',concat('{name}',if(equals({c},0),' is off or missing; ',concat(' has ',string({c}),' copies on; '))))")
    flow_problems = "concat(" + ",".join(flow_terms) + ")"

    # --- 2b. project flows: same "<name>=Started" check against EXPECTED_PROJECT_FLOWS, own dashboard bucket ---
    proj_counts = [f"sub(length(split(outputs('FlowStr'),'{name}=Started')),1)" for name in EXPECTED_PROJECT_FLOWS]
    proj_flow_fields = {}
    for i, (name, c) in enumerate(zip(EXPECTED_PROJECT_FLOWS, proj_counts)):
        proj_flow_fields[f"item/Slot{i+1}"] = PROJECT_FLOW_LABELS[name]
        proj_flow_fields[f"item/Share{i+1}"] = f"@if(not(outputs('FlowsOk')),-1,if(equals({c},1),0,100))"
        proj_flow_fields[f"item/Slot{i+4}"] = (f"@if(not(outputs('FlowsOk')),'could not list flows',"
            f"if(equals({c},1),'on',if(equals({c},0),'off or missing',concat(string({c}),' copies on'))))")
    proj_bad_terms = [f"if(and(outputs('FlowsOk'),not(equals({c},1))),1,0)" for c in proj_counts]
    proj_issue_count = proj_bad_terms[0]
    for t in proj_bad_terms[1:]:
        proj_issue_count = f"add({proj_issue_count},{t})"
    # a HD19/HD20 self_check failure (stamped on DASH_PROJHEALTH's own Slot10/Slot11 - see self_check's
    # dash_id/dash_title) shown for 24h, same convention as NotifyRecent/Slot10 does for the ticket side below.
    proj_issue_count = f"add({proj_issue_count},if(outputs('ProjNotifyRecent'),1,0))"
    proj_problem_terms = [f"if(and(outputs('FlowsOk'),not(equals({c},1))),'{PROJECT_FLOW_LABELS[name]}; ','')"
                           for name, c in zip(EXPECTED_PROJECT_FLOWS, proj_counts)]
    proj_problem_terms.append(f"if(outputs('ProjNotifyRecent'),concat('Notifications (HD19/HD20 self-checks): ',{prev_proj}?['Slot10'],'; '),'')")

    acts = seq([
        ("Previous", sp("GetItem", {"table": DASH, "id": DASH_HEALTH})),
        ("Previous_proj", {**sp("GetItem", {"table": DASH, "id": DASH_PROJHEALTH}), "runAfter": ra("Previous", status=ANY)}),
        # merges the Previous_proj branch back into the main chain right away, so everything after this point
        # (ProjNotifyRecent included, further down, auto-chained with no runAfter of its own) is guaranteed to run
        # after Previous_proj has completed, not just after Previous.
        ("Probes", {"type": "Scope", "runAfter": {**ra("Previous", status=ANY), **ra("Previous_proj", status=ANY)}, "actions": probes}),
        ("Probe_failed", {"type": "Query", "runAfter": ra("Probes", status=ANY), "inputs": {"from": "@result('Probes')",
            "where": "@not(equals(item()?['status'],'Succeeded'))"}}),
        ("Probe_names", {"type": "Select", "inputs": {"from": "@body('Probe_failed')", "select": "@" + label_expr}}),
        ("ConnBad", compose("@greater(length(body('Probe_failed')),0)")),
        ("FlowsOk", compose("@equals(actions('Probe_Flows')?['status'],'Succeeded')")),
        ("FlowList", {"type": "Select", "inputs": {"from": "@if(outputs('FlowsOk'),body('Probe_Flows')?['value'],json('[]'))",
            "select": "@concat(item()?['properties']?['displayName'],'=',item()?['properties']?['state'])"}}),
        ("FlowStr", compose("@join(body('FlowList'),';')")),
        ("FlowProblems", compose("@if(outputs('FlowsOk'),trim(" + flow_problems + "),'')")),
        ("FlowBad", compose("@and(outputs('FlowsOk'),greater(length(outputs('FlowProblems')),0))")),

        # --- 3. new tickets set up ---
        ("Recent", sp("GetItems", {"$filter": "Created ge '@{addHours(utcNow(),-24)}' and Created lt '@{addMinutes(utcNow(),-20)}'", "$top": 500})),
        ("Incomplete", {"type": "Query", "inputs": {"from": "@body('Recent')?['value']", "where":
            f"@or(empty(coalesce({it('PlannerTaskId')},'')),empty(coalesce({it('ChannelMessageId')},'')),not(contains(coalesce({it('Conversation')},''),'Ticket created')))"}}),
        ("Incomplete_lines", {"type": "Select", "inputs": {"from": "@body('Incomplete')", "select": "@" + li(
            f"'IT-',{it('ID')},' — ',{it('Title')},' (',"
            f"if(empty(coalesce({it('PlannerTaskId')},'')),'no Planner task; ',''),"
            f"if(empty(coalesce({it('ChannelMessageId')},'')),'no channel post; ',''),"
            f"if(not(contains(coalesce({it('Conversation')},''),'Ticket created')),'HD3 did not finish; ',''),')'")}}),
        ("Incomplete_ids", {"type": "Select", "inputs": {"from": "@body('Incomplete')", "select": f"@concat('IT-',{it('ID')})"}}),

        # --- 4. updates notified (HD4 alive): status differs from LastNotifiedStatus for more than 30 min ---
        ("Changed", sp("GetItems", {"$filter": "Modified ge '@{addHours(utcNow(),-48)}' and Modified lt '@{addMinutes(utcNow(),-30)}'", "$top": 500})),
        ("Unnotified", {"type": "Query", "inputs": {"from": "@body('Changed')?['value']", "where":
            f"@not(equals(coalesce({it('Status')}?['Value'],'Open'),coalesce({it('LastNotifiedStatus')},'Open')))"}}),
        ("Unnotified_ids", {"type": "Select", "inputs": {"from": "@body('Unnotified')", "select": f"@concat('IT-',{it('ID')})"}}),

        # --- 5. intake: Inbox mail HD1 has not picked up; Review folder for attention ---
        ("Inbox_old", op("shared_office365", "HttpRequest", {"Method": "GET",
            "Uri": f"https://graph.microsoft.com/v1.0/users/{MBX}/mailFolders/inbox/messages?$filter=receivedDateTime le @{{addMinutes(utcNow(),-20)}}&$select=subject,from,receivedDateTime&$top=50"}),),
        ("InboxOk", {"type": "Compose", "runAfter": ra("Inbox_old", status=ANY), "inputs": "@equals(actions('Inbox_old')?['status'],'Succeeded')"}),
        # ignore Microsoft/Power Automate system mail (connection alerts, digests) - HD1 parks those, not tickets
        ("Inbox_real", {"type": "Query", "runAfter": ra("InboxOk", status=("Succeeded", "Failed", "Skipped", "TimedOut")), "inputs": {"from": "@if(outputs('InboxOk'),coalesce(body('Inbox_old')?['value'],json('[]')),json('[]'))", "where":
            f"@and(not(equals(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'{MBX.lower()}')),"
            f"not(endsWith(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'.messaging.microsoft.com')),"
            f"not(endsWith(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'@microsoft.com')),"
            f"not(endsWith(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'@powerautomate.com')))"}}),
        ("InboxCount", compose("@if(outputs('InboxOk'),length(body('Inbox_real')),-1)")),
        ("Review_folder", {"type": "Scope", "runAfter": ra("InboxCount", status=ANY), "actions": seq([
            ("Find_folder", op("shared_office365", "HttpRequest", {"Method": "GET",
                "Uri": f"https://graph.microsoft.com/v1.0/users/{MBX}/mailFolders?$filter=displayName eq 'Review'&$select=id"})),
            ("Find_subfolder", op("shared_office365", "HttpRequest", {"Method": "GET",
                "Uri": f"https://graph.microsoft.com/v1.0/users/{MBX}/mailFolders/inbox/childFolders?$filter=displayName eq 'Review'&$select=id"})),
            ("ReviewId", compose("@first(union(coalesce(body('Find_folder')?['value'],json('[]')),coalesce(body('Find_subfolder')?['value'],json('[]'))))?['id']")),
            ("Review_mail", op("shared_office365", "HttpRequest", {"Method": "GET",
                "Uri": f"https://graph.microsoft.com/v1.0/users/{MBX}/mailFolders/@{{outputs('ReviewId')}}/messages?$filter=receivedDateTime ge @{{addHours(utcNow(),-24)}}&$select=subject,from,receivedDateTime&$top=50&$orderby=receivedDateTime desc"})),
            ("Review_real", {"type": "Query", "inputs": {"from": "@body('Review_mail')?['value']", "where":
                f"@and(not(equals(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'{MBX.lower()}')),"
                f"not(endsWith(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'.messaging.microsoft.com')),"
                f"not(endsWith(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'@microsoft.com')),"
                f"not(endsWith(toLower(coalesce({it('from')}?['emailAddress']?['address'],'')),'@powerautomate.com')))"}}),
            ("Review_lines", {"type": "Select", "inputs": {"from": "@body('Review_real')",
                "select": "@" + li(f"coalesce({it('from')}?['emailAddress']?['address'],'?'),': ',coalesce({it('subject')},'(no subject)')")}}),
        ])}),
        ("ReviewOk", {"type": "Compose", "runAfter": ra("Review_folder", status=ANY), "inputs": "@equals(actions('Review_folder')?['status'],'Succeeded')"}),
        ("ReviewCount", compose("@if(outputs('ReviewOk'),length(body('Review_lines')),-1)")),
        ("ReviewNote", compose(
            "@if(outputs('ReviewOk'),"
            "if(greater(outputs('ReviewCount'),0),concat('<ul>',join(body('Review_lines'),''),'</ul>'),'<p>None.</p>'),"
            "'<p>Could not read the Review folder (check the HD16 run).</p>')")),
        ("IntakeState", compose("@if(greater(outputs('InboxCount'),0),100,if(not(outputs('InboxOk')),-1,if(greater(outputs('ReviewCount'),0),50,0)))")),

        # --- 6. clock and auto-close (HD6) ---
        ("Working", sp("GetItems", {"$filter": f"({working}) and Created lt '@{{addMinutes(utcNow(),-90)}}'", "$top": 500})),
        ("Touched", {"type": "Query", "inputs": {"from": "@body('Working')?['value']",
            "where": f"@greater(ticks({it('Modified')}),ticks(addHours(utcNow(),-3)))"}}),
        # HD6 runs 8:00-21:00; outside that, "nothing touched in 3 h" is just the clock being off for the night
        ("ClockStalled", compose(f"@and(greaterOrEquals({hourET},11),less({hourET},22),greater(length(body('Working')?['value']),0),equals(length(body('Touched')),0))")),
        ("StuckResolved", sp("GetItems", {"$filter": f"Status eq '{S_RES}' and ResolvedOn lt '@{{addHours(utcNow(),-72)}}'", "$top": 100})),
        ("StuckResolved_lines", {"type": "Select", "inputs": {"from": "@body('StuckResolved')?['value']",
            "select": "@" + li(f"'IT-',{it('ID')},' — ',{it('Title')},' (resolved ',formatDateTime({it('ResolvedOn')},'yyyy-MM-dd'),', still not closed)'")}}),
        ("StuckResolved_ids", {"type": "Select", "inputs": {"from": "@body('StuckResolved')?['value']", "select": f"@concat('IT-',{it('ID')})"}}),
        ("ClockBad", compose("@or(outputs('ClockStalled'),greater(length(body('StuckResolved')?['value']),0))")),

        # --- 7. dashboard data (HD15 alive): newest Modified on the stats rows, business hours only ---
        ("Dash_rows", sp("GetItems", {"table": DASH, "$filter": f"ID ne {DASH_HEALTH}", "$orderby": "Modified desc", "$top": 1})),
        ("DashStale", compose(f"@and({office},greater(length(body('Dash_rows')?['value']),0),less(ticks(first(body('Dash_rows')?['value'])?['Modified']),ticks(addMinutes(utcNow(),-30))))")),

        # notification problems flagged by HD3/HD4 (they stamp Slot11 with the UTC time): shown for 24h
        ("NotifyRecent", compose(f"@greater(ticks(if(empty(coalesce({prev}?['Slot11'],'')),'2000-01-01T00:00:00Z',{prev}?['Slot11'])),ticks(addHours(utcNow(),-24)))")),
        # same thing for HD19/HD20 self-check failures, off DASH_PROJHEALTH's own Slot10/Slot11 - never DASH_HEALTH's.
        # No explicit runAfter: Previous_proj ran in parallel with Probes right at the start, long before this point
        # in the chain, so letting seq() auto-chain this straight after NotifyRecent keeps the single-chain backbone
        # (and everything IssueCount transitively needs) intact instead of detouring IssueCount's dependency off it.
        ("ProjNotifyRecent", compose(f"@greater(ticks(if(empty(coalesce({prev_proj}?['Slot11'],'')),'2000-01-01T00:00:00Z',{prev_proj}?['Slot11'])),ticks(addHours(utcNow(),-24)))")),

        ("IssueCount", compose(
            "@add(add(add(add(add(add(add("
            "if(outputs('ConnBad'),1,0),if(outputs('FlowBad'),1,0)),"
            "if(greater(length(body('Incomplete')),0),1,0)),"
            "if(greater(length(body('Unnotified')),0),1,0)),"
            "if(greater(outputs('InboxCount'),0),1,0)),"
            "if(outputs('ClockBad'),1,0)),"
            "if(outputs('DashStale'),1,0)),"
            "if(outputs('NotifyRecent'),1,0))")),
        ("Details", compose(
            "@trim(concat("
            f"if(outputs('ConnBad'),concat('Connections broken: ',join(body('Probe_names'),', '),' — Power Automate › Connections › fix or recreate (see Tech Guide).',{NL}),''),"
            f"if(outputs('FlowBad'),concat('Flows: ',outputs('FlowProblems'),{NL}),''),"
            f"if(greater(length(body('Incomplete')),0),concat('Incomplete setup (no Planner task / channel post): ',join(body('Incomplete_ids'),', '),' — check the HD3 runs.',{NL}),''),"
            f"if(greater(length(body('Unnotified')),0),concat('Status changed but requester not told: ',join(body('Unnotified_ids'),', '),' — check HD4.',{NL}),''),"
            f"if(greater(outputs('InboxCount'),0),concat(string(outputs('InboxCount')),' email(s) waiting in the IT Support inbox for over 20 min — HD1 is not ticketing them.',{NL}),''),"
            f"if(outputs('ClockStalled'),concat('Hourly clock: no open ticket updated in 3 hours — check HD6.',{NL}),''),"
            f"if(greater(length(body('StuckResolved')?['value']),0),concat('Not auto-closed: ',join(body('StuckResolved_ids'),', '),{NL}),''),"
            f"if(outputs('DashStale'),concat('Dashboard charts not refreshed for 30 min — check HD15.',{NL}),''),"
            f"if(outputs('NotifyRecent'),concat({prev}?['Slot10'],{NL}),'')))")),
        ("Update_dashboard", sp("PatchItem", {"table": DASH, "id": DASH_HEALTH, "item/Title": "Health", "item/Sub": "IT Ticket Systems Health",
            "item/Total": "@outputs('IssueCount')", "item/Refreshed": f"@{nowET('h:mm tt')}",
            "item/Slot22": "@string(div(sub(ticks(utcNow()),621355968000000000),10000))",   # epoch ms, for the stale banner
            "item/Slot1": "Connections", "item/Share1": "@if(outputs('ConnBad'),100,0)",
            "item/Slot15": "@if(outputs('ConnBad'),concat('broken: ',join(body('Probe_names'),', ')),'Outlook, Teams, Planner, OneDrive, Users, SharePoint all answering')",
            "item/Slot2": "Flows", "item/Share2": "@if(not(outputs('FlowsOk')),-1,if(outputs('FlowBad'),100,0))",
            "item/Slot16": f"@if(not(outputs('FlowsOk')),'could not list flows',if(outputs('FlowBad'),outputs('FlowProblems'),'{len(EXPECTED_FLOWS)} flows on, one copy each'))",
            "item/Slot3": "New tickets set up", "item/Share3": "@if(greater(length(body('Incomplete')),0),100,0)",
            "item/Slot17": "@concat(length(body('Recent')?['value']),' created in 24h, ',length(body('Incomplete')),' incomplete')",
            "item/Slot4": "Updates notified (HD4)", "item/Share4": "@if(greater(length(body('Unnotified')),0),100,0)",
            "item/Slot18": "@if(greater(length(body('Unnotified')),0),concat(length(body('Unnotified')),' ticket(s) changed, requester not told'),'every status change was sent')",
            "item/Slot5": "Email intake (HD1)", "item/Share5": "@outputs('IntakeState')",
            "item/Slot19": "@if(greater(outputs('InboxCount'),0),concat(outputs('InboxCount'),' email(s) stuck in the inbox'),if(not(outputs('InboxOk')),'could not read the inbox',if(greater(outputs('ReviewCount'),0),concat(outputs('ReviewCount'),' in the Review folder, last 24 h'),'inbox clear; nothing in Review')))",
            "item/Slot6": "Clock & auto-close (HD6)", "item/Share6": "@if(outputs('ClockBad'),100,0)",
            "item/Slot20": f"@if(outputs('ClockStalled'),'no open ticket updated in 3 h',if(greater(length(body('StuckResolved')?['value']),0),concat(length(body('StuckResolved')?['value']),' resolved ticket(s) past 72 h not closed'),if(less({hourET},8),'clock starts at 8:00',concat(length(body('Working')?['value']),' open tickets being tracked'))))",
            "item/Slot7": "Dashboard data (HD15)", "item/Share7": "@if(outputs('DashStale'),100,0)",
            "item/Slot21": "@if(outputs('DashStale'),'charts not refreshed for 30 min',if(greater(length(body('Dash_rows')?['value']),0),concat('refreshed ',formatDateTime(convertTimeZone(first(body('Dash_rows')?['value'])?['Modified'],'UTC','" + TZ + "'),'h:mm tt')),'no data yet'))",
            "item/Slot10": f"@if(empty(coalesce({prev}?['Slot10'],'')),'no problems reported',{prev}?['Slot10'])",
            "item/Grad": "@outputs('Details')"})),
        ("Update_project_health", sp("PatchItem", {"table": DASH, "id": DASH_PROJHEALTH, "item/Title": "ProjectHealth",
            "item/Sub": "IT Project Systems", "item/Refreshed": f"@{nowET('h:mm tt')}",
            "item/Total": f"@if(outputs('FlowsOk'),{proj_issue_count},-1)",
            "item/Grad": "@if(not(outputs('FlowsOk')),'could not list flows',trim(concat(" +
                ",".join(proj_problem_terms) + ")))",
            # echoes self_check's own HD19/HD20 stamp back for the next run to read, same convention as the
            # ticket row's item/Slot10 below - Slot11 is never touched here, only self_check writes it.
            "item/Slot10": f"@if(empty(coalesce({prev_proj}?['Slot10'],'')),'no problems reported',{prev_proj}?['Slot10'])",
            **proj_flow_fields})),
        ("Report", compose(
            "@concat("
            "'<h3>Help desk health — ',if(equals(outputs('IssueCount'),0),'all clear','PROBLEMS FOUND'),'</h3>',"
            "'<p><b>Connections:</b> ',if(outputs('ConnBad'),concat('<span style=\"color:#b00\">broken: ',join(body('Probe_names'),', '),'</span>'),'all answering'),'</p>',"
            "'<p><b>Flows:</b> ',if(not(outputs('FlowsOk')),'could not list flows',if(outputs('FlowBad'),concat('<span style=\"color:#b00\">',outputs('FlowProblems'),'</span>'),'all on, one copy each')),'</p>',"
            "'<p><b>New tickets, last 24h:</b> ',length(body('Recent')?['value']),' created; ',length(body('Incomplete')),' with incomplete setup.</p>',"
            "if(greater(length(body('Incomplete')),0),concat('<ul>',join(body('Incomplete_lines'),''),'</ul>'),''),"
            "'<p><b>Updates notified (HD4):</b> ',if(greater(length(body('Unnotified')),0),concat('<span style=\"color:#b00\">not sent for ',join(body('Unnotified_ids'),', '),'</span>'),'OK'),'</p>',"
            "'<p><b>Email intake (HD1):</b> ',if(greater(outputs('InboxCount'),0),concat('<span style=\"color:#b00\">',string(outputs('InboxCount')),' email(s) waiting in the inbox</span>'),if(outputs('InboxOk'),'inbox clear','could not read the inbox')),'</p>',"
            "'<p><b>Hourly clock (HD6):</b> ',if(outputs('ClockStalled'),'<span style=\"color:#b00\">NOT RUNNING — no open ticket has been updated in 3 hours</span>',"
            "concat('OK (',length(body('Working')?['value']),' open tickets being tracked)')),'</p>',"
            "'<p><b>Auto-close (HD6):</b> ',if(greater(length(body('StuckResolved')?['value']),0),"
            "concat('<span style=\"color:#b00\">',length(body('StuckResolved')?['value']),' resolved ticket(s) past 72h still not closed</span><ul>',join(body('StuckResolved_lines'),''),'</ul>'),'OK'),'</p>',"
            "'<p><b>Dashboard data (HD15):</b> ',if(outputs('DashStale'),'<span style=\"color:#b00\">stale for 30 min</span>','OK'),'</p>',"
            f"'<p><b>Notifications (HD3/HD4 self-checks):</b> ',if(outputs('NotifyRecent'),concat('<span style=\"color:#b00\">',{prev}?['Slot10'],'</span>'),'no problems reported in the last 24h'),'</p>',"
            "'<p><b>Emails parked in Review (last 24h, excluding system mail):</b></p>',outputs('ReviewNote'),"
            f"'<p><a href=\"{SITE}/SitePages/IT-Dashboard.aspx\">Open the IT dashboard</a></p>')")),
        ("Daily_email", cond(f"@equals({hourET},9)", {
            "Email_report": mail(ME, "@{concat('[Help desk health] ',if(equals(outputs('IssueCount'),0),'All clear',concat('PROBLEMS: ',string(outputs('IssueCount')))))}",
                                 "@{outputs('Report')}")})),
        # alert Martin (email + Teams) only when the issue count changes, 7am-9pm; overnight problems wait for 7:30
        ("Alert_if_changed", {**cond(f"@and({office},greater(outputs('IssueCount'),0),not(equals(outputs('IssueCount'),int(coalesce({prev}?['Total'],0)))))",
            seq([("Email_change", mail(ME, "@{concat('[Help desk health] PROBLEMS: ',string(outputs('IssueCount')))}", "@{outputs('Report')}")),
                 ("Teams_change", {"type": "Scope", "runAfter": ra("Email_change", status=("Succeeded", "Failed")),
                                   "actions": support_msg("Health_dm", ME, "@{outputs('Report')}")})])),
            "runAfter": ra("Daily_email", status=ANY)}),
    ])
    acts["Inbox_old"]["runAfter"] = ra("Unnotified_ids", status=ANY)
    trig = {"Recurrence": {"type": "Recurrence", "recurrence": {"frequency": "Week", "interval": 1, "timeZone": TZ,
            "schedule": {"weekDays": ALLDAYS, "hours": [str(h) for h in range(0, 24)], "minutes": [30]}}}}
    return "HD16 Health check v2", trig, acts

# ---------- packaging ----------
def used_apis(obj, acc):
    if isinstance(obj, str):
        import re as _r
        for m in _r.findall(r"\$connections'\)\['(shared_[a-z0-9]+)'\]", obj):
            acc.add(m)
        return
    if isinstance(obj, dict):
        h = obj.get("host")
        if isinstance(h, dict) and "connectionName" in h:
            acc.add(h["connectionName"])
        for v in obj.values():
            used_apis(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            used_apis(v, acc)

def build_definition(name, triggers, actions):
    return {"$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
            "contentVersion": "1.0.0.0",
            "parameters": {"$authentication": {"defaultValue": {}, "type": "SecureObject"},
                           "$connections": {"defaultValue": {}, "type": "Object"}},
            "triggers": triggers, "actions": actions, "outputs": {}}

# ---------- HD17 Fix ticket permissions (run once) ----------
def hd17():
    """One-off: make the Tickets list stop inheriting from the site (drop Visitors/Members), then give every
    existing ticket its own permissions exactly as HD3 does for new ones (Owners full control, requester Read)."""
    L = f"_api/web/lists(guid'{TICKETS}')"
    trig = {"manual": {"type": "Request", "kind": "Button", "inputs": {"schema": {"type": "object", "properties": {}, "required": []}}}}
    it = lambda k: f"items('Each_ticket')?['{k}']"
    per_item = seq([
        ("ReqEmail", compose(f"@toLower(coalesce({it('RequesterEmail')},{it('Requester')}?['EMail'],''))")),
        ("Break_item", sp_http("POST", f"{L}/items(@{{{it('Id')}}})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")),
        ("Owners_on_item", sp_http("POST", f"{L}/items(@{{{it('Id')}}})/roleassignments/addroleassignment(principalid=@{{body('Get_owner_group')?['Id']}},roledefid=1073741829)")),
        ("Has_requester", cond("@and(not(empty(outputs('ReqEmail')))," + is_internal("outputs('ReqEmail')") + ")", seq([
            ("Ensure_requester", sp_http("POST", "_api/web/ensureuser", body="{\"logonName\":\"i:0#.f|membership|@{outputs('ReqEmail')}\"}")),
            ("Requester_read", sp_http("POST", f"{L}/items(@{{{it('Id')}}})/roleassignments/addroleassignment(principalid=@{{body('Ensure_requester')?['Id']}},roledefid=1073741826)")),
        ]))),
    ])
    acts = seq([
        ("Get_owner_group", sp_http("GET", "_api/web/associatedownergroup?$select=Id")),
        ("Get_visitor_group", sp_http("GET", "_api/web/associatedvisitorgroup?$select=Id")),
        ("Get_member_group", sp_http("GET", "_api/web/associatedmembergroup?$select=Id")),
        ("Break_list", sp_http("POST", f"{L}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)")),
        ("Remove_visitors", sp_http("POST", f"{L}/roleassignments/removeroleassignment(principalid=@{{body('Get_visitor_group')?['Id']}})")),
        ("Remove_members", sp_http("POST", f"{L}/roleassignments/removeroleassignment(principalid=@{{body('Get_member_group')?['Id']}})")),
        ("Get_tickets", sp_http("GET", f"{L}/items?$select=Id,RequesterEmail,Requester/EMail&$expand=Requester&$top=5000")),
        ("Each_ticket", {"type": "Foreach", "foreach": "@body('Get_tickets')?['value']", "actions": per_item,
                         "runtimeConfiguration": {"concurrency": {"repetitions": 1}}}),
    ])
    # removing a group that is already gone must not stop the run
    acts["Remove_members"]["runAfter"] = ra("Remove_visitors", status=("Succeeded", "Failed"))
    acts["Get_tickets"]["runAfter"] = ra("Remove_members", status=("Succeeded", "Failed"))
    return "HD17 Fix ticket permissions (run once)", trig, acts

# ---------- HD18 Quick close (SharePoint) ----------
def hd18():
    """The SharePoint twin of the Teams card's Quick close. A button in the Tickets list's ID column runs this
    flow ("For a selected item"); SharePoint asks for the resolution, then the ticket is assigned to whoever
    clicked (unless already assigned), set to Closed and given the resolution in one update, so HD4 sends the
    requester a single "Resolved" message exactly as it does after the Teams Quick close.
    The trigger JSON is what Power Automate itself generates for the SharePoint "For a selected item" trigger
    (operation GetItemHybridTriggerSchema); the clicker's email arrives base64-encoded in a trigger header.
    The schema must be wrapped exactly as Power Automate's own designer canonicalizes it — a top-level
    {"type":"object","properties":{"rows":{...}},"required":["rows"]}, with an "entity" (the row's ID) sitting
    alongside "text" under items.properties — or the SharePoint "Run flow" panel renders no input box at all for
    a button-triggered flow, even though the flow itself is otherwise valid. Confirmed by touching the trigger's
    description in the designer and re-saving: Power Automate rewrote the flatter schema into this shape."""
    trig = {"For_a_selected_item": {"type": "Request", "kind": "ApiConnection", "inputs": {
        "operationId": "GetItemHybridTriggerSchema",
        "parameters": {"dataset": SITE, "table": TICKETS},
        "schema": {"type": "object", "properties": {"rows": {"type": "array", "items": {"type": "object", "properties": {
            "text": {"description": "What fixed it (the requester sees this)", "title": "Resolution", "type": "string",
                     "x-ms-content-hint": "TEXT", "x-ms-dynamically-added": True},
            "entity": {"properties": {"ID": {"title": "ID", "type": "integer", "format": "int64"}},
                       "type": "object", "required": []}}, "required": ["text"]}}}, "required": ["rows"]},
        "host": {"connection": {"name": "@parameters('$connections')['shared_sharepointonline']['connectionId']"}}},
        "splitOn": "@triggerBody()['rows']"}}
    acts = seq([
        ("Clicker", compose("@toLower(base64ToString(triggerOutputs()?['headers']?['x-ms-user-email-encoded']))")),
        ("TicketId", compose("@triggerBody()?['entity']?['ID']")),
        ("QuickText", compose("@trim(coalesce(triggerBody()?['text'],''))")),
        ("Get_item", sp("GetItem", {"id": "@outputs('TicketId')"})),
        ("Quick_close", cond("@empty(outputs('QuickText'))", {
            "Say_need_resolution_quick": scope_msg("Need_res_quick_msg", "@outputs('Clicker')",
                "<p>IT-@{outputs('TicketId')} wasn't closed — Quick close needs the resolution typed in the box.</p>")},
            {"Close_now": sp("PatchItem", {
                "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                "item/Status/Value": S_CLOSED,
                "item/AssignedTo/Claims": "@coalesce(body('Get_item')?['AssignedTo']?['Email'],outputs('Clicker'))",
                "item/Resolution": "@outputs('QuickText')"})})),
    ])
    return "HD18 Quick close (SharePoint)", trig, acts

# ---------- HD19 Project created ----------
def hd19():
    """Fires when a project is added to the IT Projects list (entered directly in SharePoint by Martin or a
    tech). Breaks the new item's permission inheritance exactly as HD3 does for tickets - Owners full control,
    then Read for every Project Sponsor on just that one item, nothing else on the list - then emails +
    Teams-messages the creator, Martin and every sponsor. Owner and Project Sponsor both allow multiple
    people (SharePoint connector exposes a multi-value Person field as an array of {Email,DisplayName,...}
    items, same keys as the old single-value field, just array-wrapped); ProjectSponsor also keeps a
    plain-text SponsorEmail fallback column from before it supported multiple people, used only when no
    Project Sponsor person is set. Stamps LastNotifiedStatus/LastNotifiedSponsorLen/LastNotifiedInternalLen
    so HD20 has a clean baseline to compare against on the next edit."""
    pt = lambda k: f"triggerOutputs()?['body/{k}']"
    ID = pt('ID')
    tell_cond = lambda who: ("@and(not(empty(outputs('" + who + "'))),not(equals(outputs('" + who + "'),'" + ME
                              + "')),") + is_internal("outputs('" + who + "')") + ")"
    sponsor_email_fallback = pt('SponsorEmail')
    actions = seq([
        ("CreatorEmail", compose(f"@toLower(coalesce({pt('Author/Email')},''))")),
        ("Sponsor_emails_select", people_select("Sponsor_emails_select", pt('ProjectSponsor'), "Email")["Sponsor_emails_select"]),
        ("Sponsor_names_select", people_select("Sponsor_names_select", pt('ProjectSponsor'), "DisplayName")["Sponsor_names_select"]),
        ("Owner_names_select", people_select("Owner_names_select", pt('Owner'), "DisplayName")["Owner_names_select"]),
        ("SponsorEmails", compose(f"@{join_emails('Sponsor_emails_select', sponsor_email_fallback)}")),
        ("SponsorNames", compose(
            f"@if(empty(body('Sponsor_emails_select')),"
            f"if(empty(coalesce({sponsor_email_fallback},'')),'None',{sponsor_email_fallback}),"
            f"{join_names('Sponsor_names_select', 'None')})")),
        ("OwnerNames", compose(f"@{join_names('Owner_names_select', 'Not yet assigned')}")),
        ("Get_owner_group", sp_http("GET", "_api/web/associatedownergroup")),
        ("Break_inheritance", sp_http("POST", f"_api/web/lists(guid'{PROJECTS}')/items(@{{{ID}}})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")),
        ("Give_owners_full_control", sp_http("POST", f"_api/web/lists(guid'{PROJECTS}')/items(@{{{ID}}})/roleassignments/addroleassignment(principalid=@{{body('Get_owner_group')?['Id']}},roledefid=1073741829)")),
        ("Give_sponsors_read", {"type": "Scope", "runAfter": ra("Give_owners_full_control"), "actions": {
            "Sponsors_read_loop": {"type": "Foreach", "runAfter": {}, "foreach": "@split(outputs('SponsorEmails'),';')",
                "actions": seq([
                    ("Sponsor_read_who", compose("@trim(item())")),
                    ("Sponsor_read_check", cond("@not(empty(outputs('Sponsor_read_who')))", seq([
                        ("Ensure_sponsor", sp_http("POST", "_api/web/ensureuser", body="{\"logonName\":\"i:0#.f|membership|@{outputs('Sponsor_read_who')}\"}")),
                        ("Give_sponsor_read", sp_http("POST", f"_api/web/lists(guid'{PROJECTS}')/items(@{{{ID}}})/roleassignments/addroleassignment(principalid=@{{body('Ensure_sponsor')?['Id']}},roledefid=1073741826)")),
                    ]))),
                ])},
        }}),
        ("Send_created_email", {**mail("@outputs('Recipients')",
            f"[PRJ-@{{{ID}}}] New project: @{{{pt('Title')}}}",
            "<div style=\"font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333\">"
            "<p>A new IT project has been created.</p>"
            f"<p><b>PROJECT:</b> PRJ-@{{{ID}}}: @{{{pt('Title')}}}<br>"
            f"<b>STATUS:</b> @{{{pt('Status/Value')}}}<br>"
            f"<b>PRIORITY:</b> @{{{pt('Priority/Value')}}}<br>"
            "<b>OWNER:</b> @{outputs('OwnerNames')}<br>"
            "<b>SPONSOR:</b> @{outputs('SponsorNames')}<br>"
            f"<b>TARGET DATE:</b> @{{if(empty(coalesce({pt('TargetDate')},'')),'Not set',formatDateTime({pt('TargetDate')},'MMMM d, yyyy'))}}</p>"
            f"<p>@{{coalesce({pt('Description')},'')}}</p>"
            f"<p><a href=\"{PRJ_ITEM_URL}@{{{ID}}}\">View this project</a></p></div>"),
            "runAfter": ra("Recipients")}),
        ("Recipients", {**compose(
            "@join(union(if(empty(outputs('CreatorEmail')),json('[]'),createArray(outputs('CreatorEmail'))),"
            "union(createArray('" + ME + "'),split(outputs('SponsorEmails'),';'))),';')"),
            "runAfter": ra("Give_sponsors_read", status=("Succeeded", "Failed"))}),
        ("Martin_new_project", {**scope_msg("Martin_new_project_msg",
            ME, f"<p><b>New project: PRJ-@{{{ID}}} @{{{pt('Title')}}}</b></p><p>Status: @{{{pt('Status/Value')}}} · Owner: @{{outputs('OwnerNames')}}</p>"),
            "runAfter": ra("Send_created_email", status=("Succeeded", "Failed"))}),
        ("Tell_creator", {**cond(tell_cond("CreatorEmail"),
            {"Creator_new_project": scope_msg("Creator_new_project_msg", "@outputs('CreatorEmail')",
                f"<p><b>New project: PRJ-@{{{ID}}} @{{{pt('Title')}}}</b></p><p>You'll get a message here each time it's updated.</p>")}),
            "runAfter": ra("Martin_new_project", status=("Succeeded", "Failed"))}),
        ("Tell_sponsors", {**notify_each("Tell_sponsors", "outputs('SponsorEmails')",
            f"<p><b>A new IT project lists you as sponsor: PRJ-@{{{ID}}} @{{{pt('Title')}}}</b></p><p>You'll get a message here each time it's updated.</p>",
            exclude=["outputs('CreatorEmail')"])["Tell_sponsors"],
            "runAfter": ra("Tell_creator", status=("Succeeded", "Failed"))}),
        ("Update_item", {**spp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{pt('Title')}",
            "item/LastNotifiedStatus": f"@{pt('Status/Value')}",
            "item/LastNotifiedSponsorLen": f"@length(coalesce({pt('SponsorUpdates')},''))",
            "item/LastNotifiedInternalLen": f"@length(coalesce({pt('InternalNotes')},''))"}),
            "runAfter": ra("Tell_sponsors", status=("Succeeded", "Failed"))}),
    ])
    actions.update(self_check("Health", ["Give_sponsors_read", "Tell_creator", "Tell_sponsors"],
        f"[Help desk health] HD19 problem on a new project",
        f"Project <b>PRJ-@{{{ID}}}</b> (@{{{pt('Title')}}}) was created, but these steps did not complete:",
        f"concat('PRJ-',{ID},' (created)')",
        leaf=("Break_inheritance", "Give_owners_full_control", "Send_created_email", "Martin_new_project",
              "Creator_new_project"),
        after=ra("Update_item", status=("Succeeded", "Failed", "Skipped")),
        dash_id=DASH_PROJHEALTH, dash_title="ProjectHealth"))
    return "HD19 Project created", {"When_an_item_is_created": sp_trigger("GetOnNewItems", PROJECTS)}, actions

# ---------- HD20 Project updated ----------
def hd20():
    """Fires on every create/edit of an IT Projects item. Compares the current Status and the length of the two
    running note logs (Sponsor Updates, Internal Notes) against what was notified last time
    (LastNotifiedStatus/LastNotifiedSponsorLen/LastNotifiedInternalLen - stamped by HD19 on creation and by this
    flow at the end of every run) to find what's actually new; a plain re-trigger off its own Update_item write
    (see CLAUDE.md) finds nothing changed and exits quietly, same as HD4 already relies on for tickets.
    A Status change or a new Sponsor Updates entry notifies the project's creator, Martin and every Project
    Sponsor (email + Teams; Owner/Project Sponsor both allow multiple people - see hd19()'s docstring for the
    multi-value Person field shape and the legacy plain-text SponsorEmail fallback). A new Internal Notes
    entry notifies only Martin and the project's Owner(s) - never the sponsor, and never the creator unless
    they are also an Owner."""
    pt = lambda k: f"triggerOutputs()?['body/{k}']"
    ID = pt('ID')
    V = pt('Status/Value')
    status_changed = f"not(equals({V},coalesce({pt('LastNotifiedStatus')},'')))"
    sponsor_len = f"length(coalesce({pt('SponsorUpdates')},''))"
    sponsor_changed = f"greater({sponsor_len},coalesce({pt('LastNotifiedSponsorLen')},0))"
    internal_len = f"length(coalesce({pt('InternalNotes')},''))"
    internal_changed = f"greater({internal_len},coalesce({pt('LastNotifiedInternalLen')},0))"
    # the two note logs are prepended with newest-first (same convention as Tickets' Conversation), so the text
    # added since the last notification is just the leading slice the log grew by. substring()'s base string is
    # padded with a trailing space first (per CLAUDE.md) - SponsorUpdates/InternalNotes can be empty, and
    # substring('',0,0) throws "start index must be less than the length of the string" in Power Automate.
    # LastNotifiedSponsorLen/LastNotifiedInternalLen are SharePoint Number columns, which the REST API always
    # returns as a float/double even for whole numbers - coalesce()/sub()/max()/min() then silently widen the
    # whole expression to Float, and substring()'s length parameter rejects a Float with InvalidTemplate ("the
    # provided value is of type 'Float'"). Found 2026-10-25 full-circle-testing HD20's Internal Notes path (the
    # Sponsor Updates path didn't surface it because NoteLine's enclosing if() short-circuits the untaken
    # branch, so the same bug was latent there too). Fixed by wrapping the whole length expression in int(...).
    _SP = f"coalesce({pt('SponsorUpdates')},'')"
    _IN = f"coalesce({pt('InternalNotes')},'')"
    new_sponsor_text = f"substring(concat({_SP},' '),0,int(min({sponsor_len},max(0,sub({sponsor_len},coalesce({pt('LastNotifiedSponsorLen')},0))))))"
    new_internal_text = f"substring(concat({_IN},' '),0,int(min({internal_len},max(0,sub({internal_len},coalesce({pt('LastNotifiedInternalLen')},0))))))"
    tell_cond = lambda who: ("@and(not(empty(outputs('" + who + "'))),not(equals(outputs('" + who + "'),'" + ME
                              + "')),") + is_internal("outputs('" + who + "')") + ")"

    sponsor_branch = seq([
        ("StatusLine", compose(f"@if({status_changed},concat('<p><b>Status changed to: ',{V},'</b></p>'),'')")),
        ("NoteLine", {**compose(f"@if({sponsor_changed},concat('<p>',{new_sponsor_text},'</p>'),'')"), "runAfter": ra("StatusLine")}),
        ("CreatorEmail", {**compose(f"@toLower(coalesce({pt('Author/Email')},''))"), "runAfter": ra("NoteLine")}),
        ("Sponsor_emails_select", {**people_select("Sponsor_emails_select", pt('ProjectSponsor'), "Email")["Sponsor_emails_select"], "runAfter": ra("CreatorEmail")}),
        ("SponsorEmails", {**compose(f"@{join_emails('Sponsor_emails_select', pt('SponsorEmail'))}"), "runAfter": ra("Sponsor_emails_select")}),
        ("Recipients", {**compose(
            "@join(union(if(empty(outputs('CreatorEmail')),json('[]'),createArray(outputs('CreatorEmail'))),"
            "union(createArray('" + ME + "'),split(outputs('SponsorEmails'),';'))),';')"),
            "runAfter": ra("SponsorEmails")}),
        ("Email_sponsor_update", {**mail("@outputs('Recipients')", f"[PRJ-@{{{ID}}}] Update on project: @{{{pt('Title')}}}",
            "<div style=\"font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333\">"
            f"<p>{MARKER}</p>@{{outputs('StatusLine')}}@{{outputs('NoteLine')}}"
            f"<p>Reply to this email to add a note.</p><p><a href=\"{PRJ_ITEM_URL}@{{{ID}}}\">View this project</a></p></div>"),
            "runAfter": ra("Recipients")}),
        ("Martin_sponsor_update", {**scope_msg("Martin_sponsor_update_msg", ME,
            f"<p><b>PRJ-@{{{ID}}} updated: @{{{pt('Title')}}}</b></p>@{{outputs('StatusLine')}}@{{outputs('NoteLine')}}"),
            "runAfter": ra("Email_sponsor_update", status=("Succeeded", "Failed"))}),
        ("Tell_creator_sponsor_update", {**cond(tell_cond("CreatorEmail"),
            {"Creator_sponsor_update": scope_msg("Creator_sponsor_update_msg", "@outputs('CreatorEmail')",
                f"<p><b>PRJ-@{{{ID}}} updated: @{{{pt('Title')}}}</b></p>@{{outputs('StatusLine')}}@{{outputs('NoteLine')}}")}),
            "runAfter": ra("Martin_sponsor_update", status=("Succeeded", "Failed"))}),
        ("Tell_sponsors_update", {**notify_each("Tell_sponsors_update", "outputs('SponsorEmails')",
            f"<p><b>PRJ-@{{{ID}}} updated: @{{{pt('Title')}}}</b></p>@{{outputs('StatusLine')}}@{{outputs('NoteLine')}}",
            exclude=["outputs('CreatorEmail')"])["Tell_sponsors_update"],
            "runAfter": ra("Tell_creator_sponsor_update", status=("Succeeded", "Failed"))}),
    ])
    internal_branch = seq([
        ("InternalLine", compose(f"@concat('<p>',{new_internal_text},'</p>')")),
        ("Owner_emails_select", {**people_select("Owner_emails_select", pt('Owner'), "Email")["Owner_emails_select"], "runAfter": ra("InternalLine")}),
        ("OwnerEmails", {**compose(f"@{join_emails('Owner_emails_select')}"), "runAfter": ra("Owner_emails_select")}),
        ("Email_internal_update", {**mail(
            "@join(union(createArray('" + ME + "'),split(outputs('OwnerEmails'),';')),';')",
            f"[PRJ-@{{{ID}}}] Internal note: @{{{pt('Title')}}}",
            "<div style=\"font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333\">"
            f"<p>{MARKER}</p>@{{outputs('InternalLine')}}<p><a href=\"{PRJ_LIST_URL}@{{{ID}}}\">View in Full Detail</a></p></div>"),
            "runAfter": ra("OwnerEmails")}),
        ("Martin_internal_update", {**scope_msg("Martin_internal_update_msg", ME,
            f"<p><b>Internal note on PRJ-@{{{ID}}}: @{{{pt('Title')}}}</b></p>@{{outputs('InternalLine')}}"),
            "runAfter": ra("Email_internal_update", status=("Succeeded", "Failed"))}),
        ("Tell_owners_internal_update", {**notify_each("Tell_owners_internal_update", "outputs('OwnerEmails')",
            f"<p><b>Internal note on PRJ-@{{{ID}}}: @{{{pt('Title')}}}</b></p>@{{outputs('InternalLine')}}")["Tell_owners_internal_update"],
            "runAfter": ra("Martin_internal_update", status=("Succeeded", "Failed"))}),
    ])
    actions = seq([
        ("Sponsor_facing_change", cond(f"@or({status_changed},{sponsor_changed})", sponsor_branch)),
        ("Internal_only_change", {**cond(f"@{internal_changed}", internal_branch),
            "runAfter": ra("Sponsor_facing_change", status=("Succeeded", "Failed", "Skipped"))}),
        ("Update_item", {**spp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{pt('Title')}", "item/LastNotifiedStatus": f"@{V}",
            "item/LastNotifiedSponsorLen": f"@{sponsor_len}", "item/LastNotifiedInternalLen": f"@{internal_len}"}),
            "runAfter": ra("Internal_only_change", status=("Succeeded", "Failed", "Skipped"))}),
    ])
    actions.update(self_check("Health", ["Sponsor_facing_change", "Internal_only_change"],
        f"[Help desk health] HD20 problem on project update",
        f"Project <b>PRJ-@{{{ID}}}</b> (@{{{pt('Title')}}}) was updated, but these steps did not complete:",
        f"concat('PRJ-',{ID},' (updated)')",
        leaf=("Email_sponsor_update", "Martin_sponsor_update", "Creator_sponsor_update",
              "Email_internal_update", "Martin_internal_update"),
        after=ra("Update_item", status=("Succeeded", "Failed", "Skipped")),
        dash_id=DASH_PROJHEALTH, dash_title="ProjectHealth"))
    # Without a trigger-level condition, Update_item's own PATCH (it runs unconditionally, even when nothing
    # changed, to keep LastNotified* in sync) bumps the item's Modified date like any other edit, re-firing this
    # same trigger - which then finds status_changed/sponsor_changed/internal_changed all false and does nothing
    # but patch the identical values again, which bumps Modified again... an infinite ~1/min self-retrigger loop
    # that never sends a duplicate notification but never stops either. HD4 avoids exactly this with a trigger
    # condition (see its trig["conditions"] below); HD20 was missing the equivalent. Found 2026-10-25 full-circle
    # testing: a test project kept re-triggering HD20 every ~30-40s for 10+ minutes after one note was saved.
    trig = sp_trigger("GetOnUpdatedItems", PROJECTS)
    trig["conditions"] = [{"expression": f"@or({status_changed},{sponsor_changed},{internal_changed})"}]
    return "HD20 Project updated", {"When_an_item_is_created_or_modified": trig}, actions

# ---------- HD21 No-update check ----------
def hd21():
    """Runs once a week, Friday at 9am. Every active (not Completed/Cancelled) project whose Modified date is
    more than 7 days old - meaning nothing, not a status change, a sponsor update or an internal note, has
    touched it in a week, since every one of those writes also updates Modified - gets one email + Teams
    message to its creator and Martin asking for a status update."""
    it = lambda k: f"items('Apply_to_each')?['{k}']"
    ID = it('ID')
    status_expr = f"{it('Status')}?['Value']"
    lnk = f"<a href=\"{PRJ_ITEM_URL}@{{{ID}}}\">View this project</a>"
    body_html = (f"<p>PRJ-@{{{ID}}}: @{{{it('Title')}}} hasn't had a status change or a new note in over a week"
                 f" (last touched @{{formatDateTime({it('Modified')},'MMMM d')}}).</p>"
                 f"<p>Status: @{{{status_expr}}}</p><p>{lnk}</p>")
    tell_cond = lambda who: ("@and(not(empty(outputs('" + who + "'))),not(equals(outputs('" + who + "'),'" + ME
                              + "')),") + is_internal("outputs('" + who + "')") + ")"
    loop_actions = seq([
        ("CreatorEmail", compose(f"@toLower(coalesce({it('Author')}?['Email'],''))")),
        ("Email_creator", {**mail("@outputs('CreatorEmail')", f"[PRJ-@{{{ID}}}] Status update needed: @{{{it('Title')}}}", body_html),
            "runAfter": ra("CreatorEmail")}),
        ("Tell_creator_no_update", {**cond(tell_cond("CreatorEmail"),
            {"Creator_no_update": scope_msg("Creator_no_update_msg", "@outputs('CreatorEmail')", body_html)}),
            "runAfter": ra("Email_creator", status=("Succeeded", "Failed"))}),
        ("Tell_martin_no_update", {**scope_msg("Martin_no_update", ME, body_html),
            "runAfter": ra("Tell_creator_no_update", status=("Succeeded", "Failed", "Skipped"))}),
    ])
    loop = {"type": "Foreach", "foreach": "@body('Get_stale_projects')?['value']", "actions": loop_actions}
    actions = seq([
        ("Get_stale_projects", spp("GetItems", {
            "$filter": "(" + " or ".join(f"Status eq '{s}'" for s in PRJ_ACTIVE) + ") and Modified lt '@{addDays(utcNow(),-7)}'",
            "$top": 200})),
        ("Apply_to_each", loop),
    ])
    return "HD21 No-update check", {"Recurrence": recurrence([9], ["Friday"])}, actions

def write_package(flows, zip_path, pkg_name):
    tmp = zip_path + "_dir"
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(os.path.join(tmp, "Microsoft.Flow", "flows"))
    resources = {}
    api_ids, conn_ids = {}, {}
    asset_paths = []
    for name, trig, acts in flows:
        apis = set()
        used_apis({"t": trig, "a": acts}, apis)
        for a in sorted(apis):
            if a not in conn_ids:
                short, disp, _ = APIS[a]
                real = "shared_" + short
                if real not in api_ids:
                    api_ids[real] = str(uuid.uuid4())
                    resources[api_ids[real]] = {"id": f"/providers/Microsoft.PowerApps/apis/{real}", "name": real,
                        "type": "Microsoft.PowerApps/apis", "suggestedCreationType": "Existing",
                        "details": {"displayName": APIS[real][1]}, "configurableBy": "System", "hierarchy": "Child", "dependsOn": []}
                api_ids[a] = api_ids[real]
                conn_ids[a] = str(uuid.uuid4())
                resources[conn_ids[a]] = {"type": "Microsoft.PowerApps/apis/connections", "suggestedCreationType": "Existing",
                    "creationType": "Existing", "details": {"displayName": SUPPORT if a == "shared_teams_1" else ME}, "configurableBy": "User",
                    "hierarchy": "Child", "dependsOn": [api_ids[a]]}
        fid = str(uuid.uuid4())
        asset_paths.append(fid)
        resources[fid] = {"type": "Microsoft.Flow/flows", "suggestedCreationType": "New",
            "creationType": "Existing, New, Update", "details": {"displayName": name}, "configurableBy": "User",
            "hierarchy": "Root", "dependsOn": sorted(set(x for a in sorted(apis) for x in (api_ids[a], conn_ids[a])))}
        flow_guid = str(uuid.uuid4())
        d = {"name": flow_guid, "id": f"/providers/Microsoft.Flow/flows/{flow_guid}", "type": "Microsoft.Flow/flows",
             "properties": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows", "displayName": name,
                "description": FLOW_DESCRIPTIONS.get(name, ""),
                "definition": build_definition(name, trig, acts),
                "connectionReferences": {a: {"connectionName": APIS[a][2], "source": "Embedded",
                    "id": f"/providers/Microsoft.PowerApps/apis/shared_{APIS[a][0]}", "tier": "NotSpecified", "apiName": APIS[a][0],
                    "isProcessSimpleApiReferenceConversionAlreadyDone": False} for a in sorted(apis)},
                "flowFailureAlertSubscribed": True, "isManaged": False}}
        fd = os.path.join(tmp, "Microsoft.Flow", "flows", fid)
        os.makedirs(fd)
        json.dump(d, open(os.path.join(fd, "definition.json"), "w"))
        json.dump({a: api_ids[a] for a in sorted(apis)}, open(os.path.join(fd, "apisMap.json"), "w"))
        json.dump({a: conn_ids[a] for a in sorted(apis)}, open(os.path.join(fd, "connectionsMap.json"), "w"))
    json.dump({"packageSchemaVersion": "1.0", "flowAssets": {"assetPaths": asset_paths}},
              open(os.path.join(tmp, "Microsoft.Flow", "flows", "manifest.json"), "w"))
    json.dump({"schema": "1.0", "details": {"displayName": pkg_name, "description": "Branthaven IT help desk flows (pilot)",
               "createdTime": "2026-10-10T12:00:00Z", "packageTelemetryId": str(uuid.uuid4()), "creator": "N/A", "sourceEnvironment": ""},
               "resources": resources}, open(os.path.join(tmp, "manifest.json"), "w"))
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(tmp):
            for f in files:
                p = os.path.join(root, f)
                z.write(p, os.path.relpath(p, tmp))
    shutil.rmtree(tmp)

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    flows = [hd1(), hd2(), hd3(), hd4(), hd5(), hd6(), hd7(), hd8(), hd9(), hd11(), hd12(), hd13(), hd14(), hd15(), hd16(), hd17(), hd18(), hd19(), hd20(), hd21()]   # HD10 retired: questions now go out as IT Support messages and replies come back via HD12
    for f in flows:
        safe = f[0].split()[0]
        write_package([f], os.path.join(OUT, f"{safe}.zip"), f[0])
    write_package(flows, os.path.join(OUT, "HelpDesk-all-flows.zip"), "IT Help Desk flows")
    # dump definitions for review
    for f in flows:
        json.dump(build_definition(*f), open(os.path.join(OUT, f[0].split()[0] + ".json"), "w"), indent=1)
    print("ok", [f[0] for f in flows])
