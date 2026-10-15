import json, uuid, os, zipfile, shutil, copy

OUT = os.path.join(os.path.dirname(__file__), 'out')
SITE = "https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2"
TICKETS = "22c09b0b-bbcd-42a9-9aeb-844d81089636"
HOLIDAYS = "430cd9b6-9cdc-45a2-8e9b-66a88cd19b74"
MBX = "IThelpdesk-test@branthaven.com"
ME = "mpersaud@branthaven.com"
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
REQ_URL = SITE + "/Lists/Tickets/Ticket.aspx?FilterField1=ID&FilterValue1="  # simple requester view

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
}

# ---------- statuses ----------
S_OPEN, S_ASSIGNED, S_PROG, S_ESC = "Open", "Assigned", "In Progress", "Escalated or Waiting Internally"
S_WAIT, S_SCHED, S_RES, S_CLOSED = "Waiting for Client", "To Be Scheduled", "Resolved Pending Confirmation", "Closed"
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
            f" &nbsp;·&nbsp; <a href=\"{mailto}\">Reply by email</a> &nbsp;·&nbsp; <a href=\"{NEWISSUE_URL}\">Report a new issue</a></p>"
            "<p style=\"color:#888;font-size:12px\">To add details or reply to IT, use Reply by email. To report a different problem, use Report a new issue. Messages typed in this chat aren't seen by IT.</p>")

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

def action_card(ID, title, req, cat, pri, src, details, pic=None):
    """Static ticket card with action buttons (Adaptive Card 1.2 for Teams mobile). Args are expressions without a leading @."""
    tid = f"@{{{ID}}}"
    def text_card(inputs, act):
        return {"type": "AdaptiveCard", "body": inputs,
                "actions": [{"type": "Action.Submit", "title": "Send", "data": {"act": act, "ticket": tid}}]}
    status_inputs = [
        {"type": "TextBlock", "text": "New status", "weight": "Bolder"},
        {"type": "Input.ChoiceSet", "id": "status_choice", "style": "compact", "value": S_PROG,
         "choices": [{"title": x, "value": x} for x in STATUSES if x != S_OPEN]},
        {"type": "TextBlock", "text": "Message (sent to the user for Waiting for Client and To Be Scheduled; "
                                      "the resolution for Resolved; an internal note for Escalated)", "wrap": True, "size": "Small", "isSubtle": True},
        {"type": "Input.Text", "id": "status_text", "isMultiline": True, "placeholder": "Optional message"}]
    return {"type": "AdaptiveCard", "version": "1.2",
        "body": [
            {"type": "TextBlock", "text": f"IT-@{{{ID}}}: @{{{title}}}", "weight": "Bolder", "size": "Medium", "wrap": True},
            {"type": "FactSet", "facts": [
                {"title": "Requester", "value": f"@{{{req}}}"},
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
            {"type": "Action.OpenUrl", "title": "View ticket", "url": f"{LIST_URL}@{{{ID}}}"},
            *([{"type": "Action.OpenUrl", "title": "View image", "url": f"{SITE}/Lists/Tickets/DispForm.aspx?ID=@{{{ID}}}", "isVisible": pic}] if pic else [])]}

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

def nowET(fmt):
    return f"convertTimeZone(utcNow(),'UTC','{TZ}','{fmt}')"

WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ALLDAYS = ["Sunday"] + WEEKDAYS + ["Saturday"]
DAY_H = 14   # business day = 7am-9pm = 14 hours

def recurrence(hours, days=WEEKDAYS):
    return {"type": "Recurrence", "recurrence": {
        "frequency": "Week", "interval": 1, "timeZone": TZ,
        "schedule": {"weekDays": days,
                     "hours": [str(h) for h in hours], "minutes": [0]}}}

def sp_trigger(opid):
    return {"recurrence": {"frequency": "Minute", "interval": 1},
            "splitOn": "@triggerOutputs()?['body/value']", "type": "OpenApiConnection",
            "inputs": {"parameters": {"dataset": SITE, "table": TICKETS},
                       "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                                "connectionName": "shared_sharepointonline", "operationId": opid},
                       "authentication": "@parameters('$authentication')"}}

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
         ("Waiting for Client", "IT Support has requested further information from you"),
         ("To Be Scheduled", "IT Support has requested to schedule a meeting with you"),
         ("Resolved Pending Confirmation", "IT Support has provided resolution and awaiting confirmation that issue has been resolved"),
         ("Closed", "Ticket has been resolved and closed")]
CONFIRM_BODY = (
    '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333">'
    "<p>Hi @{first(split(trim(coalesce(§T§Requester/DisplayName'],first(split(outputs('ReqEmail'),'@')))),' '))},</p>"
    "<p><b>SUBJECT:</b> @{§T§Title']}<br><b>DESCRIPTION:</b> @{coalesce(§T§Details'],'')}</p>"
    "@{if(outputs('HasPic'),'<p><b>IMAGE:</b> 1 image is attached to your ticket.</p>','')}"
    "<p>This email was system generated for the purpose of communicating that the referenced Ticket Number has been created to record your contact with the IT Team. "
    "To keep you informed on the progress of your Ticket, you will receive ongoing automated emails and Teams messages when your IT Ticket Status changes as follows:</p>"
    '<table style="border-collapse:collapse;margin:8px 0 12px 0">'
    f"<tr><th {_TH}>TICKET STATUS</th><th {_TH}>MEANING</th></tr>"
    + "".join(f"<tr><td {_TD}>{a}</td><td {_TD}>{b}</td></tr>" for a, b in _ROWS) +
    "</table>"
    f'<p>Please feel free to use our <a href="{PORTAL_URL}">Ticket Portal</a> anytime to view the Status of your Ticket. To add comments or attachments, reply to this email. To report a different problem, <a href="' + NEWISSUE_URL + '">submit a new request</a>.</p>'
    "<p>You will also receive Teams messages from <b>IT Support</b>. Replies typed in that chat are not seen by IT. To reply to IT, reply to this email.</p>"
    "<p><b>TICKET NUMBER:</b> IT-@{§ID§}<br>"
    "<b>TICKET STATUS:</b> @{if(empty(coalesce(§T§AssignedTo/Email'],'')),'Open','Assigned')}<br>"
    "<b>TICKET CONTACT:</b> @{coalesce(§T§Requester/DisplayName'],outputs('ReqEmail'))} (@{outputs('ReqEmail')})<br>"
    "<b>TICKET CREATED:</b> @{convertTimeZone(§T§Created'],'UTC','Eastern Standard Time','MMMM d, yyyy h:mm tt')}<br>"
    "<b>TICKET URGENCY:</b> @{§T§Priority/Value']}<br>"
    "<b>TICKET ASSIGNEE:</b> @{coalesce(§T§AssignedTo/DisplayName'],'Not yet assigned')}</p>"
    "</div>")

SIG_MARKERS = ["Regards", "Thks", "Thanks,", "Thank you,", "Cheers,", "Sent from my", "Get Outlook for", "[cid:", "[https://", "T: ", "________________"]
def strip_signature(expr):
    """Cut text at the first signature marker. expr is an expression without a leading @."""
    out = expr
    for m in SIG_MARKERS:
        out = f"first(split({out},'{m}'))"
    return f"trim({out})"

USER_CARD_TYPE = "ITHelpDeskUserCard"
CHAT_NOTICE = "Please don't reply in this chat, as messages typed here aren't seen by IT. Use Reply to IT or the answer box above."

def user_card(ID, blocks, title_expr="triggerOutputs()?['body/Title']"):
    """Card sent to the requester in their Workflows chat. Replies go by email (HD1), because Teams
    does not pass card clicks from the Workflows chat to a flow."""
    body = [{"type": "TextBlock", "text": txt, "wrap": True, **({"weight": "Bolder", "size": "Medium"} if bold else {})}
            for txt, bold in blocks]
    body.append({"type": "TextBlock", "text": CHAT_NOTICE.replace(" or the answer box above", ""), "wrap": True, "size": "Small", "isSubtle": True})
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
        return {"runAfter": after, "type": "Foreach", "foreach": "@triggerOutputs()?['body/attachments']",
                "actions": {f"Add_attachment{suffix}": sp("CreateAttachment", {
                    "itemId": item_id, "displayName": f"@items('Apply_to_each{suffix}')?['name']",
                    "body": f"@base64ToBinary(items('Apply_to_each{suffix}')?['contentBytes'])"})}}

    new_item = {
        "item/Title": title_expr,
        "item/Details": "@outputs('CleanBody')",
        "item/Requester/Claims": "@outputs('Sender')",
        "item/RequesterEmail": "@outputs('Sender')",
        "item/Source/Value": "Email",
        "item/Category/Value": "Uncategorised",
        "item/Priority/Value": f"@if(contains(toLower(coalesce({T},'')),'urgent'),'High','Medium')",
        "item/Status/Value": S_OPEN,
    }
    followup = dict(new_item)
    followup["item/Details"] = ("@if(equals(body('Get_item')?['Status']?['Value'],'Closed'),"
                                "concat('Follow-up to IT-',string(outputs('TicketId')),'. ',outputs('ReplyText')),outputs('CleanBody'))")

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
            "Tell_the_tech": chat(f"@{{{gi}?['AssignedTo']?['Email']}}",
                f"<p>@{{triggerOutputs()?['body/from']}} replied on IT-@{{outputs('TicketId')}}:</p><p>@{{outputs('ReplyText')}}</p><p>{LT}</p>")})),
        ("Has_channel_post", cond(f"@not(empty(coalesce({gi}?['ChannelMessageId'],'')))", {
            "Reply_in_channel": channel_reply(f"@{gi}?['ChannelMessageId']",
                f"<p>@{{triggerOutputs()?['body/from']}} replied on IT-@{{outputs('TicketId')}}:</p><p>@{{outputs('ReplyText')}}</p>")})),
        ("Was_resolved", cond(f"@equals({gi}?['Status']?['Value'],'{S_RES}')", {
            "Tell_requester_reopened": {"type": "Scope", "runAfter": {}, "actions": support_msg("Post_reopened", "@outputs('Sender')",
                f"<p><b>IT-@{{outputs('TicketId')}} has been reopened</b></p><p>@{{{gi}?['Title']}}</p>" + support_links("outputs('TicketId')", f"{gi}?['Title']"))}})),
    ])
    reply_no = seq([
        ("Create_followup_item", sp("PostItem", followup)),
        ("Apply_to_each_2", attach_loop("@body('Create_followup_item')?['ID']", None, "_2")),
    ])
    reply_branch = seq([
        ("TicketId", compose(f"@if(contains(coalesce({T},''),'[IT-'),int(first(split(last(split({T},'[IT-')),']'))),outputs('RefNum'))")),
        ("ReplyText", compose("@" + strip_signature(f"first(split(first(split(body('Html_to_text'),'{MARKER}')),'From: '))"))),
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
    actions = seq([
        ("Sender", compose("@toLower(triggerOutputs()?['body/from'])")),
        ("Is_it_a_real_internal_email", cond(
            f"@and(endsWith(outputs('Sender'),'@branthaven.com'),not(equals(outputs('Sender'),toLower('{MBX}'))),"
            f"not(startsWith(coalesce({T},''),'Automatic reply')),not(startsWith(coalesce({T},''),'Undeliverable')))",
            {}, seq([("Move_to_Review", move(F_REVIEW)),
                     ("Terminate", {"type": "Terminate", "inputs": {"runStatus": "Succeeded"}})]))),
        ("Html_to_text", op("shared_conversionservice", "HtmlToText", {"Content": "@triggerOutputs()?['body/body']"})),
        ("CleanBody", compose("@if(empty(" + strip_signature("body('Html_to_text')") + "),substring(body('Html_to_text'),0,min(length(body('Html_to_text')),2000))," + strip_signature("body('Html_to_text')") + ")")),
        # safety net: no [IT-n] in the subject, but the subject or body names IT-n (first occurrence; digits only)
        ("RefPart", compose(f"@concat(coalesce(split(toUpper(concat(coalesce({T},''),' ',coalesce(body('Html_to_text'),''))),'IT-')?[1],''),'     ')")),
        ("RefNum", compose("@if(not(contains('0123456789',substring(outputs('RefPart'),0,1))),0,"
                           "int(substring(outputs('RefPart'),0,if(not(contains('0123456789',substring(outputs('RefPart'),1,1))),1,"
                           "if(not(contains('0123456789',substring(outputs('RefPart'),2,1))),2,"
                           "if(not(contains('0123456789',substring(outputs('RefPart'),3,1))),3,"
                           "if(not(contains('0123456789',substring(outputs('RefPart'),4,1))),4,5)))))))")),
        ("Is_it_a_reply", cond(f"@or(contains(coalesce({T},''),'[IT-'),greater(outputs('RefNum'),0))", reply_branch, new_branch)),
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
        ("Tag", compose(f"@concat('IT-',{ID})")),
        ("Get_owner_group", sp_http("GET", "_api/web/associatedownergroup")),
        ("Break_inheritance", sp_http("POST", f"_api/web/lists/getbytitle('Tickets')/items(@{{{ID}}})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)")),
        ("Give_owners_full_control", sp_http("POST", f"_api/web/lists/getbytitle('Tickets')/items(@{{{ID}}})/roleassignments/addroleassignment(principalid=@{{body('Get_owner_group')?['Id']}},roledefid=1073741829)")),
        ("Has_requester", cond("@not(empty(outputs('ReqEmail')))", seq([
            ("Ensure_requester", sp_http("POST", "_api/web/ensureuser",
                body="{\"logonName\":\"i:0#.f|membership|@{outputs('ReqEmail')}\"}")),
            ("Give_requester_read", sp_http("POST", f"_api/web/lists/getbytitle('Tickets')/items(@{{{ID}}})/roleassignments/addroleassignment(principalid=@{{body('Ensure_requester')?['Id']}},roledefid=1073741826)")),
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
                                 f"Source: @{{{t('Source/Value')}}}\n\n@{{{t('Details')}}}\n\n{LIST_URL}@{{{ID}}}")})),
        ("Wait_for_attachment", {"type": "Wait", "inputs": {"interval": {"count": 10, "unit": "Second"}}}),
        ("Get_ticket", sp("GetItem", {"id": f"@{ID}"})),
        ("HasPic", compose("@equals(body('Get_ticket')?['{HasAttachments}'],true)")),
        ("IsOpenNow", compose(
            f"@and(greaterOrEquals(int(convertTimeZone(utcNow(),'UTC','{TZ}','HH')),7),less(int(convertTimeZone(utcNow(),'UTC','{TZ}','HH')),21))")),
        ("HoursLine", compose("@if(outputs('IsOpenNow'),'A technician will pick it up shortly.','Support hours are 7:00 am to 9:00 pm, every day. Your ticket will be actioned when support reopens at 7:00 am.')")),
        ("Tell_requester", cond("@not(empty(outputs('ReqEmail')))", seq([
            ("Send_confirmation_email", mail("@outputs('ReqEmail')",
                f"[IT-@{{{ID}}}] We've got your request: @{{{t('Title')}}}",
                CONFIRM_BODY.replace("§ID§", ID).replace("§T§", "triggerOutputs()?['body/"))),
            ("Confirm_in_Teams", {"type": "Scope", "runAfter": ra("Send_confirmation_email", status=("Succeeded", "Failed")),
                "actions": support_msg("Send_confirmation_chat", "@outputs('ReqEmail')",
                    f"<p><b>IT-@{{{ID}}} created:</b> @{{{t('Title')}}}</p>"
                    f"<p>Status: @{{if(empty(coalesce({t('AssignedTo/Email')},'')),'Open','Assigned')}} · Urgency: @{{{t('Priority/Value')}}}</p>"
                    "@{if(outputs('HasPic'),'<p>Image attached.</p>','')}"
                    "<p>You'll get a message here and an email each time the status changes.</p>" + support_links(ID, t('Title')))}),
        ]))),
        ("CreatedLine", compose(f"@concat('[',convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm'),'] Ticket created from ',coalesce({t('Source/Value')},'IT'),' by ',outputs('ReqEmail'),decodeUriComponent('%0A%0A'),coalesce({t('Conversation')},''))")),
        ("Card", compose(action_card(ID, t('Title'), "outputs('ReqEmail')", t('Category/Value'), t('Priority/Value'), t('Source/Value'),
            f"substring(coalesce({t('Details')},''),0,min(length(coalesce({t('Details')},'')),400))", pic="@outputs('HasPic')"))),
        ("Post_to_Help_Desk_channel", channel_card("@{string(outputs('Card'))}", card_type=CARD_TYPE)),
        ("Update_item", sp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{t('Title')}",
            "item/RequesterEmail": "@outputs('ReqEmail')",
            "item/PlannerTaskId": "@body('Create_a_task')?['id']",
            "item/ChannelMessageId": "@body('Post_to_Help_Desk_channel')?['id']",
            "item/Conversation": "@outputs('CreatedLine')"})),
    ])
    actions["Tell_requester"]["runAfter"] = ra("HoursLine")
    actions["CreatedLine"]["runAfter"] = ra("Tell_requester", status=("Succeeded", "Failed", "Skipped"))
    actions["Update_item"]["runAfter"] = ra("Post_to_Help_Desk_channel", status=("Succeeded", "Failed"))
    return "HD3 Ticket created", {"When_an_item_is_created": sp_trigger("GetOnNewItems")}, actions

# ---------- HD4 Ticket updated ----------
def hd4():
    t = lambda k: f"triggerOutputs()?['body/{k}']"
    ID = t('ID')
    hasTask = f"@not(empty(coalesce({t('PlannerTaskId')},'')))"
    lnk = link(ID)
    V = "variables('varStatus')"
    stamp = f"convertTimeZone(utcNow(),'UTC','{TZ}','yyyy-MM-dd HH:mm')"
    who = f"coalesce({t('Editor/DisplayName')},'IT')"
    first = "outputs('FirstName')"
    portal = f'<a href="{PORTAL_URL}">Ticket Portal</a>'
    details = ("<p><b>TICKET NUMBER:</b> IT-@{" + ID + "}<br>"
               "<b>TICKET STATUS:</b> @{" + V + "}<br>"
               "<b>TICKET ASSIGNEE:</b> @{coalesce(" + t('AssignedTo/DisplayName') + ",'Not yet assigned')}<br>"
               "<b>TICKET URGENCY:</b> @{" + t('Priority/Value') + "}</p>")
    _hist = "replace(substring(coalesce(triggerOutputs()?['body/Conversation'],''),0,min(length(coalesce(triggerOutputs()?['body/Conversation'],'')),3000)),decodeUriComponent('%0A%0A'),'<br>')"
    earlier = ("@{if(empty(coalesce(triggerOutputs()?['body/Conversation'],'')),'',concat('<p style=\"margin-top:18px\"><b>Ticket history</b></p>"
               "<div style=\"font-size:13px;color:#555;line-height:1.5\">'," + _hist + ",'</div>'))}")
    style = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333">'

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
        ("Reassigned_same_status", cond(f"@and(not(empty(coalesce({t('LastNotifiedTech')},''))),equals({V},coalesce({t('LastNotifiedStatus')},'')))", {
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
                f"<p>Reply to this email to answer. You can also view your ticket in the {portal}.</p>" + details + earlier + "</div>")),
            ("Chat_the_note", {"type": "Scope", "runAfter": ra("Email_the_note", status=("Succeeded", "Failed")),
                "actions": support_msg("Post_note", "@outputs('ReqEmail')",
                    f"<p><b>IT-@{{{ID}}}: message from @{{{who}}}</b></p><p>@{{{t('NoteToUser')}}}</p><p style=\"color:#666\">@{{{t('Title')}}}</p>" + support_links(ID, t('Title')))}),
            ("Note_in_channel", {**cond(f"@not(empty(coalesce({t('ChannelMessageId')},'')))", {
                "Reply_note_in_channel": channel_reply(f"@{t('ChannelMessageId')}",
                    f"<p>Note to user from @{{{who}}}: @{{{t('NoteToUser')}}}</p>")}),
                "runAfter": ra("Chat_the_note", status=("Succeeded", "Failed"))}),
        ]))),
    ])

    # Part C: the status changed
    note_html = f"@{{if(empty(coalesce({t('NoteToUser')},'')),'',concat('<p><b>Message from IT:</b> ',{t('NoteToUser')},'</p>'))}}"
    extra_html = (f"@{{if(equals({V},'{S_RES}'),concat('<p><b>Resolution:</b> ',coalesce({t('Resolution')},''),'</p><p>If this is not fixed, reply to this email within 2 days and your ticket will be reopened.</p>'),"
                  f"if(contains({arr([S_WAIT, S_SCHED])},{V}),'<p>Please reply to this email to respond to IT.</p>',"
                  f"if(equals({V},'{S_CLOSED}'),'<p>To raise this again, reply to this email or log a new ticket.</p>','')))}}")
    partC = seq([
        ("Has_task_C", cond(hasTask, {"Move_Planner_task": op("shared_planner", "UpdateTask_V3", {
            "id": f"@{t('PlannerTaskId')}", "body/bucketId": "@" + bucket_expr(V),
            "body/percentComplete": f"@if(contains({arr([S_RES, S_CLOSED])},{V}),100,if(equals({V},'{S_OPEN}'),0,50))"})})),
        ("Meanings", {**compose(MEANING), "runAfter": ra("Has_task_C", status=("Succeeded", "Failed", "Skipped"))}),
        ("StatusLine", compose(f"@concat('[',{stamp},' | ',{who},'] Status: ',{V},if(equals({V},'{S_RES}'),concat(' · Resolution: ',coalesce({t('Resolution')},'')),''),decodeUriComponent('%0A%0A'),variables('varConversation'))")),
        ("Set_varConversation_status", {"type": "SetVariable", "inputs": {"name": "varConversation", "value": "@outputs('StatusLine')"}}),
        ("Has_requester", cond("@not(empty(outputs('ReqEmail')))", seq([
            ("Email_status", mail("@outputs('ReqEmail')", f"[IT-@{{{ID}}}] Ticket status: @{{{V}}}: @{{{t('Title')}}}",
                style + f"<p style=\"color:#999;font-size:12px\">{MARKER}</p><p>Hi @{{{first}}},</p>"
                f"<p>The status of your ticket IT-@{{{ID}}} has changed to <b>@{{{V}}}</b>: @{{outputs('Meanings')?[{V}]}}.</p>"
                + note_html + extra_html +
                f"<p>You can view your ticket anytime in the {portal}.</p>" + details + earlier + "</div>")),
            ("Chat_status", {"type": "Scope", "runAfter": ra("Email_status", status=("Succeeded", "Failed")),
                "actions": support_msg("Post_status", "@outputs('ReqEmail')",
                    f"<p><b>IT-@{{{ID}}} is now @{{{V}}}</b></p><p>@{{outputs('Meanings')?[{V}]}}.</p><p style=\"color:#666\">@{{{t('Title')}}}</p>"
                    f"@{{if(empty(coalesce({t('NoteToUser')},'')),'',concat('<p><b>Message from IT:</b> ',{t('NoteToUser')},'</p>'))}}"
                    f"@{{if(equals({V},'{S_RES}'),concat('<p><b>Resolution:</b> ',coalesce({t('Resolution')},''),'</p><p>If this is not fixed, use Reply by email within 2 days and your ticket will be reopened.</p>'),if(contains({arr([S_WAIT, S_SCHED])},{V}),'<p>Please use Reply by email to respond to IT.</p>',''))}}"
                    + support_links(ID, t('Title')))}),
        ]))),
        ("Question_missing", {**cond(f"@and(contains({arr([S_WAIT, S_SCHED])},{V}),empty(coalesce({t('NoteToUser')},'')))", {
            "Nudge_the_tech": chat(f"@{{coalesce({t('Editor/Email')},'{ME}')}}",
                f"<p>IT-@{{{ID}}} is set to @{{{V}}}, but no message was sent to the user. Add one with Note to user.</p><p>{lnk}</p>")}),
            "runAfter": ra("Has_requester", status=("Succeeded", "Failed", "Skipped"))}),
        ("Has_channel_post", {**cond(f"@not(empty(coalesce({t('ChannelMessageId')},'')))", {
            "Reply_status_in_channel": channel_reply(f"@{t('ChannelMessageId')}",
                f"<p>IT-@{{{ID}}} is now <b>@{{{V}}}</b>"
                f"@{{if(empty(coalesce({t('AssignedTo/DisplayName')},'')),'',concat(' · IT: ',{t('AssignedTo/DisplayName')}))}}"
                f" · updated by @{{{who}}}</p>"
                f"@{{if(empty(coalesce({t('NoteToUser')},'')),'',concat('<p>Message to user: ',{t('NoteToUser')},'</p>'))}}"
                f"@{{if(equals({V},'{S_RES}'),concat('<p>Resolution: ',coalesce({t('Resolution')},''),'</p>'),'')}}")}),
            "runAfter": ra("Question_missing", status=("Succeeded", "Failed", "Skipped"))}),
    ])

    def var(name, typ, val):
        return {"type": "InitializeVariable", "inputs": {"variables": [{"name": name, "type": typ, "value": val}]}}
    ALL = ("Succeeded", "Failed", "Skipped", "TimedOut")
    actions = seq([
        ("Init_varStatus", var("varStatus", "string", f"@{t('Status/Value')}")),
        ("Init_varConversation", var("varConversation", "string", f"@coalesce({t('Conversation')},'')")),
        ("Init_varFirstResponse", var("varFirstResponse", "string", f"@coalesce({t('FirstResponseOn')},'')")),
        ("Tech", compose(f"@toLower(coalesce({t('AssignedTo/Email')},''))")),
        ("ReqEmail", compose(f"@toLower(coalesce({t('RequesterEmail')},{t('Requester/Email')},''))")),
        ("FirstName", compose(f"@first(split(trim(coalesce({t('Requester/DisplayName')},first(split(outputs('ReqEmail'),'@')),'there')),' '))")),
        ("Part_B_tech_assigned", cond(f"@and(not(empty(outputs('Tech'))),not(equals(outputs('Tech'),toLower(coalesce({t('LastNotifiedTech')},'')))))", partB)),
        ("Part_A_note_for_user", {**cond(f"@not(empty(coalesce({t('NoteToUser')},'')))", partA),
                                  "runAfter": ra("Part_B_tech_assigned", status=ALL)}),
        ("Part_C_status_changed", {**cond(f"@not(equals({V},coalesce({t('LastNotifiedStatus')},'')))", partC),
                                   "runAfter": ra("Part_A_note_for_user", status=ALL)}),
        ("Update_item", {**sp("PatchItem", {
            "id": f"@{ID}", "item/Title": f"@{t('Title')}",
            "item/Status/Value": f"@{V}",
            "item/LastNotifiedStatus": f"@{V}",
            "item/LastNotifiedTech": "@outputs('Tech')",
            "item/NoteToUser": "@null",
            "item/Conversation": "@variables('varConversation')",
            "item/FirstResponseOn": "@if(empty(variables('varFirstResponse')),null,variables('varFirstResponse'))",
            "item/ResolvedOn": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_RES}')),utcNow(),{t('ResolvedOn')})",
            "item/ClosedOn": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_CLOSED}')),utcNow(),{t('ClosedOn')})",
            "item/WaitingHours": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_WAIT}')),0,{t('WaitingHours')})",
            "item/ResolvedHours": f"@if(and(not(equals({V},coalesce({t('LastNotifiedStatus')},''))),equals({V},'{S_RES}')),0,{t('ResolvedHours')})"}),
            "runAfter": ra("Part_C_status_changed", status=ALL)}),
    ])
    trig = sp_trigger("GetOnUpdatedItems")
    trig["conditions"] = [{"expression":
        "@or(not(equals(triggerOutputs()?['body/Status/Value'],triggerOutputs()?['body/LastNotifiedStatus'])),"
        "not(equals(toLower(coalesce(triggerOutputs()?['body/AssignedTo/Email'],'')),toLower(coalesce(triggerOutputs()?['body/LastNotifiedTech'],'')))),"
        "greater(length(coalesce(triggerOutputs()?['body/NoteToUser'],'')),0))"}]
    return "HD4 Ticket updated", {"When_an_item_is_created_or_modified": trig}, actions

# ---------- HD5 Planner task completed ----------
def hd5():
    actions = seq([
        ("Get_items", sp("GetItems", {"$filter": "PlannerTaskId eq '@{triggerOutputs()?['body/id']}'", "$top": 1})),
        ("Condition", cond(f"@and(greater(length(body('Get_items')?['value']),0),not(contains({arr([S_RES, S_CLOSED])},first(body('Get_items')?['value'])?['Status']?['Value'])))", {
            "Update_item": sp("PatchItem", {"id": "@first(body('Get_items')?['value'])?['ID']",
                                            "item/Title": "@first(body('Get_items')?['value'])?['Title']",
                                            "item/Status/Value": S_RES})})),
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
    rlnk = req_link(ID)
    req = f"toLower(coalesce({it('RequesterEmail')},''))"
    biz = "outputs('BizDay')"
    new_case = seq([
        ("H", compose(f"@add(int(coalesce({it('BusinessHoursOpen')},0)),1)")),
        ("S", compose(f"@add(int(coalesce({it('SLAHours')},0)),if({biz},1,0))")),
        ("Update_H", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/BusinessHoursOpen": "@outputs('H')", "item/SLAHours": "@outputs('S')"})),
        ("Unassigned_at_4_hours", cond(f"@and({biz},empty(coalesce({it('AssignedTo')}?['Email'],'')),equals(outputs('S'),4))", {
            "Post_unassigned_4h": channel(f"<p>IT-@{{{ID}}} has had no technician for 4 hours: @{{{it('Title')}}}</p><p>{lnk}</p>")})),
        ("Unassigned_at_5_hours", cond(f"@and({biz},empty(coalesce({it('AssignedTo')}?['Email'],'')),equals(outputs('S'),5))", {
            "Tell_me_unassigned_5h": chat(ME, f"<p>IT-@{{{ID}}} is still unassigned after 5 hours: @{{{it('Title')}}}</p><p>{lnk}</p>")})),
        ("Open_a_full_business_day", cond(f"@and({biz},greaterOrEquals(outputs('S'),{DAY_H}),equals(mod(outputs('S'),{DAY_H}),0))", {
            "Tell_tech_open_days": chat(f"@{{coalesce({it('AssignedTo')}?['Email'],'{ME}')}}",
                f"<p>IT-@{{{ID}}} has been open @{{div(outputs('S'),{DAY_H})}} day(s): @{{{it('Title')}}}, requested by @{{{req}}}</p><p>{lnk}</p>")})),
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
                        f"<pre>@{{substring(coalesce({it('Conversation')},''),0,min(length(coalesce({it('Conversation')},'')),1000))}}</pre>"
                        f"<p>Reply to this email to answer. The ticket closes after 5 days without a reply.</p><p>{rlnk}</p>")),
                    ("Chat_reminder", {"type": "Scope", "runAfter": ra("Email_reminder", status=("Succeeded", "Failed")),
                        "actions": support_msg("Post_reminder", f"@{req}",
                            f"<p><b>Reminder: IT is waiting for your reply on IT-@{{{ID}}}</b></p><p>@{{{it('Title')}}}</p>"
                            "<p>The ticket closes after 5 days without a reply.</p>" + support_links(ID, it('Title')))}),
                ]))),
            ]))),
    ])
    resolved_case = seq([
        ("ResolvedAge", compose(f"@div(sub(ticks(utcNow()),ticks(coalesce({it('ResolvedOn')},utcNow()))),36000000000)")),
        ("Window_48h", cond("@greaterOrEquals(outputs('ResolvedAge'),48)", {
            "Close_resolved": sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/Status/Value": S_CLOSED})},
            {"Reminder_24h": cond(f"@and(greaterOrEquals(outputs('ResolvedAge'),24),lessOrEquals(int(coalesce({it('ResolvedHours')},0)),0),not(empty({req})))", seq([
                ("Mark_reminded", sp("PatchItem", {"id": f"@{ID}", "item/Title": f"@{it('Title')}", "item/ResolvedHours": 1})),
                ("Email_resolved_reminder", mail(f"@{req}", f"[IT-@{{{ID}}}] Reminder: is your issue fixed? @{{{it('Title')}}}",
                    f"<p>{MARKER}</p><p>IT marked ticket IT-@{{{ID}}} as resolved: @{{{it('Title')}}}</p>"
                    f"<p><b>Resolution:</b> @{{coalesce({it('Resolution')},'')}}</p>"
                    "<p>If it is fixed, you don't need to do anything. If it is not fixed, reply to this email within the next day and your ticket will be reopened. Otherwise it closes automatically 48 hours after it was resolved.</p>"
                    f"<p>{rlnk}</p>")),
                ("Chat_resolved_reminder", {"type": "Scope", "runAfter": ra("Email_resolved_reminder", status=("Succeeded", "Failed")),
                    "actions": support_msg("Post_resolved_reminder", f"@{req}",
                        f"<p><b>Reminder: is IT-@{{{ID}}} fixed?</b></p><p>@{{{it('Title')}}}</p>"
                        "<p>If it is fixed, no action is needed. If not, use Reply by email within the next day to reopen it. The ticket closes automatically 48 hours after it was resolved.</p>" + support_links(ID, it('Title')))}),
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
            "Overdue": "@length(body('Filter_this_tech_overdue'))"}}}),
    ])
    by_cat_loop = {"type": "Foreach", "foreach": "@outputs('Categories')", "actions": seq([
        ("Filter_this_category", {"type": "Query", "inputs": {"from": "@body('Opened')?['value']",
            "where": "@equals(item()?['Category']?['Value'],items('Apply_to_each_category'))"}}),
        ("Append_category", {"type": "AppendToArrayVariable", "inputs": {"name": "varByCategory", "value": {
            "Category": "@items('Apply_to_each_category')", "Tickets": "@length(body('Filter_this_category'))"}}}),
    ])}
    fname = "HelpDesk-@{" + nowET('yyyy-MM-dd') + "}"
    report = (
        "<html><body style=\"font-family:Arial,sans-serif\">"
        "<h2>IT help desk daily report, @{" + nowET('dddd MMMM d, yyyy') + "}</h2>"
        "<table border=\"1\" cellpadding=\"6\" style=\"border-collapse:collapse\"><tr><th>Opened today</th><th>Closed today</th><th>Open now</th><th>Average time to close (hours)</th></tr>"
        "<tr><td>@{length(body('Opened')?['value'])}</td><td>@{length(body('ResolvedToday')?['value'])}</td><td>@{length(body('OpenNow')?['value'])}</td>"
        "<td>@{formatNumber(outputs('Average'),'0.0')}</td></tr></table>"
        "<h3>Needs attention: overdue</h3>" + or_none("overdue", "Nothing overdue") +
        "<h3>Needs attention: unassigned</h3>" + or_none("unassigned", "Nothing unassigned") +
        "<h3>Waiting on users</h3>" + or_none("waiting", "Nobody is waiting") +
        "<h3>Open tickets by technician</h3>@{body('Table_by_tech')}"
        "<h3>Today's tickets by category</h3>@{body('Table_by_category')}"
        "</body></html>")
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
         ("Email_me_the_report", {**mail(ME, "Daily help desk report, @{" + nowET('dddd MMMM d') + "}", "@{outputs('ReportHtml')}",
             attachments=[{"Name": fname + ".pdf", "ContentBytes": "@body('Convert_file')"}]),
             "runAfter": ra("Create_file_Reports", status=("Succeeded", "Failed"))}),
         ])
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
            "Case_status": {"case": "status", "actions": seq([
                ("Set_status", sp("PatchItem", {
                    "id": "@outputs('TicketId')", "item/Title": "@body('Get_item')?['Title']",
                    "item/Status/Value": "@outputs('StatusChoice')",
                    "item/AssignedTo/Claims": "@coalesce(body('Get_item')?['AssignedTo']?['Email'],outputs('Responder'))",
                    "item/NoteToUser": f"@if(and(contains({arr([S_WAIT, S_SCHED])},outputs('StatusChoice')),not(empty(outputs('Text')))),outputs('Text'),null)",
                    "item/Resolution": f"@if(equals(outputs('StatusChoice'),'{S_RES}'),outputs('Text'),body('Get_item')?['Resolution'])"})),
                ("Internal_note", cond(f"@and(equals(outputs('StatusChoice'),'{S_ESC}'),not(empty(outputs('Text'))),not(empty(coalesce(body('Get_item')?['ChannelMessageId'],''))))", {
                    "Post_internal_note": channel_reply("@body('Get_item')?['ChannelMessageId']",
                        "<p><b>Internal note</b> (not sent to the user): @{outputs('Text')}</p>")})),
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
        {"type": "TextBlock", "text": "Answer in the box above and select Send answer. Messages typed in this chat aren't seen by IT.", "wrap": True, "size": "Small", "isSubtle": True}],
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
            "Tell_the_tech": chat(f"@{{{tk}?['AssignedTo']?['Email']}}",
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
    person = seq([
        ("Get_sender", op("shared_office365users", "UserProfile_V2", {"id": f"@{{{M}?['from']?['user']?['id']}}"})),
        ("SenderEmail", compose(f"@{sender}")),
        ("Not_IT_Support", cond(f"@not(equals(outputs('SenderEmail'),toLower('{SUPPORT}')))", seq([
            ("Msg_to_text", op("shared_conversionservice", "HtmlToText", {"Content": f"@{{coalesce({M}?['body']?['content'],'')}}"})),
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
                ("PendingFresh", compose("@if(and(not(empty(body('Get_pending')?['value'])),less(sub(ticks(utcNow()),ticks(first(body('Get_pending')?['value'])?['Created'])),864000000000)),first(body('Get_pending')?['value']),null)")),
                ("IsBare", compose("@and(not(empty(outputs('Ticket'))),not(empty(outputs('PendingFresh'))),equals(replace(replace(replace(toUpper(outputs('ReplyText')),' ',''),'#',''),'.',''),concat('IT-',string(outputs('Ticket')?['ID']))))")),
                ("FinalText", compose("@if(outputs('IsBare'),outputs('PendingFresh')?['ReplyText'],outputs('ReplyText'))")),
                ("Ticket_found", cond(f"@not(empty({tk}))", found, not_found)),
            ]))),
        ]))),
    ])
    actions = seq([
        ("Msg", compose("@coalesce(triggerBody()?['value']?[0],triggerBody())")),
        ("From_a_person", cond(f"@not(empty(coalesce({M}?['from']?['user']?['id'],'')))", person)),
    ])
    trig = {"When_someone_messages_IT_Support": {"type": "OpenApiConnectionWebhook",
        "inputs": {"parameters": {"threadType": "groupchat"},
                   "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_teams", "connectionName": "shared_teams_1",
                            "operationId": "WebhookNewMessageTrigger"},
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
                "definition": build_definition(name, trig, acts),
                "connectionReferences": {a: {"connectionName": APIS[a][2], "source": "Embedded",
                    "id": f"/providers/Microsoft.PowerApps/apis/shared_{APIS[a][0]}", "tier": "NotSpecified", "apiName": APIS[a][0],
                    "isProcessSimpleApiReferenceConversionAlreadyDone": False} for a in sorted(apis)},
                "flowFailureAlertSubscribed": False, "isManaged": False}}
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
    flows = [hd1(), hd2(), hd3(), hd4(), hd5(), hd6(), hd7(), hd8(), hd9(), hd11(), hd12(), hd13(), hd14()]   # HD10 retired: questions now go out as IT Support messages and replies come back via HD12
    for f in flows:
        safe = f[0].split()[0]
        write_package([f], os.path.join(OUT, f"{safe}.zip"), f[0])
    write_package(flows, os.path.join(OUT, "HelpDesk-all-flows.zip"), "IT Help Desk flows")
    # dump definitions for review
    for f in flows:
        json.dump(build_definition(*f), open(os.path.join(OUT, f[0].split()[0] + ".json"), "w"), indent=1)
    print("ok", [f[0] for f in flows])
