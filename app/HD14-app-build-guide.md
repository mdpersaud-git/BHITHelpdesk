# Branthaven "Report an IT issue" app (Power Apps, runs in Teams)

The app is the front end. Flow HD14 creates the ticket (same fields as HD2), then HD3 sends the confirmation email and the IT Support chat message as usual. Requesters never get edit access to the Tickets list.

## 1. Import the flow
1. make.powerautomate.com > My flows > Import > Import Package (Legacy) > upload HD14-app-intake.zip, Create as new, map SharePoint to your connection.
2. Open HD14 > Edit > confirm the trigger is "PowerApps (V2)" and that Create_item and Respond_to_the_app have no red warnings. Save.
3. On the HD14 details page, under "Run only users", Edit: set the SharePoint connection to "Use this connection (mpersaud@branthaven.com)". This is what keeps requesters from needing access to the list.

## 2. Create the app
make.powerapps.com > Create > Blank app > Blank canvas app > name "Report an IT issue", Tablet format.

Build one form screen (scrForm) and one thank-you screen (scrDone).

### scrForm controls (names matter for the formulas)
- Rectangle `recHeader`: X 0, Y 0, Width Parent.Width, Height 120, Fill RGBA(0,0,0,1)
- Image `imgLogo`: upload the Branthaven BH logo; X 30, Y 20, Width 80, Height 80
- Label `lblTitle`: Text "How can IT help?", Color White, Size 28, X 130, Y 30
- Label `lblSub`: Text "Tell us what's going on. You'll get a confirmation email and an update in Teams.", Color RGBA(255,255,255,0.8), X 130, Y 72
- Label + Text input `txtWhat`: label "What do you need help with?"; HintText "A short summary"
- Label + Dropdown `ddCategory`: label "Category"
  Items: ["Access Request", "Application", "Bookings", "BuilderGo/OnLocation Warranty", "Email", "Horizon", "InFormXL", "Landline", "Laptop / Desktop", "Mobile (iPad / Cell)", "Mouse / Keyboard / Monitor", "New User Setup", "NewStar", "Other", "Password or Sign-In", "Printer", "SaleFish", "SaleXpress", "Teams", "VPN"]
  Default: keyword auto-select from the summary + details — formula in `ddCategory-Default.fx` (app names first, then generic IT terms; falls back to Other). The user can still change it.
  (Must match the Category choices on the Tickets list exactly. Check the list and add/remove to match, e.g. Access Request, Application.)
- Label + Radio `rdUrgency`: label "How urgent is it?"
  Items: ["I can't work at all","It's slowing me down","No rush"]; Default "It's slowing me down"
- Label + Text input `txtDetails`: label "Details"; Mode TextMode.MultiLine; Height 160
- Button `btnSubmit`: Text "Send to IT", Fill RGBA(0,0,0,1), Color White, RadiusTopLeft/TopRight/BottomLeft/BottomRight 8
  - DisplayMode:
    If(IsBlank(txtWhat.Text) || IsBlank(ddCategory.Selected.Value) || IsBlank(txtDetails.Text), DisplayMode.Disabled, DisplayMode.Edit)
  - OnSelect:
    Set(varTicket, HD14Teamsappintake.Run(txtWhat.Text, ddCategory.Selected.Value, rdUrgency.Selected.Value, txtDetails.Text));
    Navigate(scrDone)

Add the flow first: Power Automate icon (left bar) > Add flow > HD14 Teams app intake. The formula name is the flow name without spaces.

### scrDone controls
- Same black header rectangle and logo
- Label `lblDone`: Text "Ticket IT-" & varTicket.ticketid & " created"; Size 26
- Label: "A confirmation email is on its way, and IT Support will message you in Teams as it moves along. To add details, reply to that email."
- Button "Report another issue": OnSelect Reset(txtWhat); Reset(ddCategory); Reset(rdUrgency); Reset(txtDetails); Navigate(scrForm)

## 3. Add to Teams
1. Save and Publish the app.
2. App details > Add to Teams > Add to Teams (or "Add to a team"). Test it yourself in Teams first.
3. When ready for rollout: share the app with all employees (Share > the Branthaven group), then in Teams admin center pin it for everyone (Teams apps > Setup policies > Pinned apps).

## Notes
- Screenshots can't be attached from the app (no file upload in this build). Tell requesters to reply to the confirmation email with a screenshot; HD1 attaches it.
- The requester identity comes from their own sign-in (not typed), so tickets can't be raised in someone else's name.
