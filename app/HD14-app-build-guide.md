# Branthaven "IT Help Desk" app (Power Apps canvas app, runs in Teams)

The app is the front end. Flow HD14 creates the ticket (same fields as HD2), then HD3 sends the confirmation email and the IT Support chat message as usual. Requesters never get edit access to the Tickets list.

Live app: **IT Help Desk** (environment `branthavenhomes (default)`, App ID `3aadc4f4-b800-40d6-9aa1-0a14a8335fc2`). Last visual redesign: 21 Oct 2026 — matches `app-mockup.png`.

## 1. Import the flow
1. make.powerautomate.com > My flows > Import > Import Package (Legacy) > upload HD14-app-intake.zip, Create as new, map SharePoint to your connection.
2. Open HD14 > Edit > confirm the trigger is "PowerApps (V2)" and that Create_item and Respond_to_the_app have no red warnings. Save.
3. On the HD14 details page, under "Run only users", Edit: set the SharePoint connection to "Use this connection (IT_Support@branthaven.com)". This is what keeps requesters from needing access to the list.

**If tickets save to SharePoint but no Teams post / email goes out**, or the app shows `HD14Teamsappintake.Run failed … 502 BadGateway`: check make.powerautomate.com > Connections. A tenant sign-in policy change (e.g. MFA, error AADSTS50079) silently breaks the IT_Support connections (SharePoint, Office 365 Outlook, Office 365 Users, Planner, Teams). Click **Reconnect** on each red one and sign in as IT_Support. Happened 21 Oct 2026.

## 2. Create the app
make.powerapps.com > Create > Blank app > Blank canvas app > name "IT Help Desk", Tablet format.

The app is **one screen** (`Screen1`). The form and the confirmation share the screen; a global variable `varTicket` (the HD14 response) decides which is showing. Everything is positioned by formula so it reflows in Teams.

### App.OnStart (theme)
```
Set(varDark, Lower(Param("theme")) = "dark");
Set(varBg,      If(varDark, RGBA(18,18,20,1),  RGBA(237,238,245,1)));
Set(varCardBg,  If(varDark, RGBA(30,30,33,1),  Color.White));
Set(varText,    If(varDark, Color.White,       RGBA(26,26,26,1)));
Set(varSubtext, If(varDark, RGBA(170,170,170,1), RGBA(110,110,120,1)));
Set(varBorder,  If(varDark, RGBA(55,55,60,1),  RGBA(225,226,235,1)));
Set(varSegBg,   If(varDark, Color.White,       RGBA(26,26,26,1)));
Set(varSegText, If(varDark, RGBA(26,26,26,1),  Color.White));
Set(varCategory, "Other"); Set(varUrgency, "It's slowing me down");
```
The header is always dark regardless of theme. The confirmation panel (HtmlText2) uses fixed light colours.

### App.Formulas (named formulas — responsive column)
```
colW = Min(App.Width, Max(710, App.Width - 64), 1500);
colX = (App.Width - colW) / 2;
isPhone = App.Width < 520;
```

### Layout containers (Screen1)
| Control | Type | Key properties |
|---|---|---|
| `Screen1` | Screen | Fill `varBg` |
| `recHeader` | Rectangle | X 0, Y 0, Width `App.Width`, Fill `RGBA(26,26,26,1)` (always dark), Height `Max(120, lblSub.Y+lblSub.Height+16)` |
| `lblTitle` | Label | "How can IT help?", Color `Color.White`, Size 28, X `colX+40`, Width `If(isPhone, colW-70, colW-80-HtmlText1_1.Width)` |
| `lblSub` | Label | "Tell us what's going on. You'll get a confirmation email and an update in Teams.", Color `RGBA(200,200,205,1)`, AutoHeight true, same X/Width as lblTitle |
| `Image1` | Image | **The Branthaven wordmark.** Image = data URI of `LogoWhite.svg` (see below). Width 220, Height 20, ImagePosition Fit, X `txtWhat.X+txtWhat.Width-Self.Width`, Y `lblTitle.Y+lblTitle.Height/2-Self.Height/2` |
| `HtmlText1_1` | HTML text | Old text logo — kept hidden (`Visible false`) because lblTitle/lblSub widths reference it |
| `recCard` | Rectangle | White card. Fill `varCardBg`, BorderThickness 0, Y `lblWhat.Y-24`, X `App.Width/2-Self.Width/2`, Width `If(IsBlank(varTicket), txtWhat.Width+48, Min(txtWhat.Width+48, 760))`, Height `If(IsBlank(varTicket), btnSubmit.Y+btnSubmit.Height+24-Self.Y, btnAnother.Y+btnAnother.Height+40-Self.Y)` — Send to back |
| `recDone` | Rectangle | Confirmation cover. Same X/Y/Width/Height as recCard, Fill `varCardBg`, BorderThickness 0, Visible `!IsBlank(varTicket)`. Z-order: above every form control, below HtmlText2 and the three confirmation buttons |
| `HtmlText1` | HTML text | Footer. X 0, Width `App.Width`, Y `recCard.Y+recCard.Height`, Height 60, all Padding 0. HtmlText: `"<div style='text-align:center;font-family:Open Sans,sans-serif;font-size:12px;color:#8a8a8a;padding:10px 0;'>Developed by <b style='color:#1a1a1a;letter-spacing:0.5px;'>BRANTHAVEN</b></div>"` |
| `Rectangle1` | Rectangle | Legacy fixed panel — `Visible false`, leave in place |

**Logo image:** build the data URI from `LogoWhite.svg` — `"data:image/svg+xml;charset=utf8," & <URL-encoded SVG>` and paste it as the Image property. (Uploading the SVG as app media works too; the data URI just avoids the media upload step.) The site's wordmark is outlined artwork, not a font, so don't try to recreate it with text.

### Form controls
All form controls have `Visible: IsBlank(varTicket)` so they disappear on confirmation. Field labels are **Size 10, FontWeight Bold, uppercase text**, Color `varText`. Inputs have `BorderStyle.Solid`, BorderThickness 1, BorderColor `varBorder`. Width `txtWhat.Width` / X `txtWhat.X` for everything below the summary box.

- `lblWhat` "WHAT DO YOU NEED HELP WITH? *" + Text input `txtWhat`: HintText "A short summary", **DelayOutput false** (must stay false for live category detection)
- `lblCat` "CATEGORY *" + Dropdown `ddCategory`, ChevronBackground `varCardBg`, ChevronFill `varSubtext`
  - Items: `["Access Request", "Application", "Bookings", "BuilderGo/OnLocation Warranty", "Email", "Horizon", "InFormXL", "Landline", "Laptop / Desktop", "Mobile (iPad / Cell)", "Mouse / Keyboard / Monitor", "New User Setup", "NewStar", "OneDrive / SharePoint", "Other", "Password or Sign-In", "Printer", "SaleFish", "SaleXpress", "Teams", "VPN"]` — must match the Category choices on the Tickets list exactly.
  - **Default:** the keyword formula in `ddCategory-Default.fx`. It reads `txtWhat.Text & txtDetails.Text` on every keystroke (app names first, then generic IT terms, falls back to Other). The user can still change it. Do not replace this with an OnChange/variable version — that loses as-you-type detection.
- `Label2` — the **"auto-detected" badge** next to the category label: Text "auto-detected", Size 9, Fill `If(varDark, RGBA(40,60,90,1), RGBA(232,240,254,1))`, Color blue, Width 96, Height 20, X `lblCat.X+90`, Y centred on lblCat, Visible `IsBlank(varTicket) && ddCategory.Selected.Value <> "Other" && Len(txtWhat.Text & txtDetails.Text) > 0`
- `lblUrg` "HOW URGENT IS IT?" + segmented control of three Buttons `Button1`, `Button1_1`, `Button1_2` (texts "I can't work at all", "It's slowing me down", "No rush"). Each: Width `txtWhat.Width/3`, X chained (`Button1.X+Button1.Width`, etc.), BorderStyle Solid, BorderColor `varBorder`, outer corners radius 10 and inner corners 0 so they read as one strip. Fill `If(varUrgency = <its text>, varSegBg, varCardBg)`, Color likewise; OnSelect `Set(varUrgency, "<its text>")`. (The original `rdUrgency` Radio still exists hidden; HD14 is passed `varUrgency`.)
- `lblDet` "DETAILS *" + Text input `txtDetails`: Mode MultiLine, HintText "What happened, and when did it start?"
- `AddMediaWithImage1` (Add picture): the screenshot drop-zone, dashed border, X `txtWhat.X`, Width `txtWhat.Width`. Its button `AddMediaButton1` has Height `If(IsBlank(Self.Media), 90, 240)` so the box opens up once a picture is attached; child `UploadedImage1` same X/Width, ImagePosition **Fit** (never Fill — it stretched screenshots into a strip), Y `AddMediaButton1.Y + 36`, Height `If(IsBlank(AddMediaButton1.Media), 48, 196)`. Everything below chains off it so it moves down with the picture: `Label1.Y = AddMediaButton1.Y + AddMediaButton1.Height + 10`, `lblOnBehalf.Y = Label1.Y + 40`, `txtOnBehalf.Y = lblOnBehalf.Y + 24`, `btnSubmit.Y = txtOnBehalf.Y + 54`. `Label1`: "One picture only. To send more, reply to your confirmation email.", Color `varSubtext`, Visible `IsBlank(varTicket) && Parent.Width >= 520`
- `lblOnBehalf` "ON BEHALF OF ANOTHER USER (OPTIONAL)" + Text input `txtOnBehalf`, HintText "Colleague's email, e.g. jane@branthaven.com (leave blank for yourself)"
- Button `btnSubmit` "Send to IT": Fill `RGBA(26,26,26,1)`, Color White, radius 8
  - DisplayMode: `If(IsBlank(txtWhat.Text) || IsBlank(ddCategory.Selected.Value) || IsBlank(txtDetails.Text), DisplayMode.Disabled, DisplayMode.Edit)`
  - OnSelect: `Set(varTicket, HD14Teamsappintake.Run(txtWhat.Text, ddCategory.Selected.Value, varUrgency, txtDetails.Text, txtOnBehalf.Text, <screenshot>))` (see the flow for the exact parameter list)

Add the flow first: Power Automate icon (left bar) > Add flow > HD14 Teams app intake. The formula name is the flow name without spaces.

### Confirmation (same screen, shown when `varTicket` is set)
- `HtmlText2` — the whole confirmation panel in one HTML control. X `recCard.X`, Width `recCard.Width`, Y `recCard.Y+56`, Height 350, Visible `!IsBlank(varTicket)`. Content, centred:
  1. green check in a pale-green circle (56px, `#dcfce7` / `#16a34a`)
  2. `"We've got it, " & First(Split(User().FullName, " ")).Value`
  3. black pill with `"IT-" & varTicket.ticketid`
  4. a small uppercase label **TICKET DETAILS** (10px bold, letter-spacing 1px, `#1a1a1a`), then `txtWhat.Text` in grey on the next line (`<span …>TICKET DETAILS</span><br>` inside the same grey div, margin-top 14px)
  5. grey box (`#f6f6f8`, border `#e1e2eb`, radius 10, max-width 460) with three numbered steps:
     1. A confirmation is on its way to you by email and Teams.
     2. IT Support will pick it up shortly — you'll see every status change in Teams and email.
     3. Need to add something? Reply to the email or the IT Support chat.
- Three outlined buttons, each Width 188, Height 40, Size 12, Fill `varCardBg`, Color `varText`, BorderColor `varBorder` (Report another issue uses `varText`), HoverFill `RGBA(240,240,243,1)`, Y `HtmlText2.Y+HtmlText2.Height+4`, Visible `!IsBlank(varTicket)`; X measured from the card centre `recCard.X+recCard.Width/2`:
  - `btnAnother_1` "View this ticket" — centre − 294 (opens the ticket as before)
  - `btnAnother` "Report another issue" — centre − 196. OnSelect: `Set(varSuggestedCategory, ""); Reset(txtWhat); Reset(ddCategory); Reset(rdUrgency); Reset(txtDetails); Reset(txtOnBehalf); Reset(AddMediaButton1); Set(varShot, Blank()); Set(varTicket, Blank())`
  - `btnAnother_2` "My open tickets ↗" — centre + 106 (same action as the old "My Tickets")
- Hidden leftovers from the first confirmation design, keep `Visible false`: `lblDoneTitle`, `lblDoneMsg`, `lblTicketBadge`, `Icon1`.

Because the card's height shrinks to the buttons on confirmation and every form control hides, the screen no longer scrolls, so Teams snaps back to the top.

## 3. Add to Teams
1. Save and Publish the app.
2. App details > Add to Teams > Add to Teams (or "Add to a team"). Test it yourself in Teams first.
3. When ready for rollout: share the app with all employees (Share > the Branthaven group), then in Teams admin center pin it for everyone (Teams apps > Setup policies > Pinned apps).
4. After publishing, Teams users must fully quit and reopen Teams to pick up the new version.

## Notes
- The requester identity comes from their own sign-in (not typed), so tickets can't be raised in someone else's name — except via the "On behalf of" field, which is for IT and managers logging for a colleague.
- Studio gotcha: the tree-view search box keeps stale characters; if a formula edit lands on `Screen1.BackgroundImage` instead of the control, clear the search box (X) before selecting the control. Check App checker > Formulas = 0 errors before each publish.
