# Working rules for this repo

- **Read the matching .md before changing anything.** The guides are the source of truth: `app/HD14-app-build-guide.md` for the Power Apps canvas app, `flows/` and `docs/IT-Tech-Guide.md` for the Power Automate flows, `docs/End-User-Guide.md` for what users are told. If the live app or flow disagrees with the guide, say so before changing either.
- **Update the guide in the same change.** Any edit to the app, a flow, the category list, or the notifications must land in the relevant .md in the same commit.
- Don't replace a documented formula with a "simpler" one (e.g. `ddCategory.Default` must stay the live keyword formula in `app/ddCategory-Default.fx`).
- Category names must match in three places: the Tickets list choice column, `ddCategory.Items`, and `generator/gen.py`.
