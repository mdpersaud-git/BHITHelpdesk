// One-time backfill: give every existing ticket its own permissions
// (Owners = Full Control, requester = Read), the same as HD3 does for new tickets.
//
// How to run:
//   1. Open https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2 in Edge/Chrome, signed in as Martin.
//   2. Press F12 -> Console tab, paste this whole file, press Enter.
//   3. Wait for "DONE" and read the per-ticket lines. Safe to re-run.
//
// It does NOT touch the list-level permissions: do "Stop Inheriting Permissions" on the list first
// (see docs/IT-Tech-Guide.md -> "Who can see a ticket").

(async () => {
  const site = 'https://branthavenhomes.sharepoint.com/sites/ITHelpDeskv2';
  const LIST = `${site}/_api/web/lists(guid'22c09b0b-bbcd-42a9-9aeb-844d81089636')`;
  const json = { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' };
  const digest = (await (await fetch(`${site}/_api/contextinfo`, { method: 'POST', headers: json })).json()).FormDigestValue;
  const post = (url, body) => fetch(url, { method: 'POST', headers: { ...json, 'X-RequestDigest': digest }, body });
  const get = async (url) => (await fetch(url, { headers: json })).json();

  const owners = (await get(`${site}/_api/web/associatedownergroup?$select=Id`)).Id;
  const items = (await get(`${LIST}/items?$select=Id,RequesterEmail,Requester/EMail&$expand=Requester&$top=500`)).value;
  console.log(`${items.length} tickets, owners group ${owners}`);

  for (const it of items) {
    const email = (it.RequesterEmail || (it.Requester && it.Requester.EMail) || '').toLowerCase();
    const I = `${LIST}/items(${it.Id})`;
    let line = `IT-${it.Id}: `;
    let r = await post(`${I}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
    line += `break ${r.status}, `;
    r = await post(`${I}/roleassignments/addroleassignment(principalid=${owners},roledefid=1073741829)`);
    line += `owners ${r.status}, `;
    if (email.endsWith('@branthaven.com')) {
      const u = await (await post(`${site}/_api/web/ensureuser`, JSON.stringify({ logonName: `i:0#.f|membership|${email}` }))).json();
      r = await post(`${I}/roleassignments/addroleassignment(principalid=${u.Id},roledefid=1073741826)`);
      line += `requester ${email} read ${r.status}`;
    } else {
      line += `no internal requester (${email || 'blank'})`;
    }
    console.log(line);
  }
  console.log('DONE');
})();
