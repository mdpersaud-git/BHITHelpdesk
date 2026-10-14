import json,re,sys,glob
def walk(acts, names, errs, path):
    for n,a in acts.items():
        names.add(n)
        for r in (a.get("runAfter") or {}):
            if r not in acts: errs.append(f"{path}/{n}: runAfter missing {r}")
        for k in ("actions",):
            if isinstance(a.get(k),dict): walk(a[k],names,errs,path+"/"+n)
        if isinstance(a.get("else"),dict) and a["else"].get("actions"): walk(a["else"]["actions"],names,errs,path+"/"+n+"(else)")
        if a.get("type")=="Switch":
            for c,v in a["cases"].items(): walk(v["actions"],names,errs,path+"/"+n+"/"+c)
            walk(a["default"]["actions"],names,errs,path+"/"+n+"/default")
for f in sorted(glob.glob(sys.argv[1]+"/HD*.json")):
    d=json.load(open(f)); names=set(d["triggers"].keys()); errs=[]
    walk(d["actions"],names,errs,"")
    s=json.dumps(d)
    for m in set(re.findall(r"(?:body|outputs|items|variables)\('([^']+)'\)",s)):
        if m not in names and not m.startswith("var"): errs.append("ref missing "+m)
    print(f.split('/')[-1], "OK" if not errs else errs)
