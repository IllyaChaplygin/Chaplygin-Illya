import sys, json, time, re
sys.path.insert(0,'.')
from formd import get, companies, cik_name

ROOTS = json.load(open('roots.json'))
out = {}
for label, roots in ROOTS.items():
    ents = {}
    for r in roots:
        for cik, nm in companies(r):
            ents[cik] = nm
        time.sleep(0.4)
    rec = {'label': label, 'roots': roots, 'entities': len(ents), 'filings': []}
    for cik, nm in ents.items():
        s = get(f'https://data.sec.gov/submissions/CIK{cik}.json')
        time.sleep(0.25)
        if not s: continue
        try: j = json.loads(s)
        except: continue
        rc = j.get('filings',{}).get('recent',{})
        for form, acc, date in zip(rc.get('form',[]), rc.get('accessionNumber',[]), rc.get('filingDate',[])):
            if not form.startswith('D'): continue
            rec['filings'].append({'cik':cik,'entity':nm,'form':form,'acc':acc,'date':date})
    out[label] = rec
    print(f"{label}: entities={rec['entities']} formD={len(rec['filings'])}", flush=True)
    json.dump(out, open('harvest_raw.json','w'), indent=1)
print("DONE")
