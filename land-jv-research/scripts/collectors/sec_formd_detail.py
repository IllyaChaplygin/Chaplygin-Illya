import json, re, sys, time
sys.path.insert(0,'.')
from formd import get

raw = json.load(open('snap1.json'))
try: done = json.load(open('stage2.json'))
except: done = {}

def parse(cik, acc):
    a = acc.replace('-','')
    url = f'https://www.sec.gov/Archives/edgar/data/{int(cik)}/{a}/primary_doc.xml'
    x = get(url)
    if not x: return None
    def g(t, s=None):
        m = re.search(rf'<{t}>(.*?)</{t}>', s if s else x, re.S)
        return m.group(1).strip() if m else None
    exm = re.search(r'<federalExemptionsExclusions>(.*?)</federalExemptionsExclusions>', x, re.S)
    exemptions = re.findall(r'<item>([^<]+)</item>', exm.group(1)) if exm else []
    ind = g('industryGroupType')
    return {
      'entity': g('entityName'), 'industry': ind,
      'offering': g('totalOfferingAmount'), 'sold': g('totalAmountSold'),
      'minInvest': g('minimumInvestmentAccepted'),
      'investors': g('totalNumberAlreadyInvested'),
      'exemptions': exemptions,
      'is506c': '06c' in exemptions, 'is506b': '06b' in exemptions,
      'url': url,
    }

SINCE = '2019-01-01'
for label, rec in raw.items():
    if label in done: continue
    rows=[]
    fl = [f for f in rec['filings'] if f['date'] >= SINCE]
    for f in fl:
        p = parse(f['cik'], f['acc'])
        time.sleep(0.16)
        if p: p['date']=f['date']; p['form']=f['form']; rows.append(p)
    done[label]=rows
    json.dump(done, open('stage2.json','w'))
    nc = sum(1 for r in rows if r['is506c']); nb = sum(1 for r in rows if r['is506b'])
    print(f"{label}: parsed={len(rows)} 506c={nc} 506b={nb}", flush=True)
print("STAGE2 DONE")
