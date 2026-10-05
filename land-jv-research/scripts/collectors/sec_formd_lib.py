import json, re, sys, time, urllib.request, urllib.parse
UA = 'Chaplygin Research chaplygin.illya002@gmail.com'
def get(url, retries=3):
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept-Encoding':'gzip, deflate'})
            r = urllib.request.urlopen(req, timeout=40)
            data = r.read()
            if r.headers.get('Content-Encoding')=='gzip':
                import gzip; data = gzip.decompress(data)
            return data.decode('utf-8', 'replace')
        except Exception as e:
            if i==retries-1: return None
            time.sleep(2*(i+1))

def companies(name):
    """Find CIKs whose company name matches."""
    url = ('https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&company='
           + urllib.parse.quote(name) + '&type=D&dateb=&owner=include&count=100&output=atom')
    x = get(url)
    if not x: return []
    ciks = sorted(set(re.findall(r'<cik>(\d{10})</cik>', x, re.I)))
    res = []
    for c in ciks:
        res.append((c, cik_name(c)))
    return res

_NAME_CACHE = {}
def cik_name(cik):
    if cik in _NAME_CACHE: return _NAME_CACHE[cik]
    s = get(f'https://data.sec.gov/submissions/CIK{cik}.json')
    nm = None
    if s:
        try:
            import json as _j; nm = _j.loads(s).get('name')
        except Exception: pass
    _NAME_CACHE[cik] = nm or '?'
    return _NAME_CACHE[cik]

NUM = lambda s: int(s) if s and s.isdigit() else None
def formd_filings(cik, since='2023-01-01'):
    sub = get(f'https://data.sec.gov/submissions/CIK{cik}.json')
    if not sub: return []
    try: j = json.loads(sub)
    except: return []
    recent = j.get('filings',{}).get('recent',{})
    rows=[]
    for form, acc, date, pdoc in zip(recent.get('form',[]), recent.get('accessionNumber',[]),
                                     recent.get('filingDate',[]), recent.get('primaryDocument',[])):
        if not form.startswith('D'): continue
        if date < since: continue
        rows.append((form, acc, date, pdoc, j.get('name')))
    return rows

def parse_formd(cik, acc):
    a = acc.replace('-','')
    url = f'https://www.sec.gov/Archives/edgar/data/{int(cik)}/{a}/primary_doc.xml'
    x = get(url)
    if not x: return None
    g = lambda t: (re.search(rf'<{t}>(.*?)</{t}>', x, re.S).group(1).strip() if re.search(rf'<{t}>(.*?)</{t}>', x, re.S) else None)
    return {
      'entityName': g('entityName'),
      'industryGroup': g('industryGroupType'),
      'totalOfferingAmount': g('totalOfferingAmount'),
      'totalAmountSold': g('totalAmountSold'),
      'totalRemaining': g('totalRemaining'),
      'minimumInvestmentAccepted': g('minimumInvestmentAccepted'),
      'investorsAlready': g('totalNumberAlreadyInvested'),
      'hasNonAccredited': g('hasNonAccreditedInvestors'),
      'yearOfInc': g('value'),
      'url': url,
    }
