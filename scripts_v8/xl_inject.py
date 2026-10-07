import zipfile,sys,re,shutil,os
from lxml import etree
sys.path.insert(0,os.path.dirname(__file__))
from xl_eval import evaluate
src,dst=sys.argv[1],sys.argv[2]
vals=evaluate(src)
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
zin=zipfile.ZipFile(src)
wbx=etree.fromstring(zin.read('xl/workbook.xml'))
rels=etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
rid2t={r.get('Id'):r.get('Target') for r in rels}
sheetfile={}
for s in wbx.find('m:sheets',NS):
    t=rid2t[s.get('{%s}id'%NS['r'])]; t=t.lstrip('/'); t=t if t.startswith('xl/') else 'xl/'+t
    sheetfile[t]=s.get('name')
zout=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
miss=[]; n=0
for item in zin.infolist():
    data=zin.read(item.filename)
    if item.filename in sheetfile:
        name=sheetfile[item.filename]; root=etree.fromstring(data)
        for c in root.iter('{%s}c'%NS['m']):
            f=c.find('m:f',NS)
            if f is None: continue
            v=vals.get((name.upper(),c.get('r')))
            old=c.find('m:v',NS)
            if old is not None: c.remove(old)
            if v is None or (isinstance(v,str) and v.startswith('#')) or type(v).__name__=='XlError':
                miss.append((name,c.get('r'),str(v)[:30])); continue
            ve=etree.SubElement(c,'{%s}v'%NS['m'])
            if isinstance(v,bool): c.set('t','b'); ve.text='1' if v else '0'
            elif isinstance(v,(int,float)) or hasattr(v,'item'):
                try: ve.text=repr(float(v)) if not float(v).is_integer() else str(int(float(v)))
                except Exception: miss.append((name,c.get('r'),str(v)[:30])); c.remove(ve); continue
                if 't' in c.attrib: del c.attrib['t']
            else: c.set('t','str'); ve.text=str(v)
            n+=1
        data=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    zout.writestr(item,data)
zout.close(); print("injected",n,"missing",len(miss),miss[:12])
