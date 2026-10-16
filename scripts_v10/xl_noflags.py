import zipfile,sys,re,shutil,os
from lxml import etree
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
AFTER=['smartTags','drawing','legacyDrawing','legacyDrawingHF','drawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']
def fix(path):
    tmp=path+'.tmp'; zin=zipfile.ZipFile(path); zout=zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED); n=0
    for it in zin.infolist():
        data=zin.read(it.filename)
        if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',it.filename):
            root=etree.fromstring(data)
            for old in root.findall(f'{{{NS}}}ignoredErrors'): root.remove(old)
            dim=root.find(f'{{{NS}}}dimension'); ref=dim.get('ref') if dim is not None else 'A1:AZ200'
            last=ref.split(':')[-1]; m=re.match(r'([A-Z]+)(\d+)',last); sq=f"A1:{m.group(1)}{int(m.group(2))+20}" if m else 'A1:AZ200'
            ie=etree.Element(f'{{{NS}}}ignoredErrors'); e=etree.SubElement(ie,f'{{{NS}}}ignoredError'); e.set('sqref',sq)
            for a in ('evalError','formula','formulaRange','unlockedFormula','numberStoredAsText','emptyCellReference','twoDigitTextYear','listDataValidation','calculatedColumn'): e.set(a,'1')
            pos=None
            for i,ch in enumerate(root):
                if etree.QName(ch).localname in AFTER: pos=i; break
            if pos is None: root.append(ie)
            else: root.insert(pos,ie)
            data=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True); n+=1
        zout.writestr(it,data)
    zout.close(); zin.close(); shutil.move(tmp,path); return n
if __name__=='__main__':
    for p in sys.argv[1:]: print(p.split('/')[-1],'sheets patched:',fix(p))
