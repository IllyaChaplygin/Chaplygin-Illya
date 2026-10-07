import formulas,os,sys,re,openpyxl,warnings
warnings.filterwarnings("ignore")
def evaluate(path):
    xl=formulas.ExcelModel().loads(path).finish(); sol=xl.calculate()
    out={}
    base=os.path.basename(path).upper()
    for k,v in sol.items():
        m=re.match(r"^'\[(.+?)\](.+)'!([A-Z]+\d+)$",k)
        if not m or m.group(1).upper()!=base: continue
        val=v.value
        try: val=val[0][0]
        except Exception: pass
        out[(m.group(2).upper(),m.group(3))]=val
    return out
if __name__=="__main__":
    p=sys.argv[1]; vals=evaluate(p)
    wb=openpyxl.load_workbook(p)
    for n in wb.sheetnames[2:]:
        print('==',n)
        for co in ("F6","F7","F9","F10","F11","F12","F34","F35","F36","F37","F38","F39","F40","F41","F42","F43","F44","F45","F54","F57","F59","F60","F61","D71","G71","H71","I71","C77","E77","E79","G80","F89","B13"):
            print(co,vals.get((n.upper(),co)))
