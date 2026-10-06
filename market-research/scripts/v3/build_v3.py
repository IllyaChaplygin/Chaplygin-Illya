import sys
sys.path.insert(0, '.')
exec(open('lib.py').read())
exec(open('data.py').read().replace("if __name__ == '__main__':", "if False:"))
exec(open('s_customs.py').read())
exec(open('s_formats.py').read())
exec(open('s_charts.py').read())
exec(open('s_rest.py').read())

orig = slide_ids()          # 23 слайдів користувача
o_cover, o_retail, o_chan = orig[0], orig[19], orig[20]
keep = {id(o_cover), id(o_retail), id(o_chan)}

def last(): return slide_ids()[-1]
new = {}
def mk(name, fn, *a):
    fn(*a); new[name] = last()

mk('kpi', s_kpi); mk('comp', s_composition); mk('dyn', s_dynamics); mk('imp', s_importers)
mk('over', s_overview)
for f in FMT_ORDER: mk('fmt_' + f, s_format, f)
mk('heat', s_heat); mk('top', s_top13); mk('lad1', s_ladder1); mk('lad2', s_ladder2)
mk('stand', s_stand); mk('fin', s_fin); mk('tbl', s_price_table)
before = len(slide_ids()); cat_slides(); cats = slide_ids()[before:]

# канали: замінити зображення-чарт на нативний
chan_slide = [s for s in prs.slides if s.slide_id == o_chan.id][0]
channels_chart(chan_slide)

order = [o_cover, new['kpi'], new['comp'], new['dyn'], new['imp'], new['over']] + \
        [new['fmt_' + f] for f in FMT_ORDER] + \
        [new['heat'], new['top'], new['lad1'], new['lad2'], o_retail, o_chan, new['stand'], new['fin'], new['tbl']] + cats
for el in orig:
    if id(el) not in keep and el not in order:
        drop_slide(el)
reorder(order)
renumber()
OUT = '/home/user/Chaplygin-Illya/market-research/Gotovyi_Rys_Final_v2.pptx'
prs.save(OUT); print('saved', len(prs.slides._sldIdLst), 'slides')
