import sys
sys.path.insert(0, '.')
exec(open('lib.py').read())
exec(open('data.py').read().replace("if __name__ == '__main__':", "if False:"))
exec(open('s_customs.py').read())
exec(open('s_formats.py').read())
exec(open('s_charts.py').read())
exec(open('s_rest.py').read())
exec(open('s_end.py').read())

orig = slide_ids()          # 23 слайди колоди користувача; лишаємо лише обкладинку
o_cover = orig[0]


def last(): return slide_ids()[-1]


order = [o_cover]


def mk(fn, *a):
    fn(*a); order.append(last())


mk(s_kpi); mk(s_composition); mk(s_dynamics); mk(s_importers); mk(s_overview)
for f in FMT_ORDER:
    mk(s_format, f)                       # слайд формату, одразу за ним — його каталог
    before = len(slide_ids()); cat_slides(f); order += slide_ids()[before:]
for fn in (s_heat, s_top13, s_ladder1, s_ladder2, s_ladder3, s_retail_matrix, s_retail_table, s_channels,
           s_fin, s_price_table, s_stand_pouch, s_stand_cup, s_stand_matrix, s_stand_neighbors):
    mk(fn)
for el in orig:
    if el is not o_cover:
        drop_slide(el)
reorder(order)
renumber()
OUT = '/home/user/Chaplygin-Illya/market-research/Gotovyi_Rys_Final_v2.pptx'
prs.save(OUT); print('saved', len(prs.slides._sldIdLst), 'slides')
