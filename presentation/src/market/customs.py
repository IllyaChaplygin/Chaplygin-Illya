# -*- coding: utf-8 -*-
"""Customs-base analysis (UKT ZED, 2025 + Jan-May 2026) for the snack categories.
Source: user's file with four exact codes. Declarations are not split into SKUs, so
every figure is a declaration-level total; the snack subset is cut by description
keywords (stated openly on the slides)."""
import glob, json, os, re
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = glob.glob('/root/.claude/uploads/c53f34b4-ae43-52db-aeb2-01bdcda48cc3/439e51ff-*')[0]
CODES = {'2008999990': 'Морські водорості приготовані (нори-снеки, топінг, сендвічі)',
         '1905905500': 'Екструдовані / експандовані снеки (рисові крекери, темпура)',
         '1902301000': 'Локшина швидкого приготування (рис)',
         '1904901000': 'Готовий рис'}
BRAND_TOKENS = {'Haelove / Delisse': 'HAELOVE|DELISSE', 'Tao Kae Noi': 'TAO KAE NOI|TAOKAENOI', 'Ock Dong Ja': r'OCK.?DONG|ОК.?ДОНГ',
                'Akura': 'AKURA', 'Royal Tiger': 'ROYAL TIGER', 'Clearspring': 'CLEARSPRING', 'Metro Chef': 'METRO CHEF',
                'Kimnori': 'KIMNORI', 'Yamchan': 'YAMCHAN|ЯМЧАН'}


def short(n):
    n = re.sub(r'\s+\d{5}\b.*$', '', n)                 # strip trailing address
    n = n.replace('ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ', 'ТОВ').replace('Товариство з обмеженою відповідаль- ністю', 'ТОВ')
    n = re.sub(r'["«»]', '', n).strip()
    return n[:46]


def main():
    df = pd.read_excel(SRC, sheet_name='Сирі дані', dtype={'hs_code_normalized': str})
    df['U'] = df.product_description.str.upper().fillna('')
    df['kg'] = df.analysis_weight_kg
    df['usd'] = df.analysis_invoice_value_usd
    out = dict(period='2025 + січень–травень 2026', months=sorted(df.report_month.unique().tolist()))
    # --- overview per code
    ov = []
    for c, nm in CODES.items():
        d = df[df.hs_code_normalized == c]
        top = d.groupby('importer_display_name').kg.sum().sort_values(ascending=False)
        co = d.groupby('origin_country_normalized_ua').kg.sum().sort_values(ascending=False)
        ov.append(dict(code=c, name=nm, tons=d.kg.sum() / 1000, usd=d.usd.sum(), usd_kg=d.usd.sum() / d.kg.sum(), decl=int(d.declaration_number.nunique()),
                       importers=int(d.importer_display_name.nunique()), top_importer=short(top.index[0]), top_share=top.iloc[0] / d.kg.sum(),
                       top_origin=co.index[0], origin_share=co.iloc[0] / d.kg.sum()))
    out['overview'] = ov
    # --- seaweed snacks inside 2008 99 99 90
    d = df[df.hs_code_normalized == '2008999990'].copy()
    d['usdkg'] = d.usd / d.kg
    sea = d.U.str.contains('ВОДОРОСТ|НОРІ|NORI|SEAWEED|LAVER|KIMNORI|ГІМ|GIM')
    snack = d.U.str.contains('СНЕК|ЧИПС|SNACK|CHIPS|ХРУСТ|CRISP|ПІДСМАЖ|ЗАКУСК|ТОПІНГ|TOPPING|РУЛОН|ROLL|ТЕМПУР')
    sushi = d.U.str.contains('СУШІ|SUSHI|ЯКІНОРІ|YAKI')
    ok = sea & snack & (d.usdkg > 8) & (d.usdkg < 120)
    A, B = d[ok & ~sushi].copy(), d[ok & sushi].copy()
    A['tier'] = 'A'; B['tier'] = 'B'
    S = pd.concat([A, B])
    def pack(x):
        return dict(tons=x.kg.sum() / 1000, usd=x.usd.sum(), usd_kg=(x.usd.sum() / x.kg.sum()) if x.kg.sum() else 0, rows=int(len(x)), decl=int(x.declaration_number.nunique()))
    out['seaweed'] = dict(A=pack(A), B=pack(B), all=pack(S), code_total_tons=d.kg.sum() / 1000)
    months = out['months']
    out['seaweed']['monthly_A'] = [round(A[A.report_month == m].kg.sum() / 1000, 2) for m in months]
    out['seaweed']['monthly_B'] = [round(B[B.report_month == m].kg.sum() / 1000, 2) for m in months]
    imp = S.groupby('importer_display_name').agg(a=('kg', lambda s: s[S.loc[s.index, 'tier'] == 'A'].sum()), tot=('kg', 'sum'), usd=('usd', 'sum'))
    imp = imp.sort_values('tot', ascending=False)
    rows = []
    for name, r in imp.head(12).iterrows():
        sub = S[S.importer_display_name == name]
        brands = [b for b, p in BRAND_TOKENS.items() if sub.U.str.contains(p).any()]
        rows.append(dict(name=short(name), tons=r.tot / 1000, tons_pure=r.a / 1000, share=r.tot / S.kg.sum(), usd_kg=r.usd / r.tot, brands=brands,
                         origin=sub.origin_country_normalized_ua.value_counts().index[0]))
    out['seaweed']['importers'] = rows
    og = S.groupby('origin_country_normalized_ua').agg(kg=('kg', 'sum'), usd=('usd', 'sum')).sort_values('kg', ascending=False)
    out['seaweed']['origins'] = [dict(country=c, tons=r.kg / 1000, share=r.kg / S.kg.sum(), usd_kg=r.usd / r.kg) for c, r in og.iterrows()]
    mf = S.groupby('manufacturer_normalized').agg(kg=('kg', 'sum'), usd=('usd', 'sum'), country=('origin_country_normalized_ua', 'first')).sort_values('kg', ascending=False)
    out['seaweed']['manufacturers'] = [dict(name=n[:40], tons=r.kg / 1000, usd_kg=r.usd / r.kg, country=r.country) for n, r in mf.head(10).iterrows()]
    # --- rice snacks / rice crackers inside 1905 90 55 00
    d2 = df[df.hs_code_normalized == '1905905500'].copy()
    rice = d2.U.str.contains('РИС')
    cracker = d2.U.str.contains('КРЕКЕР|CRACKER|SENBEI|АРАРЕ|ARARE|NORIMAKI')
    sea2 = d2.U.str.contains('ВОДОРОСТ|НОРІ|NORI|SEAWEED')
    R = d2[rice]
    rc = d2[rice & cracker]
    out['rice'] = dict(code_total_tons=d2.kg.sum() / 1000, rice_all=pack(R), rice_cracker=pack(rc), nori_anything=pack(d2[sea2]),
                       importers=[dict(name=short(n), tons=v / 1000) for n, v in R.groupby('importer_display_name').kg.sum().sort_values(ascending=False).head(8).items()],
                       manufacturers=[dict(name=n[:40], tons=v / 1000, country=R[R.manufacturer_normalized == n].origin_country_normalized_ua.iloc[0]) for n, v in R.groupby('manufacturer_normalized').kg.sum().sort_values(ascending=False).head(8).items()],
                       cracker_importers=[dict(name=short(n), tons=v / 1000) for n, v in rc.groupby('importer_display_name').kg.sum().sort_values(ascending=False).head(5).items()])
    json.dump(out, open(os.path.join(HERE, 'customs.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=float)
    print(json.dumps(out['seaweed']['A'], default=float), json.dumps(out['seaweed']['B'], default=float))
    print(json.dumps(out['rice'], ensure_ascii=False, default=float)[:900])


if __name__ == '__main__':
    main()
