"""Generate the AN6004 food-security Tableau workbook (FoodSecurity_Dashboard.twbx).

Usage:  python build_dashboard.py [output.twbx] [active dashboard name]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from twb import DS, Inst, run, NL, fmt, style_rule
from builder import Datasource, Parameter, Workbook, Worksheet, Dashboard, package_twbx

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV = 'food_security_2017_2023_CLEAN.csv'
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'FoodSecurity_Dashboard.twbx')
ACTIVE = sys.argv[2] if len(sys.argv) > 2 else '1 · Where things stand'

# --------------------------------------------------------------------------
# Design tokens
# --------------------------------------------------------------------------
PAPER = '#f4f0e8'      # dashboard background
CARD = '#ffffff'       # card background
LINE = '#e4ddd1'       # card border / hairlines
NAVY = '#14213d'       # header band + ink
INK = '#1f2a44'
MUTED = '#6b7280'
SOFT = '#a39e94'
GOLD = '#e9a23b'
RED = '#c8473b'
WINE = '#7e1f2b'
TEAL = '#2a9d8f'
GREY_DOT = '#9aa3b2'
REGION_COLORS = {'Africa': '#c8473b', 'Americas': '#2a9d8f', 'Asia': '#e9a23b',
                 'Europe': '#4a6fa5', 'Oceania': '#8e6bb0'}
BAND_COLORS = ['#f6e7d3', '#f0bf94', '#e3875f', '#c24b3a', '#7e1f2b']
BAND_LABELS = ['Under 10%', '10–25%', '25–50%', '50–75%', '75% or more']

SEMI = 'Tableau Semibold'
BOOK = 'Tableau Book'
W, H = 1300, 820

# --------------------------------------------------------------------------
# Data source, parameter and calculated fields
# --------------------------------------------------------------------------
ds = Datasource(os.path.join(ROOT, CSV), 'Food Security 2017-2023', 'Data/Datasources', CSV)
ds['ISO3'].semantic = '[Country].[ISO3166_2]'
ds['Healthy_Diet_Unaffordable_Pct'].caption = 'Healthy diet unaffordable (%)'
ds['Undernourishment_Pct'].caption = 'Undernourishment (%)'
for f in ('Healthy_Diet_Unaffordable_Pct', 'Undernourishment_Pct', 'Healthy_Diet_2017_Pct',
          'Healthy_Diet_2023_Pct', 'PoU_2017_Pct', 'PoU_2023_Pct'):
    ds[f].default_format = 'n0.0"%"'
for f in ('Healthy_Diet_Change_pp', 'PoU_Change_pp'):
    ds[f].default_format = 'n+0.0" pp";-0.0" pp"'

year_p = Parameter('Parameter 1', 'Year', 'string', '"2023"', ['"%d"' % y for y in range(2017, 2024)])

PCT = 'n0.0"%"'
c_isyear = ds.calc('Is Selected Year', 'STR([Year]) = [Parameters].[Parameter 1]', 'boolean', role='dimension')
c_diet_med = ds.calc('Median: healthy diet unaffordable', 'MEDIAN([Healthy_Diet_Unaffordable_Pct])', 'real', default_format=PCT)
c_pou_med = ds.calc('Median: undernourishment', 'MEDIAN([Undernourishment_Pct])', 'real', default_format=PCT)
c_africa_med = ds.calc('Median: Africa healthy diet unaffordable',
                       "MEDIAN(IF [Region] = 'Africa' THEN [Healthy_Diet_Unaffordable_Pct] END)", 'real', default_format=PCT)
c_is_top = ds.calc('Is highest share',
                   '[Healthy_Diet_Unaffordable_Pct] = {FIXED [Year] : MAX([Healthy_Diet_Unaffordable_Pct])}',
                   'boolean', role='dimension')
c_med17 = ds.calc('Median 2017', 'MEDIAN(IF [Year] = 2017 THEN [Healthy_Diet_Unaffordable_Pct] END)', 'real', default_format=PCT)
c_med23 = ds.calc('Median 2023', 'MEDIAN(IF [Year] = 2023 THEN [Healthy_Diet_Unaffordable_Pct] END)', 'real', default_format=PCT)
c_improved = ds.calc('Countries improved',
                     "COUNTD(IF [Healthy_Diet_Change_Class] = 'Improved' THEN [ISO3] END)", 'integer', default_format='n0')
c_worsened = ds.calc('Countries worsened',
                     "COUNTD(IF [Healthy_Diet_Change_Class] = 'Worsened' THEN [ISO3] END)", 'integer', default_format='n0')
c_watch_n = ds.calc('Countries on watchlist',
                    "COUNTD(IF [Both_Elevated_And_Both_Worsened] = 'Yes' THEN [ISO3] END)", 'integer', default_format='n0')
c_watch = ds.calc('Watchlist',
                  "IF [Both_Elevated_And_Both_Worsened] = 'Yes' THEN 'High burden & worsening on both' ELSE 'Other countries' END",
                  'string', role='dimension')
c_watch_lbl = ds.calc('Watchlist label', "IF [Both_Elevated_And_Both_Worsened] = 'Yes' THEN [Country] END",
                      'string', role='dimension')
c_band = ds.calc('Healthy diet band',
                 "IF [Healthy_Diet_Unaffordable_Pct] < 10 THEN '%s' ELSEIF [Healthy_Diet_Unaffordable_Pct] < 25 THEN '%s' "
                 "ELSEIF [Healthy_Diet_Unaffordable_Pct] < 50 THEN '%s' ELSEIF [Healthy_Diet_Unaffordable_Pct] < 75 THEN '%s' "
                 "ELSE '%s' END" % tuple(BAND_LABELS), 'string', role='dimension')
c_point = ds.calc('Dumbbell point',
                  "IF [Year] = 2017 THEN '2017' ELSE '2023 ' + [Healthy_Diet_Change_Class] END", 'string', role='dimension')
c_top_w = ds.calc('Top 10 worsened', '{FIXED [ISO3] : MIN([Diet_Change_Rank_Worst])} <= 10', 'boolean', role='dimension')
c_top_b = ds.calc('Top 10 improved', '{FIXED [ISO3] : MIN([Diet_Change_Rank_Best])} <= 10', 'boolean', role='dimension')
c_chg_lbl = ds.calc('Change label',
                    'IF ([Healthy_Diet_Change_pp] >= 0 AND [Year] = 2023) OR ([Healthy_Diet_Change_pp] < 0 AND [Year] = 2017) '
                    'THEN [Healthy_Diet_Change_pp] END', 'real', default_format='n+0.0" pp";-0.0" pp"')
c_diet_copy = ds.calc('Healthy diet unaffordable (line)', '[Healthy_Diet_Unaffordable_Pct]', 'real', default_format=PCT)
c_end_region = ds.calc('Trend end label', 'IF MIN([Year]) = 2023 THEN ATTR([Region]) END', 'string', ftype='nominal')
c_end_val = ds.calc('Trend end value', 'IF MIN([Year]) = 2023 THEN MEDIAN([Healthy_Diet_Unaffordable_Pct]) END', 'real', default_format=PCT)
c_start_val = ds.calc('Trend start value', 'IF MIN([Year]) = 2017 THEN MEDIAN([Healthy_Diet_Unaffordable_Pct]) END', 'real', default_format=PCT)
c_year_num = ds.calc('Year (continuous)', '[Year]', 'integer', role='dimension', ftype='quantitative', default_format='n0')
c_burden = ds.calc('Burden 2023',
                   "CASE [Current_Burden_Group_2023] WHEN 'Both elevated' THEN 'High on both' "
                   "WHEN 'Diet affordability elevated only' THEN 'Costly diet' "
                   "WHEN 'PoU elevated only' THEN 'High hunger' ELSE 'Neither high' END",
                   'string', role='dimension')
c_burden_o = ds.calc('Burden order',
                     "CASE [Current_Burden_Group_2023] WHEN 'Both elevated' THEN 1 WHEN 'Diet affordability elevated only' THEN 2 "
                     "WHEN 'PoU elevated only' THEN 3 ELSE 4 END", 'integer')
c_prog = ds.calc('Progress 2017-2023',
                 "CASE [Progress_Group_2017_2023] WHEN 'Both worsened' THEN 'Worse' "
                 "WHEN 'Divergent (opposite direction)' THEN 'Mixed' "
                 "WHEN 'Includes small change' THEN 'Flat' ELSE 'Better' END",
                 'string', role='dimension')
c_prog_o = ds.calc('Progress order',
                   "CASE [Progress_Group_2017_2023] WHEN 'Both worsened' THEN 1 WHEN 'Divergent (opposite direction)' THEN 2 "
                   "WHEN 'Includes small change' THEN 3 ELSE 4 END", 'integer')
c_has_summary = ds.calc('Has 2017 and 2023', 'NOT ISNULL([Current_Burden_Group_2023])', 'boolean', role='dimension')

ds.color_maps += [
    (Inst(c_watch), {'"High burden & worsening on both"': WINE, '"Other countries"': GREY_DOT}),
    (Inst(c_band), {'"%s"' % l: c for l, c in zip(BAND_LABELS, BAND_COLORS)}),
    (Inst(c_point), {'"2017"': '#c9c2b6', '"2023 Worsened"': RED, '"2023 Improved"': TEAL,
                     '"2023 Small change"': MUTED}),
    (Inst(ds['Region']), {'"%s"' % k: v for k, v in REGION_COLORS.items()}),
]

wb = Workbook(ds, [year_p])
wb.wb_style.append(style_rule('all', fmt('font-family', BOOK)))

# Column instances ---------------------------------------------------------
I = Inst
iso = I(ds['ISO3'])
country = I(ds['Country'])
region = I(ds['Region'])
year = I(ds['Year'])
isyear = I(c_isyear)
diet = I(ds['Healthy_Diet_Unaffordable_Pct'], 'Avg')
pou = I(ds['Undernourishment_Pct'], 'Avg')
band = I(c_band)
watch = I(c_watch)
watch_lbl = I(c_watch_lbl)
diet_med = I(c_diet_med, 'User')
pou_med = I(c_pou_med, 'User')
GEO = '[%s].[Geometry (generated)]' % DS
a_country = I(ds['Country'], 'Attribute')
a_region = I(ds['Region'], 'Attribute')
a_year = I(ds['Year'], 'Attribute')


def base(ws, bg=CARD):
    ws.rule('table', fmt('background-color', bg))
    ws.rule('worksheet', fmt('font-family', BOOK), fmt('color', MUTED), fmt('font-size', '8'),
            fmt('display-field-labels', 'false', scope='rows'), fmt('display-field-labels', 'false', scope='cols'))
    ws.rule('gridline', fmt('line-visibility', 'off', scope='rows'), fmt('line-visibility', 'off', scope='cols'))
    ws.rule('zeroline', fmt('line-visibility', 'off', scope='rows'), fmt('line-visibility', 'off', scope='cols'))
    return ws


def title(text, sub=None):
    r = [run(text, 11, INK, SEMI)]
    if sub:
        r += [NL, run(sub, 8, MUTED, BOOK)]
    return r


def axis_fixed(ws, inst, lo, hi, scope='cols', show=True, title_text=None):
    ws.rule('axis', "<encoding attr='space' class='0' field=%s field-type='quantitative' max='%s' min='%s' "
                    "range-type='fixed' scope=%s type='space' />" % (qa(inst.full), hi, lo, qa(scope)))
    if not show:
        ws.rule('axis', fmt('display', 'false', class_='0', field=inst.full, scope=scope))
    if title_text is not None:
        ws.rule('axis', fmt('title', title_text, class_='0', field=inst.full, scope=scope))


def qa(v):
    return "'" + v.replace("'", '&apos;') + "'"


def tooltip_country():
    return [run('<%s>' % a_country.full, 11, INK, SEMI), NL,
            run('<%s> · <%s>' % (a_region.full, a_year.full), 8, MUTED), NL, NL,
            run("Can't afford a healthy diet: ", 9, MUTED), run('<%s>' % diet.full, 9, INK, SEMI), NL,
            run('Undernourished: ', 9, MUTED), run('<%s>' % pou.full, 9, INK, SEMI)]


def kpi_sheet(name, value_runs, caption, extra_insts=(), filters_true=(), filter_insts=()):
    ws = base(Worksheet(wb, name))
    p = ws.pane(mark='Text')
    for inst in extra_insts:
        ws.use(inst)
        p.add('text', inst)
    for f in filters_true:
        ws.filter_true(f)
    p.label = value_runs
    ws.caption = caption
    p.mark_fmt.append(fmt('mark-labels-show', 'true'))
    ws.rule('cell', fmt('text-align', 'left'), fmt('vertical-align', 'center'))
    return ws


# ==========================================================================
# Dashboard 1 sheets
# ==========================================================================
k1 = kpi_sheet('D1 KPI Median Diet', [run('<%s>' % diet_med.full, 22, RED, SEMI)],
               "can't afford a healthy diet | in the median country", [diet_med], [isyear])
africa = I(c_africa_med, 'User')
k2 = kpi_sheet('D1 KPI Africa', [run('<%s>' % africa.full, 22, WINE, SEMI)],
               "can't afford a healthy diet in | the median African country", [africa], [isyear])
k3 = kpi_sheet('D1 KPI Undernourishment', [run('<%s>' % pou_med.full, 22, INK, SEMI)],
               'are undernourished | in the median country', [pou_med], [isyear])
istop = I(c_is_top)
k4 = kpi_sheet('D1 KPI Highest', [run('<%s>' % diet.full + chr(10), 22, RED, SEMI),
                                  run('<%s>' % a_country.full, 9, INK, SEMI),
                                  run('  highest share', 8, MUTED, BOOK)],
               None, [diet, a_country], [isyear, istop])

# Map ----------------------------------------------------------------------
m = base(Worksheet(wb, 'D1 Map', title=title("Share of people who can't afford a healthy diet",
                                              'Darker = less affordable. Hover a country for details.')))
m.uses_map = True
m.use(iso, band, diet, pou, a_country, a_region, a_year)
m.filter_true(isyear)
m.filter_exclude(iso, ['"TWN"'])   # not recognised by Tableau's geocoder
p = m.pane(mark='Multipolygon')
p.add('color', band).add('lod', iso).add('tooltip', a_country).add('tooltip', a_region).add('tooltip', a_year)
p.add('tooltip', diet).add('tooltip', pou).add('geometry', GEO)
p.tooltip = tooltip_country()
m.rows = '[%s].[Latitude (generated)]' % DS
m.cols = '[%s].[Longitude (generated)]' % DS
m.rule('map', fmt('washout', '0.0'))

# Scatter ------------------------------------------------------------------
sc = base(Worksheet(wb, 'D1 Scatter', title=title('Unaffordable diets go with hunger',
                                                  'Each dot is a country · wine = high on both and worse on both')))
sc.use(country, watch, watch_lbl, diet, pou, a_country, a_region, a_year)
sc.filter_true(isyear)
p = sc.pane(mark='Circle')
p.add('color', watch).add('lod', country).add('text', watch_lbl).add('tooltip', a_country).add('tooltip', a_region).add('tooltip', a_year)
p.tooltip = tooltip_country()
p.label = [run('<%s>' % watch_lbl.full, 8, WINE, SEMI)]
p.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('mark-labels-cull', 'true'), fmt('size', '0.6'),
               fmt('mark-transparency', '200')]
p.trend = True
sc.cols = diet.full
sc.rows = pou.full
axis_fixed(sc, diet, 0, 100, 'cols', title_text="Can't afford a healthy diet (%)")
axis_fixed(sc, pou, 0, 60, 'rows', title_text='Undernourished (%)')
sc.rule('axis', fmt('text-format', 'n0"%"', class_='0', field=diet.full, scope='cols'), fmt('text-format', 'n0"%"', class_='0', field=pou.full, scope='rows'))

# Strip plot by region -----------------------------------------------------
st = base(Worksheet(wb, 'D1 Region Strip', title=title('Every country, by region',
                                                       'Dots = countries · bar = regional median')))
st.use(country, region, diet, diet_med, a_country, a_region, a_year)
st.filter_true(isyear)
st.sort_by(region, diet_med, 'DESC')
st.rows = region.full
st.cols = '(%s + %s)' % (diet.full, diet_med.full)
st.pane()
p1 = st.pane(mark='Circle', pid=1, x_axis=diet)
p1.add('color', region).add('lod', country).add('tooltip', a_country).add('tooltip', a_region).add('tooltip', a_year)
p1.tooltip = tooltip_country()[:4] + [run("Can't afford a healthy diet: ", 9, MUTED), run('<%s>' % diet.full, 9, INK, SEMI)]
p1.mark_fmt += [fmt('size', '0.45'), fmt('mark-transparency', '150')]
p2 = st.pane(mark='GanttBar', pid=2, x_axis=diet_med)
p2.add('text', diet_med)
p2.label = [run('<%s>' % diet_med.full, 8, INK, SEMI)]
p2.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('mark-color', INK), fmt('size', '2')]
st.rule('axis', "<encoding attr='space' class='0' field=%s field-type='quantitative' fold='true' scope='cols' "
                "synchronized='true' type='space' />" % qa(diet_med.full),
        fmt('display', 'false', class_='0', field=diet_med.full, scope='cols'))
axis_fixed(st, diet, 0, 100, 'cols', show=False)
st.rule('header', fmt('font-size', '8'), fmt('color', INK))

# ==========================================================================
# Dashboard 2 sheets
# ==========================================================================
m17, m23 = I(c_med17, 'User'), I(c_med23, 'User')
k5 = kpi_sheet('D2 KPI Median Change', [run('<%s>' % m17.full, 18, SOFT, SEMI), run(' → ', 14, SOFT, BOOK),
                                         run('<%s>' % m23.full, 18, INK, SEMI)],
               "median share that can't afford | a healthy diet, 2017 → 2023", [m17, m23])
ki = I(c_improved, 'User')
k6 = kpi_sheet('D2 KPI Improved', [run('<%s>' % ki.full, 22, TEAL, SEMI)],
               'countries where a healthy diet | became more affordable (≥ 1 pp)', [ki])
kw = I(c_worsened, 'User')
k7 = kpi_sheet('D2 KPI Worsened', [run('<%s>' % kw.full, 22, RED, SEMI)],
               'countries where a healthy diet | became less affordable (≥ 1 pp)', [kw])
kn = I(c_watch_n, 'User')
k8 = kpi_sheet('D2 KPI Watchlist', [run('<%s>' % kn.full, 22, WINE, SEMI)],
               'countries high on both measures | and worse on both since 2017', [kn])

# Regional trend -----------------------------------------------------------
yr_c = I(c_year_num)
end_r, end_v, start_v = I(c_end_region, 'User'), I(c_end_val, 'User'), I(c_start_val, 'User')
tr = base(Worksheet(wb, 'D2 Regional Trend', title=title('Africa: no progress since 2017',
                                                         'Median %, by region · Oceania (2 countries) omitted')))
tr.use(yr_c, region, diet_med, end_r, end_v, start_v)
tr.filter_members(region, ['"Africa"', '"Americas"', '"Asia"', '"Europe"'])
tr.cols = yr_c.full
tr.rows = diet_med.full
p = tr.pane(mark='Line')
p.add('color', region).add('text', end_r).add('text', end_v)
p.label = [run(' <%s>' % end_r.full, 8, INK, SEMI), run(' <%s>' % end_v.full, 8, INK)]
p.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('size', '2.5')]
axis_fixed(tr, yr_c, 2016.5, 2024.6, 'cols', title_text='')
axis_fixed(tr, diet_med, 0, 70, 'rows', title_text='')
tr.rule('axis', fmt('text-format', 'n0', field=yr_c.full, scope='cols'))

# Dumbbells ----------------------------------------------------------------
pt = I(c_point)
chg = I(ds['Healthy_Diet_Change_pp'], 'Avg')
chg_lbl = I(c_chg_lbl, 'Avg')
diet_line = I(c_diet_copy, 'Avg')


def dumbbell(name, flag, direction, ttl, sub, line_color):
    ws = base(Worksheet(wb, name, title=title(ttl, sub)))
    ws.use(country, year, pt, I(flag), diet, diet_line, chg, chg_lbl)
    ws.filter_members(year, ['2017', '2023'])
    ws.filter_true(I(flag))
    ws.sort_by(country, chg, direction)
    ws.rows = country.full
    ws.cols = '(%s + %s)' % (diet.full, diet_line.full)
    ws.pane()
    a = ws.pane(mark='Circle', pid=1, x_axis=diet)
    a.add('color', pt).add('lod', year).add('text', chg_lbl).add('tooltip', chg)
    a.label = [run('<%s>' % chg_lbl.full, 8, INK, SEMI)]
    a.tooltip = [run('<%s>' % country.full, 11, INK, SEMI), NL,
                 run("<%s>: <%s> can't afford a healthy diet" % (year.full, diet.full), 9, MUTED), NL,
                 run('Change 2017→2023: <%s>' % chg.full, 9, INK, SEMI)]
    a.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('size', '0.9')]
    b = ws.pane(mark='Line', pid=2, x_axis=diet_line)
    b.add('path', year)
    b.mark_fmt += [fmt('mark-color', line_color), fmt('size', '2.5')]
    ws.rule('axis', "<encoding attr='space' class='0' field=%s field-type='quantitative' fold='true' scope='cols' "
                    "synchronized='true' type='space' />" % qa(diet_line.full),
            fmt('display', 'false', class_='0', field=diet_line.full, scope='cols'))
    axis_fixed(ws, diet, 0, 100, 'cols', title_text=' ')
    ws.rule('header', fmt('font-size', '9'), fmt('color', INK))
    return ws


dw = dumbbell('D2 Dumbbell Worsened', c_top_w, 'DESC', '10 biggest setbacks',
              'grey = 2017 · red = 2023', '#efc3bd')
dbt = dumbbell('D2 Dumbbell Improved', c_top_b, 'ASC', '10 biggest gains',
               'grey = 2017 · teal = 2023', '#b5ddd7')

# Burden x progress matrix ------------------------------------------------
bur, bur_o = I(c_burden), I(c_burden_o, 'Min')
prog, prog_o = I(c_prog), I(c_prog_o, 'Min')
cnt = I(ds['ISO3'], 'CountD')
mx = base(Worksheet(wb, 'D2 Matrix', title=title('Burden now vs. direction of change',
                                                 '2023 burden × change since 2017 · top-left = watchlist')))
mx.use(bur, bur_o, prog, prog_o, cnt, I(c_has_summary), year)
mx.filter_members(year, ['2023'])
mx.filter_true(I(c_has_summary))
mx.sort_by(bur, bur_o, 'ASC', 'rows')
mx.sort_by(prog, prog_o, 'ASC', 'columns')
mx.rows = bur.full
mx.cols = prog.full
p = mx.pane(mark='Square')
p.add('color', cnt).add('text', cnt)
p.label = [run('<%s>' % cnt.full, 11, INK, SEMI)]
p.mark_fmt += [fmt('mark-labels-show', 'true')]
mx.table_rules.append(style_rule('mark', "<encoding attr='color' field=%s type='custom-interpolated'>\n"
                                         "<color-palette custom='true' name='' type='ordered-sequential'>\n"
                                         "<color>#f7f1e8</color>\n<color>#e8c9a8</color>\n<color>#c98a6a</color>\n"
                                         "</color-palette>\n</encoding>" % qa(cnt.full)))
mx.rule('header', fmt('font-size', '9'), fmt('color', INK))

# ==========================================================================
# Dashboards
# ==========================================================================
M = 32                       # outer margin
GAP = 16
HEAD = 116                   # header band height
KPI_Y, KPI_H = HEAD + 14, 100
BODY = KPI_Y + KPI_H + 14
BOTTOM = 764


def header(db, kicker, headline, sub):
    db.box(0, 0, W, HEAD, NAVY)
    db.text(M - 4, 8, 1030, HEAD - 12, [run(kicker, 8, GOLD, SEMI), NL, run(headline, 16, '#ffffff', SEMI), NL,
                                        run(sub, 9, '#c3cbe0', BOOK)], bg=NAVY)


def kpis(db, sheets, colors, y=KPI_Y, h=KPI_H):
    w = (W - 2 * M - 3 * GAP) / 4
    for i, (ws, c) in enumerate(zip(sheets, colors)):
        x = M + i * (w + GAP)
        db.box(x, y, w, h, CARD, LINE)
        db.box(x, y, w, 4, c)
        if ws.caption is None:
            db.sheet(x + 10, y + 8, w - 20, h - 12, ws, show_title=False, bg=CARD)
        else:
            db.sheet(x + 10, y + 6, w - 20, 38, ws, show_title=False, bg=CARD)
            lines = [l.strip() for l in ws.caption.split('|')]
            runs = []
            for j, l in enumerate(lines):
                runs += ([NL] if j else []) + [run(l, 8, MUTED, BOOK)]
            db.text(x + 14, y + 42, w - 28, h - 44, runs, bg=CARD)


def card(db, x, y, w, h, ws):
    db.box(x, y, w, h, CARD, LINE)
    db.sheet(x + 12, y + 10, w - 24, h - 18, ws, bg=CARD)


def footer(db, text):
    db.box(M, BOTTOM + 10, W - 2 * M, 1, LINE)
    a, b = text.split(' | ')
    db.text(M - 4, BOTTOM + 13, W - 2 * M + 8, H - BOTTOM - 15, [run(a, 7, MUTED, BOOK), NL, run(b, 7, MUTED, BOOK)])


SOURCE = ("Source: FAO, IFAD, UNICEF, WFP & WHO, The State of Food Security and Nutrition in the World (via Our World in Data). "
          "Taiwan is not mapped. pp = percentage points. | "
          "Undernourishment is a 3-year average; values ≤ 2.5% are reported as 2.5%. "
          "'High' = top quartile of 2023 values (diet ≥ 55.2%, undernourishment ≥ 13.1%), not a UN threshold.")

d1 = Dashboard(wb, '1 · Where things stand', W, H, PAPER)
header(d1, 'FOOD SECURITY 2017–2023  ·  1 / 2  ·  WHERE THINGS STAND',
       "1 in 4 people can't afford a healthy diet. In Africa, it's 6 in 10.",
       'Share of people unable to afford a healthy diet, and share undernourished, across 138 countries. Pick a year on the right.')
d1.param(W - M - 170, 34, 170, 48, year_p, bg='#ffffff')
kpis(d1, [k1, k2, k3, k4], [RED, WINE, INK, RED])
MW = 700
d1.box(M, BODY, MW, BOTTOM - BODY, CARD, LINE)
d1.sheet(M + 12, BODY + 10, MW - 24, BOTTOM - BODY - 44, m, bg=CARD)
d1.text(M + 16, BOTTOM - 32, MW - 32, 26,
        sum(([run('■ ', 12, c), run(l + '     ', 8, MUTED)] for c, l in zip(BAND_COLORS, BAND_LABELS)), []))
RX1 = M + MW + GAP
RW1 = W - M - RX1
card(d1, RX1, BODY, RW1, 290, sc)
card(d1, RX1, BODY + 290 + GAP, RW1, BOTTOM - BODY - 290 - GAP, st)
footer(d1, SOURCE)

d2 = Dashboard(wb, '2 · Who is falling behind', W, H, PAPER)
header(d2, 'FOOD SECURITY 2017–2023  ·  2 / 2  ·  WHO IS FALLING BEHIND',
       '63 countries made healthy diets more affordable. 38 slipped back.',
       'Change in the share of people unable to afford a healthy diet, 2017 → 2023, and who is getting worse on both measures.')
kpis(d2, [k5, k6, k7, k8], [INK, TEAL, RED, WINE])
LW = 440
card(d2, M, BODY, LW, 280, tr)
card(d2, M, BODY + 280 + GAP, LW, BOTTOM - BODY - 280 - GAP, mx)
RX = M + LW + GAP
RW = W - M - RX
HW = (RW - GAP) / 2
card(d2, RX, BODY, HW, BOTTOM - BODY, dw)
card(d2, RX + HW + GAP, BODY, HW, BOTTOM - BODY, dbt)
footer(d2, SOURCE)

xml = wb.xml(active=ACTIVE)
package_twbx(xml, OUT, 'FoodSecurity_Dashboard.twb', ds, os.path.join(ROOT, CSV), os.path.dirname(os.path.abspath(OUT)))
print('wrote', OUT)
