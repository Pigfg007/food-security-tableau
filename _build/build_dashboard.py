"""Generate the AN6004 food-security Tableau workbook (FoodSecurity_Dashboard.twbx), v3.

v3 follows TABLEAU修改.docx: three dashboards
  1 Where Is the Burden?          (map / region distribution / top-10 ranking, Metric parameter)
  2 Related, but Not the Same     (2023 scatter, Sample parameter, dynamic n and r)
  3 Who Is Falling Behind?        (change scatter, burden x progress matrix -> country detail,
                                   one dumbbell switched by a parameter, regional trend)

Data: the merged CLEAN file (Country x Year, 2017-2023, no 2024 data). Country-level fields
(change, burden group, progress group) repeat on every year row, so every sheet that uses them
filters Year = 2023 and aggregates with MIN / MEDIAN / COUNTD, never SUM.

Usage:  python build_dashboard.py [output.twbx] [active dashboard name]
"""
import csv
import os
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from twb import DS, Inst, run, NL, fmt, style_rule
from builder import Datasource, Parameter, Workbook, Worksheet, Dashboard, package_twbx

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV = 'food_security_2017_2023_CLEAN.csv'
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'FoodSecurity_Dashboard.twbx')
D1N = '1 · Where Is the Burden?'
D2N = '2 · Related, but Not the Same'
D3N = '3 · Who Is Falling Behind?'
ACTIVE = sys.argv[2] if len(sys.argv) > 2 else D1N

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
SLATE = '#4a6fa5'
GREY_DOT = '#9aa3b2'
REGION_COLORS = {'Africa': '#c8473b', 'Americas': '#2a9d8f', 'Asia': '#e9a23b',
                 'Europe': '#4a6fa5', 'Oceania': '#8e6bb0'}
MAP_SCALE = ['#f6e7d3', '#e3875f', '#c24b3a', '#7e1f2b']

SEMI = 'Tableau Semibold'
BOOK = 'Tableau Book'
W, H = 1300, 820
PCT = 'n0.0"%"'
PP = 'n+0.0" pp";-0.0" pp"'

# --------------------------------------------------------------------------
# Data source
# --------------------------------------------------------------------------
ds = Datasource(os.path.join(ROOT, CSV), 'Food Security 2017-2023', 'Data/Datasources', CSV)
ds['ISO3'].semantic = '[Country].[ISO3166_2]'
ds['Healthy_Diet_Unaffordable_Pct'].caption = 'Healthy diet unaffordable (%)'
ds['Undernourishment_Pct'].caption = 'Undernourishment (%)'
for f in ('Healthy_Diet_Unaffordable_Pct', 'Undernourishment_Pct', 'Healthy_Diet_2017_Pct',
          'Healthy_Diet_2023_Pct', 'PoU_2017_Pct', 'PoU_2023_Pct'):
    ds[f].default_format = PCT
for f in ('Healthy_Diet_Change_pp', 'PoU_Change_pp'):
    ds[f].default_format = PP

with open(os.path.join(ROOT, CSV), newline='', encoding='utf-8') as fh:
    _rows = list(csv.DictReader(fh))
COUNTRIES = sorted({r['Country'] for r in _rows})
REGIONS = sorted({r['Region'] for r in _rows})

# --------------------------------------------------------------------------
# Parameters
# --------------------------------------------------------------------------
HD_LBL, POU_LBL = 'Healthy diet unaffordable', 'Undernourishment (PoU)'
ALL_SAMPLE, EXCL_SAMPLE = 'All countries', 'Exclude PoU at 2.5% reporting floor'
SETBACKS, GAINS = '10 biggest setbacks', '10 biggest gains'

year_p = Parameter('Parameter 1', 'Year', 'string', '"2023"', ['"%d"' % y for y in range(2017, 2024)])
metric_p = Parameter('Parameter 2', 'Metric', 'string', '"%s"' % HD_LBL, ['"%s"' % HD_LBL, '"%s"' % POU_LBL])
region_p = Parameter('Parameter 3', 'Region', 'string', '"All regions"',
                     ['"All regions"'] + ['"%s"' % r for r in REGIONS])
country_p = Parameter('Parameter 4', 'Country', 'string', '"All countries"',
                      ['"All countries"'] + ['"%s"' % c for c in COUNTRIES])
sample_p = Parameter('Parameter 5', 'Sample', 'string', '"%s"' % ALL_SAMPLE,
                     ['"%s"' % ALL_SAMPLE, '"%s"' % EXCL_SAMPLE])
db_p = Parameter('Parameter 6', 'Show', 'string', '"%s"' % SETBACKS, ['"%s"' % SETBACKS, '"%s"' % GAINS])
P_YEAR, P_METRIC, P_REGION, P_COUNTRY, P_SAMPLE, P_DB = (p.full for p in
                                                         (year_p, metric_p, region_p, country_p, sample_p, db_p))

# --------------------------------------------------------------------------
# Calculated fields
# --------------------------------------------------------------------------
# Shared selection logic
c_isyear = ds.calc('Is Selected Year', 'STR([Year]) = %s' % P_YEAR, 'boolean', role='dimension')
c_region_ok = ds.calc('In selected region', "%s = 'All regions' OR [Region] = %s" % (P_REGION, P_REGION),
                      'boolean', role='dimension')
c_insel = ds.calc('In selection',
                  "(%s = 'All regions' OR [Region] = %s) AND (%s = 'All countries' OR [Country] = %s)"
                  % (P_REGION, P_REGION, P_COUNTRY, P_COUNTRY), 'boolean', role='dimension')
c_d1 = ds.calc('D1 in view', '[Is Selected Year] AND [In selection]', 'boolean', role='dimension')
c_metric = ds.calc('Selected metric', "IF %s = '%s' THEN [Healthy_Diet_Unaffordable_Pct] ELSE [Undernourishment_Pct] END"
                   % (P_METRIC, HD_LBL), 'real', default_format=PCT)
c_y23 = ds.calc('Is 2023 with 2017 pair', '[Year] = 2023 AND NOT ISNULL([Current_Burden_Group_2023])',
                'boolean', role='dimension')

# D1 ----------------------------------------------------------------------
c_diet_med = ds.calc('Median: healthy diet unaffordable', 'MEDIAN([Healthy_Diet_Unaffordable_Pct])', 'real', default_format=PCT)
c_pou_med = ds.calc('Median: undernourishment', 'MEDIAN([Undernourishment_Pct])', 'real', default_format=PCT)
c_n = ds.calc('Countries in view', 'COUNTD([Country])', 'integer', default_format='n0')
c_metric_med = ds.calc('Median: selected metric', 'MEDIAN([Selected metric])', 'real', default_format=PCT)
c_is_top = ds.calc('Is highest (selected metric)',
                   '[In selection] AND [Selected metric] = {FIXED [Year] : MAX(IF [In selection] THEN [Selected metric] END)}',
                   'boolean', role='dimension')
_sel = 'IF [In selection] THEN [Selected metric] END'
c_q1 = ds.calc('Region Q1', '{FIXED [Region], [Year] : PERCENTILE(%s, 0.25)}' % _sel, 'real', default_format=PCT)
c_qm = ds.calc('Region median', '{FIXED [Region], [Year] : MEDIAN(%s)}' % _sel, 'real', default_format=PCT)
c_q3 = ds.calc('Region Q3', '{FIXED [Region], [Year] : PERCENTILE(%s, 0.75)}' % _sel, 'real', default_format=PCT)
c_half = ds.calc('Box half', "IF [Selected metric] <= [Region median] THEN 'Lower half' ELSE 'Upper half' END",
                 'string', role='dimension')
c_bstart = ds.calc('Box start', "IF [Box half] = 'Lower half' THEN [Region Q1] ELSE [Region median] END",
                   'real', default_format=PCT)
c_blen = ds.calc('Box length', "IF [Box half] = 'Lower half' THEN [Region median] - [Region Q1] "
                               "ELSE [Region Q3] - [Region median] END", 'real', default_format=PCT)

# D2 ----------------------------------------------------------------------
c_d2 = ds.calc('D2 in sample', "[Year] = 2023 AND [In selected region] AND "
                               "(%s = '%s' OR [PoU_At_2p5_Reporting_Floor] = 'No')" % (P_SAMPLE, ALL_SAMPLE),
               'boolean', role='dimension')
c_r = ds.calc('Correlation r', 'CORR([Healthy_Diet_Unaffordable_Pct], [Undernourishment_Pct])', 'real', default_format='n0.000')
EXC_DIET = 'Costly diet only'
EXC_POU = 'High hunger only'
c_grp = ds.calc('Scatter group',
                "CASE [Current_Burden_Group_2023] WHEN 'Both elevated' THEN 'High on both' "
                "WHEN 'Diet affordability elevated only' THEN '%s' "
                "WHEN 'PoU elevated only' THEN '%s' ELSE 'Neither high' END" % (EXC_DIET, EXC_POU),
                'string', role='dimension')
c_is_exc = ds.calc('Is exception', "[Scatter group] = '%s' OR [Scatter group] = '%s'" % (EXC_DIET, EXC_POU),
                   'boolean', role='dimension')
c_exc_lbl = ds.calc('Exception label', 'IF [Is exception] THEN [Country] END', 'string', role='dimension')
c_exc_n = ds.calc('Exception countries', "COUNTD(IF [Is exception] THEN [Country] END)", 'integer', default_format='n0')

# D3 ----------------------------------------------------------------------
c_med17 = ds.calc('Median 2017 (paired)', 'MEDIAN([Healthy_Diet_2017_Pct])', 'real', default_format=PCT)
c_med23 = ds.calc('Median 2023 (paired)', 'MEDIAN([Healthy_Diet_2023_Pct])', 'real', default_format=PCT)
c_improved = ds.calc('Countries improved',
                     "COUNTD(IF [Healthy_Diet_Change_Class] = 'Improved' THEN [Country] END)", 'integer', default_format='n0')
c_worsened = ds.calc('Countries worsened',
                     "COUNTD(IF [Healthy_Diet_Change_Class] = 'Worsened' THEN [Country] END)", 'integer', default_format='n0')
c_watch_n = ds.calc('Countries on watchlist',
                    "COUNTD(IF [Both_Elevated_And_Both_Worsened] = 'Yes' THEN [Country] END)", 'integer', default_format='n0')
c_quad = ds.calc('Change quadrant',
                 "IF [Healthy_Diet_Change_pp] <= 0 AND [PoU_Change_pp] <= 0 THEN 'Both improved' "
                 "ELSEIF [Healthy_Diet_Change_pp] > 0 AND [PoU_Change_pp] > 0 THEN 'Both worsened' "
                 "ELSEIF [Healthy_Diet_Change_pp] <= 0 THEN 'Diet better, hunger worse' "
                 "ELSE 'Diet worse, hunger better' END", 'string', role='dimension')
c_burden = ds.calc('Burden 2023',
                   "CASE [Current_Burden_Group_2023] WHEN 'Both elevated' THEN 'High on both' "
                   "WHEN 'Diet affordability elevated only' THEN 'Costly diet only' "
                   "WHEN 'PoU elevated only' THEN 'High hunger only' ELSE 'Neither high' END",
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
c_floor = ds.calc('Floor note', "IF [PoU_2017_or_2023_At_Floor] = 'Yes' THEN '  ⚑ floor' ELSE '' END",
                  'string', role='dimension')
def _chg(f):
    return "(IF [%s] >= 0 THEN '+' ELSE '' END) + STR(ROUND([%s], 1))" % (f, f)


def _span(a, b, chg):
    return ("STR(ROUND([%s], 1)) + '→' + STR(ROUND([%s], 1)) + ' (' + %s + ')'" % (a, b, _chg(chg)))


c_det_diet = ds.calc('Healthy diet 2017→2023', _span('Healthy_Diet_2017_Pct', 'Healthy_Diet_2023_Pct',
                                                      'Healthy_Diet_Change_pp'), 'string', role='dimension')
c_det_pou = ds.calc('PoU 2017→2023', _span('PoU_2017_Pct', 'PoU_2023_Pct', 'PoU_Change_pp'),
                    'string', role='dimension')
c_point = ds.calc('Dumbbell point',
                  "IF [Year] = 2017 THEN '2017' ELSE '2023 ' + [Healthy_Diet_Change_Class] END", 'string', role='dimension')
c_db_show = ds.calc('In dumbbell',
                    "IF %s = '%s' THEN {FIXED [Country] : MIN([Diet_Change_Rank_Worst])} <= 10 "
                    "ELSE {FIXED [Country] : MIN([Diet_Change_Rank_Best])} <= 10 END" % (P_DB, SETBACKS),
                    'boolean', role='dimension')
c_db_sort = ds.calc('Dumbbell sort', "IF %s = '%s' THEN -[Healthy_Diet_Change_pp] ELSE [Healthy_Diet_Change_pp] END"
                    % (P_DB, SETBACKS), 'real')
c_chg_lbl = ds.calc('Change label',
                    'IF ([Healthy_Diet_Change_pp] >= 0 AND [Year] = 2023) OR ([Healthy_Diet_Change_pp] < 0 AND [Year] = 2017) '
                    'THEN [Healthy_Diet_Change_pp] END', 'real', default_format=PP)
c_diet_copy = ds.calc('Healthy diet unaffordable (line)', '[Healthy_Diet_Unaffordable_Pct]', 'real', default_format=PCT)
c_end_region = ds.calc('Trend end label', 'IF MIN([Year]) = 2023 THEN ATTR([Region]) END', 'string', ftype='nominal')
c_end_val = ds.calc('Trend end value', 'IF MIN([Year]) = 2023 THEN MEDIAN([Healthy_Diet_Unaffordable_Pct]) END', 'real', default_format=PCT)
c_year_num = ds.calc('Year (continuous)', '[Year]', 'integer', role='dimension', ftype='quantitative', default_format='n0')



def internal_refs(formula):
    """Formulas above name other calcs by caption; the .twb must use their internal names."""
    for c in sorted(ds.calcs, key=lambda c: -len(c.caption)):
        formula = formula.replace('[' + c.caption + ']', c.ref)
    return formula


for c in ds.calcs:
    c.formula = internal_refs(c.formula)

ds.color_maps += [
    (Inst(c_grp), {'"High on both"': WINE, '"%s"' % EXC_DIET: GOLD, '"%s"' % EXC_POU: SLATE,
                   '"Neither high"': GREY_DOT}),
    (Inst(c_quad), {'"Both improved"': TEAL, '"Both worsened"': WINE,
                    '"Diet better, hunger worse"': GOLD, '"Diet worse, hunger better"': SLATE}),
    (Inst(c_half), {'"Lower half"': '#ece4d6', '"Upper half"': '#dccfba'}),
    (Inst(c_point), {'"2017"': '#c9c2b6', '"2023 Worsened"': RED, '"2023 Improved"': TEAL,
                     '"2023 Small change"': MUTED}),
    (Inst(ds['Region']), {'"%s"' % k: v for k, v in REGION_COLORS.items()}),
]

wb = Workbook(ds, [year_p, metric_p, region_p, country_p, sample_p, db_p])
wb.wb_style.append(style_rule('all', fmt('font-family', BOOK)))

# Column instances ---------------------------------------------------------
I = Inst
iso = I(ds['ISO3'])
country = I(ds['Country'])
region = I(ds['Region'])
year = I(ds['Year'])
diet = I(ds['Healthy_Diet_Unaffordable_Pct'], 'Avg')
pou = I(ds['Undernourishment_Pct'], 'Avg')
metric = I(c_metric, 'Avg')
GEO = '[%s].[Geometry (generated)]' % DS
a_country = I(ds['Country'], 'Attribute')
a_region = I(ds['Region'], 'Attribute')
a_year = I(ds['Year'], 'Attribute')


def base(ws, bg=CARD, zero=False):
    ws.rule('table', fmt('background-color', bg))
    ws.rule('worksheet', fmt('font-family', BOOK), fmt('color', MUTED), fmt('font-size', '8'),
            fmt('display-field-labels', 'false', scope='rows'), fmt('display-field-labels', 'false', scope='cols'))
    ws.rule('gridline', fmt('line-visibility', 'off', scope='rows'), fmt('line-visibility', 'off', scope='cols'))
    if zero:
        ws.rule('zeroline', fmt('stroke-color', '#b9b2a6', scope='rows'), fmt('stroke-color', '#b9b2a6', scope='cols'))
    else:
        ws.rule('zeroline', fmt('line-visibility', 'off', scope='rows'), fmt('line-visibility', 'off', scope='cols'))
    return ws


def title(text, sub=None):
    r = [run(text, 11, INK, SEMI)]
    if sub:
        r += [NL, run(sub, 8, MUTED, BOOK)]
    return r


def qa(v):
    return "'" + v.replace("'", '&apos;') + "'"


def axis_fixed(ws, inst, lo, hi, scope='cols', show=True, title_text=None):
    ws.rule('axis', "<encoding attr='space' class='0' field=%s field-type='quantitative' max='%s' min='%s' "
                    "range-type='fixed' scope=%s type='space' />" % (qa(inst.full), hi, lo, qa(scope)))
    if not show:
        ws.rule('axis', fmt('display', 'false', class_='0', field=inst.full, scope=scope))
    if title_text is not None:
        ws.rule('axis', fmt('title', title_text, class_='0', field=inst.full, scope=scope))


def axis_title(ws, inst, text, scope):
    ws.rule('axis', fmt('title', text, class_='0', field=inst.full, scope=scope))


def tooltip_country():
    return [run('<%s>' % a_country.full, 11, INK, SEMI), NL,
            run('<%s> · <%s>' % (a_region.full, a_year.full), 8, MUTED), NL, NL,
            run("Can't afford a healthy diet: ", 9, MUTED), run('<%s>' % diet.full, 9, INK, SEMI), NL,
            run('Undernourished (PoU): ', 9, MUTED), run('<%s>' % pou.full, 9, INK, SEMI)]


def kpi_sheet(name, value_runs, caption, extra_insts=(), filters_true=()):
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


def continuous_color(ws, inst, colors):
    ws.table_rules.append(style_rule('mark', "<encoding attr='color' field=%s type='custom-interpolated'>\n"
                                             "<color-palette custom='true' name='' type='ordered-sequential'>\n%s\n"
                                             "</color-palette>\n</encoding>"
                                     % (qa(inst.full), '\n'.join('<color>%s</color>' % c for c in colors))))


# ==========================================================================
# Dashboard 1: Where Is the Burden?
# ==========================================================================
d1f = I(c_d1)
k1n = kpi_sheet('D1 KPI Countries', [run('<%s>' % I(c_n, 'User').full, 22, INK, SEMI)],
                'countries in view | (year, region, country controls)', [I(c_n, 'User')], [d1f])
k1d = kpi_sheet('D1 KPI Median Diet', [run('<%s>' % I(c_diet_med, 'User').full, 22, RED, SEMI)],
                "can't afford a healthy diet | median country, not pop.-weighted",
                [I(c_diet_med, 'User')], [d1f])
k1p = kpi_sheet('D1 KPI Median PoU', [run('<%s>' % I(c_pou_med, 'User').full, 22, INK, SEMI)],
                'undernourished (PoU) | median country, not pop.-weighted',
                [I(c_pou_med, 'User')], [d1f])
k1h = kpi_sheet('D1 KPI Highest', [run('<%s>' % metric.full, 22, WINE, SEMI),
                                   run('  <%s>' % a_country.full, 10, INK, SEMI)],
                'highest on the selected metric | (selected year and filters)', [metric, a_country],
                [I(c_isyear), I(c_is_top)])

# Map (Metric parameter) ----------------------------------------------------
m = base(Worksheet(wb, 'D1 Map', title=title('Where the burden sits (selected metric)',
                                              'Darker = higher share. Switch the metric at the top right. Hover for details.')))
m.uses_map = True
m.use(iso, metric, diet, pou, a_country, a_region, a_year)
m.filter_true(d1f)
m.filter_exclude(iso, ['"TWN"'])   # not recognised by Tableau's geocoder
p = m.pane(mark='Multipolygon')
p.add('color', metric).add('lod', iso).add('tooltip', a_country).add('tooltip', a_region).add('tooltip', a_year)
p.add('tooltip', diet).add('tooltip', pou).add('geometry', GEO)
p.tooltip = tooltip_country()
m.rows = '[%s].[Latitude (generated)]' % DS
m.cols = '[%s].[Longitude (generated)]' % DS
m.rule('map', fmt('washout', '0.0'))
continuous_color(m, metric, MAP_SCALE)

# Region distribution: box (Q1-median-Q3 drawn as two Gantt halves) + one dot per country
half = I(c_half)
bstart = I(c_bstart, 'Min')
blen = I(c_blen, 'Min')
q1, qm, q3 = I(c_q1, 'Min'), I(c_qm, 'Min'), I(c_q3, 'Min')
mmed = I(c_metric_med, 'User')
bx = base(Worksheet(wb, 'D1 Region Distribution',
                    title=title('Distribution by region', 'Dot = country · box = middle 50%, shade changes at the median')))
bx.use(country, region, half, bstart, blen, q1, qm, q3, metric, mmed, a_country, a_region, a_year)
bx.filter_true(d1f)
bx.sort_by(region, mmed, 'DESC')
bx.rows = region.full
bx.cols = '(%s + %s)' % (bstart.full, metric.full)
bx.pane()
pb = bx.pane(mark='GanttBar', pid=1, x_axis=bstart)
pb.add('color', half).add('lod', half).add('size', blen).add('tooltip', q1).add('tooltip', qm).add('tooltip', q3)
pb.tooltip = [run('<%s>' % region.full, 11, INK, SEMI), NL,
              run('Q1 <%s>  ·  median <%s>  ·  Q3 <%s>' % (q1.full, qm.full, q3.full), 9, INK)]
pd = bx.pane(mark='Circle', pid=2, x_axis=metric)
pd.add('color', region).add('lod', country).add('tooltip', a_country).add('tooltip', a_region).add('tooltip', a_year)
pd.tooltip = [run('<%s>' % a_country.full, 11, INK, SEMI), NL, run('<%s> · <%s>' % (a_region.full, a_year.full), 8, MUTED),
              NL, run('Selected metric: ', 9, MUTED), run('<%s>' % metric.full, 9, INK, SEMI)]
pd.mark_fmt += [fmt('size', '0.45'), fmt('mark-transparency', '170')]
bx.rule('axis', "<encoding attr='space' class='0' field=%s field-type='quantitative' fold='true' scope='cols' "
                "synchronized='true' type='space' />" % qa(metric.full),
        fmt('display', 'false', class_='0', field=metric.full, scope='cols'))
axis_title(bx, bstart, '% (selected metric)', 'cols')
bx.rule('axis', fmt('text-format', 'n0"%"', class_='0', field=bstart.full, scope='cols'))
bx.rule('header', fmt('font-size', '8'), fmt('color', INK))

# Country ranking: Top 10 on the selected metric --------------------------
rk = base(Worksheet(wb, 'D1 Country Ranking', title=title('Top 10 countries (selected metric)')))
rk.use(country, metric, region, a_region)
rk.filter_true(d1f)
rk.filter_top(country, 10, internal_refs('MAX(IF [D1 in view] THEN [Selected metric] END)'))
rk.sort_by(country, metric, 'DESC')
rk.rows = country.full
rk.cols = metric.full
p = rk.pane(mark='Bar')
p.add('color', region).add('text', metric).add('tooltip', a_region)
p.label = [run('<%s>' % metric.full, 8, INK, SEMI)]
p.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('size', '0.6')]
rk.rule('axis', fmt('display', 'false', class_='0', field=metric.full, scope='cols'))
rk.rule('header', fmt('font-size', '8'), fmt('color', INK))

# ==========================================================================
# Dashboard 2: Related, but Not the Same (2023 only)
# ==========================================================================
d2f = I(c_d2)
k2n = kpi_sheet('D2 KPI n', [run('n = ', 16, SOFT, SEMI), run('<%s>' % I(c_n, 'User').full, 22, INK, SEMI)],
                'countries in the current sample | (2023, current controls)', [I(c_n, 'User')], [d2f])
k2r = kpi_sheet('D2 KPI r', [run('r = ', 16, SOFT, SEMI), run('<%s>' % I(c_r, 'User').full, 22, RED, SEMI)],
                'Pearson r for the current sample | all 0.762 · excl. floor 0.652',
                [I(c_r, 'User')], [d2f])
k2e = kpi_sheet('D2 KPI Exceptions', [run('<%s>' % I(c_exc_n, 'User').full, 22, GOLD, SEMI)],
                'exceptions: high on one measure | only (top quartile, 2023)', [I(c_exc_n, 'User')], [d2f])

grp = I(c_grp)
exc_lbl = I(c_exc_lbl)
sc = base(Worksheet(wb, 'D2 Scatter', title=title('Unaffordable diets and hunger move together, not in lock-step',
                                                  'Each dot is a country (2023) · line = linear trend · '
                                                  'gold / blue = exceptions, labelled')))
sc.use(country, grp, exc_lbl, diet, pou, a_country, a_region, a_year)
sc.filter_true(d2f)
p = sc.pane(mark='Circle')
p.add('color', grp).add('lod', country).add('text', exc_lbl).add('tooltip', a_country).add('tooltip', a_region)
p.add('tooltip', a_year)
p.tooltip = tooltip_country() + [NL, run('<%s>' % grp.full, 8, MUTED)]
p.label = [run('<%s>' % exc_lbl.full, 8, INK, SEMI)]
p.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('mark-labels-cull', 'true'), fmt('size', '0.7'),
               fmt('mark-transparency', '210')]
p.trend = True
sc.cols = diet.full
sc.rows = pou.full
axis_fixed(sc, diet, 0, 100, 'cols', title_text="Can't afford a healthy diet (%)")
axis_fixed(sc, pou, 0, 60, 'rows', title_text='Undernourishment, PoU (%)')
sc.rule('axis', fmt('text-format', 'n0"%"', class_='0', field=diet.full, scope='cols'),
        fmt('text-format', 'n0"%"', class_='0', field=pou.full, scope='rows'))

ex = base(Worksheet(wb, 'D2 Exceptions', title=title('The exceptions', 'High on one measure, not the other (2023)')))
ex.use(grp, country, diet, pou)
ex.filter_true(d2f)
ex.filter_true(I(c_is_exc))
ex.sort_by(country, diet, 'DESC')
ex.rows = '(%s / %s)' % (grp.full, country.full)
p = ex.pane(mark='Text')
p.add('text', diet).add('text', pou)
p.label = [run('diet <%s>   ·   PoU <%s>' % (diet.full, pou.full), 8, INK)]
p.mark_fmt.append(fmt('mark-labels-show', 'true'))
ex.rule('header', fmt('font-size', '8'), fmt('color', INK))
ex.rule('cell', fmt('text-align', 'left'))

# ==========================================================================
# Dashboard 3: Who Is Falling Behind?
# ==========================================================================
y23 = I(c_y23)
m17, m23 = I(c_med17, 'User'), I(c_med23, 'User')
k3m = kpi_sheet('D3 KPI Median Change', [run('<%s>' % m17.full, 18, SOFT, SEMI), run(' → ', 14, SOFT, BOOK),
                                         run('<%s>' % m23.full, 18, INK, SEMI)],
                "median healthy diet unaffordable | 2017 → 2023 (137 countries)", [m17, m23], [y23])
ki = I(c_improved, 'User')
k3i = kpi_sheet('D3 KPI Improved', [run('<%s>' % ki.full, 22, TEAL, SEMI)],
                'healthy diet more affordable | (improved ≥ 1 pp)', [ki], [y23])
kw = I(c_worsened, 'User')
k3w = kpi_sheet('D3 KPI Worsened', [run('<%s>' % kw.full, 22, RED, SEMI)],
                'healthy diet less affordable | (worsened ≥ 1 pp)', [kw], [y23])
kn = I(c_watch_n, 'User')
k3b = kpi_sheet('D3 KPI Watchlist', [run('<%s>' % kn.full, 22, WINE, SEMI)],
                'both elevated & both worsened | (high now, worse on both)', [kn], [y23])

# Change scatter --------------------------------------------------------------
quad = I(c_quad)
dchg = I(ds['Healthy_Diet_Change_pp'], 'Min')
pchg = I(ds['PoU_Change_pp'], 'Min')
cs = base(Worksheet(wb, 'D3 Change Scatter',
                    title=title('Change 2017 → 2023',
                                '↙ both improved (teal) · ↗ both worsened (wine) · other quadrants: one up, one down')),
          zero=True)
cs.use(country, quad, dchg, pchg, a_country, a_region)
cs.filter_true(y23)
p = cs.pane(mark='Circle')
p.add('color', quad).add('lod', country).add('tooltip', a_country).add('tooltip', a_region)
p.tooltip = [run('<%s>' % a_country.full, 11, INK, SEMI), NL, run('<%s>' % a_region.full, 8, MUTED), NL, NL,
             run('Healthy diet change: ', 9, MUTED), run('<%s>' % dchg.full, 9, INK, SEMI), NL,
             run('PoU change: ', 9, MUTED), run('<%s>' % pchg.full, 9, INK, SEMI), NL, run('<%s>' % quad.full, 8, MUTED)]
p.mark_fmt += [fmt('size', '0.6'), fmt('mark-transparency', '200')]
cs.cols = dchg.full
cs.rows = pchg.full
axis_title(cs, dchg, 'Healthy diet, change (pp)', 'cols')
axis_title(cs, pchg, 'PoU, change (pp)', 'rows')

# Burden x progress matrix --------------------------------------------------
bur, bur_o = I(c_burden), I(c_burden_o, 'Min')
prog, prog_o = I(c_prog), I(c_prog_o, 'Min')
cnt = I(ds['Country'], 'CountD')
mx = base(Worksheet(wb, 'D3 Matrix', title=title('Current burden × progress',
                                                 'Countries per cell · click a cell to list them below · Worse / Better = both measures, Flat = small change')))
mx.use(bur, bur_o, prog, prog_o, cnt)
mx.filter_true(y23)
mx.sort_by(bur, bur_o, 'ASC', 'rows')
mx.sort_by(prog, prog_o, 'ASC', 'columns')
mx.rows = bur.full
mx.cols = prog.full
p = mx.pane(mark='Square')
p.add('color', cnt).add('text', cnt)
p.label = [run('<%s>' % cnt.full, 11, INK, SEMI)]
p.mark_fmt += [fmt('mark-labels-show', 'true')]
continuous_color(mx, cnt, ['#f7f1e8', '#e8c9a8', '#c98a6a'])
mx.rule('header', fmt('font-size', '8'), fmt('color', INK))

# Country detail (target of the matrix filter action) ----------------------
cd = base(Worksheet(wb, 'D3 Country Detail',
                    title=title('Country detail', 'Columns: healthy diet % and PoU %, 2017→2023 (change, pp) · '
                                                  '⚑ = PoU at 2.5% reporting floor')))
det_diet, det_pou = I(c_det_diet), I(c_det_pou)
floor = I(c_floor, 'Attribute')
cd.use(country, dchg, det_diet, det_pou, floor)
cd.filter_true(y23)
cd.sort_by(country, dchg, 'DESC')
cd.rows = '(%s / (%s / %s))' % (country.full, det_diet.full, det_pou.full)
cd.zoom = 'fit-width'
p = cd.pane(mark='Text')
p.add('text', floor)
p.label = [run('<%s>' % floor.full, 8, RED, SEMI)]
p.mark_fmt.append(fmt('mark-labels-show', 'true'))
cd.rule('header', fmt('font-size', '8'), fmt('color', INK))
cd.rule('cell', fmt('text-align', 'left'))

# One dumbbell, switched by the Show parameter -----------------------------
pt = I(c_point)
chg = I(ds['Healthy_Diet_Change_pp'], 'Avg')
chg_lbl = I(c_chg_lbl, 'Avg')
diet_line = I(c_diet_copy, 'Avg')
dsort = I(c_db_sort, 'Avg')
db = base(Worksheet(wb, 'D3 Dumbbell', title=title('Biggest setbacks / gains', 'Healthy diet unaffordable, 2017 → 2023')))
db.use(country, year, pt, I(c_db_show), diet, diet_line, chg, chg_lbl, dsort)
db.filter_members(year, ['2017', '2023'])
db.filter_true(I(c_db_show))
db.sort_by(country, dsort, 'ASC')
db.rows = country.full
db.cols = '(%s + %s)' % (diet.full, diet_line.full)
db.pane()
a = db.pane(mark='Circle', pid=1, x_axis=diet)
a.add('color', pt).add('lod', year).add('text', chg_lbl).add('tooltip', chg)
a.label = [run('<%s>' % chg_lbl.full, 8, INK, SEMI)]
a.tooltip = [run('<%s>' % country.full, 11, INK, SEMI), NL,
             run("<%s>: <%s> can't afford a healthy diet" % (year.full, diet.full), 9, MUTED), NL,
             run('Change 2017→2023: <%s>' % chg.full, 9, INK, SEMI)]
a.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('size', '0.9')]
b = db.pane(mark='Line', pid=2, x_axis=diet_line)
b.add('path', year)
b.mark_fmt += [fmt('mark-color', '#d9d2c6'), fmt('size', '2.5')]
db.rule('axis', "<encoding attr='space' class='0' field=%s field-type='quantitative' fold='true' scope='cols' "
                "synchronized='true' type='space' />" % qa(diet_line.full),
        fmt('display', 'false', class_='0', field=diet_line.full, scope='cols'))
axis_fixed(db, diet, 0, 100, 'cols', show=False)
db.rule('header', fmt('font-size', '8'), fmt('color', INK))

# Regional trend (Regional file: Country x Year) ---------------------------
yr_c = I(c_year_num)
end_r, end_v = I(c_end_region, 'User'), I(c_end_val, 'User')
diet_med = I(c_diet_med, 'User')
tr = base(Worksheet(wb, 'D3 Regional Trend', title=title('Regional median, healthy diet')))
tr.use(yr_c, region, diet_med, end_r)
tr.filter_members(region, ['"Africa"', '"Americas"', '"Asia"', '"Europe"'])
tr.cols = yr_c.full
tr.rows = diet_med.full
p = tr.pane(mark='Line')
p.add('color', region).add('text', end_r)
p.label = [run(' <%s>' % end_r.full, 8, INK, SEMI)]
p.mark_fmt += [fmt('mark-labels-show', 'true'), fmt('size', '2.5')]
axis_fixed(tr, yr_c, 2016.5, 2024.8, 'cols', title_text='')
axis_fixed(tr, diet_med, 0, 70, 'rows', title_text='')
tr.rule('axis', fmt('text-format', 'n0', field=yr_c.full, scope='cols'))

# ==========================================================================
# Dashboards
# ==========================================================================
M = 32                       # outer margin
GAP = 16
HEAD = 116                   # header band height
KPI_Y, KPI_H = HEAD + 14, 100
BODY = KPI_Y + KPI_H + 14
BOTTOM = 764


def header(db_, kicker, headline, sub, text_w=1030):
    db_.box(0, 0, W, HEAD, NAVY)
    db_.text(M - 4, 8, text_w, HEAD - 12, [run(kicker, 8, GOLD, SEMI), NL, run(headline, 15, '#ffffff', SEMI), NL,
                                           run(sub, 9, '#c3cbe0', BOOK)], bg=NAVY)


def controls(db_, params, cols=2, w=200, h=48):
    """Parameter controls in the header band, right-aligned, `cols` per row."""
    x0 = W - M - cols * w - (cols - 1) * 8
    for i, p_ in enumerate(params):
        r, c = divmod(i, cols)
        db_.param(x0 + c * (w + 8), 8 + r * (h + 4), w, h, p_, bg='#ffffff')


def kpis(db_, items, colors, y=KPI_Y, h=KPI_H):
    """items: worksheets, or ('text', runs) for a plain text card."""
    w = (W - 2 * M - 3 * GAP) / 4
    for i, (ws, c) in enumerate(zip(items, colors)):
        x = M + i * (w + GAP)
        db_.box(x, y, w, h, CARD, LINE)
        db_.box(x, y, w, 4, c)
        if isinstance(ws, tuple):
            db_.text(x + 14, y + 12, w - 28, h - 16, ws[1], bg=CARD)
        elif ws.caption is None:
            db_.sheet(x + 10, y + 8, w - 20, h - 12, ws, show_title=False, bg=CARD)
        else:
            db_.sheet(x + 10, y + 6, w - 20, 38, ws, show_title=False, bg=CARD)
            lines = [l.strip() for l in ws.caption.split('|')]
            runs = []
            for j, l in enumerate(lines):
                runs += ([NL] if j else []) + [run(l, 8, MUTED, BOOK)]
            db_.text(x + 14, y + 42, w - 28, h - 44, runs, bg=CARD)


def card(db_, x, y, w, h, ws):
    db_.box(x, y, w, h, CARD, LINE)
    db_.sheet(x + 12, y + 10, w - 24, h - 18, ws, bg=CARD)


def footer(db_, text):
    db_.box(M, BOTTOM + 10, W - 2 * M, 1, LINE)
    a_, b_ = text.split(' | ')
    db_.text(M - 4, BOTTOM + 13, W - 2 * M + 8, H - BOTTOM - 15, [run(a_, 7, MUTED, BOOK), NL, run(b_, 7, MUTED, BOOK)])


SOURCE = ("Source: FAO, IFAD, UNICEF, WFP & WHO (SOFI, via Our World in Data), 2017–2023; 2024 diet data excluded. "
          "Country medians, not population-weighted. Taiwan not mapped. | "
          "PoU ≤ 2.5% is reported as 2.5% (reporting floor). 'High' = top quartile of 2023 "
          "(diet ≥ 55.2%, PoU ≥ 13.1%), not a UN threshold. pp = percentage points.")

# --- D1 ---------------------------------------------------------------------
d1 = Dashboard(wb, D1N, W, H, PAPER)
header(d1, 'FOOD SECURITY 2017–2023  ·  1 / 3  ·  WHERE IS THE BURDEN?',
       "Median country, 2023: 24.5% can't afford a healthy diet",
       'Country medians, not a population-weighted global share. Heaviest burden: Africa.',
       text_w=W - 2 * M - 2 * 200 - 24)
controls(d1, [year_p, metric_p, region_p, country_p])
kpis(d1, [k1n, k1d, k1p, k1h], [INK, RED, INK, WINE])
MW = 640
card(d1, M, BODY, MW, BOTTOM - BODY, m)
RX1 = M + MW + GAP
RW1 = W - M - RX1
TOP_H = 290
card(d1, RX1, BODY, RW1, TOP_H, bx)
card(d1, RX1, BODY + TOP_H + GAP, RW1, BOTTOM - BODY - TOP_H - GAP, rk)
footer(d1, SOURCE)

# --- D2 ---------------------------------------------------------------------
d2 = Dashboard(wb, D2N, W, H, PAPER)
header(d2, 'FOOD SECURITY 2023  ·  2 / 3  ·  RELATED, BUT NOT THE SAME',
       "Diet costs and hunger mostly go together",
       'All countries n = 138, r ≈ 0.762 · excluding the 2.5% PoU floor n = 80, r ≈ 0.652',
       text_w=W - 2 * M - 2 * 230 - 24)
controls(d2, [sample_p, region_p], cols=2, w=230)
kpis(d2, [k2n, k2r, k2e,
          ('text', [run('Association, not causation.', 10, INK, SEMI), NL,
                    run('Related, but not the same thing:', 8, MUTED), NL,
                    run('diet cost vs. calorie deficit.', 8, MUTED)])],
     [INK, RED, GOLD, LINE])
SW = 820
card(d2, M, BODY, SW, BOTTOM - BODY, sc)
card(d2, M + SW + GAP, BODY, W - 2 * M - SW - GAP, BOTTOM - BODY, ex)
footer(d2, SOURCE)

# --- D3 ---------------------------------------------------------------------
d3 = Dashboard(wb, D3N, W, H, PAPER)
header(d3, 'FOOD SECURITY 2017–2023  ·  3 / 3  ·  WHO IS FALLING BEHIND?',
       'Diets got more affordable in 63 countries; 38 slipped back',
       '137 countries with both years · 9 are high on both measures and worse on both',
       text_w=W - 2 * M - 220 - 16)
controls(d3, [db_p], cols=1, w=220)
kpis(d3, [k3m, k3i, k3w, k3b], [INK, TEAL, RED, WINE])
CW = 360
card(d3, M, BODY, CW, BOTTOM - BODY, cs)
MX = M + CW + GAP
MWD = 470
MXH = 200
card(d3, MX, BODY, MWD, MXH, mx)
card(d3, MX, BODY + MXH + GAP, MWD, BOTTOM - BODY - MXH - GAP, cd)
RX3 = MX + MWD + GAP
RW3 = W - M - RX3
DBH = 300
card(d3, RX3, BODY, RW3, DBH, db)
card(d3, RX3, BODY + DBH + GAP, RW3, BOTTOM - BODY - DBH - GAP, tr)
footer(d3, SOURCE)

# Matrix cell -> Country Detail (filter action; clearing the selection shows all) ----
_others = [ws.name for ws in (k3m, k3i, k3w, k3b, cs, mx, db, tr)]
wb.actions.append(
    "<action caption='Matrix cell → country detail' name='[Action1_MatrixToDetail]'>\n"
    "<activation auto-clear='true' type='on-select' />\n"
    "<source dashboard=%s type='sheet' worksheet=%s />\n"
    "<command command='tsc:tsl-filter'>\n"
    "<param name='exclude' value=%s />\n"
    "<param name='special-fields' value='all' />\n"
    "<param name='target' value=%s />\n"
    "</command>\n</action>" % (qa(D3N), qa(mx.name), qa(','.join(_others)), qa(D3N)))

xml = wb.xml(active=ACTIVE)

# Reuse the packaged extract when tableauhyperapi is not installed (the CSV columns are unchanged).
hyper = None
try:
    import tableauhyperapi  # noqa: F401
except ImportError:
    src = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'FoodSecurity_Dashboard_v2_backup.twbx')
    with zipfile.ZipFile(src) as z:
        hyper = os.path.join(tempfile.mkdtemp(), 'food_security.hyper')
        with open(hyper, 'wb') as fh:
            fh.write(z.read(ds.hyper_path))
package_twbx(xml, OUT, 'FoodSecurity_Dashboard.twb', ds, os.path.join(ROOT, CSV),
             os.path.dirname(os.path.abspath(OUT)), hyper=hyper)
print('wrote', OUT)
