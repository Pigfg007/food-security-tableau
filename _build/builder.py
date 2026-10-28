"""Workbook / worksheet / dashboard builders on top of twb.py."""
import csv
import os
from twb import (DS, Field, Inst, q, uid, esc, formatted_text, fmt, style_rule, run, NL)


# --------------------------------------------------------------------------
# Datasource
# --------------------------------------------------------------------------
class Datasource:
    def __init__(self, csv_path, caption, directory, filename):
        self.caption = caption
        self.directory = directory
        self.filename = filename
        self.table = os.path.splitext(filename)[0]
        with open(csv_path, newline='', encoding='utf-8') as f:
            r = csv.reader(f)
            header = next(r)
            rows = list(r)
        self.columns = []   # (name, datatype)
        for i, h in enumerate(header):
            vals = [row[i] for row in rows if row[i] != '']
            dt = 'string'
            if vals and all(_is_int(v) for v in vals):
                dt = 'integer'
            elif vals and all(_is_float(v) for v in vals):
                dt = 'real'
            self.columns.append((h, dt))
        self.fields = {}
        self.calcs = []
        self.color_maps = []   # (field_inst_name, {value: color})
        for name, dt in self.columns:
            if dt == 'real':
                f = Field(name, dt, 'measure', 'quantitative')
            elif dt == 'integer':
                f = Field(name, dt, 'dimension', 'ordinal')
            else:
                f = Field(name, dt, 'dimension', 'nominal')
            self.fields[name] = f
        self._calc_n = 1000
        self.hyper_path = 'Data/Extracts/food_security.hyper'

    def __getitem__(self, k):
        return self.fields[k]

    def calc(self, caption, formula, datatype, role='measure', ftype=None, default_format=None):
        self._calc_n += 1
        if ftype is None:
            ftype = 'quantitative' if role == 'measure' else 'nominal'
        f = Field('Calculation_%d' % self._calc_n, datatype, role, ftype, caption=caption,
                  formula=formula, default_format=default_format)
        self.calcs.append(f)
        self.fields[caption] = f
        return f

    def all_fields(self):
        return [self.fields[n] for n, _ in self.columns] + self.calcs

    def xml(self):
        tid = self.table + '_0F00D5EC0000000000000000000000AA'
        cols = '\n'.join("<column datatype=%s name=%s ordinal='%d' />" % (q(dt), q(n), i)
                         for i, (n, dt) in enumerate(self.columns))
        relation = ("<relation connection='textscan.foodsec01' name=%s table=%s type='table'>\n"
                    "<columns character-set='UTF-8' header='yes' locale='en_US' separator=','>\n%s\n"
                    "</columns>\n</relation>") % (q(self.filename), q('[' + self.table + '#csv]'), cols)
        md = []
        md.append("""<metadata-record class='capability'>
<remote-name />
<remote-type>0</remote-type>
<parent-name>[%s]</parent-name>
<remote-alias />
<aggregation>Count</aggregation>
<contains-null>true</contains-null>
<attributes>
<attribute datatype='string' name='character-set'>&quot;UTF-8&quot;</attribute>
<attribute datatype='string' name='collation'>&quot;en_US&quot;</attribute>
<attribute datatype='string' name='field-delimiter'>&quot;,&quot;</attribute>
<attribute datatype='string' name='header-row'>&quot;true&quot;</attribute>
<attribute datatype='string' name='locale'>&quot;en_US&quot;</attribute>
<attribute datatype='string' name='single-char'>&quot;&quot;</attribute>
</attributes>
</metadata-record>""" % esc(self.filename))
        rt = {'string': 129, 'integer': 20, 'real': 5}
        agg = {'string': 'Count', 'integer': 'Sum', 'real': 'Sum'}
        for i, (n, dt) in enumerate(self.columns):
            extra = "<collation flag='0' name='LEN_RUS' />\n" if dt == 'string' else ''
            md.append("""<metadata-record class='column'>
<remote-name>%s</remote-name>
<remote-type>%d</remote-type>
<local-name>[%s]</local-name>
<parent-name>[%s]</parent-name>
<remote-alias>%s</remote-alias>
<ordinal>%d</ordinal>
<local-type>%s</local-type>
<aggregation>%s</aggregation>
<contains-null>true</contains-null>
%s<object-id>[%s]</object-id>
</metadata-record>""" % (esc(n), rt[dt], esc(n), esc(self.filename), esc(n), i, dt, agg[dt], extra, esc(tid)))
        out = []
        out.append("<datasource caption=%s inline='true' name=%s version='18.1'>" % (q(self.caption), q(DS)))
        out.append("<connection class='federated'>\n<named-connections>")
        out.append("<named-connection caption=%s name='textscan.foodsec01'>" % q(self.table))
        out.append("<connection class='textscan' directory=%s filename=%s password='' server='' />"
                   % (q(self.directory), q(self.filename)))
        out.append("</named-connection>\n</named-connections>")
        out.append(relation)
        out.append('<metadata-records>\n' + '\n'.join(md) + '\n</metadata-records>')
        out.append('</connection>')
        out.append("<aliases enabled='yes' />")
        for f in self.all_fields():
            raw_default = (f.formula is None and f.semantic is None and f.caption is None
                           and f.default_format is None)
            if not raw_default or f.formula is not None:
                out.append(f.xml())
        out.append("<column caption=%s datatype='table' name=%s role='measure' type='quantitative' />"
                   % (q(self.filename), q('[__tableau_internal_object_id__].[' + tid + ']')))
        for inst, _ in self.color_maps:
            out.append(inst.xml())
        out.append(self.extract_xml(tid))
        out.append("<layout dim-ordering='alphabetic' measure-ordering='alphabetic' show-structure='true' />")
        if self.color_maps:
            out.append('<style>\n<style-rule element=\'mark\'>')
            for inst, m in self.color_maps:
                out.append("<encoding attr='color' field=%s type='palette'>" % q(inst.iname))
                for val, col in m.items():
                    out.append("<map to=%s>\n<bucket>%s</bucket>\n</map>" % (q(col), esc(val)))
                out.append('</encoding>')
            out.append('</style-rule>\n</style>')
        out.append("<semantic-values>\n<semantic-value key='[Country].[Name]' value='&quot;Singapore&quot;' />\n</semantic-values>")
        out.append("<object-graph>\n<objects>\n<object caption=%s id=%s>\n<properties context=''>\n%s\n</properties>\n<properties context='extract'>\n<relation name='Extract' table='[Extract].[Extract]' type='table' />\n</properties>\n</object>\n</objects>\n</object-graph>"
                   % (q(self.filename), q(tid), relation))
        out.append('</datasource>')
        return '\n'.join(out)


    def extract_xml(self, tid):
        rt = {'string': 129, 'integer': 20, 'real': 5}
        agg = {'string': 'Count', 'integer': 'Sum', 'real': 'Sum'}
        md = []
        for i, (n, dt) in enumerate(self.columns):
            extra = "<collation flag='0' name='LEN_RUS' />\n" if dt == 'string' else ''
            md.append("""<metadata-record class='column'>
<remote-name>%s</remote-name>
<remote-type>%d</remote-type>
<local-name>[%s]</local-name>
<parent-name>[Extract]</parent-name>
<remote-alias>%s</remote-alias>
<ordinal>%d</ordinal>
<family>%s</family>
<local-type>%s</local-type>
<aggregation>%s</aggregation>
<contains-null>true</contains-null>
%s<object-id>[%s]</object-id>
</metadata-record>""" % (esc(n), rt[dt], esc(n), esc(n), i, esc(self.filename), dt, agg[dt], extra, esc(tid)))
        return ("<extract count='-1' enabled='true' object-id='' units='records'>\n"
                "<connection access_mode='readonly' author-locale='en_US' class='hyper' dbname=%s default-settings='hyper' "
                "schema='Extract' sslmode='' tablename='Extract' update-time='10/15/2026 09:00:00 AM' username='tableau_internal_user'>\n"
                "<relation name='Extract' table='[Extract].[Extract]' type='table' />\n<metadata-records>\n%s\n</metadata-records>\n"
                "</connection>\n</extract>") % (q(self.hyper_path), '\n'.join(md))

    def write_hyper(self, csv_path, out_path):
        from tableauhyperapi import (HyperProcess, Telemetry, Connection, CreateMode, TableDefinition,
                                     SqlType, TableName, Inserter)
        types = {'string': SqlType.text(), 'integer': SqlType.big_int(), 'real': SqlType.double()}
        conv = {'string': lambda v: v, 'integer': int, 'real': float}
        table = TableDefinition(TableName('Extract', 'Extract'),
                                [TableDefinition.Column(n, types[dt]) for n, dt in self.columns])
        with open(csv_path, newline='', encoding='utf-8') as f:
            r = csv.reader(f)
            next(r)
            rows = [[None if v == '' else conv[dt](v) for v, (n, dt) in zip(row, self.columns)] for row in r]
        if os.path.exists(out_path):
            os.remove(out_path)
        with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU, parameters={'log_dir': os.path.dirname(out_path)}) as hp:
            with Connection(hp.endpoint, out_path, CreateMode.CREATE_AND_REPLACE) as c:
                c.catalog.create_schema('Extract')
                c.catalog.create_table(table)
                with Inserter(c, table) as ins:
                    ins.add_rows(rows)
                    ins.execute()


def _is_int(v):
    try:
        int(v)
        return True
    except ValueError:
        return False


def _is_float(v):
    try:
        float(v)
        return True
    except ValueError:
        return False


# --------------------------------------------------------------------------
# Parameters
# --------------------------------------------------------------------------
class Parameter:
    def __init__(self, name, caption, datatype, value, members, alias_map=None):
        self.name = name
        self.caption = caption
        self.datatype = datatype
        self.value = value
        self.members = members
        self.alias_map = alias_map

    @property
    def full(self):
        return '[Parameters].[' + self.name + ']'

    def xml(self):
        ptype = 'quantitative'   # what the working v2 used; 'nominal' made every sheet invalid
        out = ["<column caption=%s datatype=%s name=%s param-domain-type='list' role='measure' type=%s value=%s>"
               % (q(self.caption), q(self.datatype), q('[' + self.name + ']'), q(ptype), q(self.value)),
               "<calculation class='tableau' formula=%s />" % q(self.value), '<members>']
        for m in self.members:
            out.append('<member value=%s />' % q(m))
        out.append('</members>\n</column>')
        return '\n'.join(out)


# --------------------------------------------------------------------------
# Worksheet
# --------------------------------------------------------------------------
class Pane:
    def __init__(self, mark='Automatic', pid=None, x_axis=None, y_axis=None, x_index=None,
                 y_index=None):
        self.mark = mark
        self.pid = pid
        self.x_axis = x_axis
        self.y_axis = y_axis
        self.x_index = x_index
        self.y_index = y_index
        self.enc = []           # (kind, inst)
        self.label = None       # runs
        self.tooltip = None     # runs
        self.mark_fmt = []      # format xml strings for element mark
        self.trend = False
        self.extra_rules = []

    def add(self, kind, inst):
        self.enc.append((kind, inst))
        return self

    def xml(self):
        a = ''
        if self.pid is not None:
            a += " id='%d'" % self.pid
        a += " selection-relaxation-option='selection-relaxation-allow'"
        if self.x_axis is not None:
            a += ' x-axis-name=' + q(self.x_axis.full)
        if self.x_index is not None:
            a += " x-index='%d'" % self.x_index
        if self.y_axis is not None:
            a += ' y-axis-name=' + q(self.y_axis.full)
        if self.y_index is not None:
            a += " y-index='%d'" % self.y_index
        out = ['<pane' + a + '>', "<view>\n<breakdown value='auto' />\n</view>",
               '<mark class=%s />' % q(self.mark)]
        if self.enc:
            out.append('<encodings>')
            for kind, inst in self.enc:
                col = inst if isinstance(inst, str) else inst.full
                out.append('<%s column=%s />' % (kind, q(col)))
            out.append('</encodings>')
        if self.trend:
            out.append("<trendline enable-confidence-bands='false' enable-instant-analytics='true' enabled='true' exclude-color='true' exclude-intercept='false' fit='linear' />")
        if self.tooltip:
            out.append('<customized-tooltip>\n' + formatted_text(self.tooltip, cdata=True) + '\n</customized-tooltip>')
        if self.label:
            out.append('<customized-label>\n' + formatted_text(self.label, cdata=True) + '\n</customized-label>')
        rules = []
        if self.mark_fmt:
            rules.append(style_rule('mark', *self.mark_fmt))
        rules += self.extra_rules
        if rules:
            out.append('<style>\n' + '\n'.join(rules) + '\n</style>')
        out.append('</pane>')
        return '\n'.join(out)


class Worksheet:
    def __init__(self, wb, name, title=None):
        self.wb = wb
        self.name = name
        self.title = title       # runs or None
        self.rows = ''
        self.cols = ''
        self.panes = []
        self.filters = []        # xml strings
        self.slices = []         # inst full names
        self.sorts = []          # xml strings
        self.table_rules = []    # style-rule xml strings
        self.insts = {}
        self.uses_map = False
        self.extra_view = []
        self.annotations = []
        self.zoom = 'entire-view'
        self.rules = {}
        wb.worksheets.append(self)

    def rule(self, element, *items):
        self.rules.setdefault(element, []).extend(items)
        return self

    def use(self, *insts):
        for i in insts:
            if isinstance(i, Inst):
                self.insts[i.iname] = i
        return insts[0] if len(insts) == 1 else insts

    def pane(self, **kw):
        p = Pane(**kw)
        self.panes.append(p)
        return p

    # -- filters --------------------------------------------------------
    def filter_members(self, inst, members):
        self.use(inst)
        if len(members) == 1:
            m = members[0]
            body = ("<groupfilter function='member' level=%s member=%s user:ui-domain='database' "
                    "user:ui-enumeration='inclusive' user:ui-marker='enumerate' />") % (q(inst.iname), q(m))
        else:
            inner = '\n'.join("<groupfilter function='member' level=%s member=%s />" % (q(inst.iname), q(m))
                              for m in members)
            body = ("<groupfilter function='union' user:ui-domain='database' user:ui-enumeration='inclusive' "
                    "user:ui-marker='enumerate'>\n%s\n</groupfilter>") % inner
        self.filters.append("<filter class='categorical' column=%s>\n%s\n</filter>" % (q(inst.full), body))
        self.slices.append(inst.full)

    def filter_exclude(self, inst, members):
        self.use(inst)
        inner = '\n'.join("<groupfilter function='member' level=%s member=%s />" % (q(inst.iname), q(m)) for m in members)
        self.filters.append("<filter class='categorical' column=%s>\n<groupfilter function='except' user:ui-domain='database' "
                            "user:ui-enumeration='exclusive' user:ui-marker='enumerate'>\n"
                            "<groupfilter function='level-members' level=%s />\n%s\n</groupfilter>\n</filter>"
                            % (q(inst.full), q(inst.iname), inner))
        self.slices.append(inst.full)

    def filter_top(self, inst, n, expression, direction='DESC'):
        """Top-N filter on a dimension, ranked by an aggregate expression."""
        self.use(inst)
        self.filters.append(
            "<filter class='categorical' column=%s>\n"
            "<groupfilter count='%d' end='top' function='end' units='records' user:ui-marker='end' user:ui-top-by-field='true'>\n"
            "<groupfilter direction=%s expression=%s function='order' user:ui-marker='order'>\n"
            "<groupfilter function='level-members' level=%s user:ui-enumeration='all' user:ui-marker='enumerate' />\n"
            "</groupfilter>\n</groupfilter>\n</filter>"
            % (q(inst.full), n, q(direction), q(expression), q(inst.iname)))
        self.slices.append(inst.full)

    def filter_true(self, inst):
        self.filter_members(inst, ['true'])

    def sort_by(self, dim, measure, direction='DESC', shelf='rows'):
        self.use(dim, measure)
        self.sorts.append("<shelf-sort-v2 dimension-to-sort=%s direction=%s is-on-innermost-dimension='true' "
                          "measure-to-sort-by=%s shelf=%s />" % (q(dim.full), q(direction), q(measure.full), q(shelf)))

    def manual_sort(self, dim, order):
        self.use(dim)
        b = '\n'.join('<bucket>%s</bucket>' % esc('"' + v + '"') for v in order)
        self.extra_view.append("<sort class='manual' column=%s direction='ASC'>\n<dictionary>\n%s\n</dictionary>\n</sort>"
                               % (q(dim.full), b))

    def xml(self, ds, params=()):
        out = ['<worksheet name=%s>' % q(self.name)]
        if self.title:
            out.append('<layout-options>\n<title>\n' + formatted_text(self.title) + '\n</title>\n</layout-options>')
        out.append('<table>\n<view>')
        pds = "<datasource name='Parameters' />\n" if params else ''
        out.append('<datasources>\n%s<datasource caption=%s name=%s />\n</datasources>' % (pds, q(ds.caption), q(DS)))
        if self.uses_map:
            out.append("<mapsources>\n<mapsource name='Tableau' />\n</mapsources>")
        if params:
            out.append("<datasource-dependencies datasource='Parameters'>\n%s\n</datasource-dependencies>"
                       % '\n'.join(p.xml() for p in params))
        out.append('<datasource-dependencies datasource=%s>' % q(DS))
        for f in ds.all_fields():
            out.append(f.xml())
        for i in sorted(self.insts.values(), key=lambda x: x.iname):
            out.append(i.xml())
        out.append('</datasource-dependencies>')
        out += self.filters
        if self.sorts:
            out.append('<shelf-sorts>\n' + '\n'.join(self.sorts) + '\n</shelf-sorts>')
        out += self.extra_view
        if self.slices:
            out.append('<slices>\n' + '\n'.join('<column>%s</column>' % esc(s) for s in self.slices) + '\n</slices>')
        out.append("<aggregation value='true' />\n</view>")
        rules = [style_rule(el, *items) for el, items in self.rules.items()] + self.table_rules
        out.append('<style>\n' + '\n'.join(rules) + '\n</style>' if rules else '<style />')
        out.append('<panes>\n' + '\n'.join(p.xml() for p in self.panes) + '\n</panes>')
        out.append('<rows>%s</rows>' % esc(self.rows) if self.rows else '<rows />')
        out.append('<cols>%s</cols>' % esc(self.cols) if self.cols else '<cols />')
        if self.annotations:
            out.append('<annotations>\n' + '\n'.join(self.annotations) + '\n</annotations>')
        out.append('</table>')
        out.append('<simple-id uuid=%s />' % q(uid()))
        out.append('</worksheet>')
        return '\n'.join(out)

    def window_xml(self, hidden=False, maximized=False):
        h = " hidden='true'" if hidden else ''
        if maximized:
            h += " maximized='true'"
        return ("<window class='worksheet'%s name=%s>\n<cards>\n<edge name='left'>\n<strip size='160'>\n"
                "<card type='pages' />\n<card type='filters' />\n<card type='marks' />\n</strip>\n</edge>\n"
                "<edge name='top'>\n<strip size='2147483647'>\n<card type='columns' />\n</strip>\n"
                "<strip size='2147483647'>\n<card type='rows' />\n</strip>\n<strip size='31'>\n"
                "<card type='title' />\n</strip>\n</edge>\n</cards>\n<viewpoint>\n<zoom type=%s />\n</viewpoint>\n"
                "<simple-id uuid=%s />\n</window>"
                % (h, q(self.name), q(self.zoom), q(uid())))


# --------------------------------------------------------------------------
# Dashboard (absolute layout in pixels, converted to Tableau's 100000 units)
# --------------------------------------------------------------------------
class Dashboard:
    def __init__(self, wb, name, w, h, bg):
        self.wb = wb
        self.name = name
        self.w = w
        self.h = h
        self.bg = bg
        self.zones = []
        self._id = 2
        wb.dashboards.append(self)

    def _nid(self):
        self._id += 1
        return self._id

    def _geom(self, x, y, w, h):
        X = round(x * 100000 / self.w)
        Y = round(y * 100000 / self.h)
        W = round((x + w) * 100000 / self.w) - X
        H = round((y + h) * 100000 / self.h) - Y
        return "h='%d' w='%d' x='%d' y='%d'" % (H, W, X, Y)

    @staticmethod
    def _style(bg=None, border=None, bw=1, pad=0, margin=0):
        s = ['<zone-style>']
        if border:
            s.append(fmt('border-color', border))
            s.append(fmt('border-style', 'solid'))
            s.append(fmt('border-width', str(bw)))
        else:
            s.append(fmt('border-color', '#000000'))
            s.append(fmt('border-style', 'none'))
            s.append(fmt('border-width', '0'))
        s.append(fmt('margin', str(margin)))
        s.append(fmt('padding', str(pad)))
        if bg:
            s.append(fmt('background-color', bg))
        s.append('</zone-style>')
        return '\n'.join(s)

    def box(self, x, y, w, h, bg, border=None, bw=1):
        self.zones.append("<zone %s id='%d' type-v2='empty'>\n%s\n</zone>"
                          % (self._geom(x, y, w, h), self._nid(), self._style(bg, border, bw)))

    def text(self, x, y, w, h, runs, bg=None, pad=0):
        self.zones.append("<zone %s id='%d' type-v2='text'>\n%s\n%s\n</zone>"
                          % (self._geom(x, y, w, h), self._nid(), formatted_text(runs), self._style(bg, pad=pad)))

    def sheet(self, x, y, w, h, ws, show_title=True, bg=None, border=None, pad=0):
        st = '' if show_title else " show-title='false'"
        self.zones.append("<zone %s id='%d' name=%s%s>\n%s\n</zone>"
                          % (self._geom(x, y, w, h), self._nid(), q(ws.name), st, self._style(bg, border, pad=pad)))
        self.sheets_used = getattr(self, 'sheets_used', [])
        self.sheets_used.append(ws)

    def param(self, x, y, w, h, p, mode='compact', bg=None, show_title=True):
        st = '' if show_title else " show-title='false'"
        self.zones.append("<zone %s id='%d' mode=%s param=%s%s type-v2='paramctrl'>\n%s\n</zone>"
                          % (self._geom(x, y, w, h), self._nid(), q(mode), q(p.full), st, self._style(bg)))

    def xml(self):
        out = ['<dashboard name=%s>' % q(self.name)]
        out.append('<style>\n' + style_rule('table', fmt('background-color', self.bg)) + '\n</style>')
        out.append("<size maxheight='%d' maxwidth='%d' minheight='%d' minwidth='%d' />" % (self.h, self.w, self.h, self.w))
        out.append('<zones>')
        out.append("<zone h='100000' id='2' type-v2='layout-basic' w='100000' x='0' y='0'>")
        out += self.zones
        out.append(self._style(self.bg))
        out.append('</zone>')
        out.append('</zones>')
        out.append('<simple-id uuid=%s />' % q(uid()))
        out.append('</dashboard>')
        return '\n'.join(out)

    def window_xml(self, maximized=False):
        m = " maximized='true'" if maximized else ''
        vps = '\n'.join("<viewpoint name=%s>\n<zoom type=%s />\n</viewpoint>" % (q(ws.name), q(ws.zoom))
                        for ws in getattr(self, 'sheets_used', []))
        return ("<window class='dashboard'%s name=%s>\n<viewpoints>\n%s\n</viewpoints>\n<active id='-1' />\n"
                "<simple-id uuid=%s />\n</window>" % (m, q(self.name), vps, q(uid())))


# --------------------------------------------------------------------------
# Workbook
# --------------------------------------------------------------------------
class Workbook:
    def __init__(self, ds, params):
        self.ds = ds
        self.params = params
        self.worksheets = []
        self.dashboards = []
        self.wb_style = []
        self.actions = []

    def xml(self, active=None, hide_sheets=False):
        out = ["<?xml version='1.0' encoding='utf-8' ?>",
               "<workbook original-version='18.1' source-build='2026.2.1 (20262.26.0708.1337)' source-platform='win' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>",
               "<document-format-change-manifest>\n<AnimationOnByDefault />\n<IntuitiveSorting />\n<IntuitiveSorting_SP2 />\n<MapboxVectorStylesAndLayers />\n<MarkAnimation />\n<ObjectModelEncapsulateLegacy />\n<ObjectModelExtractV2 />\n<ObjectModelTableType />\n<SchemaViewerObjectModel />\n<SheetIdentifierTracking />\n<SortTagCleanup />\n<WindowsPersistSimpleIdentifiers />\n</document-format-change-manifest>",
               "<preferences>\n<preference name='ui.encoding.shelf.height' value='24' />\n<preference name='ui.shelf.height' value='26' />\n</preferences>"]
        if self.wb_style:
            out.append('<style>\n' + '\n'.join(self.wb_style) + '\n</style>')
        out.append('<datasources>')
        if self.params:
            out.append("<datasource hasconnection='false' inline='true' name='Parameters' version='18.1'>\n<aliases enabled='yes' />")
            for p in self.params:
                out.append(p.xml())
            out.append('</datasource>')
        out.append(self.ds.xml())
        out.append('</datasources>')
        out.append("<mapsources>\n<mapsource name='Tableau' />\n</mapsources>")
        if self.actions:
            out.append('<actions>\n' + '\n'.join(self.actions) + '\n</actions>')
        out.append('<worksheets>')
        for ws in self.worksheets:
            # No per-sheet Parameters dependencies: Tableau 2026.2 marks every sheet
            # "no valid data source" when they are written this way.
            out.append(ws.xml(self.ds))
        out.append('</worksheets>')
        out.append('<dashboards>')
        for d in self.dashboards:
            out.append(d.xml())
        out.append('</dashboards>')
        out.append("<windows source-height='44'>")
        for ws in self.worksheets:
            out.append(ws.window_xml(hidden=hide_sheets, maximized=(ws.name == active)))
        for d in self.dashboards:
            out.append(d.window_xml(maximized=(d.name == active)))
        out.append('</windows>')
        out.append('</workbook>')
        return '\n'.join(out)


def package_twbx(wb_xml, twbx_path, twb_name, ds, csv_path, workdir, hyper=None):
    """Write a .twbx: the .twb plus the .hyper extract and the source CSV.

    Pass `hyper` to reuse an existing extract (same columns) instead of
    regenerating it with tableauhyperapi.
    """
    import zipfile
    import tempfile
    if hyper is None:
        workdir = tempfile.mkdtemp()
        hyper = os.path.join(workdir, 'food_security.hyper')
        ds.write_hyper(csv_path, hyper)
    with zipfile.ZipFile(twbx_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr(twb_name, wb_xml.encode('utf-8'))
        z.write(hyper, ds.hyper_path)
        z.write(csv_path, ds.directory + '/' + ds.filename)
