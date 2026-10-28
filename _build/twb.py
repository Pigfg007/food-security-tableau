"""Small helper library for generating Tableau .twb XML from Python.

Only the subset of the Tableau XML schema used by build_dashboard.py is covered.
"""
import uuid
from xml.sax.saxutils import escape, quoteattr

DS = 'federated.foodsec2017to2023'


def q(s):
    """XML attribute value (single quoted like Tableau writes it)."""
    s = escape(str(s), {"'": '&apos;', '"': '&quot;'})
    return "'" + s + "'"


def uid():
    return '{' + str(uuid.uuid4()).upper() + '}'


def esc(s):
    return escape(str(s))


# --------------------------------------------------------------------------
# Fields
# --------------------------------------------------------------------------
class Field:
    """A datasource column (raw or calculated)."""

    def __init__(self, name, datatype, role, ftype, caption=None, formula=None,
                 semantic=None, default_format=None):
        self.name = name              # e.g. 'Year' or 'Calculation_101'
        self.datatype = datatype      # string, integer, real, boolean
        self.role = role              # dimension / measure
        self.ftype = ftype            # nominal / ordinal / quantitative
        self.caption = caption
        self.formula = formula
        self.semantic = semantic
        self.default_format = default_format

    @property
    def ref(self):
        return '[' + self.name + ']'

    def xml(self):
        a = []
        if self.caption:
            a.append('caption=' + q(self.caption))
        a.append('datatype=' + q(self.datatype))
        if self.default_format:
            a.append('default-format=' + q(self.default_format))
        a.append('name=' + q(self.ref))
        a.append('role=' + q(self.role))
        if self.semantic:
            a.append('semantic-role=' + q(self.semantic))
        a.append('type=' + q(self.ftype))
        head = '<column ' + ' '.join(a)
        if self.formula is not None:
            return (head + '>\n  <calculation class=\'tableau\' formula=' + q(self.formula)
                    + ' />\n</column>')
        return head + ' />'


DERIV_PREFIX = {
    'None': 'none', 'Sum': 'sum', 'Avg': 'avg', 'Median': 'med', 'Min': 'min',
    'Max': 'max', 'CountD': 'ctd', 'Count': 'cnt', 'Attribute': 'attr', 'User': 'usr',
}


class Inst:
    """A column-instance (field + aggregation) used on a shelf."""

    def __init__(self, field, derivation='None', kind=None):
        self.field = field
        self.derivation = derivation
        if kind is None:
            if derivation == 'None':
                kind = {'nominal': 'nk', 'ordinal': 'ok', 'quantitative': 'qk'}[field.ftype]
                if field.role == 'measure':
                    kind = 'qk'
            else:
                kind = 'nk' if (field.datatype == 'string' and derivation == 'User') else 'qk'
                if derivation == 'Attribute':
                    kind = {'nominal': 'nk', 'ordinal': 'ok', 'quantitative': 'qk'}[field.ftype]
        self.kind = kind

    @property
    def iname(self):
        return '[' + DERIV_PREFIX[self.derivation] + ':' + self.field.name + ':' + self.kind + ']'

    @property
    def full(self):
        return '[' + DS + '].' + self.iname

    @property
    def itype(self):
        return {'nk': 'nominal', 'ok': 'ordinal', 'qk': 'quantitative'}[self.kind]

    def xml(self):
        return ('<column-instance column=' + q(self.field.ref) + ' derivation=' + q(self.derivation)
                + ' name=' + q(self.iname) + ' pivot=\'key\' type=' + q(self.itype) + ' />')

    def __str__(self):
        return self.full


def P(*parts):
    """Plain text formatted-text run helper -> list of (text, attrs)."""
    return list(parts)


def run(text, size=None, color=None, font=None, bold=False, align=None, italic=False):
    a = {}
    if bold:
        a['bold'] = 'true'
    if align is not None:
        a['fontalignment'] = str(align)
    if color:
        a['fontcolor'] = color
    if font:
        a['fontname'] = font
    if size:
        a['fontsize'] = str(size)
    if italic:
        a['italic'] = 'true'
    return (text, a)


NL = ('Æ\n', {})


def formatted_text(runs, cdata=False):
    out = ['<formatted-text>']
    for text, a in runs:
        attrs = ''.join(' ' + k + '=' + q(v) for k, v in sorted(a.items()))
        if cdata:
            body = '<![CDATA[' + text + ']]>'
        else:
            body = esc(text).replace('\n', '&#10;')
        out.append('<run' + attrs + '>' + body + '</run>')
    out.append('</formatted-text>')
    return '\n'.join(out)


def fmt(attr, value, **kw):
    a = ''.join(' ' + k.rstrip('_').replace('_', '-') + '=' + q(v) for k, v in kw.items())
    return '<format attr=' + q(attr) + a + ' value=' + q(value) + ' />'


def style_rule(element, *items):
    return '<style-rule element=' + q(element) + '>\n' + '\n'.join(items) + '\n</style-rule>'
