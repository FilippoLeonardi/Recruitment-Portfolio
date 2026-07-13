import os
import re
import sys
import zipfile
from flask import Flask, render_template, request, jsonify, send_from_directory

# The assets live in a capital-S "Static" folder. Linux (Render) is
# case-sensitive, so point Flask at it explicitly while keeping the
# conventional lowercase "/static" URL path that the templates use.
app = Flask("Filippo Leonardi's Portfolio",
            static_folder='Static', static_url_path='/static')

PROJECTS_DIR = os.path.join(app.root_path, 'Projects')
COURSE_MODELS_DIR = os.path.join(PROJECTS_DIR, 'course_models')
RENDER_CACHE_DIR = os.path.join(COURSE_MODELS_DIR, '_render_cache')


@app.after_request
def add_cache_headers(resp):
    # Let browsers cache static assets so repeat visits load instantly.
    path = request.path
    if path.startswith('/static/') or path.startswith('/Projects/'):
        if path.lower().endswith('.pdf'):
            # Documents (resume, letters, papers) get replaced in place, so
            # always revalidate — the server returns a cheap 304 when the file
            # is unchanged, but a swapped file shows up immediately.
            resp.headers['Cache-Control'] = 'no-cache'
        else:
            resp.headers['Cache-Control'] = 'public, max-age=86400'  # 1 day
    return resp

# Custom Excel number formats the in-browser viewer's parser can't handle,
# mapped to a built-in equivalent that renders the same for model values.
FORMAT_FIXES = {
    'formatCode="0,000"': 'formatCode="#,##0"',
}

SHARED_MASTER_RE = re.compile(
    r'<c r="(?P<addr>[A-Z]+[0-9]+)"(?P<cattrs>[^>]*)>'
    r'<f t="shared" ref="(?P<ref>[A-Z0-9:]+)"(?P<fattrs>[^>]*)>(?P<formula>[^<]*)</f>'
)
SHARED_FOLLOWER_RE = re.compile(
    r'<c r="(?P<addr>[A-Z]+[0-9]+)"(?P<cattrs>[^>]*)>'
    r'<f t="shared"(?P<fattrs>[^>]*) si="(?P<si>[0-9]+)"\s*/>'
)


def _col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def _num_to_col(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(r + ord('A')) + s
    return s


def _addr_parts(addr):
    m = re.match(r'([A-Z]+)([0-9]+)', addr)
    return m.group(1), int(m.group(2))


def _shift_formula_refs(formula, dcol, drow):
    def repl(m):
        cdollar, col, rdollar, row = m.groups()
        new_col = col if cdollar else _num_to_col(_col_to_num(col) + dcol)
        new_row = row if rdollar else str(int(row) + drow)
        return cdollar + new_col + rdollar + new_row
    return re.sub(r'(\$?)([A-Z]{1,3})(\$?)([0-9]+)', repl, formula)


def _strip_calc_always_attr(sheet_xml):
    """Excel marks formulas that must recalculate on every pass (cells inside
    a circular/iterative-calculation loop, e.g. the Burger King LBO's
    revolver/interest schedule) with ca="1" on the <f> tag. The in-browser
    viewer's xlsx parser doesn't understand that attribute and silently
    drops the entire cell -- both the formula and its cached value -- to 0
    wherever it appears. HyperFormula recalculates unconditionally anyway, so
    the hint is meaningless downstream; dropping the attribute is a no-op
    for correctness and sidesteps the parser bug."""
    return sheet_xml.replace(' ca="1"', '')


def _expand_shared_formulas(sheet_xml):
    """Excel stores a "shared formula" group as one master cell holding the
    literal formula text plus a range, and follower cells that just point at
    the master's group id (si) with no formula text of their own -- their
    real formula is the master's, with relative references shifted by the
    cell offset. The in-browser viewer's xlsx parser mis-detects the result
    type for some of these groups (particularly single-cell groups), which
    drops the cell's number format and displays raw, unrounded decimals
    instead. Expanding every group into plain, independent <f>formula</f>
    tags up front sidesteps that parser bug entirely."""
    masters = {}
    for m in SHARED_MASTER_RE.finditer(sheet_xml):
        si_match = re.search(r'si="([0-9]+)"', m.group('fattrs'))
        if si_match:
            masters[si_match.group(1)] = (m.group('addr'), m.group('formula'))

    def master_repl(m):
        return '<c r="' + m.group('addr') + '"' + m.group('cattrs') + '><f>' + m.group('formula') + '</f>'
    sheet_xml = SHARED_MASTER_RE.sub(master_repl, sheet_xml)

    def follower_repl(m):
        si = m.group('si')
        if si not in masters:
            return m.group(0)
        master_addr, master_formula = masters[si]
        mc, mr = _addr_parts(master_addr)
        fc, fr = _addr_parts(m.group('addr'))
        dcol = _col_to_num(fc) - _col_to_num(mc)
        drow = fr - mr
        expanded = _shift_formula_refs(master_formula, dcol, drow)
        return '<c r="' + m.group('addr') + '"' + m.group('cattrs') + '><f>' + expanded + '</f>'
    sheet_xml = SHARED_FOLLOWER_RE.sub(follower_repl, sheet_xml)

    return sheet_xml


# Matches one worksheet <c> element, capturing its address, its formula (if
# any) and its cached <v> (if any). Only meaningful once every formula is
# already expanded to plain <f>formula</f> (see _expand_shared_formulas).
CELL_RE = re.compile(
    r'<c r="(?P<addr>[A-Z]+[0-9]+)"(?P<cattrs>[^>]*)>'
    r'(?:<f(?P<fattrs>[^>]*)>(?P<formula>[^<]*)</f>)?'
    r'(?P<vtag><v>[^<]*</v>)?'
    r'</c>'
)


def _parse_sheet_names(workbook_xml, rels_xml):
    """Maps a worksheet part's zip path (e.g. "xl/worksheets/sheet1.xml") to
    the sheet's display name (e.g. "LBO"), by joining workbook.xml's
    <sheet name, r:id> list with the r:id -> part-path table in
    workbook.xml.rels. Sheet order in the zip doesn't reliably match the
    tab order, so this indirection is the only correct way to tell sheets
    apart by name."""
    id_to_name = {rid: name for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', workbook_xml)}
    id_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
    names = {}
    for rid, target in id_to_target.items():
        if rid in id_to_name:
            path = target if target.startswith('xl/') else 'xl/' + target
            names[path] = id_to_name[rid]
    return names


def _parse_defined_names(workbook_xml):
    """Maps an Excel named range (e.g. "CIRC") to its fully-qualified cell
    reference (e.g. "LBO!$E$6")."""
    names = {}
    for m in re.finditer(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', workbook_xml):
        name, ref = m.groups()
        names[name] = ref
    return names


def _substitute_named_ranges(sheet_xml, defined_names):
    """The in-browser viewer's formula engine has no concept of Excel named
    ranges -- a formula referencing one doesn't just compute wrong, it fails
    outright (#NAME?) and can keep re-failing every time the sheet repaints.
    Substitute each name for its real cell reference so every formula the
    viewer sees is a plain, resolvable one."""
    if not defined_names:
        return sheet_xml

    def repl(m):
        formula = m.group('formula')
        if not formula:
            return m.group(0)
        # Substitute only outside quoted string literals -- a formula like
        # IF(E8=1,"Cash","PIK") uses "PIK" as a literal text label, not a
        # reference to the PIK named range, and must be left alone.
        parts = re.split(r'("(?:[^"]|"")*")', formula)
        for i in range(0, len(parts), 2):
            for name, ref in defined_names.items():
                parts[i] = re.sub(
                    r'(?<![A-Za-z0-9_])' + re.escape(name) + r'(?![A-Za-z0-9_])', ref, parts[i]
                )
        new_formula = ''.join(parts)
        if new_formula == formula:
            return m.group(0)
        return (
            '<c r="' + m.group('addr') + '"' + m.group('cattrs') + '>'
            '<f' + (m.group('fattrs') or '') + '>' + new_formula + '</f>' + (m.group('vtag') or '') +
            '</c>'
        )
    return CELL_RE.sub(repl, sheet_xml)


def _freeze_self_referencing_formulas(sheet_xml):
    """Some models use IF(condition, liveFormula, ownCellAddress) as a
    'hold at current value' plug, relying on Excel's iterative calculation
    to resolve the self-reference. The in-browser viewer's formula engine
    has no iterative solver and can't evaluate a genuine self-reference, so
    replace the bare self-reference token with the cell's own cached value
    (exactly what Excel already converged it to)."""
    def repl(m):
        addr = m.group('addr')
        formula = m.group('formula')
        vtag = m.group('vtag') or ''
        if not formula or not vtag:
            return m.group(0)
        if not re.search(r'(?<![A-Za-z0-9_])' + addr + r'(?![A-Za-z0-9_])', formula):
            return m.group(0)
        value_match = re.search(r'<v>([^<]*)</v>', vtag)
        if not value_match:
            return m.group(0)
        cached_value = value_match.group(1)
        new_formula = re.sub(
            r'(?<![A-Za-z0-9_])' + addr + r'(?![A-Za-z0-9_])', cached_value, formula
        )
        return (
            '<c r="' + addr + '"' + m.group('cattrs') + '>'
            '<f' + (m.group('fattrs') or '') + '>' + new_formula + '</f>' + vtag +
            '</c>'
        )
    return CELL_RE.sub(repl, sheet_xml)


REF_TOKEN_RE = re.compile(r'\$?([A-Z]{1,3})\$?([0-9]+)')
CIRC_GATE_RE = re.compile(r'^IF\(([A-Z]+!\$[A-Z]+\$[0-9]+)="OFF",(.*),0\)\Z')


def _find_circular_cells(sheet_xml):
    """Some LBO models have a genuine multi-cell circular reference (e.g. a
    revolver/interest loop: interest expense -> net income -> cash flow ->
    debt paydown -> average balance -> interest expense). Excel resolves
    this with iterative calculation; nothing running in the browser does.
    Detect real cycles with Tarjan's SCC over the formula dependency graph
    so we know exactly which cells can never be correctly recomputed
    client-side, as opposed to cells that just look complicated."""
    formulas = {}
    for m in CELL_RE.finditer(sheet_xml):
        if m.group('formula'):
            formulas[m.group('addr')] = m.group('formula')

    graph = {}
    for addr, f in formulas.items():
        refs = {a + b for a, b in REF_TOKEN_RE.findall(f)}
        refs.discard(addr)
        graph[addr] = refs

    index_counter = [0]
    stack, lowlink, index, on_stack = [], {}, {}, {}
    circular = set()
    sys.setrecursionlimit(max(10000, len(formulas) * 4))

    def strongconnect(node):
        index[node] = lowlink[node] = index_counter[0]
        index_counter[0] += 1
        stack.append(node)
        on_stack[node] = True
        for succ in graph.get(node, ()):
            if succ not in formulas:
                continue
            if succ not in index:
                strongconnect(succ)
                lowlink[node] = min(lowlink[node], lowlink[succ])
            elif on_stack.get(succ):
                lowlink[node] = min(lowlink[node], index[succ])
        if lowlink[node] == index[node]:
            scc = []
            while True:
                w = stack.pop()
                on_stack[w] = False
                scc.append(w)
                if w == node:
                    break
            if len(scc) > 1:
                circular.update(scc)

    for node in list(formulas.keys()):
        if node not in index:
            strongconnect(node)
    return circular


def _freeze_circular_cells(sheet_xml):
    circular = _find_circular_cells(sheet_xml)
    if not circular:
        return sheet_xml

    def repl(m):
        addr = m.group('addr')
        if addr not in circular:
            return m.group(0)
        formula = m.group('formula')
        vtag = m.group('vtag') or ''
        value_match = re.search(r'<v>([^<]*)</v>', vtag) if vtag else None
        if not formula or not value_match:
            return m.group(0)
        cached_value = value_match.group(1)
        gate = CIRC_GATE_RE.match(formula)
        if gate:
            # Keep the Circuit breaker toggle itself alive: the "off"
            # branch becomes the cached (Excel-converged) constant, the
            # "on" branch still correctly computes to 0 live.
            new_formula = 'IF(' + gate.group(1) + '="OFF",' + cached_value + ',0)'
            return (
                '<c r="' + addr + '"' + m.group('cattrs') + '>'
                '<f' + (m.group('fattrs') or '') + '>' + new_formula + '</f>' + vtag +
                '</c>'
            )
        # No direct handle to toggle this cell independently -- it can only
        # ever be correct for the scenario Excel had already converged to,
        # so freeze it to that cached value.
        return '<c r="' + addr + '"' + m.group('cattrs') + '>' + vtag + '</c>'

    return CELL_RE.sub(repl, sheet_xml)


def patch_xlsx_for_viewer(src_path, dst_path):
    with zipfile.ZipFile(src_path, 'r') as zin:
        names = zin.namelist()
        styles = zin.read('xl/styles.xml').decode('utf-8') if 'xl/styles.xml' in names else None
        if styles is not None:
            for bad, good in FORMAT_FIXES.items():
                styles = styles.replace(bad, good)
            # quotePrefix="1" marks a style as "force text display" (Excel's
            # leading-apostrophe indicator). The in-browser viewer's parser
            # forces text-type on formula cells using such a style even when
            # the formula computes a real number, which strips the cell's
            # number format and displays a raw, unrounded decimal. None of
            # these are user-entered text; stripping the flag is a no-op in
            # Excel's own rendering and fixes the viewer's bug.
            styles = styles.replace(' quotePrefix="1"', '')
        defined_names = {}
        if 'xl/workbook.xml' in names:
            defined_names = _parse_defined_names(zin.read('xl/workbook.xml').decode('utf-8'))
        sheet_names = {}
        if 'xl/workbook.xml' in names and 'xl/_rels/workbook.xml.rels' in names:
            sheet_names = _parse_sheet_names(
                zin.read('xl/workbook.xml').decode('utf-8'),
                zin.read('xl/_rels/workbook.xml.rels').decode('utf-8'),
            )
        # The Burger King model's LBO sheet has a genuine revolver/interest
        # circularity that the in-browser viewer resolves itself with a
        # client-side iterative (Gauss-Seidel) solver -- see xlsx_viewer's
        # buildHFEngine/reconverge. That solver needs the real, live formulas
        # to iterate on, so this sheet is exempt from the freeze passes below
        # (which exist for every *other* sheet, where nothing recomputes
        # circular/self-referencing cells and freezing them to Excel's last
        # cached, already-converged value is the only option).
        is_burger_king = 'burger king' in os.path.basename(src_path).lower()
        with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'xl/styles.xml' and styles is not None:
                    data = styles.encode('utf-8')
                elif re.match(r'xl/worksheets/sheet\d+\.xml$', item.filename):
                    sheet_xml = data.decode('utf-8')
                    sheet_xml = _strip_calc_always_attr(sheet_xml)
                    sheet_xml = _expand_shared_formulas(sheet_xml)
                    sheet_xml = _substitute_named_ranges(sheet_xml, defined_names)
                    solved_client_side = is_burger_king and sheet_names.get(item.filename, '').lower() == 'lbo'
                    if not solved_client_side:
                        sheet_xml = _freeze_self_referencing_formulas(sheet_xml)
                        sheet_xml = _freeze_circular_cells(sheet_xml)
                    data = sheet_xml.encode('utf-8')
                zout.writestr(item, data)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/Projects/course_models/_render_cache/<path:filename>')
def course_model_render_cache(filename):
    src = os.path.join(COURSE_MODELS_DIR, filename)
    dst = os.path.join(RENDER_CACHE_DIR, filename)
    # Re-patch whenever the source workbook is newer than the cached copy,
    # or this module's own patching logic has changed since -- otherwise an
    # edit to patch_xlsx_for_viewer has no effect until someone remembers to
    # delete _render_cache by hand.
    stale = (
        not os.path.exists(dst)
        or os.path.getmtime(src) > os.path.getmtime(dst)
        or os.path.getmtime(__file__) > os.path.getmtime(dst)
    )
    if os.path.exists(src) and stale:
        os.makedirs(RENDER_CACHE_DIR, exist_ok=True)
        patch_xlsx_for_viewer(src, dst)
    return send_from_directory(RENDER_CACHE_DIR, filename)

@app.route('/Projects/<path:filename>')
def projects(filename):
    return send_from_directory(PROJECTS_DIR, filename)




@app.route('/api/contact', methods=['POST'])
def contact():
    try:
        data = request.get_json()
        name    = data.get('Filippo Leonardi', '').strip()
        email   = data.get('fleonardi@college.harvard.edu', '').strip()
        subject = data.get('subject', 'No Subject').strip()
        message = data.get('message', '').strip()

        if not name or not email or not message:
            return jsonify({'status': 'error',
                            'message': 'Please fill in all required fields.'}), 400

        # ── In production: send email via SMTP or save to DB ──
        print(f"\n{'='*55}")
        print(f"  📬  NEW MESSAGE FROM PORTFOLIO")
        print(f"{'='*55}")
        print(f"  Name   : {name}")
        print(f"  Email  : {email}")
        print(f"  Subject: {subject}")
        print(f"  Message: {message}")
        print(f"{'='*55}\n")

        return jsonify({'status': 'success',
                        'message': "Thanks for reaching out! I'll get back to you soon. 🚀"})

    except Exception as e:
        return jsonify({'status': 'error',
                        'message': 'Something went wrong. Please try again.'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
