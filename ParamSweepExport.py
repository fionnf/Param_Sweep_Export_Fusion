# ParamSweepExport.py — clr (unitless) + dpt (mm) sweep, letter code, STL, CSV
import adsk.core, adsk.fusion, adsk.cam, traceback, os, time, csv

# ====== CONFIG ======
PARAM1_NAME   = 'clr'   # unitless ratio
PARAM2_NAME   = 'dpt'   # dimension in mm

# Give numbers or strings. For clr: no units. For dpt: numbers are treated as mm.
CLR_VALUES = [1.02, 1.02, 1.04, 1.06, 1.08, 1.10]
DPT_VALUES = [0.5, 0.6, 0.7, 0.8]

TEXT_PLACEHOLDER = 'XXX'              # sketch text to replace
COMPONENT_PREFIX = 'c.5.2'        # filename prefix

OUTPUT_DIR = '/Users/fionnferreira/Library/CloudStorage/GoogleDrive-fionnferreira@gmail.com/My Drive/PhD/Experiments/DEMO'
CSV_NAME   = 'variants.csv'
STL_BINARY = True
# ====================

_app = None
_ui  = None

def _get_design():
    app = adsk.core.Application.get()
    ui  = app.userInterface

    # Make sure a document is open
    doc = app.activeDocument
    if not doc:
        raise RuntimeError('Open a design document (.f3d) first and save it.')

    # Ensure we’re in the Design workspace (FusionSolidEnvironment)
    try:
        ws = ui.workspaces.itemById('FusionSolidEnvironment')
        if ws: ws.activate()
    except:
        pass

    # Pull the Design product from the active document’s products
    prods = doc.products
    design = adsk.fusion.Design.cast(prods.itemByProductType('DesignProductType'))
    if not design:
        raise RuntimeError('Active document has no Design product (are you in a Drawing/Electronics/empty doc?).')
    return design

def _ensure_dir(p):
    if not os.path.isdir(p):
        os.makedirs(p, exist_ok=True)

def _find_user_param(design, name):
    p = design.allParameters.itemByName(name)
    if not p:
        raise RuntimeError(f'User parameter "{name}" not found.')
    return p

def _fmt_clr_for_expr(v):
    # Accept float/int/str; clr is unitless
    if isinstance(v, (int, float)):
        return f'{v:.2f}'
    s = str(v).strip()
    return s  # assume already unitless
def _fmt_clr_for_name(v):
    # Safe for filenames: no spaces
    if isinstance(v, (int, float)):
        return f'{v:.2f}'
    return str(v).strip()

def _fmt_dpt_for_expr(v):
    # Accept float/int/str; ensure mm expression for Fusion
    if isinstance(v, (int, float)):
        return f'{v:g} mm'
    s = str(v).strip()
    return s if s.endswith('mm') or s.endswith(' mm') else f'{s} mm'
def _fmt_dpt_for_name(v):
    # For filenames, like 0.5mm (no space)
    if isinstance(v, (int, float)):
        return f'{v:g}mm'
    s = str(v).strip().replace(' ', '')
    return s if s.endswith('mm') else f'{s}mm'

def _update_param_expression(param, expr_str):
    param.expression = expr_str

def _do_recompute(design):
    design.computeAll()
    adsk.doEvents()
    time.sleep(0.05)

def _index_to_label(idx):
    label = ''
    n = idx + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        label = chr(65 + r) + label
    return label

def _collect_placeholder_texts(design, placeholder):
    hits = []
    def scan_comp(comp):
        for sk in comp.sketches:
            for i in range(sk.sketchTexts.count):
                st = sk.sketchTexts.item(i)
                if (st.text or '').strip() == placeholder:
                    hits.append(st)
    root = design.rootComponent
    scan_comp(root)
    for occ in root.allOccurrences:
        scan_comp(occ.component)
    return hits

def _set_all_placeholders(items, new_text):
    for st in items: st.text = new_text

def _export_root_stl(design, filepath):
    exp = design.exportManager
    target = design.rootComponent
    try:
        opts = exp.createSTLExportOptions(target, filepath)
    except:
        bodies = [b for b in target.bRepBodies if b.isSolid]
        if not bodies:
            raise RuntimeError('No solid bodies in root to export.')
        opts = exp.createSTLExportOptions(bodies, filepath)
    opts.isBinaryFormat = STL_BINARY
    exp.execute(opts)

def _safe_filename(s):
    bad = '<>:"/\\|?* '
    return ''.join(c if c not in bad else '_' for c in s)

def _log(msg):
    if _ui: _ui.messageBox(str(msg))

def run(context):
    global _app, _ui
    _app = adsk.core.Application.get()
    _ui  = _app.userInterface
    try:
        design = _get_design()
        _ensure_dir(OUTPUT_DIR)

        p_clr = _find_user_param(design, PARAM1_NAME)
        p_dpt = _find_user_param(design, PARAM2_NAME)

        # cache originals
        orig_clr = p_clr.expression
        orig_dpt = p_dpt.expression

        placeholders = _collect_placeholder_texts(design, TEXT_PLACEHOLDER)
        if not placeholders:
            raise RuntimeError(f'No SketchText equals "{TEXT_PLACEHOLDER}".')

        rows = []
        combos = [(c, d) for c in CLR_VALUES for d in DPT_VALUES]

        for i, (clr_v, dpt_v) in enumerate(combos):
            code = _index_to_label(i)

            # set params
            _update_param_expression(p_clr, _fmt_clr_for_expr(clr_v))
            _update_param_expression(p_dpt, _fmt_dpt_for_expr(dpt_v))
            _do_recompute(design)

            # replace placeholder with letter code
            _set_all_placeholders(placeholders, code)
            _do_recompute(design)

            clr_name = _safe_filename(_fmt_clr_for_name(clr_v))
            dpt_name = _safe_filename(_fmt_dpt_for_name(dpt_v))
            fname = f'{COMPONENT_PREFIX}_{_safe_filename(PARAM1_NAME)}{clr_name}_{_safe_filename(PARAM2_NAME)}{dpt_name}_{code}.stl'
            out_path = os.path.join(OUTPUT_DIR, fname)

            _export_root_stl(design, out_path)

            rows.append({
                'code': code,
                PARAM1_NAME: _fmt_clr_for_expr(clr_v).replace(' mm','').strip(),  # store just number for clr
                PARAM2_NAME: _fmt_dpt_for_expr(dpt_v),                             # store with " mm"
                'filename': fname
            })

        # CSV
        csv_path = os.path.join(OUTPUT_DIR, CSV_NAME)
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=['code', PARAM1_NAME, PARAM2_NAME, 'filename'])
            w.writeheader()
            for r in rows: w.writerow(r)

        # restore
        _update_param_expression(p_clr, orig_clr)
        _update_param_expression(p_dpt, orig_dpt)
        _set_all_placeholders(placeholders, TEXT_PLACEHOLDER)
        _do_recompute(design)

        _log(f'Exported {len(rows)} variants → {OUTPUT_DIR}\nCSV: {CSV_NAME}')

    except Exception as e:
        if _ui:
            _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context): pass