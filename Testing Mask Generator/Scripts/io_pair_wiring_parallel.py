import csv
import re
import sys
import math
import pathlib

import gdstk

HERE = pathlib.Path(__file__).parent
INPUT_DIR = HERE.parent / "io_pair_inputs"
OUTPUT_DIR = HERE.parent / "io_pair_outputs"
GDS_DIR = OUTPUT_DIR / "gds"                # chip mask GDS files
SCHEMATIC_DIR = OUTPUT_DIR / "schematics"  # per-chip circuit diagrams (SVG)
CALIB_DIR = OUTPUT_DIR / "calibration"     # calibration coupon GDS + CSV

PAD_SIZE = 80.0
WIRE_WIDTH = 3.0           # aluminium trace width
WIRE_SPACE = 15.0           # gap between adjacent comb teeth
COIL_GAP = 38.0            # gap from the pad edge to the first comb tooth
VIA_SIZE = 8.0             # via from a 2nd-layer wire up to its top-layer pad

PLANE_RING_MARGIN = 70.0   # keep the shared plane this far inside the ring
FINGER_OVERLAP = 60.0      # how far each output finger / return reaches into the plane
PLANE_SPLIT_GAP = 60.0     # isolation gap between the two half-planes when splitting
SPLIT_HALVES = False       # set in main(): halve each coupon into two 4-resistor sets

# Aluminium-on-Ti resistor: R = SHEET_RES * L / W.
METAL_THICKNESS_UM = 0.1     # 1000 angstrom
AL_RESISTIVITY = 3.243e-8    # ohm*m
SHEET_RES = AL_RESISTIVITY / (METAL_THICKNESS_UM * 1e-6)   # ohm/square (~0.324)
COIL_BASE_R = 4000.0          # ohms for the smallest coil in a group
MAX_BINARY_INPUTS = 8        # split groups bigger than this into separate coupons

CALIB_COUNT = 4              # number of calibration resistors (first N ladder steps)

MOSAIC_GAP = 300.0           # spacing between dies in the combined tiled GDS

WAFER_DIAM_UM = 150000.0     # 6-inch (150 mm) fabrication wafer
WAFER_EDGE_EXCLUSION_UM = 3000.0  # unusable edge ring (edge bead, handling, non-uniformity)
WAFER_STREET_UM = MOSAIC_GAP # gap between stepped reticle fields on the wafer
WAFER_LINE_UM = 500.0        # width of the wafer-edge outline ring

ADD_LABELS = True
LABEL_SIZE = 16.8            # pad-text height (um); IO tag + pad name inside each pad
ADD_DIE_OUTLINE = True
DIE_MARGIN_BUFFER = 100.0    # clearance beyond the farthest resistor (um)
DIE_W = DIE_H = 0.0          # set in main()

METAL_BASE, METAL_DT = 1, 0
VIA_BASE, VIA_DT = 20, 0
LABEL_LAYER, LABEL_DT = 101, 0
IO_LABEL_DT = 1
BOUNDARY_LAYER, BOUNDARY_DT = 100, 0
WAFER_LAYER, WAFER_DT = 104, 0   # wafer-edge overlay -- reference only, not fabricated


def safe_path(path):
    """`path`, or a `_new` sibling if it's locked (open in Excel / a GDS viewer)."""
    try:
        if path.exists():
            with open(path, "a"):
                pass
        return path
    except PermissionError:
        alt = path.with_name(path.stem + "_new" + path.suffix)
        print(f"  ! {path.name} is locked (open elsewhere); writing {alt.name} instead.")
        return alt


def metal_layer(k):
    return METAL_BASE + k


def via_layer(k):
    return VIA_BASE + k


COL_PAD, COL_SIGNAL, COL_X, COL_Y = ("pad", "signal", "x (um)", "y (um)")
COL_IO = "i/o"     # grouping column: INPUTn / OUTPUTn, blank = unused
_IO_RE = re.compile(r"(INPUT|OUTPUT)\s*(\d+)$")
GLOBAL_OUT = "*"   # a bare 'OUTPUT' pad: common to every group, on every chip
SPLIT_OUT = "+"    # an 'OUTPUTSPLIT' pad: on every chip, halved across the two layers
SINGLE = "S"       # INPUTSINGLE / OUTPUTSINGLE: one shorted continuity coupon, own layer


def classify_io(io_cell):
    """'INPUT3' -> ('input', 3); 'OUTPUT7' -> ('output', 7); a bare 'OUTPUT' ->
    ('output', GLOBAL_OUT) -- a common output wired to every group; 'OUTPUTSPLIT' ->
    ('output', SPLIT_OUT) -- on every chip, halved across the two layers;
    'INPUTSINGLE'/'OUTPUTSINGLE' -> (role, SINGLE) -- one shorted continuity coupon on
    its own layer; else ('', None)."""
    s = io_cell.strip().upper()
    if s == "OUTPUT":
        return ("output", GLOBAL_OUT)
    if s == "OUTPUTSPLIT":
        return ("output", SPLIT_OUT)
    if s == "INPUTSINGLE":
        return ("input", SINGLE)
    if s == "OUTPUTSINGLE":
        return ("output", SINGLE)
    m = _IO_RE.match(s)
    if not m:
        return ("", None)
    return ("input" if m.group(1) == "INPUT" else "output", int(m.group(2)))


# ----------------------------------------------------------------------
# CSV / XLSX parsing
# ----------------------------------------------------------------------
def load_rows(path):
    """Read the pinout as a list of string-cell rows, from .csv or .xlsx/.xlsm.
    For a workbook, use the first sheet whose cells hold the X/Y header row
    (so a cover/notes sheet is skipped); every cell is stringified, None -> ''."""
    path = pathlib.Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl                              # optional; only for .xlsx
        except ImportError:
            raise SystemExit(f"Reading {path.name} needs openpyxl: pip install openpyxl "
                             f"(or export the sheet to .csv).")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        best = None
        for ws in wb.worksheets:
            rows = [["" if c is None else str(c) for c in r]
                    for r in ws.iter_rows(values_only=True)]
            if best is None:
                best = rows                              # fall back to the first sheet
            if any(COL_X in [c.strip().lower() for c in row]
                   and COL_Y in [c.strip().lower() for c in row] for row in rows):
                return rows
        return best or []
    with open(path, newline="") as f:
        return list(csv.reader(f))


def read_pads(csv_path):
    rows = load_rows(csv_path)
    header_idx = None
    for i, row in enumerate(rows):
        cells = [c.strip().lower() for c in row]
        if COL_X in cells and COL_Y in cells:
            header_idx, header = i, cells
            break
    if header_idx is None:
        raise ValueError(f"No header row with '{COL_X}' and '{COL_Y}'.")
    ix, iy = header.index(COL_X), header.index(COL_Y)
    ipad = header.index(COL_PAD) if COL_PAD in header else None
    isig = header.index(COL_SIGNAL) if COL_SIGNAL in header else None
    iio = header.index(COL_IO) if COL_IO in header else None
    if iio is None:
        raise ValueError(f"No grouping column '{COL_IO}' (INPUTn / OUTPUTn).")
    pads = []
    for row in rows[header_idx + 1:]:
        if len(row) <= max(ix, iy):
            continue
        try:
            x, y = float(row[ix]), float(row[iy])
        except ValueError:
            continue
        name = row[ipad].strip() if ipad is not None and ipad < len(row) else ""
        signal = row[isig].strip() if isig is not None and isig < len(row) else ""
        io_cell = row[iio].strip() if iio < len(row) else ""
        role, group = classify_io(io_cell)
        pads.append({"name": name, "signal": signal, "x": x, "y": y,
                     "io": role, "group": group})
    return pads

def dist(a, b):
    return math.dist((a["x"], a["y"]), (b["x"], b["y"]))

def place_pads(pads, raw_bounds, margins):
    """Shift the pad ring so each side sits `margins[edge]` from the die edge.
    Die spans x in [0, W], y in [0, -H]."""
    minx, _, _, maxy = raw_bounds
    ox = margins["left"] - minx
    oy = -margins["top"] - maxy
    for p in pads:
        p["x"] += ox
        p["y"] += oy


def assign_groups(inputs, outputs):
    """Group pads by IO number. A group needs inputs and at least one output -- its own
    OUTPUTn, a bare 'OUTPUT' (common to every group), or an 'OUTPUTSPLIT' pad (added
    to every chip's layers at build time). Returns [{"num", "inputs", "outputs"}]
    sorted by number. OUTPUTSPLIT pads are NOT attached here -- build_chip halves them
    across each chip's two layers."""
    common = [p for p in outputs if p["group"] == GLOBAL_OUT]
    has_split = any(p["group"] == SPLIT_OUT for p in outputs)
    in_by, out_by = {}, {}
    for p in inputs:
        in_by.setdefault(p["group"], []).append(p)
    for p in outputs:
        if p["group"] not in (GLOBAL_OUT, SPLIT_OUT):
            out_by.setdefault(p["group"], []).append(p)
    groups = []
    for n in sorted(in_by):
        outs = out_by.get(n, []) + common          # own outputs first, then the common
        if outs or has_split:                      # split pads guarantee output access
            groups.append({"num": n, "inputs": in_by[n], "outputs": outs})
    return groups


def split_coupons(groups, max_in):
    """Split groups into "coupons" (one metal layer each) of <= max_in inputs. Each
    carries the group's outputs and restarts the ladder, so it decodes on its own.
    Coupons keep the group `num` so sub-coupons stay off the same chip."""
    coupons = []
    for g in groups:
        ins = ordered_inputs(g)
        chunks = [ins[i:i + max_in] for i in range(0, len(ins), max_in)] or [[]]
        for j, chunk in enumerate(chunks):
            coupons.append({"num": g["num"], "sub": j, "inputs": chunk,
                            "outputs": g["outputs"]})
    return coupons


def pack_chips(coupons, per_chip):
    """Assign coupons to chips (`per_chip` layers each), never two coupons of the
    same group on one chip (they share output pads, which would tie them)."""
    if per_chip <= 1:
        return [[c] for c in coupons]
    rem, chips = list(coupons), []
    while rem:
        a = rem.pop(0)
        j = next((i for i, b in enumerate(rem) if b["num"] != a["num"]), None)
        chips.append([a, rem.pop(j)] if j is not None else [a])
    return chips


# ----------------------------------------------------------------------
# Geometry helpers
# ----------------------------------------------------------------------
def polyline_len(pts):
    return sum(math.dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def edge_of(p, bounds):
    """Nearest ring edge a pad sits on."""
    minx, maxx, miny, maxy = bounds
    d = {"left": abs(p["x"] - minx), "right": abs(p["x"] - maxx),
         "bottom": abs(p["y"] - miny), "top": abs(p["y"] - maxy)}
    return min(d, key=d.get)


def inward_along(edge):
    """(inward unit vector, along-edge unit vector) for an edge."""
    return {"bottom": ((0, 1), (1, 0)), "top": ((0, -1), (1, 0)),
            "left": ((1, 0), (0, 1)), "right": ((-1, 0), (0, 1))}[edge]


ROUTE_PITCH = WIRE_WIDTH + WIRE_SPACE   # tooth pitch (traces sit WIRE_SPACE apart)


def decode_margin(resistances):
    """Smallest gap between any two subset sums of the CONDUCTANCES, as a fraction of
    the all-landed sum -- the fraction of full scale a meter must resolve to tell every
    landed/open pattern apart. Uses the ACTUAL extracted resistances, not the targets,
    so it reflects what the routing really built; 0 means two patterns read the same
    and the coupon cannot be decoded."""
    g = [1.0 / r for r in resistances if r > 0]
    if len(g) < 2 or len(g) > 20:
        return 1.0
    sums = {0.0}
    for w in g:
        sums |= {s + w for s in sums}
    s = sorted(sums)
    return min(b - a for a, b in zip(s, s[1:])) / sum(g)


def input_target_r(k):
    """Target resistance (ohms) for the k-th input: binary COIL_BASE_R * 2**(k-1)."""
    return COIL_BASE_R * (2 ** (k - 1))


def input_target_len(k):
    """Trace length (um) giving the k-th input its target resistance."""
    return input_target_r(k) * WIRE_WIDTH / SHEET_RES


def ordered_inputs(group):
    """Group inputs ranked by distance to the first output (rank k -> target R). A
    group whose only outputs are OUTPUTSPLIT pads has no dedicated output here, so it
    ranks by the input centroid instead -- a stable reference that does not depend on
    which split pads a chip happens to get, keeping ranks identical across the sizing
    and drawing passes."""
    outs = group["outputs"]
    if outs:
        ref = outs[0]
    else:
        n = len(group["inputs"])
        ref = {"x": sum(p["x"] for p in group["inputs"]) / n,
               "y": sum(p["y"] for p in group["inputs"]) / n}
    return sorted(group["inputs"], key=lambda p: dist(p, ref))


def near_edge_coords(pads, bounds):
    """Per edge, sorted along-coords of every pad on that edge line (within a pad of
    it), so a return can find a real gap -- includes corner pads."""
    minx, maxx, miny, maxy = bounds
    tol = PAD_SIZE
    out = {"left": [], "right": [], "top": [], "bottom": []}
    for p in pads:
        if abs(p["x"] - minx) < tol:
            out["left"].append(p["y"])
        if abs(p["x"] - maxx) < tol:
            out["right"].append(p["y"])
        if abs(p["y"] - miny) < tol:
            out["bottom"].append(p["x"])
        if abs(p["y"] - maxy) < tol:
            out["top"].append(p["x"])
    for e in out:
        out[e].sort()
    return out


CORNER_INSET = PAD_SIZE + WIRE_SPACE             # keep combs clear of the ring corners
LEAD_RADIUS = PAD_SIZE / 2.0 + COIL_GAP / 2.0    # pad->comb lead lane (under the teeth)


def to_xy(e, t, r, bounds):
    """Point at along-coord `t`, outward radius `r` (from the edge line) on edge `e`.
    r > 0 is outside the ring (the comb); r < 0 is inside, toward the plane."""
    minx, maxx, miny, maxy = bounds
    if e == "left":
        return (minx - r, t)
    if e == "right":
        return (maxx + r, t)
    if e == "bottom":
        return (t, miny - r)
    return (t, maxy + r)                                  # top


def edge_along_range(e, bounds):
    """Usable along-edge interval for edge `e`, inset from the corners."""
    minx, maxx, miny, maxy = bounds
    lo, hi = (miny, maxy) if e in ("left", "right") else (minx, maxx)
    return lo + CORNER_INSET, hi - CORNER_INSET


def edge_gaps(e, edge_coords):
    """Centres of the pad-to-pad gaps on edge `e` that a return can actually thread.

    The return runs BETWEEN two probe pads, so both must clear it by WIRE_SPACE -- the
    gap has to hold the wire plus that clearance on each side. Narrower gaps are dropped
    rather than squeezed through: routing one would leave metal closer to a probe pad
    than the process allows, which is a spacing violation even though nothing shorts."""
    c = edge_coords[e]
    need = PAD_SIZE + WIRE_WIDTH + 2 * WIRE_SPACE      # centre-to-centre, not edge-to-edge
    return [(c[i] + c[i + 1]) / 2.0 for i in range(len(c) - 1)
            if c[i + 1] - c[i] >= need - 1e-6]


def edge_inputs(group, bounds, target_of):
    """Same-group inputs bucketed by edge, each sorted by along-coord and tagged
    (along, target_len, pad). `target_of` maps id(pad) -> target trace length, so a
    split coupon's two halves can each run their own ladder."""
    by_edge = {}
    for p in group["inputs"]:
        e = edge_of(p, bounds)
        _, v = inward_along(e)
        a = p["x"] * v[0] + p["y"] * v[1]
        by_edge.setdefault(e, []).append((a, target_of[id(p)], p))
    for e in by_edge:
        by_edge[e].sort(key=lambda it: it[0])
    return by_edge


def assign_return_gaps(a, gaps):
    """One DISTINCT return gap per input, in pad order, each lying strictly between its
    own pad's neighbouring INPUTS. Without this two combs could claim the slots either
    side of a third input and leave it none, forcing its return to double back over its
    own comb. Assignments strictly increase, so the reserved pad-to-gap spans stay
    ordered and disjoint. Returns [gap or None] aligned with `a`."""
    n = len(a)

    def window(i):
        return (a[i - 1] if i else -math.inf,
                a[i + 1] if i < n - 1 else math.inf)

    def strandable(i, last):
        """Greedy-smallest for inputs i.. -- if that fails, `last` stranded someone."""
        for j in range(i, n):
            lo, hi = window(j)
            nxt = next((g for g in gaps if g > last and lo < g < hi), None)
            if nxt is None:
                return True
            last = nxt
        return False

    out, last = [None] * n, -math.inf
    for i in range(n):
        lo, hi = window(i)
        cand = [g for g in gaps if g > last and lo < g < hi]
        if not cand:
            continue                                   # caller falls back for this one
        # Prefer the nearest slot BELOW the pad. Bands are handed out low-to-high, so a
        # low gap leaves the whole band above the pad free for the fold to grow into --
        # and taking the low slot in pad order keeps the choices strictly increasing.
        rank = lambda g: (0 if g < a[i] else 1, abs(g - a[i]))
        pick = next((g for g in sorted(cand, key=rank)
                     if not strandable(i + 1, g)), cand[0])
        out[i], last = pick, pick
    return out


def solve_bands(items, edge_lo, edge_hi, must=None):
    """Partition [edge_lo, edge_hi] into one band per input (pad order), each strictly
    containing that input's reserved span `must[i]` -- its pad, and the return gap it
    has to reach -- with widths ~proportional to trace length so the comb depth
    (= len*pitch/width) is levelled across the edge. Disjoint bands that each hold
    their own span mean combs never touch. Returns [(lo, hi)] aligned with `items`."""
    n = len(items)
    a = [it[0] for it in items]
    req_lo = [must[i][0] if must else a[i] for i in range(n)]
    req_hi = [must[i][1] if must else a[i] for i in range(n)]
    eps = WIRE_SPACE + WIRE_WIDTH                          # keep neighbouring combs apart
    edge_lo = min(edge_lo, min(req_lo) - eps)             # a pad may sit in the corner
    edge_hi = max(edge_hi, max(req_hi) + eps)            #   inset -> still span it
    if n == 1:
        return [(edge_lo, edge_hi)]
    need = [it[1] * ROUTE_PITCH for it in items]          # wire area = width * depth

    def cuts_for(depth):
        """Boundaries giving every comb width >= need/depth and still holding its span;
        None if the pad spacing makes it impossible."""
        cuts, prev = [], edge_lo
        for i in range(n - 1):
            c = min(max(prev + need[i] / depth, req_hi[i] + eps), req_lo[i + 1] - eps)
            if (c - prev < need[i] / depth - 1e-6
                    or not (prev <= req_lo[i] and req_hi[i] <= c)):
                return None
            cuts.append(c)
            prev = c
        if (edge_hi - prev < need[-1] / depth - 1e-6
                or not (prev <= req_lo[-1] and req_hi[-1] <= edge_hi)):
            return None
        return cuts

    hi = max(max(need), 1.0)                # surely feasible (~1 um wide combs)
    while cuts_for(hi) is None and hi < 1e15:
        hi *= 2.0
    lo = 1e-9
    for _ in range(50):                     # binary-search the smallest feasible depth
        mid = (lo + hi) / 2.0
        if cuts_for(mid) is None:
            lo = mid
        else:
            hi = mid
    cuts = cuts_for(hi) or [(a[i] + a[i + 1]) / 2.0 for i in range(n - 1)]
    edges = [edge_lo] + cuts + [edge_hi]
    return [(edges[i], edges[i + 1]) for i in range(n)]


CANOPY_CLEAR = 3.0 * ROUTE_PITCH        # radial gap between the base combs and a canopy


CHANNEL_MARGIN = 2.0 * ROUTE_PITCH     # clearance between a channel and the first pad


def channel_fold(target_lat, width):
    """How to lay `target_lat` of lateral wire on a storey `width` wide, entering at the
    low end and FINISHING at the high end so the return channel is reachable.

    One crossing costs `width`; every extra there-and-back costs 2*width. Those alone
    quantise the length in steps of 2*width, which a small resistor cannot land between,
    so the remainder is taken up by a short there-and-back of depth `d`. Returns
    (round_trips, partial_reach, radial_depth)."""
    extra = max(0.0, target_lat - width)
    k = int(extra // (2.0 * width))
    d = (extra - 2.0 * k * width) / 2.0
    runs = 1 + 2 * k + (2 if d > 1e-6 else 0)
    return k, d, runs * ROUTE_PITCH


def plan_edge_channels(items, edge_lo, edge_hi, gaps):
    """Route every comb on this edge as a FULL-WIDTH storey.

    The storey scheme reserves a slot at each input's own pad on every storey below it,
    so a band ends up pinched between neighbouring pads. With pads a few hundred um
    apart and an 18 um pitch that forces very deep folds and leaves most of the area
    outboard of the ring empty -- measured fill was under 20 percent.

    Here the slots move OFF the pads into two channels at the ENDS of the edge: risers
    at the low end, returns at the high end. Each input runs out to its own lead lane
    under the storeys, along to its riser column, up to its storey, folds across the
    WHOLE edge, then drops down a return gap past the last pad. Every storey gets
    essentially the full edge width instead of one pad pitch.

    Nothing can cross, by construction:
      - lead lanes are ordered by pad position, so a lead never passes over a lane whose
        span reaches back past its own pad;
      - riser columns are ordered the same way, so a riser only meets lanes whose span
        starts to its right;
      - storey order is REVERSED against pad position, so a comb only passes over risers
        that stop below it;
      - returns drop beyond the last pad, clear of every lane and every comb.

    Returns a per-input plan, or None if the channels do not fit -- the caller then
    falls back to the storey scheme."""
    n = len(items)
    a = [it[0] for it in items]
    pitch = ROUTE_PITCH
    r0 = PAD_SIZE / 2.0 + COIL_GAP
    need_ch = n * pitch + CHANNEL_MARGIN
    if a[0] - edge_lo < need_ch or edge_hi - a[-1] < need_ch:
        return None
    # returns drop through real pad gaps past the last pad, one each, pitch apart
    tail = [g for g in gaps if a[-1] + CHANNEL_MARGIN < g < edge_hi]
    rets, last = [], -math.inf
    for g in tail:
        if g - last >= pitch:
            rets.append(g)
            last = g
        if len(rets) == n:
            break
    if len(rets) < n:
        return None
    ret_lo = rets[0] - pitch                       # combs stop short of the return bank
    cols = [edge_lo + CHANNEL_MARGIN / 2.0 + i * pitch for i in range(n)]
    lanes = [r0 + i * pitch for i in range(n)]
    if cols[-1] + pitch >= a[0] or ret_lo <= cols[-1] + pitch:
        return None
    # storey order reversed against pad order: rightmost pad sits innermost
    base = r0 + n * pitch + CANOPY_CLEAR
    fold = [channel_fold(items[j][1], ret_lo - cols[j]) for j in range(n)]
    rs, r = {}, base
    for j in reversed(range(n)):
        rs[j] = r
        r += fold[j][2] + CANOPY_CLEAR
    # Return columns go in STOREY order, innermost nearest the combs. A comb runs from
    # its far end across to its own column, so it passes over the columns nearer than
    # its own -- those must belong to lower storeys, whose descents stop beneath it.
    return [{"kind": "channel", "lane": lanes[j], "col": cols[j],
             "ret": rets[n - 1 - j], "r": rs[j], "hi": ret_lo} for j in range(n)]


def build_channel_coil(pad, e, ch, target_len, bounds, plane):
    """Fold one resistor across a full-width storey, reached by the end channels.
    Length is trimmed with a partial last run so the target is met exactly."""
    u, v = inward_along(e)
    cu = (-u[0], -u[1])
    a_pad = pad["x"] * v[0] + pad["y"] * v[1]
    r0 = PAD_SIZE / 2.0 + COIL_GAP
    lane, col, ret, rs, hi = ch["lane"], ch["col"], ch["ret"], ch["r"], ch["hi"]
    width = hi - col
    P = lambda tt, rr: to_xy(e, tt, rr, bounds)

    def build(lat):
        k, d, _ = channel_fold(lat, width)
        pts = [(pad["x"], pad["y"]), P(a_pad, lane),      # out to its own lead lane
               P(col, lane), P(col, rs)]                  # along to its riser, then up
        r = rs
        pts.append(P(hi, r))                              # first crossing
        for _ in range(k):                                # there and back, full width
            r += ROUTE_PITCH
            pts += [P(hi, r), P(col, r)]
            r += ROUTE_PITCH
            pts += [P(col, r), P(hi, r)]
        if d > 1e-6:                                      # short there and back, trims
            r += ROUTE_PITCH
            pts += [P(hi, r), P(hi - d, r)]
            r += ROUTE_PITCH
            pts += [P(hi - d, r), P(hi, r)]
        pts += [P(ret, r), P(ret, r0)]                    # over to the return gap, down
        rr = return_through_gap(0.0, 0.0, P(ret, r0), cu, plane)
        return pts, rr

    # The lead, riser and return add a fixed overhead that is a large fraction of a
    # SMALL resistor, so solve for the lateral length rather than assuming it.
    lat, coil, rtn = target_len, None, None
    for _ in range(24):                                   # length is linear in `lat`
        coil, rtn = build(lat)
        err = target_len - (polyline_len(coil) + polyline_len(rtn))
        if abs(err) < 0.5:
            break
        lat = max(0.0, lat + err)
    r_ohm, clen = _coil_result(pad, coil, rtn, plane)
    return coil, rtn, r_ohm, clen


def plan_edge(items, edge_lo, edge_hi, gaps):
    """Plan every input on one edge: a distinct return gap each, then a stack of storeys.

    A comb's depth is its wire area over its band width, and on one storey every band is
    pinched between the ADJACENT INPUT PADS -- about one pad pitch. That is what makes a
    64k or 128k resistor tower over the die: it has metres of wire and a 150 um slot.

    So the big ones are lifted onto their own storeys above the rest. The trick is which
    inputs have to be represented on a given storey: only the one folding there and the
    ones stacked ABOVE it, which still need a slot to rise through. Everything smaller is
    already finished below and is simply absent -- so a storey's band is bounded by just
    a couple of pads and stretches to the ends of the edge, tens of pad pitches wide. The
    same wire then folds shallow instead of deep. That is the dead space being reclaimed:
    the empty area outboard of the short combs, up to the die edge the tallest one sets.

    Nothing crosses: an input rises radially at its OWN pad and drops at its OWN gap,
    both inside the slot reserved for it on every storey it passes through, and folds
    away from that gap so its own descent never meets its own teeth.

    Costs every split -- the `m` biggest each get a storey, the rest share the base --
    and keeps the shallowest. Returns [(band, gap, rc)] aligned with `items`; rc is None
    for a base-storey input, else the radius its fold starts at."""
    ch = plan_edge_channels(items, edge_lo, edge_hi, gaps)
    if ch is not None:                # full-width storeys whenever the channels fit
        return ch
    n = len(items)
    a = [it[0] for it in items]
    need = [it[1] * ROUTE_PITCH for it in items]
    gap_of = assign_return_gaps(a, gaps)
    eps = WIRE_SPACE + WIRE_WIDTH
    r0 = PAD_SIZE / 2.0 + COIL_GAP
    req = [(min(a[i], gap_of[i]), max(a[i], gap_of[i])) if gap_of[i] is not None
           else (a[i], a[i]) for i in range(n)]
    grow = [1.0 if (gap_of[i] is None or gap_of[i] <= a[i]) else -1.0 for i in range(n)]
    small_first = sorted(range(n), key=lambda i: need[i])

    best = None
    for m in range(n + 1):
        base_set = set(small_first[:n - m])        # the rest share the base storey
        stacked = small_first[n - m:]              # the m biggest, one storey each
        # Base storey: every input appears. A stacked input asks for no depth, just the
        # slot its riser and its drop back down need.
        base = [(a[i], items[i][1] if i in base_set else 0.0, items[i][2])
                for i in range(n)]
        bb = solve_bands(base, edge_lo, edge_hi, req)
        plan, ok, h = {}, True, 0.0
        for i in base_set:
            lo, hi = bb[i]
            # A base comb ends ON its gap and grows away from it, so only the band on
            # the far side of the gap is usable -- cost it the way it is actually built.
            gp = gap_of[i] if gap_of[i] is not None else a[i]
            w = ((hi - gp) if gp <= a[i] else (gp - lo)) - eps
            if w <= ROUTE_PITCH:
                ok = False
                break
            h = max(h, need[i] / w)
            plan[i] = (bb[i], None)
        if not ok:
            continue
        r = r0 + h + (CANOPY_CLEAR if base_set else 0.0)
        for idx, i in enumerate(stacked):
            # Only this input and the ones still to be stacked above it are present, and
            # those only as a slot to rise through -- so the builder takes everything
            # between its neighbours' slots, out to the ends of the edge.
            part = sorted(stacked[idx:], key=lambda j: a[j])
            at = part.index(i)
            lo = edge_lo if at == 0 else (req[part[at - 1]][1] + req[i][0]) / 2.0
            hi = edge_hi if at == len(part) - 1 else (req[i][1] + req[part[at + 1]][0]) / 2.0
            band = (lo, hi)
            usable = (hi - a[i] if grow[i] > 0 else a[i] - lo) - eps
            if usable <= ROUTE_PITCH:
                ok = False
                break
            plan[i] = (band, r)
            r += need[i] / usable + ROUTE_PITCH + CANOPY_CLEAR
        if ok and (best is None or r < best[0]):
            best = (r, plan)

    if best is None:                                   # nothing fits; keep it simple
        bb = solve_bands(items, edge_lo, edge_hi, req)
        return [(bb[i], gap_of[i], None) for i in range(n)]
    plan = best[1]
    return [(plan[i][0], gap_of[i], plan[i][1]) for i in range(n)]


def return_through_gap(px, py, far_end, cu, plane):
    """Polyline from the comb's far end straight INWARD onto the plane. The far end
    already sits in a pad gap, so this drops perpendicular through it."""
    px0, py0, px1, py1 = plane
    u = (-cu[0], -cu[1])                                   # inward
    if u == (0, 1):                                        # bottom -> plane bottom
        land = (far_end[0], py0)
    elif u == (0, -1):                                     # top
        land = (far_end[0], py1)
    elif u == (1, 0):                                      # left
        land = (px0, far_end[1])
    else:                                                  # right
        land = (px1, far_end[1])
    into = (land[0] + u[0] * FINGER_OVERLAP, land[1] + u[1] * FINGER_OVERLAP)
    return [far_end, land, into]


# A right-angle bend conducts ~CORNER_SQUARES squares, not the ~1 a centre-line count
# gives it (current crowds the inside of the turn). A wide-shallow comb has ~2 bends
# per tooth, so this is a real correction.
CORNER_SQUARES = 0.56


def _overlap(a0, a1, b0, b1):
    lo, hi = max(min(a0, a1), min(b0, b1)), min(max(a0, a1), max(b0, b1))
    return max(0.0, hi - lo)


def _seg_outside(p, q, nodes):
    """Length of axis-aligned segment p->q lying OUTSIDE every node rectangle."""
    length = math.dist(p, q)
    if length < 1e-9:
        return 0.0
    horiz, vert = abs(p[1] - q[1]) < 1e-6, abs(p[0] - q[0]) < 1e-6
    if not (horiz or vert):
        return length                                  # no diagonal traces in practice
    covered = 0.0
    for rx0, ry0, rx1, ry1 in nodes:
        if horiz and ry0 - 1e-6 <= p[1] <= ry1 + 1e-6:
            covered = max(covered, _overlap(p[0], q[0], rx0, rx1))
        elif vert and rx0 - 1e-6 <= p[0] <= rx1 + 1e-6:
            covered = max(covered, _overlap(p[1], q[1], ry0, ry1))
    return max(0.0, length - covered)


def _in_node(p, nodes):
    return any(rx0 - 1e-6 <= p[0] <= rx1 + 1e-6 and ry0 - 1e-6 <= p[1] <= ry1 + 1e-6
              for rx0, ry0, rx1, ry1 in nodes)


def _is_corner(a, b, c):
    d1, d2 = (b[0] - a[0], b[1] - a[1]), (c[0] - b[0], c[1] - b[1])
    return (abs(d1[0] * d2[0] + d1[1] * d2[1]) < 1e-6
            and (d1[0] or d1[1]) and (d2[0] or d2[1]))


def resistive_length(pts, nodes):
    """Effective resistive length (um) of an axis-aligned polyline: centre-line length
    OUTSIDE the equipotential `nodes` (the wide pad/plane carry ~0 ohm), minus a
    per-bend correction (CORNER_SQUARES). R = SHEET_RES * this / WIRE_WIDTH, i.e. the
    actual conducting metal including lead and return, not the nominal coil length."""
    clean = [pts[0]]
    for p in pts[1:]:
        if abs(p[0] - clean[-1][0]) > 1e-9 or abs(p[1] - clean[-1][1]) > 1e-9:
            clean.append(p)
    straight = sum(_seg_outside(clean[i], clean[i + 1], nodes)
                   for i in range(len(clean) - 1))
    bends = sum(1 for i in range(1, len(clean) - 1)
                if _is_corner(clean[i - 1], clean[i], clean[i + 1])
                and not _in_node(clean[i], nodes))
    return max(0.0, straight - (1.0 - CORNER_SQUARES) * WIRE_WIDTH * bends)


def _coil_result(pad, coil, ret, plane):
    """Actual extracted resistance and total length for a finished coil."""
    hp = PAD_SIZE / 2.0
    nodes = [(pad["x"] - hp, pad["y"] - hp, pad["x"] + hp, pad["y"] + hp), plane]
    r_ohm = SHEET_RES * resistive_length(coil + ret, nodes) / WIRE_WIDTH
    return r_ohm, polyline_len(coil) + polyline_len(ret)


def build_band_coil(pad, e, band, target_len, gap, bounds, plane):
    """Fold one input's resistor into a radial-teeth comb filling its `band` (wide and
    shallow); tooth depth is solved so the trace length = `target_len`. A short lead
    joins the pad to the comb, and the far tooth lands on `gap` -- the slot reserved for
    this input alone -- so the return drops straight inward to the plane. Everything
    stays in the band, so combs never touch. Returns (coil, return, R_ohm, length)."""
    lo, hi = band
    u, v = inward_along(e)
    cu = (-u[0], -u[1])                                    # outward, away from the ring
    a_pad = pad["x"] * v[0] + pad["y"] * v[1]
    r0 = PAD_SIZE / 2.0 + COIL_GAP
    blo = lo + (WIRE_SPACE + WIRE_WIDTH) / 2.0
    bhi = hi - (WIRE_SPACE + WIRE_WIDTH) / 2.0
    if gap is None:                        # no slot of its own: sit right beside the pad
        sd = 1.0 if (bhi - a_pad) >= (a_pad - blo) else -1.0
        gap = min(max(a_pad + sd * ROUTE_PITCH, blo), bhi)
    gap = min(max(gap, blo), bhi)
    s = 1.0 if gap >= a_pad else -1.0                      # comb ends ON the gap
    # Teeth: as many as the band holds behind the gap, but capped by length so a short
    # resistor stays compact near its pad instead of being stretched thin.
    span = (gap - blo) if s > 0 else (bhi - gap)
    fit = max(2, int(max(span, 0.0) / ROUTE_PITCH) // 2 * 2)
    cap = max(2, int(target_len / (4.0 * ROUTE_PITCH)) // 2 * 2)   # depth >= ~4 pitches
    m = max(2, min(fit, cap))
    t0 = min(max(gap - s * (m - 1) * ROUTE_PITCH, blo), bhi)

    def build(depth):
        P = lambda t, r: to_xy(e, t, r, bounds)
        pts = [(pad["x"], pad["y"]), P(a_pad, LEAD_RADIUS),
               P(t0, LEAD_RADIUS), P(t0, r0)]              # pad -> lead lane -> tooth 0
        out = True
        for j in range(m):
            t = t0 + s * j * ROUTE_PITCH
            if out:                                        # tooth in -> out
                pts += [P(t, r0), P(t, r0 + depth)]
                if j < m - 1:
                    pts.append(P(t + s * ROUTE_PITCH, r0 + depth))
            else:                                          # tooth out -> in
                pts += [P(t, r0 + depth), P(t, r0)]
                if j < m - 1:
                    pts.append(P(t + s * ROUTE_PITCH, r0))
            out = not out
        ret = return_through_gap(0.0, 0.0, P(gap, r0), cu, plane)
        return pts, ret

    depth, coil, ret, clen = target_len / m, None, None, 0.0
    for _ in range(3):                                     # length is linear in depth
        coil, ret = build(depth)
        clen = polyline_len(coil) + polyline_len(ret)
        depth = max(ROUTE_PITCH, depth + (target_len - clen) / m)
    r_ohm, clen = _coil_result(pad, coil, ret, plane)
    return coil, ret, r_ohm, clen


def build_canopy_coil(pad, e, band, target_len, gap, rc, bounds, plane):
    """Fold a big input's resistor in the CANOPY -- a storey starting at radius `rc`,
    clear of every base comb, whose band is bounded only by the other canopy inputs and
    so is far wider than the pad-limited band below. Wider band, shallower fold: this is
    what stops a 64k or 128k resistor from driving the die size.

    The path never crosses anything: it rises radially at its OWN pad (inside its own
    base band, which carries no comb), folds AWAY from its gap, then runs back over the
    top of its own teeth and drops down at the gap -- which lies on the empty side of the
    pad, inside this input's own bands the whole way down. Teeth count is forced odd so
    the fold finishes at the outer radius and the run back cannot retrace a tooth."""
    lo, hi = band
    u, v = inward_along(e)
    cu = (-u[0], -u[1])
    a_pad = pad["x"] * v[0] + pad["y"] * v[1]
    r0 = PAD_SIZE / 2.0 + COIL_GAP
    blo = lo + (WIRE_SPACE + WIRE_WIDTH) / 2.0
    bhi = hi - (WIRE_SPACE + WIRE_WIDTH) / 2.0
    s = 1.0 if (gap is None or gap <= a_pad) else -1.0     # grow away from the gap
    if gap is None:
        gap = a_pad - s * ROUTE_PITCH
    span = (bhi - a_pad) if s > 0 else (a_pad - blo)
    fit = max(3, int(max(span, 0.0) / ROUTE_PITCH))
    cap = max(3, int(target_len / (4.0 * ROUTE_PITCH)))
    m = max(3, min(fit, cap)) // 2 * 2 + 1                 # odd -> ends at the outer edge

    def build(depth):
        P = lambda t, r: to_xy(e, t, r, bounds)
        rtop = rc + depth + ROUTE_PITCH
        pts = [(pad["x"], pad["y"]), P(a_pad, rc)]         # riser, straight up its pad
        out = True
        for j in range(m):
            t = a_pad + s * j * ROUTE_PITCH
            if out:
                pts += [P(t, rc), P(t, rc + depth)]
                if j < m - 1:
                    pts.append(P(t + s * ROUTE_PITCH, rc + depth))
            else:
                pts += [P(t, rc + depth), P(t, rc)]
                if j < m - 1:
                    pts.append(P(t + s * ROUTE_PITCH, rc))
            out = not out
        t_end = a_pad + s * (m - 1) * ROUTE_PITCH
        pts += [P(t_end, rtop), P(gap, rtop), P(gap, r0)]  # over the top, down the gap
        ret = return_through_gap(0.0, 0.0, P(gap, r0), cu, plane)
        return pts, ret

    depth, coil, ret = max(ROUTE_PITCH, target_len / m), None, None
    for _ in range(4):                                     # length is linear in depth
        coil, ret = build(depth)
        clen = polyline_len(coil) + polyline_len(ret)
        depth = max(ROUTE_PITCH, depth + (target_len - clen) / (m + 1))
    r_ohm, clen = _coil_result(pad, coil, ret, plane)
    return coil, ret, r_ohm, clen


def build_input_coil(pad, e, plan, target_len, bounds, plane):
    """Build one input's resistor from its `plan` -- from plan_edge."""
    if isinstance(plan, dict):
        return build_channel_coil(pad, e, plan, target_len, bounds, plane)
    band, gap, rc = plan
    if rc is None:
        return build_band_coil(pad, e, band, target_len, gap, bounds, plane)
    return build_canopy_coil(pad, e, band, target_len, gap, rc, bounds, plane)


# ----------------------------------------------------------------------
# Draw one group on one layer
# ----------------------------------------------------------------------
def rect(x0, y0, x1, y1, layer):
    return gdstk.rectangle((x0, y0), (x1, y1), layer=layer, datatype=METAL_DT)


def pad_poly(p, layer):
    h = PAD_SIZE / 2.0
    return rect(p["x"] - h, p["y"] - h, p["x"] + h, p["y"] + h, layer)


def via_polys(pt, k):
    """Via stack joining metal k and metal k+1 at pt (cut + both metals)."""
    h = VIA_SIZE / 2.0
    x, y = pt
    return [gdstk.rectangle((x - h, y - h), (x + h, y + h), layer=Lr, datatype=d)
            for Lr, d in ((via_layer(k), VIA_DT),
                          (metal_layer(k), METAL_DT),
                          (metal_layer(k + 1), METAL_DT))]


def central_plane(bounds):
    """The shared-node PLANE filling the ring interior, inset PLANE_RING_MARGIN to
    clear the pads. Returns (px0, py0, px1, py1)."""
    minx, maxx, miny, maxy = bounds
    return (minx + PLANE_RING_MARGIN, miny + PLANE_RING_MARGIN,
            maxx - PLANE_RING_MARGIN, maxy - PLANE_RING_MARGIN)


MIN_SPLIT_INPUTS = 4       # below this a coupon is already easy to read; don't split


def plan_split(group, bounds, all_pads):
    """Divide a coupon in two, each half getting its own plane, its own share of the
    OUTPUT pads and its own restarted ladder. Halving the inputs is what buys decode
    margin: 4 resistors need only 1 part in 15 resolved instead of 1 part in 255, and
    the top resistor drops from 128k to 8k.

    The cut has to be SPATIAL, because each input returns straight inward and must land
    on its own half's plane. Pads whose return or finger lands on a plane SIDE pin the
    divider down, so it is placed in the clear band between the two halves -- never
    through a pad, which would leave a return stranded in the isolation gap.

    Tries a horizontal and a vertical cut and keeps the better balanced. Returns
    (axis, divider) or None if no clean cut exists. Output candidates include the
    OUTPUTSPLIT pads, so the divider is the same in the sizing and drawing passes even
    though a 2-layer chip only hands half of them to each layer."""
    # A continuity coupon is already one shorted node; halving it would leave each
    # input reachable from only half the returns, which its CSV does not express. The
    # calibration coupon is a star onto ONE common pad, and splitting it would restart
    # the ladder per half and give duplicate values instead of one rung each.
    if (group.get("single") or group.get("calib")
            or len(group["inputs"]) < MIN_SPLIT_INPUTS):
        return None
    ins = group["inputs"]
    outs = list(group["outputs"]) + [p for p in all_pads if p["io"] == "output"
                                     and p["group"] == SPLIT_OUT]
    if not outs:
        return None
    minx, maxx, miny, maxy = bounds
    best = None
    for axis in ("y", "x"):
        key = (lambda p: p["y"]) if axis == "y" else (lambda p: p["x"])
        lo, hi = (miny, maxy) if axis == "y" else (minx, maxx)
        mid = (lo + hi) / 2.0
        li = [p for p in ins if key(p) < mid]
        hi_i = [p for p in ins if key(p) >= mid]
        if not (li and hi_i):
            continue
        # only pads whose return or finger lands on a plane SIDE pin the divider
        side = ("left", "right") if axis == "y" else ("bottom", "top")
        on_side = [p for p in ins + outs if edge_of(p, bounds) in side]
        cl = [key(p) for p in on_side if key(p) < mid]
        ch = [key(p) for p in on_side if key(p) >= mid]
        d_lo = (max(cl) + PAD_SIZE) if cl else lo
        d_hi = (min(ch) - PAD_SIZE) if ch else hi
        if d_hi - d_lo < PLANE_SPLIT_GAP + 2 * FINGER_OVERLAP:
            continue                                  # no clear band for the divider
        div = (d_lo + d_hi) / 2.0
        if not (any(key(p) < div for p in outs) and any(key(p) >= div for p in outs)):
            continue                                  # a half would have no output
        cand = (abs(len(li) - len(hi_i)), -min(len(li), len(hi_i)), axis, div)
        if best is None or cand[:2] < best[:2]:
            best = cand
    return None if best is None else (best[2], best[3])


def split_planes(bounds, axis, divider):
    """The two isolated half-planes either side of `divider`."""
    px0, py0, px1, py1 = central_plane(bounds)
    g = PLANE_SPLIT_GAP / 2.0
    if axis == "y":
        return (px0, py0, px1, divider - g), (px0, divider + g, px1, py1)
    return (px0, py0, divider - g, py1), (divider + g, py0, px1, py1)


def coupon_planes(group, bounds, all_pads):
    """Plane geometry for one coupon. Returns (planes, split); split is None when the
    coupon is drawn whole, else (axis, divider, plane_low, plane_high)."""
    whole = central_plane(bounds)
    sp = plan_split(group, bounds, all_pads) if SPLIT_HALVES else None
    if sp is None:
        return [whole], None
    axis, divider = sp
    pl, ph = split_planes(bounds, axis, divider)
    return [pl, ph], (axis, divider, pl, ph)


def half_of(p, split):
    """Which half a pad belongs to: (half_id, its plane)."""
    axis, divider, pl, ph = split
    v = p["y"] if axis == "y" else p["x"]
    return (1, pl) if v < divider else (2, ph)


def input_targets(group, input_sets=None):
    """Target trace length per input. Each half RESTARTS the ladder, so a split coupon
    runs ranks 1..n/2 twice -- that is exactly what shrinks the biggest resistor and
    widens the decode margin. Split halves rank on their own input centroid, which is
    known in both the sizing and the drawing pass, so the two agree."""
    if not input_sets:
        return {id(p): input_target_len(k)
                for k, p in enumerate(ordered_inputs(group), 1)}
    out = {}
    for ins in input_sets:
        for k, p in enumerate(ordered_inputs({"inputs": ins, "outputs": []}), 1):
            out[id(p)] = input_target_len(k)
    return out


def _plane_label(cell, txt, plane, layer_idx):
    """Centred coupon label on the plane; 2-layer chips stack their two labels."""
    px0, py0, px1, py1 = plane
    size = min(200.0, max(40.0, (px1 - px0) * 0.8 / (0.62 * len(txt)))) * 0.7
    cx = (px0 + px1) / 2.0 - 0.62 * size * len(txt) / 2.0
    cy = (py0 + py1) / 2.0 - size / 2.0 + (1 - 2 * layer_idx) * size * 1.6
    cell.add(*gdstk.text(txt, size, (cx, cy), layer=LABEL_LAYER, datatype=LABEL_DT))


def draw_group(group, layer_idx, bounds, cell, chip_idx, all_pads, extra_outputs=(),
               single=False):
    """One group on metal `layer_idx`: a central PLANE is the shared node; each INPUT
    grows a resistor comb outside the ring, folded to fill its own along-edge band,
    returning inward through a pad gap to the plane. Each OUTPUT joins the plane with
    a pad-width finger. `extra_outputs` are this layer's OUTPUTSPLIT taps -- drawn as
    extra output fingers but kept OUT of input ranking (which uses the group's own
    outputs, or its input centroid), so ranks match the sizing pass. `single` draws a
    continuity coupon instead: no combs -- every INPUTSINGLE and OUTPUTSINGLE pad is
    shorted straight to the plane, so each independent input is a per-pin continuity
    check against the shared return. 2nd-layer groups get a via at each pad. Returns
    (wiring_polys, csv_rows)."""
    L = metal_layer(layer_idx)
    needs_via = layer_idx > 0
    hp = PAD_SIZE / 2.0
    all_outs = list(group["outputs"]) + list(extra_outputs)
    polys = []

    def add(obj):
        cell.add(obj)
        polys.extend(obj.to_polygons() if isinstance(obj, gdstk.FlexPath) else [obj])

    def via_at(pt):
        for vp in via_polys(pt, 0):                 # stitch metal L down to top
            add(vp)

    # One plane, or two isolated half-planes when the coupon is split in two.
    planes, split = coupon_planes(group, bounds, all_pads)
    if split is not None:                 # both halves must end up usable on this layer
        lo_i = [p for p in group["inputs"] if half_of(p, split)[0] == 1]
        hi_i = [p for p in group["inputs"] if half_of(p, split)[0] == 2]
        lo_o = [p for p in all_outs if half_of(p, split)[0] == 1]
        hi_o = [p for p in all_outs if half_of(p, split)[0] == 2]
        if not all((lo_i, hi_i, lo_o, hi_o)):
            planes, split = [central_plane(bounds)], None
    for pp in planes:
        add(rect(pp[0], pp[1], pp[2], pp[3], L))
    plane = planes[0]
    plane_of = ({id(p): half_of(p, split)[1]
                 for p in list(group["inputs"]) + list(all_outs)} if split
                else {id(p): plane for p in list(group["inputs"]) + list(all_outs)})
    edge_coords = near_edge_coords(all_pads, bounds)

    def finger(o):                        # pad-width tab joining o to ITS OWN plane
        px0, py0, px1, py1 = plane_of.get(id(o), plane)
        e = edge_of(o, bounds)
        cx, cy = o["x"], o["y"]
        if e in ("bottom", "top"):
            fx0, fx1 = cx - hp, cx + hp
            add(rect(fx0, cy, fx1, py0 + FINGER_OVERLAP, L) if e == "bottom"
                else rect(fx0, py1 - FINGER_OVERLAP, fx1, cy, L))
        else:
            fy0, fy1 = cy - hp, cy + hp
            add(rect(cx, fy0, px0 + FINGER_OVERLAP, fy1, L) if e == "left"
                else rect(px1 - FINGER_OVERLAP, fy0, cx, fy1, L))
        if needs_via:
            via_at((cx, cy))

    if single:                                       # continuity coupon: short all pads
        # Returns are OUTPUTSINGLE pads plus any OUTPUTSPLIT half on this layer.
        if ADD_LABELS:
            ins = ", ".join(p["name"] for p in group["inputs"])
            rets = ", ".join(p["name"] for p in all_outs)
            _plane_label(cell, f"Single inputs: {ins}; Return: {rets}", plane, layer_idx)
        for o in group["inputs"] + all_outs:
            finger(o)
        rows = [[chip_idx, layer_idx + 1, 0, inp["name"], inp["signal"],
                 ";".join(o["name"] for o in all_outs), "0.00", "0",
                 "continuity"] for inp in group["inputs"]]
        return polys, rows

    # Each half is its own decode unit, so it gets its own label and its own outputs.
    units = ([(1, lo_i, lo_o, split[2]), (2, hi_i, hi_o, split[3])] if split
             else [(0, group["inputs"], all_outs, plane)])
    if ADD_LABELS:
        for hid, ins, outs, pp in units:
            tag = f"Half {hid} - " if hid else ""
            names = ", ".join(p["name"] for p in
                              ordered_inputs({"inputs": ins, "outputs": outs}))
            _plane_label(cell, f"{tag}Inputs: {names}, Outputs: "
                               f"{', '.join(o['name'] for o in outs)}", pp, layer_idx)

    # Bands are solved per EDGE across the whole coupon, so the two halves' combs stay
    # disjoint even where both have inputs on the same edge.
    target_of = input_targets(group, [u[1] for u in units] if split else None)
    plan_of = {}
    for e, items in edge_inputs(group, bounds, target_of).items():
        elo, ehi = edge_along_range(e, bounds)
        plans = plan_edge(items, elo, ehi, edge_gaps(e, edge_coords))
        for (_, _, pad), pl in zip(items, plans):
            plan_of[id(pad)] = (e, pl)

    rows = []
    for hid, ins, outs, pp in units:
        r_vals, hrows = [], []
        for inp in ordered_inputs({"inputs": ins, "outputs": outs}):
            e, pl = plan_of[id(inp)]
            coil, ret, r, clen = build_input_coil(inp, e, pl, target_of[id(inp)],
                                                  bounds, pp)
            add(gdstk.FlexPath(coil, WIRE_WIDTH, layer=L, datatype=METAL_DT))
            add(gdstk.FlexPath(ret, WIRE_WIDTH, layer=L, datatype=METAL_DT))
            if needs_via:
                via_at((inp["x"], inp["y"]))
            r_vals.append(r)
            hrows.append([chip_idx, layer_idx + 1, hid, inp["name"], inp["signal"],
                          ";".join(o["name"] for o in outs), f"{r:.2f}", f"{clen:.0f}"])
        # all-contacting parallel resistance for THIS half, on each of its rows
        gpar = 1.0 / sum(1.0 / r for r in r_vals) if r_vals else float("inf")
        for row in hrows:
            row.append(f"{gpar:.3f}")
        rows.extend(hrows)
        for o in outs:
            finger(o)
    return polys, rows


# ----------------------------------------------------------------------
# Verify (cross-net overlap on a shared layer -- should be none)
# ----------------------------------------------------------------------
def group_bbox(polys):
    xmin = ymin = math.inf
    xmax = ymax = -math.inf
    for p in polys:
        bb = p.bounding_box()
        if bb is None:
            continue
        (a, b), (c, d) = bb
        xmin, ymin = min(xmin, a), min(ymin, b)
        xmax, ymax = max(xmax, c), max(ymax, d)
    return xmin, ymin, xmax, ymax


def find_shorts(geo, num_layers):
    layers = tuple(metal_layer(k) for k in range(num_layers))
    ids = list(geo.keys())
    bbox = {nid: group_bbox(polys) for nid, polys in geo.items()}
    by_layer = {(nid, L): [p for p in polys if p.layer == L]
                for nid, polys in geo.items() for L in layers}
    shorts = []
    for a in range(len(ids)):
        xa0, ya0, xa1, ya1 = bbox[ids[a]]
        for b in range(a + 1, len(ids)):
            xb0, yb0, xb1, yb1 = bbox[ids[b]]
            if xa1 < xb0 or xb1 < xa0 or ya1 < yb0 or yb1 < ya0:
                continue
            for L in layers:
                pa, pb = by_layer[(ids[a], L)], by_layer[(ids[b], L)]
                if pa and pb and gdstk.boolean(pa, pb, "and"):
                    shorts.append((ids[a], ids[b], L))
    return shorts


# ----------------------------------------------------------------------
# Calibration coupon (human multimeter measurement)
# ----------------------------------------------------------------------
def calib_coupon(out_pads, n_res):
    """The calibration coupon as a COUPON on the normal pad ring, so the probe card
    lands on it exactly as it lands on a test chip.

    The resistors hang off the OUTPUT pads, because those are NOT shorted to each other
    externally -- each one is its own net on the card, so each resistor can be read on
    its own. A handful of spread-out output pads carry one known resistor each to the
    central plane; EVERY remaining output pad ties straight to that plane as a COMMON
    tap. Measure any COMMON to resistor pad k and you read resistor k alone, since no
    other pad carries current. Having many COMMON taps is redundancy: they are all the
    same node, so if a probe misses one you simply read from another.

    Values are the same binary ladder the test chips use, so every rung has a matching
    calibration resistor. The pads are picked SPREAD around the ring rather than in
    list order: bunching them on one edge would leave each comb a narrow band, forcing
    a deep fold and growing the die for every chip on the wafer. Returns the coupon
    dict, or None if there are too few output pads."""
    if len(out_pads) < 2:
        return None
    n = max(1, min(n_res, len(out_pads) - 1))
    cx = sum(p["x"] for p in out_pads) / len(out_pads)
    cy = sum(p["y"] for p in out_pads) / len(out_pads)
    ring = sorted(out_pads, key=lambda p: math.atan2(p["y"] - cy, p["x"] - cx))
    step = len(ring) / float(n + 1)                  # evenly spaced around the ring
    driven = [ring[int((i + 1) * step) % len(ring)] for i in range(n)]
    seen = {id(p) for p in driven}
    commons = [p for p in out_pads if id(p) not in seen]
    return {"num": "CAL", "sub": 0, "calib": True,
            "inputs": driven, "outputs": commons}


def label_calib_values(cell, rows, driven, bounds):
    """Write each resistor's value beside the OUTPUT pad it hangs off, just inboard of
    the ring so it does not sit under the comb, and large enough to read on the layout
    when you are picking which pin to probe."""
    if not ADD_LABELS:
        return
    r_of = {row[3]: row[6] for row in rows}
    size = PAD_SIZE * 0.75
    for pad in driven:
        val = r_of.get(pad["name"])
        if val is None:
            continue
        u, _ = inward_along(edge_of(pad, bounds))          # points into the die
        txt = f"{float(val):,.0f}ohm"
        x = pad["x"] + u[0] * PAD_SIZE * 1.1
        y = pad["y"] + u[1] * PAD_SIZE * 1.1
        if u[0] < 0:                                       # right edge: run text leftward
            x -= 0.62 * size * len(txt)
        if u[1] != 0:                                      # top/bottom: centre it
            x -= 0.62 * size * len(txt) / 2.0
            y -= size / 2.0 if u[1] < 0 else -size / 2.0
        cell.add(*gdstk.text(txt, size, (x, y), layer=LABEL_LAYER, datatype=LABEL_DT))


def mosaic_place(entries, gap, cols):
    """Row-major placement of `(cell, w, h)` entries into `cols` columns. Returns
    (origins, field_w, field_h): origins[i] is the top-left corner of entry i
    (each cell spans [x, x+w] x [y-h, y], matching the die outline convention),
    field_w is the widest row, field_h the total stack height."""
    origins, x, y, row_h, max_w = [], 0.0, 0.0, 0.0, 0.0
    for i, (c, w, h) in enumerate(entries):
        if i % cols == 0 and i:
            max_w = max(max_w, x - gap)               # drop the trailing gap
            x, y, row_h = 0.0, y - row_h - gap, 0.0
        origins.append((x, y))
        x += w + gap
        row_h = max(row_h, h)
    max_w = max(max_w, x - gap)                        # last row
    return origins, max_w, row_h - y                   # y <= 0, so height = row_h - y


def wafer_field_origins(fw, fh, diam, street, edge_excl, offset=(0.0, 0.0)):
    """Top-left origins of every step-and-repeat field whose WHOLE footprint fits
    inside the usable wafer radius (diam/2 - edge_excl), for an array shifted by
    `offset` from the wafer centre. The array is periodic in `step`, so sweeping
    offsets over one period finds the best centring."""
    r = diam / 2.0 - edge_excl
    step_x, step_y = fw + street, fh + street
    ox0, oy0 = offset
    ni = math.ceil(diam / 2.0 / step_x) + 2
    nj = math.ceil(diam / 2.0 / step_y) + 2
    out = []
    for i in range(-ni, ni + 1):
        for j in range(-nj, nj + 1):
            ox = i * step_x - fw / 2.0 + ox0
            oy = j * step_y + fh / 2.0 + oy0
            corners = [(ox, oy), (ox + fw, oy), (ox, oy - fh), (ox + fw, oy - fh)]
            if all(math.hypot(cx, cy) <= r for cx, cy in corners):
                out.append((ox, oy))
    return out


def best_field_layout(entries, gap, diam, street, edge_excl):
    """Pick the mosaic column count -- and the array-centring offset -- that packs
    the most complete fields onto the wafer. Every field holds all the dies, so
    maximising fields maximises good dies. This is the shot-map optimisation a
    real stepper flow does: sweep the layout aspect ratio and the wafer-centre
    offset instead of guessing columns from sqrt(n). Returns
    (cols, field_w, field_h, field_origins)."""
    n = len(entries)
    fracs = (0.0, 0.25, 0.5, 0.75)                     # offset samples over one period
    best = None
    for cols in range(1, n + 1):
        _, fw, fh = mosaic_place(entries, gap, cols)
        step_x, step_y = fw + street, fh + street
        for a in fracs:
            for b in fracs:
                fields = wafer_field_origins(fw, fh, diam, street, edge_excl,
                                             (-a * step_x, -b * step_y))
                # Prefer more fields; break ties toward the squarer, tighter field.
                key = (len(fields), -max(fw, fh))
                if best is None or key > best[0]:
                    best = (key, cols, fw, fh, fields)
    _, cols, fw, fh, fields = best
    return cols, fw, fh, fields


def build_mosaic(entries, gap, cols):
    """Tile every (cell, w, h) entry -- each chip plus the calibration coupon --
    into one MOSAIC cell of `cols` columns, so the whole chip-set is a single
    reusable field. Cells keep their own layers/positions; this only adds a top
    MOSAIC cell with a translated Reference per die. Returns (Library, field_w,
    field_h); the library holds every referenced cell too, so the GDS is
    self-contained."""
    origins, fw, fh = mosaic_place(entries, gap, cols)
    lib = gdstk.Library(unit=1e-6, precision=1e-9)
    lib.add(*[c for c, _, _ in entries])
    top = lib.new_cell("MOSAIC")
    for (c, w, h), origin in zip(entries, origins):
        top.add(gdstk.Reference(c, origin=origin))
    return lib, fw, fh


def build_wafer(field_lib, field_origins, diam, layer, dt):
    """Step-and-repeat the field_lib's MOSAIC cell (every chip + the calibration
    coupon) across a circular wafer, one Reference per pre-solved field origin.
    Added LAST, on an overlay layer that is NOT fabricated: the physical wafer
    edge (full radius). Fields are still packed inside the edge-exclusion limit;
    that boundary just isn't drawn. Returns (Library, fields_placed)."""
    field = next(c for c in field_lib.cells if c.name == "MOSAIC")
    lib = gdstk.Library(unit=1e-6, precision=1e-9)
    lib.add(*field_lib.cells)
    top = lib.new_cell("WAFER")
    for origin in field_origins:
        top.add(gdstk.Reference(field, origin=origin))
    r = diam / 2.0
    top.add(gdstk.ellipse((0.0, 0.0), r, inner_radius=r - WAFER_LINE_UM,
                          tolerance=20.0, layer=layer, datatype=dt))
    return lib, len(field_origins)


# ----------------------------------------------------------------------
# Schematic (a circuit diagram per chip, one block per layer)
# ----------------------------------------------------------------------
def _resistor_zig(x0, x1, y, teeth=6, amp=7.0):
    """SVG polyline points for a horizontal resistor symbol from x0 to x1."""
    lead = (x1 - x0) * 0.13
    a, b = x0 + lead, x1 - lead
    step = (b - a) / teeth
    pts = [(x0, y), (a, y)]
    for i in range(teeth):
        pts.append((a + (i + 0.5) * step, y + (amp if i % 2 == 0 else -amp)))
    pts += [(b, y), (x1, y)]
    return " ".join(f"{px:.1f},{py:.1f}" for px, py in pts)


def build_schematic_svg(chip_idx, rows):
    """A circuit diagram for one chip: for each layer (coupon) every INPUT pad is
    drawn as a node wired through its resistor (zigzag) to the shared OUTPUT node.
    `rows` are that chip's CSV rows; returns the SVG text."""
    by_layer = {}
    for r in rows:
        by_layer.setdefault((int(r[1]), int(r[2])), []).append(r)

    x_name, x_term, x_r0, x_r1, x_bus, x_out = 14, 150, 162, 300, 430, 452
    row_h, head_h, foot_h, gap = 46, 40, 60, 30
    width = 640
    body, y = [], 20.0
    for key in sorted(by_layer):
        layer, half = key
        rs = by_layer[key]
        outs = rs[0][5]
        gpar = rs[0][8]
        is_single = gpar == "continuity"             # SINGLE coupon: shorts, no ladder
        subtitle = "continuity" if is_single else f"all-parallel {gpar} Ω"
        name = f"Layer {layer}" + (f" half {half}" if half else "")
        top = y
        body.append(f'<text x="{x_name}" y="{top:.0f}" class="ttl">'
                    f'Chip {chip_idx} — {name}  ({subtitle})</text>')
        row_y = top + head_h
        ys = [row_y + i * row_h for i in range(len(rs))]
        for (row, ry) in zip(rs, ys):
            _, _, _, pad, sig, _, r_ohm, _, _ = row
            body.append(f'<rect x="{x_term-9:.0f}" y="{ry-9:.0f}" width="18" '
                        f'height="18" class="pad"/>')
            body.append(f'<text x="{x_name}" y="{ry-2:.0f}" class="lbl">{pad}</text>')
            if sig:
                body.append(f'<text x="{x_name}" y="{ry+11:.0f}" '
                            f'class="sig">{_esc(sig)}</text>')
            body.append(f'<line x1="{x_term+9:.0f}" y1="{ry:.0f}" x2="{x_r0:.0f}" '
                        f'y2="{ry:.0f}" class="wire"/>')
            if is_single:                            # direct short to the return node
                body.append(f'<line x1="{x_r0:.0f}" y1="{ry:.0f}" x2="{x_bus:.0f}" '
                            f'y2="{ry:.0f}" class="wire"/>')
            else:
                body.append(f'<polyline points="{_resistor_zig(x_r0, x_r1, ry)}" '
                            f'class="wire"/>')
                body.append(f'<text x="{(x_r0+x_r1)/2:.0f}" y="{ry-12:.0f}" '
                            f'class="val" text-anchor="middle">{r_ohm} Ω</text>')
                body.append(f'<line x1="{x_r1:.0f}" y1="{ry:.0f}" x2="{x_bus:.0f}" '
                            f'y2="{ry:.0f}" class="wire"/>')
        # shared node: vertical bus tying every input, then out to the return/output pad
        oy = ys[-1] + foot_h - 24
        body.append(f'<line x1="{x_bus:.0f}" y1="{ys[0]:.0f}" x2="{x_bus:.0f}" '
                    f'y2="{oy:.0f}" class="bus"/>')
        body.append(f'<line x1="{x_bus:.0f}" y1="{oy:.0f}" x2="{x_out:.0f}" '
                    f'y2="{oy:.0f}" class="wire"/>')
        body.append(f'<circle cx="{x_out:.0f}" cy="{oy:.0f}" r="4" class="node"/>')
        body.append(f'<text x="{x_out+10:.0f}" y="{oy+4:.0f}" class="lbl">'
                    f'{"RETURN" if is_single else "OUTPUT"}: {_esc(outs)}</text>')
        y = ys[-1] + foot_h + gap
    height = y
    style = ("<style>.wire{stroke:#222;stroke-width:1.6;fill:none}"
             ".bus{stroke:#222;stroke-width:2.4;fill:none}"
             ".pad{fill:#fff;stroke:#222;stroke-width:1.6}.node{fill:#222}"
             ".lbl{font:12px sans-serif;fill:#111}.sig{font:10px sans-serif;fill:#666}"
             ".val{font:11px sans-serif;fill:#0645ad}"
             ".ttl{font:600 14px sans-serif;fill:#000}</style>")
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" '
            f'height="{height:.0f}" viewBox="0 0 {width} {height:.0f}">'
            f'{style}<rect width="{width}" height="{height:.0f}" fill="#fff"/>'
            + "".join(body) + "</svg>")


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
INPUT_EXTS = (".csv", ".xlsx", ".xlsm")


def find_input_csv():
    """The single pinout in inputs/ (any .csv or .xlsx name; Excel lock files
    beginning '~$' are ignored). Errors if none or several."""
    files = sorted(p for p in INPUT_DIR.glob("*")
                   if p.suffix.lower() in INPUT_EXTS and not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"No .csv or .xlsx found in {INPUT_DIR} -- put your pinout there.")
    if len(files) > 1:
        names = ", ".join(p.name for p in files)
        raise SystemExit(f"Multiple input files in {INPUT_DIR} ({names}); keep just one, "
                         f"or pass the path: py io_pair_wiring_parallel.py <file>")
    return files[0]


def parse_args(argv):
    """(csv_path, layers, split). `--layers 1|2` and `--split`/`--no-split` skip their
    prompts; either is None when not given. With no path, uses the single file in
    inputs/."""
    layers, split, pos = None, None, []
    i = 1
    while i < len(argv):
        if argv[i] in ("--layers", "-l") and i + 1 < len(argv):
            layers = argv[i + 1]
            i += 2
        elif argv[i] == "--split":
            split, i = True, i + 1
        elif argv[i] == "--no-split":
            split, i = False, i + 1
        else:
            pos.append(argv[i])
            i += 1
    csv_path = pathlib.Path(pos[0]) if pos else find_input_csv()
    return csv_path, (int(layers) if layers in ("1", "2") else None), split


def ask_split(default=False):
    """Ask whether to halve each coupon into two independent 4-resistor sets."""
    if not sys.stdin.isatty():
        return default
    prompt = ("\nSplit each coupon into two independent halves?\n"
              "  n = no  - one set of resistors per layer, one reading\n"
              "  y = yes - two halves per layer, each with its own plane and half the\n"
              "            output pads, so each half is read on its own. Halving the\n"
              "            resistor count per reading widens the decode margin a lot\n"
              "            and shrinks the biggest resistor.\n"
              "            Needs the two halves' output pads NOT shorted to each other\n"
              "            externally, or the halves rejoin off-chip.\n"
              f"Enter y or n [{'y' if default else 'n'}]: ")
    while True:
        try:
            ans = input(prompt).strip().lower()
        except EOFError:
            return default
        if ans == "":
            return default
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("  Please enter y or n.")


def ask_layers(default=2):
    """1 layer -> one group per chip (no vias); 2 -> two groups (2nd on metal 2,
    via-stitched). Non-interactive runs use `default`."""
    if not sys.stdin.isatty():
        return default
    prompt = ("\nBuild chips with how many metal layers?\n"
              "  1 = one layer  (one IO group per chip, no vias)\n"
              "  2 = two layers (two IO groups per chip)\n"
              f"Enter 1 or 2 [{default}]: ")
    while True:
        try:
            ans = input(prompt).strip()
        except EOFError:
            return default
        if ans == "":
            return default
        if ans in ("1", "2"):
            return int(ans)
        print("  Please enter 1 or 2.")


def size_and_place(pads, groups, min_die=(0.0, 0.0)):
    """Build each coil at its raw position to measure per-edge protrusion, set the
    global die size (>= farthest comb + DIE_MARGIN_BUFFER, ring aspect kept, and
    >= `min_die` in each dimension so the chips match the calibration coupon), and
    centre the ring. Returns (edge_depth, ring_w, ring_h, scale)."""
    rxs = [p["x"] for p in pads]
    rys = [p["y"] for p in pads]
    raw_bounds = (min(rxs), max(rxs), min(rys), max(rys))
    raw_edges = near_edge_coords(pads, raw_bounds)
    raw_plane = central_plane(raw_bounds)
    ring_w, ring_h = max(rxs) - min(rxs), max(rys) - min(rys)
    edge_depth = {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}
    for g in groups:
        if g.get("single"):                 # continuity coupon has fingers, no combs
            continue
        planes, split = coupon_planes(g, raw_bounds, pads)
        sets = ([[p for p in g["inputs"] if half_of(p, split)[0] == h] for h in (1, 2)]
                if split else None)
        target_of = input_targets(g, sets)
        for e, items in edge_inputs(g, raw_bounds, target_of).items():
            elo, ehi = edge_along_range(e, raw_bounds)
            plans = plan_edge(items, elo, ehi, edge_gaps(e, raw_edges))
            for (_, tl, pad), pl in zip(items, plans):
                coil, _, _, _ = build_input_coil(
                    pad, e, pl, tl, raw_bounds,
                    half_of(pad, split)[1] if split else raw_plane)
                xs = [p[0] for p in coil]
                ys = [p[1] for p in coil]
                prot = {"left": raw_bounds[0] - min(xs), "right": max(xs) - raw_bounds[1],
                        "bottom": raw_bounds[2] - min(ys), "top": max(ys) - raw_bounds[3]}[e]
                edge_depth[e] = max(edge_depth[e], prot + WIRE_WIDTH / 2.0)

    buf = DIE_MARGIN_BUFFER
    content_w = ring_w + edge_depth["left"] + edge_depth["right"] + 2 * buf
    content_h = ring_h + edge_depth["top"] + edge_depth["bottom"] + 2 * buf
    scale = max(content_w / ring_w, content_h / ring_h)
    global DIE_W, DIE_H
    DIE_W, DIE_H = scale * ring_w, scale * ring_h
    DIE_W, DIE_H = max(DIE_W, min_die[0]), max(DIE_H, min_die[1])  # match calibration
    margins = {                                 # centre the content in the die
        "left": edge_depth["left"] + buf + (DIE_W - content_w) / 2,
        "right": edge_depth["right"] + buf + (DIE_W - content_w) / 2,
        "top": edge_depth["top"] + buf + (DIE_H - content_h) / 2,
        "bottom": edge_depth["bottom"] + buf + (DIE_H - content_h) / 2,
    }
    place_pads(pads, raw_bounds, margins)
    return edge_depth, ring_w, ring_h, scale


def split_for_chip(split_outs, n_layers):
    """Halve the OUTPUTSPLIT pads across the chip's layers, alternating so each half
    is spread out (list them in perimeter order to spread by position). A half serves
    as extra output taps for a parallel layer, or as the shorted return for a single
    continuity layer. A one-layer chip takes them all. Returns {layer_idx: [pads]}."""
    split_outs = list(split_outs)
    if not split_outs:
        return {}
    if n_layers <= 1:
        return {0: split_outs}
    return {0: split_outs[0::2], 1: split_outs[1::2]}


def build_chip(ci, chip_groups, pads, bounds, split_outs=()):
    """One chip's GDS: every pad on the top layer (wired ones tagged to their group),
    each group's combs on its own metal layer, then the die outline LAST. OUTPUTSPLIT
    pads are halved across the parallel layers -- each half wired to that layer's plane
    as extra output taps. A single/continuity coupon shorts its INPUTSINGLE and
    OUTPUTSINGLE pads to its plane. Returns (library, geo, csv_rows); geo maps
    net-id -> polygons."""
    top = metal_layer(0)
    lib = gdstk.Library(unit=1e-6, precision=1e-9)
    cell = lib.new_cell(f"CHIP{ci}")
    geo = {}
    split_by_layer = split_for_chip(split_outs, len(chip_groups))
    pad_net = {}
    for group in chip_groups:
        for p in group["inputs"] + group["outputs"]:
            pad_net[id(p)] = ("group", group["num"])
    for li, extra in split_by_layer.items():       # split pads join their layer's net
        if li < len(chip_groups):
            for p in extra:
                pad_net[id(p)] = ("group", chip_groups[li]["num"])
    for p in pads:
        key = pad_net.get(id(p), ("pads",))
        poly = pad_poly(p, top)
        cell.add(poly)
        geo.setdefault(key, []).append(poly)
        if ADD_LABELS and p["name"]:
            inset = LABEL_SIZE * 0.4                  # tuck labels inside the pad
            lx = p["x"] - PAD_SIZE / 2 + inset
            ty = p["y"] + PAD_SIZE / 2 - inset - LABEL_SIZE  # top text line, inside
            if p["io"]:                               # IO tag on top, pad name below
                cell.add(*gdstk.text(p["io"].capitalize(), LABEL_SIZE, (lx, ty),
                                     layer=LABEL_LAYER, datatype=IO_LABEL_DT))
                cell.add(*gdstk.text(p["name"], LABEL_SIZE, (lx, ty - LABEL_SIZE * 1.2),
                                     layer=LABEL_LAYER, datatype=LABEL_DT))
            else:
                cell.add(*gdstk.text(p["name"], LABEL_SIZE, (lx, ty),
                                     layer=LABEL_LAYER, datatype=LABEL_DT))
    rows = []
    for li, group in enumerate(chip_groups):
        wpolys, grows = draw_group(group, li, bounds, cell, ci, pads,
                                   split_by_layer.get(li, []),
                                   single=group.get("single", False))
        geo.setdefault(("group", group["num"]), []).extend(wpolys)
        rows.extend(grows)
    if ADD_DIE_OUTLINE:                          # bottom layer, built LAST
        cell.add(gdstk.rectangle((0.0, 0.0), (DIE_W, -DIE_H),
                                 layer=BOUNDARY_LAYER, datatype=BOUNDARY_DT))
    return lib, geo, rows


def main():
    csv_path, layers, split = parse_args(sys.argv)
    if layers is None:
        layers = ask_layers()
    if split is None:
        split = ask_split()
    global SPLIT_HALVES
    SPLIT_HALVES = split
    groups_per_chip = layers
    pads = read_pads(csv_path)
    single_in = [p for p in pads if p["io"] == "input" and p["group"] == SINGLE]
    single_out = [p for p in pads if p["io"] == "output" and p["group"] == SINGLE]
    has_split = any(p["group"] == SPLIT_OUT for p in pads)
    if single_out and not single_in:                 # OUTPUTSINGLE only pairs with INPUTSINGLE
        raise SystemExit("OUTPUTSINGLE pads only pair with INPUTSINGLE: add INPUTSINGLE "
                         "pads, or remove the OUTPUTSINGLE ones.")
    if single_in and not (single_out or has_split):  # INPUTSINGLE needs a shared return
        raise SystemExit("INPUTSINGLE pads need a shared return: add OUTPUTSINGLE or "
                         "OUTPUTSPLIT pads for the continuity coupon.")
    inputs = [p for p in pads if p["io"] == "input" and p["group"] != SINGLE]
    outputs = [p for p in pads if p["io"] == "output" and p["group"] != SINGLE]
    if not (inputs or single_in) or not (outputs or single_out):
        raise SystemExit("Need at least one input and one output.")
    groups = assign_groups(inputs, outputs)
    split_outs = [p for p in outputs if p["group"] == SPLIT_OUT]
    if any(p["group"] == GLOBAL_OUT for p in outputs) and groups_per_chip > 1:
        print("  ! A common OUTPUT pad is shared by every group; using 1 group per "
              "chip so two groups can't tie their shared output together.")
        groups_per_chip = 1
    if split_outs and groups_per_chip < 2:
        print("  ! OUTPUTSPLIT needs 2 layers to halve across; with 1 layer per chip "
              "all OUTPUTSPLIT pads go on the single layer.")
    coupons = split_coupons(groups, MAX_BINARY_INPUTS)   # big groups -> sub-coupons
    if single_in:                                    # one shorted continuity coupon
        coupons.append({"num": SINGLE, "sub": 0, "single": True,
                        "inputs": single_in, "outputs": single_out})

    # The calibration coupon sits on the same pad ring and folds its combs outside it
    # exactly like a test chip, so it is sized together with them and every die matches.
    all_outs = [p for p in pads if p["io"] == "output"]
    calib = calib_coupon(all_outs, CALIB_COUNT)
    if calib is None:
        print("  ! Fewer than 2 output pads -- no calibration coupon built.")
    edge_depth, ring_w, ring_h, scale = size_and_place(
        pads, coupons + ([calib] if calib else []))

    print(f"{len(inputs)} inputs, {len(outputs)} outputs; die {DIE_W:.0f} x {DIE_H:.0f} um ")
    print("Coil protrusion per edge (um): " +
          ", ".join(f"{e} {edge_depth[e]:.0f}" for e in ("left", "right", "top", "bottom")))
    print(f"Aluminium {METAL_THICKNESS_UM} um thick, W={WIRE_WIDTH} um -> "
          f"sheet res {SHEET_RES:.3f} ohm/sq; {WIRE_WIDTH/SHEET_RES:.1f} um per ohm.")
    big = max((len(c["inputs"]) for c in coupons if not c.get("single")), default=0)
    if big and (input_target_r(big) > 1e6 or max(DIE_W, DIE_H) > 5e4):
        print(f"  ! BINARY ladder still needs a {input_target_r(big):,.0f} ohm "
              f"resistor (a {big}-input coupon) and a {DIE_W/1000:.0f} x "
              f"{DIE_H/1000:.0f} mm die -- lower MAX_BINARY_INPUTS.")
    matched = [g["num"] for g in groups]
    wired_in = sum(len(g["inputs"]) for g in groups)
    print(f"Matched group numbers {matched}: {len(groups)} groups, "
          f"{wired_in} inputs wired (unmatched numbers dropped).")
    n_par = sum(1 for c in coupons if not c.get("single"))
    if n_par > len(groups):
        print(f"  Split into {n_par} layers (<= {MAX_BINARY_INPUTS} inputs per layer).")
    for g in groups:
        nsub = sum(1 for c in coupons if c["num"] == g["num"])
        extra = f" -> {nsub} layers total" if nsub > 1 else ""
        print(f"  group {g['num']}: {len(g['inputs'])} inputs, "
              f"{len(g['outputs'])} outputs{extra}")
    if single_in:
        ret = f"{len(single_out)} OUTPUTSINGLE" + (" + OUTPUTSPLIT half" if split_outs else "")
        print(f"  SINGLE continuity coupon: {len(single_in)} INPUTSINGLE shorted to its "
              f"return [{ret}] on one layer, per-pin continuity.")
    if split_outs:
        top_n = len(split_outs[0::2])
        print(f"  {len(split_outs)} OUTPUTSPLIT pads on every chip, halved per chip: "
              f"{top_n} on the top layer, {len(split_outs) - top_n} on the bottom.")

    xs = [p["x"] for p in pads]
    ys = [p["y"] for p in pads]
    bounds = (min(xs), max(xs), min(ys), max(ys))
    for d in (OUTPUT_DIR, GDS_DIR, SCHEMATIC_DIR, CALIB_DIR):
        d.mkdir(parents=True, exist_ok=True)

    chips = pack_chips(coupons, groups_per_chip)
    print(f"{layers}-layer chips: up to {groups_per_chip} layers per chip "
          f"-> {len(chips)} chips.")
    # Name each GDS by its group(s); show the sub-index only for split groups.
    nsub = {}
    for c in coupons:
        nsub[c["num"]] = nsub.get(c["num"], 0) + 1
    tag = lambda c: f"{c['num']}.{c['sub']}" if nsub[c["num"]] > 1 else f"{c['num']}"
    all_rows = []
    mosaic_entries = []
    for ci, chip_coupons in enumerate(chips, 1):
        lib, geo, rows = build_chip(ci, chip_coupons, pads, bounds, split_outs)
        all_rows.extend(rows)
        tags = [tag(c) for c in chip_coupons]
        gds_out = safe_path(GDS_DIR / f"groups_{'_'.join(tags)}.gds")
        lib.write_gds(gds_out)
        safe_path(SCHEMATIC_DIR / f"schematic_{'_'.join(tags)}.svg").write_text(
            build_schematic_svg(ci, rows), encoding="utf-8")   # circuit diagram
        mosaic_entries.append((lib.cells[0], DIE_W, DIE_H))

        shorts = find_shorts(geo, len(chip_coupons))
        if shorts:
            print(f"  ! chip {ci} ({'_'.join(tags)}): {len(shorts)} cross-net overlaps")

    csv_out = safe_path(OUTPUT_DIR / f"{csv_path.stem}_parallel.csv")
    with open(csv_out, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["chip", "layer", "half", "input_pad", "input_signal",
                    "output_pads", "actual_R_ohm", "total_len_um",
                    "group_parallel_R_ohm"])
        w.writerows(all_rows)

    # Calibration coupon: same pad ring as the test chips, so the probe card lands on
    # it directly. split_outs is deliberately NOT passed -- an OUTPUTSPLIT tap joining
    # the plane would short out the very resistors we are trying to measure.
    if calib:
        calib_lib, _, calib_rows = build_chip(len(chips) + 1, [calib], pads, bounds)
        label_calib_values(calib_lib.cells[0], calib_rows, calib["inputs"], bounds)
        calib_gds = safe_path(CALIB_DIR / "calibration_resistors.gds")
        calib_lib.write_gds(calib_gds)
        with open(safe_path(CALIB_DIR / "calibration_resistors.csv"), "w",
                  newline="") as f:
            w = csv.writer(f)
            w.writerow(["target_R_ohm", "actual_R_ohm", "measure_from", "measure_to",
                        "trace_len_um"])
            commons = ";".join(p["name"] for p in calib["outputs"])
            for k, row in enumerate(calib_rows, 1):
                w.writerow([f"{input_target_r(k):.0f}", row[6], commons, row[3], row[7]])
        mosaic_entries.append((calib_lib.cells[0], DIE_W, DIE_H))
        print(f"  Calibration coupon: {len(calib['inputs'])} resistors on pads "
              f"{', '.join(p['name'] for p in calib['inputs'])}; "
              f"{len(calib['outputs'])} COMMON taps for redundancy.")

    # Decode margin, from the resistances actually built (not the ladder targets), so
    # routing error is included -- check it here, not at the bench.
    per_coupon = {}
    for row in all_rows:
        if row[8] != "continuity":
            per_coupon.setdefault((row[0], row[1], row[2]), []).append(float(row[6]))
    if per_coupon:
        worst_k, worst_m = min(((k, decode_margin(v)) for k, v in per_coupon.items()),
                               key=lambda kv: kv[1])
        print(f"Decode margin (worst coupon, chip {worst_k[0]} layer {worst_k[1]}): "
              f"{worst_m*100:.4f}% of full scale.")
        if worst_m < 0.001:
            print(f"  ! Margin below 0.1% -- a {1/worst_m:,.0f}:1 measurement is needed "
                  f"to decode. Lower MAX_BINARY_INPUTS ({MAX_BINARY_INPUTS}) so each "
                  f"coupon has fewer inputs.")

    # Combined mask: every chip plus the calibration coupon tiled into one field,
    # then that field stepped-and-repeated to fill a 6-inch wafer, with the wafer
    # edge drawn last as an overlay ring. The field's column count and the array
    # centring are solved to pack the most complete fields onto the usable wafer
    # (inside the edge-exclusion ring) -- a shot-map optimisation, not sqrt(n).
    cols, fw, fh, field_origins = best_field_layout(
        mosaic_entries, MOSAIC_GAP, WAFER_DIAM_UM, WAFER_STREET_UM,
        WAFER_EDGE_EXCLUSION_UM)
    mosaic_lib, fw, fh = build_mosaic(mosaic_entries, MOSAIC_GAP, cols)
    wafer_lib, n_fields = build_wafer(mosaic_lib, field_origins, WAFER_DIAM_UM,
                                      WAFER_LAYER, WAFER_DT)
    wafer_gds = safe_path(OUTPUT_DIR / "all_tiled_chips.gds")
    wafer_lib.write_gds(wafer_gds)
    print(f"Wrote {wafer_gds.name}: {len(mosaic_entries)} dies in a {cols}-col field "
          f"({fw/1000:.1f} x {fh/1000:.1f} mm) -> {n_fields} fields, "
          f"{n_fields * len(mosaic_entries)} dies on a {WAFER_DIAM_UM/1000:.0f} mm "
          f"wafer ({WAFER_EDGE_EXCLUSION_UM/1000:.0f} mm edge exclusion).")


if __name__ == "__main__":
    main()
