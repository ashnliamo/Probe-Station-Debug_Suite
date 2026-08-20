import csv
import sys
import pathlib

HERE = pathlib.Path(__file__).parent
DECODE_DIR = HERE.parent / "decode_inputs"

CHIP, LAYER, PAD, R_COL = "chip", "layer", "input_pad", "actual_R_ohm"
HALF = "half"      # split coupons: each READING is its own decode unit; 0 when not
                   # split. Named 'half' for CSV compatibility, but a coupon may be cut
                   # into more than two readings when it exceeds MAX_BINARY_INPUTS.
MATCH_TOL = 0.05   # accept a subset if its predicted R is within this of measured


# ----------------------------------------------------------------------
# Input
# ----------------------------------------------------------------------
def _header(path):
    try:
        with open(path, newline="") as f:
            return [h.strip().lower() for h in next(csv.reader(f))]
    except (StopIteration, OSError):
        return []


def find_decode_csv():
    """The pinout key CSV in decode_inputs/ -- the one with an 'input_pad' column, so
    it can sit alongside a calibration CSV. Errors if none or several."""
    hits = [p for p in sorted(DECODE_DIR.glob("*.csv")) if PAD in _header(p)]
    if not hits:
        raise SystemExit(f"No pinout .csv (with an input_pad column) in {DECODE_DIR} -- "
                         f"put the generator's *_parallel.csv there.")
    if len(hits) > 1:
        names = ", ".join(p.name for p in hits)
        raise SystemExit(f"Multiple pinout .csv files in {DECODE_DIR} ({names}); keep "
                         f"one, or pass the path: py find_missing_probes.py <file.csv>")
    return hits[0]


def load(path):
    if not path.exists():
        raise SystemExit(f"Input CSV not found: {path}\n"
                         f"Put pinout_grouped_parallel.csv in {path.parent}\\")
    with open(path, newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        raise SystemExit(f"No rows in {path}")
    for col in (CHIP, LAYER, PAD, R_COL):
        if col not in rows[0]:
            raise SystemExit(f"{path} has no '{col}' column.")
    return rows


def row_half(r):
    """This row's reading number, or 0 for an unsplit coupon or an older CSV with no
    column. Any integer is accepted: a coupon may have more than two readings."""
    try:
        return int((r.get(HALF) or "0").strip() or "0")
    except ValueError:
        return 0


def ask_choice(prompt, choices):
    choices = sorted(choices)
    while True:
        s = input(prompt).strip()
        try:
            v = int(s)
        except ValueError:
            print(f"  Enter a whole number from {choices}.")
            continue
        if v in choices:
            return v
        print(f"  Available: {choices}")


def parse_ohms(s):
    """'1.2k'/'470'/'3.3M'/'OPEN' -> ohms (inf for open), or None if unparseable."""
    s = s.strip().lower().replace(",", "").replace("ohm", "").replace("Ω", "").strip()
    if s in ("", "open", "ol", "inf", "overrange", "over"):
        return float("inf")
    mult = 1.0
    if s.endswith("meg"):
        mult, s = 1e6, s[:-3]
    elif s.endswith("k"):
        mult, s = 1e3, s[:-1]
    elif s.endswith("m"):
        mult, s = 1e6, s[:-1]
    try:
        return float(s) * mult
    except ValueError:
        return None


# ----------------------------------------------------------------------
# Calibration: per-rung scale read from the generator's calibration CSV
# ----------------------------------------------------------------------
CAL_TARGET, CAL_CALC, CAL_MEAS = "target_r_ohm", "actual_r_ohm", "measured_r_ohm"


def find_calibration_csv():
    """A calibration CSV in decode_inputs/ -- the one with a target_R_ohm column."""
    for p in sorted(DECODE_DIR.glob("*.csv")):
        if CAL_TARGET in _header(p):
            return p
    return None


def load_calibration(path):
    """Per-rung calibration read from the generator's calibration CSV. It needs a
    measured_R_ohm column filled in from the calibration chip; each rung's scale is
    measured / calculated (actual_R_ohm), i.e. the real sheet-resistance ratio. Every
    test resistor is later multiplied by the scale of its nearest rung. Returns
    [(nominal, scale)] ascending, or None if no measurements are present."""
    with open(path, newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return None
    col = {k.strip().lower(): k for k in rows[0]}
    if CAL_MEAS not in col or CAL_TARGET not in col:
        return None
    calc_col = col.get(CAL_CALC, col[CAL_TARGET])
    cal = []
    for r in rows:
        meas = parse_ohms(r[col[CAL_MEAS]])
        calc = float(r[calc_col]) if r[calc_col].strip() else 0.0
        if meas not in (None, float("inf")) and calc > 0:
            cal.append((float(r[col[CAL_TARGET]]), meas / calc))
    return sorted(cal) or None


def apply_calibration(resistors, cal):
    """Scale each resistor by its nearest rung's calibration factor (measured/calc).
    Returns a new (pad, R) list sorted by corrected resistance."""
    noms = [n for n, _ in cal]
    scale = dict(cal)
    def corrected(R):
        return R * scale[min(noms, key=lambda n: abs(n - R))]
    return sorted(((pad, corrected(R)) for pad, R in resistors), key=lambda t: t[1])


# ----------------------------------------------------------------------
# Decode
# ----------------------------------------------------------------------
def parallel(resistances):
    g = sum(1.0 / r for r in resistances if r > 0)
    return 1.0 / g if g > 0 else float("inf")


def decode_margin(conductances):
    """Smallest gap between any two subset sums, relative to the all-on sum: the
    reading must beat this fraction for the missing set to be unique."""
    n = len(conductances)
    sums = sorted(sum(conductances[i] for i in range(n) if mask >> i & 1)
                  for mask in range(1 << n))
    total = sums[-1] or 1.0
    return min((sums[i + 1] - sums[i] for i in range(len(sums) - 1)),
               default=total) / total


def rank_subsets(resistors, r_meas):
    """Score every subset by how well its REMOVAL explains r_meas (the rest reads
    1/(G_all - Gsub)). Returns (pred_R, missing_idx) list, best fit first."""
    g = [1.0 / R for _, R in resistors]
    G_all = sum(g)
    G_meas = 1.0 / r_meas if r_meas not in (0, float("inf")) else (
        0.0 if r_meas == float("inf") else G_all)
    G_missing = G_all - G_meas                       # conductance that dropped out
    n = len(resistors)
    scored = []
    for mask in range(1 << n):
        gsub = sum(g[i] for i in range(n) if mask >> i & 1)
        g_left = G_all - gsub
        pred = 1.0 / g_left if g_left > 1e-12 else float("inf")
        idx = tuple(i for i in range(n) if mask >> i & 1)
        scored.append((abs(gsub - G_missing), pred, idx))
    scored.sort(key=lambda t: t[0])
    return [(pred, idx) for _, pred, idx in scored]


def report(resistors, r_meas):
    pads = [p for p, _ in resistors]
    r_all = parallel([R for _, R in resistors])
    margin = decode_margin([1.0 / R for _, R in resistors])
    print(f"\nAll-contacting parallel resistance: {r_all:,.2f} ohm "
          f"({len(resistors)} resistors).")
    margin_ohm = margin * r_all                      # resolution near the all-on reading
    print(f"Decode margin {margin*100:.2f}% (~{margin_ohm:,.2f} ohm): the reading must "
          f"be accurate to better than this (incl. contact resistance) for the missing "
          f"set to be unique.")

    if r_meas != float("inf") and r_meas < r_all * (1 - MATCH_TOL):
        print(f"  ! Measured {r_meas:,.2f} ohm is BELOW the all-contacting value -- "
              "with every probe landed the reading can't go lower. Check the probe "
              "setup (a short, or the wrong chip/layer).")

    ranked = rank_subsets(resistors, r_meas)
    pred, idx = ranked[0]
    missing = [pads[i] for i in idx]
    present = [pads[i] for i in range(len(pads)) if i not in idx]
    err = abs(pred - r_meas) / r_meas if r_meas not in (0, float("inf")) else (
        0.0 if pred == float("inf") else 1.0)

    print(f"\nMeasured: {('OPEN' if r_meas==float('inf') else f'{r_meas:,.2f} ohm')}"
          f"  ->  best fit predicts {('OPEN' if pred==float('inf') else f'{pred:,.2f} ohm')}"
          f"  ({err*100:.1f}% off)")
    if not missing:
        print("  => ALL probes in contact (no resistor missing).")
    else:
        print(f"  => {len(missing)} probes NOT in contact (missing): "
              f"{', '.join(missing)}")
    print(f"     in contact: {', '.join(present) if present else '(none)'}")

    if err > MATCH_TOL:
        print(f"  ! Best fit is {err*100:.1f}% off (> {MATCH_TOL*100:.0f}%); treat the "
              "result as approximate -- the reading may sit between resistor sums.")
    # flag ambiguity: a different missing set predicting nearly the same resistance
    for pred2, idx2 in ranked[1:]:
        if set(idx2) == set(idx):
            continue
        e2 = abs(pred2 - r_meas) / r_meas if r_meas not in (0, float("inf")) else (
            0.0 if pred2 == float("inf") else 1.0)
        if e2 <= MATCH_TOL:
            alt = [pads[i] for i in idx2] or ["(none)"]
            print(f"  ! Ambiguous: missing {{{', '.join(alt)}}} fits about as well "
                  f"({e2*100:.1f}% off).")
        break


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------
def main():
    path = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else find_decode_csv()
    rows = load(path)
    print(f"Loaded {len(rows)} coils from {path}")
    combos = sorted({(int(r[CHIP]), int(r[LAYER]), row_half(r)) for r in rows})
    print(f"Chips available: {sorted({c for c, _, _ in combos})}")

    cal_path = find_calibration_csv()
    cal = load_calibration(cal_path) if cal_path else None
    if cal:
        print(f"Calibrating from {cal_path.name}: per-rung scale "
              + ", ".join(f"{n:,.0f} ohm x{s:.3f}" for n, s in cal))
    else:
        # A wrong sheet resistance scales every resistor by the same factor, and the
        # decode compares against ABSOLUTE values, so it cannot absorb that itself. A
        # reading tolerates only half its decode margin of uniform error before it
        # names a probe that is actually landed.
        why = (f"{cal_path.name} has no measured_R_ohm values yet"
               if cal_path else
               f"no calibration CSV in {DECODE_DIR.name}")
        print(f"\n  ! NOT CALIBRATED -- {why}.")
        print(f"  ! Results are only trustworthy if the fabricated film happens to "
              f"match the\n"
              f"    resistivity the generator assumed. Measure the calibration chip, "
              f"fill in the\n"
              f"    measured_R_ohm column, and copy that CSV into "
              f"{DECODE_DIR.name}/.")

    while True:                             # one decode per chip/layer; blank quits
        raw = input("\nChip # (blank to quit): ").strip()
        if raw == "":
            break
        try:
            chip = int(raw)
        except ValueError:
            print("  Enter a whole number."); continue
        chip_layers = sorted({l for c, l, _ in combos if c == chip})
        if not chip_layers:
            print(f"  No chip {chip}. Available: {sorted({c for c, _, _ in combos})}")
            continue
        layer = ask_choice(f"Layer # {chip_layers}: ", chip_layers)
        # A split coupon has several independent readings on this layer -- each is
        # measured and decoded on its own, between its OWN output pads and the rail.
        halves = sorted({h for c, l, h in combos if c == chip and l == layer})
        half = ask_choice(f"Reading # {halves}: ", halves) if halves != [0] else 0

        sel = [r for r in rows if int(r[CHIP]) == chip and int(r[LAYER]) == layer
               and row_half(r) == half]
        resistors = sorted(((r[PAD], float(r[R_COL])) for r in sel),
                           key=lambda t: t[1])
        if cal:
            resistors = apply_calibration(resistors, cal)
            note = f" (calibrated from {cal_path.name})"
        else:
            note = " (UNCALIBRATED)"
        where = f"Chip {chip}, layer {layer}" + (f", reading {half}" if half else "")
        print(f"\n{where}: {len(resistors)} input resistors{note}")
        print(f"   measure between the input rail and: {sel[0]['output_pads']}")
        for pad, R in resistors:
            print(f"   {pad:>12}  {R:10,.2f} ohm")

        ohms = None
        while ohms is None:
            ohms = parse_ohms(input("\nMeasured resistance: "))
            if ohms is None:
                print("  Couldn't read that. Try e.g. 470, 1.2k, 3.3M, or OPEN.")
        report(resistors, ohms)


if __name__ == "__main__":
    main()
