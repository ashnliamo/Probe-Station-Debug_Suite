import csv
import math
import pathlib

# Every length is in micrometres EXCEPT the SCH_* block at the bottom, which is
# in mils because Altium's schematic editor works in mils on a 100 mil grid.

# Folders, resolved relative to this file so the project can be moved freely.
HERE = pathlib.Path(__file__).parent
INPUT_DIR = HERE.parent / "auto_probe_pcb_inputs"      # holds exactly one .csv
OUTPUT_DIR = HERE.parent / "auto_probe_pcb_outputs"    # generated files go here

# CSV column headers the reader looks for (matched case-insensitively).
COL_PAD, COL_SIGNAL, COL_X, COL_Y = ("pad", "signal", "x (um)", "y (um)")
COL_NETCLASS = "net class"          # groups lands into schematic components
UNCLASSIFIED = "UNCLASSIFIED"       # bucket for pads with a blank net class

# --- die ------------------------------------------------------------------
DIE_X = 8170.73                     # die width; hardcoded, NOT read from the CSV
DIE_Y = 5155.584                    # die height; hardcoded, NOT read from the CSV

# --- probe lands ----------------------------------------------------------
PROBE_PAD_SIDE = 1000               # land outer DIAMETER (pads are round)
PROBE_PAD_CLEARANCE = 200           # bare copper gap between adjacent lands
PROBE_PAD_PITCH = PROBE_PAD_SIDE + PROBE_PAD_CLEARANCE   # derived: land spacing
PROBE_WIRE_WIDTH = 150              # drawn width of the reference needle tracks

# --- two staggered rows of lands per edge ---------------------------------
LAND_ROWS = 2                       # lands per edge stack in this many rows
NEEDLE_CLEARANCE = 127              # needle-to-land margin (0.005 in)
ROW_GAP = PROBE_PAD_PITCH           # derived: radial spacing between the rows
ROW_PITCH = max(PROBE_PAD_PITCH,    # derived: spacing within one row; the wider
                PROBE_PAD_SIDE + PROBE_WIRE_WIDTH + 2 * NEEDLE_CLEARANCE)
STAGGER_STEP = ROW_PITCH / float(LAND_ROWS)   # derived: step along the edge

# --- board ----------------------------------------------------------------
BOARD_CENTER = (4000 * 25.4, 3000 * 25.4)     # card center on the Altium sheet
VIA_DRILL = 508                     # plated hole, 0.020 in: Accuprobe's floor
BOARD_WIDTH = 114500                # card width  (114.5 mm)
BOARD_HEIGHT = 188500               # card height (188.5 mm)
APERTURE_CLEARANCE = 3175           # die corner to the edge of the round cutout

# --- keep-out (also sets the land ring's aspect ratio) --------------------
KEEP_OUT_WIDTH = 44450              # ring-assembly keep-out rectangle, width
KEEP_OUT_HEIGHT = 38100             # ring-assembly keep-out rectangle, height
LAND_KEEPOUT_GAP = 1500             # smallest allowed land-to-keep-out gap

# --- silkscreen and reference markers -------------------------------------
LABEL_SIZE = 400                    # silk text height
LABEL_THICK = 60                    # silk stroke width
LABEL_GAP = 508                     # land edge to the start of its label
MARKER_LINE = 150                   # keep-out marker line width

# Non-fabricated drawing layers used for reference geometry.
ALTIUM_MARKER_LAYER  = "eMechanical15"   # keep-out rectangle
ALTIUM_REF_LAYER     = "eMechanical13"   # probe needles

# Shared by the PcbLib footprint and the SchLib component; the two names must
# match or Altium cannot resolve the symbol's footprint model.
PROBECARD_NAME = "ProbeCard"

# --- schematic symbol: MILS, not micrometres ------------------------------
SCH_PIN_PITCH = 100                 # pin spacing; 100 mil is Altium's grid
SCH_PIN_LENGTH = 300                # pin stub length; wires connect at its tip
SCH_BLOCK_WIDTH = 1200              # component body width


def read_pads(csv_path):
    """Parse the pinout CSV into one dict per pad."""
    with open(csv_path, newline="") as f:
        rows = list(csv.reader(f))
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
    inet = header.index(COL_NETCLASS) if COL_NETCLASS in header else None
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
        netclass = row[inet].strip() if inet is not None and inet < len(row) else ""
        pads.append({"name": name, "signal": signal, "x": x, "y": y,
                     "netclass": netclass or UNCLASSIFIED})
    return pads


def die_center(pads):
    """center of the pad bounding box, used as the die center."""
    xs = [p["x"] for p in pads]
    ys = [p["y"] for p in pads]
    return (min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0


def group_by_edge(pads, cx, cy):
    """Sort pads into T/B/L/R buckets by the die edge each sits nearest."""
    edges = {"T": [], "B": [], "L": [], "R": []}
    for p in pads:
        nx = (p["x"] - cx) / (DIE_X / 2.0)
        ny = (p["y"] - cy) / (DIE_Y / 2.0)
        if abs(nx) >= abs(ny):
            edge = "R" if nx > 0 else "L"
        else:
            edge = "T" if ny > 0 else "B"
        edges[edge].append(p)
    return edges


def smooth_die_pad_spacing(targets, pitch):
    """Spread lands along an edge while holding the minimum pitch."""
    n = len(targets)
    if n == 0:
        return []
    a = [targets[i] - i * pitch for i in range(n)]
    blocks = []
    for x in a:
        blocks.append([x, 1])
        while len(blocks) >= 2 and blocks[-2][0] > blocks[-1][0]:
            m2, c2 = blocks.pop()
            m1, c1 = blocks.pop()
            blocks.append([(m1 * c1 + m2 * c2) / (c1 + c2), c1 + c2])
    z = []
    for m, c in blocks:
        z.extend([m] * c)
    return [z[i] + i * pitch for i in range(n)]


def required_scale(edges): 
    """Smallest keep-out multiple that fits every edge's row of lands."""
    s = 1.0 + 2.0 * LAND_KEEPOUT_GAP / min(KEEP_OUT_WIDTH, KEEP_OUT_HEIGHT)
    for edge, pads in edges.items():
        if not pads:
            continue
        need = (len(pads) - 1) * STAGGER_STEP + PROBE_PAD_SIDE
        ko_dim = KEEP_OUT_WIDTH if edge in ("T", "B") else KEEP_OUT_HEIGHT
        s = max(s, need / ko_dim)
    return s


def place_probes(edges, cx, cy, scale):
    """Place each land on the keep-out ring, alternating rows."""
    half_ax = scale * KEEP_OUT_WIDTH / 2.0
    half_ay = scale * KEEP_OUT_HEIGHT / 2.0
    probes = []
    for edge, pads in edges.items():
        if not pads:
            continue
        if edge in ("T", "B"):
            pads = sorted(pads, key=lambda p: p["x"])
            us = smooth_die_pad_spacing([p["x"] for p in pads], STAGGER_STEP)
        else:
            pads = sorted(pads, key=lambda p: p["y"])
            us = smooth_die_pad_spacing([p["y"] for p in pads], STAGGER_STEP)
        for i, (p, u) in enumerate(zip(pads, us)):
            row = i % LAND_ROWS
            if edge == "T":
                x, y = u, cy + half_ay + row * ROW_GAP
            elif edge == "B":
                x, y = u, cy - half_ay - row * ROW_GAP
            elif edge == "R":
                x, y = cx + half_ax + row * ROW_GAP, u
            else:
                x, y = cx - half_ax - row * ROW_GAP, u
            probes.append({"x": x, "y": y, "edge": edge, "row": row,
                           "id": f"{edge}{i + 1}", "die": p})
    return probes


def compute_layout(pads):
    """center the die on the board, group pads by edge, place the lands."""
    cx0, cy0 = die_center(pads)
    dx, dy = BOARD_CENTER[0] - cx0, BOARD_CENTER[1] - cy0
    for p in pads:
        p["x"] += dx
        p["y"] += dy
    cx, cy = die_center(pads)
    edges = group_by_edge(pads, cx, cy)
    scale = required_scale(edges)
    probes = place_probes(edges, cx, cy, scale)
    return {"cx": cx, "cy": cy, "edges": edges, "scale": scale, "probes": probes}


def aperture_radius():
    """Radius of the round board cutout, clearing the die's corners."""
    return math.hypot(DIE_X / 2.0, DIE_Y / 2.0) + APERTURE_CLEARANCE


def marker_dims():
    """Keep-out rectangle size."""
    return KEEP_OUT_WIDTH, KEEP_OUT_HEIGHT


def check_fit(L):
    """Tightest clearance between any land's copper and the card outline."""
    cx, cy = L["cx"], L["cy"]
    half_w = BOARD_WIDTH / 2.0
    half_h = BOARD_HEIGHT / 2.0
    pad_half = PROBE_PAD_SIDE / 2.0
    worst = min(min(half_w - abs(pr["x"] - cx), half_h - abs(pr["y"] - cy))
                for pr in L["probes"]) - pad_half
    frame_lr = half_w - (L["scale"] * KEEP_OUT_WIDTH / 2.0)


def _altium_label(pr):
    """Position and rotation of one land's silkscreen label."""
    off = ((LAND_ROWS - 1 - pr["row"]) * ROW_GAP
           + PROBE_PAD_SIDE / 2.0 + LABEL_GAP) / 1000.0
    h = LABEL_SIZE / 1000.0
    x, y = pr["x"] / 1000.0, pr["y"] / 1000.0
    sig = pr["die"]["signal"] or pr["id"]
    ln = max(1, len(sig)) * h * 1.1
    edge = pr["edge"]
    if edge == "T":
        return (x - h / 2, y + off, 90.0, sig)
    if edge == "B":
        return (x - h / 2, y - off - ln, 90.0, sig)
    if edge == "R":
        return (x + off, y - h / 2, 0.0, sig)
    return (x - off - ln, y - h / 2, 0.0, sig)


def mm(v):
    """Format a millimetre value for DelphiScript."""
    return f"{v:.4f}"


def _pas_str(s):
    """Quote and escape a string literal for DelphiScript."""
    return "'" + s.replace("'", "''") + "'"


def _pad_proc_pas(proc_name, add_call):
    """Pad-creation procedure shared by the board and footprint scripts."""
    pad = mm(PROBE_PAD_SIDE / 1000.0)
    hole = mm(VIA_DRILL / 1000.0)
    return f"""Procedure {proc_name}(XMM, YMM : Double; Desig : String);
Var Pad;
Begin
    Pad := PCBServer.PCBObjectFactory(ePadObject, eNoDimension, eCreate_Default);
    Pad.X := MMsToCoord(XMM);
    Pad.Y := MMsToCoord(YMM);
    Pad.Layer := eMultiLayer;
    Pad.TopShape := eRounded;
    Pad.MidShape := eRounded;
    Pad.BotShape := eRounded;
    Pad.TopXSize := MMsToCoord({pad}); Pad.TopYSize := MMsToCoord({pad});
    Pad.MidXSize := MMsToCoord({pad}); Pad.MidYSize := MMsToCoord({pad});
    Pad.BotXSize := MMsToCoord({pad}); Pad.BotYSize := MMsToCoord({pad});
    Pad.HoleSize := MMsToCoord({hole});
    Pad.Plated := True;
    Pad.Name := Desig;
    {add_call};
End;"""


def _part_name(netclass):
    """Component and footprint name for a net class, e.g. ProbeCard_Power."""
    safe = "".join(c if c.isalnum() else "_" for c in netclass)
    while "__" in safe:
        safe = safe.replace("__", "_")
    safe = safe.strip("_")
    return f"{PROBECARD_NAME}_{safe}" if safe else PROBECARD_NAME


def write_altium_script(path, L):
    """Write the PcbDoc script: board, aperture, lands, labels, needles."""
    cx, cy = L["cx"], L["cy"]
    HW = BOARD_WIDTH / 2.0
    HH = BOARD_HEIGHT / 2.0
    APR = aperture_radius()
    MW, MH = marker_dims()
    TEXTH = LABEL_SIZE / 1000.0
    TEXTW = LABEL_THICK / 1000.0
    PROBEW = PROBE_WIRE_WIDTH / 1000.0

    o = []
    w = o.append
    w(f"""

Var
    Board : IPCB_Board;

Procedure RegisterObj(Obj);
Begin
    Board.AddPCBObject(Obj);
    PCBServer.SendMessageToRobots(Board.I_ObjectAddress, c_Broadcast,
        PCBM_BoardRegisteration, Obj.I_ObjectAddress);
End;

{_pad_proc_pas("AddLand", "RegisterObj(Pad)")}

Procedure AddText(XMM, YMM, Rot : Double; S : String);
Var T;
Begin
    T := PCBServer.PCBObjectFactory(eTextObject, eNoDimension, eCreate_Default);
    T.XLocation := MMsToCoord(XMM);
    T.YLocation := MMsToCoord(YMM);
    T.Layer := eTopOverlay;
    T.Size := MMsToCoord({mm(TEXTH)});
    T.Width := MMsToCoord({mm(TEXTW)});
    T.Rotation := Rot;
    T.Text := S;
    RegisterObj(T);
End;

Procedure AddRefTrack(X1, Y1, X2, Y2, Wid : Double);
Var Tr;
Begin
    Tr := PCBServer.PCBObjectFactory(eTrackObject, eNoDimension, eCreate_Default);
    Tr.X1 := MMsToCoord(X1); Tr.Y1 := MMsToCoord(Y1);
    Tr.X2 := MMsToCoord(X2); Tr.Y2 := MMsToCoord(Y2);
    Tr.Width := MMsToCoord(Wid);
    Tr.Layer := {ALTIUM_REF_LAYER};
    RegisterObj(Tr);
End;

Procedure AddProbe(X1, Y1, X2, Y2 : Double);
Begin
    AddRefTrack(X1, Y1, X2, Y2, {mm(PROBEW)});
End;

// --- everything belonging to one probe: the solder land, its silk label,
// --- and the probe needle running from the land to the die-pad coordinate.
Procedure EmitLand(LX, LY : Double; Desig : String;
                   TX, TY, Rot : Double; Sig : String; DX, DY : Double);
Begin
    AddLand(LX, LY, Desig);
    AddText(TX, TY, Rot, Sig);
    AddProbe(LX, LY, DX, DY);
End;

// --- board shape: rectangular outline.
// --- Segments[i] returns the record BY VALUE, so "Segments[i].vx := ..."
// --- edits a throwaway copy and does nothing. Build a local TPolySegment
// --- and assign it back into Segments[i].
Procedure SetOutlinePoint(Idx : Integer; XMM, YMM : Double);
Var Seg : TPolySegment;
Begin
    Seg := Board.BoardOutline.Segments[Idx];
    Seg.Kind := ePolySegmentLine;
    Seg.vx := MMsToCoord(XMM);
    Seg.vy := MMsToCoord(YMM);
    Board.BoardOutline.Segments[Idx] := Seg;
End;

Procedure SetRectBoard(X1, Y1, X2, Y2 : Double);
Begin
    Board.BoardOutline.Invalidate;
    Board.BoardOutline.PointCount := 4;
    SetOutlinePoint(0, X1, Y1);
    SetOutlinePoint(1, X2, Y1);
    SetOutlinePoint(2, X2, Y2);
    SetOutlinePoint(3, X1, Y2);
    Board.BoardOutline.Validate;
    Board.ViewManager_FullUpdate;
End;

// --- circular aperture: board-cutout region approximated by a 72-gon.
// --- The region-contour API is the part most likely to need adjustment
// --- for your Altium version.
Procedure AddCircleCutout(CXmm, CYmm, Rmm : Double);
Var Rgn, C, i, ang;
Begin
    Rgn := PCBServer.PCBObjectFactory(eRegionObject, eNoDimension, eCreate_Default);
    Rgn.SetState_Kind(eRegionKind_BoardCutout);
    C := PCBServer.PCBContourFactory;
    For i := 0 To 71 Do
    Begin
        ang := i * 6.28318530717959 / 72.0;
        C.AddPoint(MMsToCoord(CXmm + Rmm * Cos(ang)), MMsToCoord(CYmm + Rmm * Sin(ang)));
    End;
    Rgn.SetOutlineContour(C);
    Rgn.Layer := eTopLayer;
    RegisterObj(Rgn);
End;

// --- die-aspect clearance marker on a non-fabricated mechanical layer.
Procedure AddMarkerTrack(X1, Y1, X2, Y2, Wid : Double);
Var Tr;
Begin
    Tr := PCBServer.PCBObjectFactory(eTrackObject, eNoDimension, eCreate_Default);
    Tr.X1 := MMsToCoord(X1); Tr.Y1 := MMsToCoord(Y1);
    Tr.X2 := MMsToCoord(X2); Tr.Y2 := MMsToCoord(Y2);
    Tr.Width := MMsToCoord(Wid);
    Tr.Layer := {ALTIUM_MARKER_LAYER};
    RegisterObj(Tr);
End;

Procedure AddMarkerRect(X1, Y1, X2, Y2, Wid : Double);
Begin
    AddMarkerTrack(X1, Y1, X2, Y1, Wid);
    AddMarkerTrack(X2, Y1, X2, Y2, Wid);
    AddMarkerTrack(X2, Y2, X1, Y2, Wid);
    AddMarkerTrack(X1, Y2, X1, Y1, Wid);
End;
""")

    w("Procedure BuildAll;\nBegin")
    for pr in L["probes"]:
        d = pr["die"]
        lx, ly, rot, sig = _altium_label(pr)
        w(f"    EmitLand({mm(pr['x']/1000)}, {mm(pr['y']/1000)}, {_pas_str(pr['id'])}, "
          f"{mm(lx)}, {mm(ly)}, {mm(rot)}, {_pas_str(sig)}, "
          f"{mm(d['x']/1000)}, {mm(d['y']/1000)});")
    w("End;\n")

    w(f"""Procedure GenerateProbeCard;
Begin
    If PCBServer = Nil Then
    Begin
        ShowMessage('PCBServer is nil -- the PCB editor is not loaded.');
        Exit;
    End;
    Board := PCBServer.GetCurrentPCBBoard;
    If Board = Nil Then
    Begin
        ShowMessage('No PCB is open. Create a blank PCB first: ' +
            'File > New > PCB, make that tab active, then run this script again.');
        Exit;
    End;
    PCBServer.PreProcess;
    SetRectBoard({mm((cx-HW)/1000)}, {mm((cy-HH)/1000)}, {mm((cx+HW)/1000)}, {mm((cy+HH)/1000)});
    AddCircleCutout({mm(cx/1000)}, {mm(cy/1000)}, {mm(APR/1000)});
    AddMarkerRect({mm((cx-MW/2)/1000)}, {mm((cy-MH/2)/1000)}, {mm((cx+MW/2)/1000)}, {mm((cy+MH/2)/1000)}, {mm(MARKER_LINE/1000)});
    BuildAll;
    PCBServer.PostProcess;
    Board.ViewManager_FullUpdate;
    Client.SendMessage('PCB:Zoom', 'Action=Redraw', 255, Client.CurrentView);
    ShowMessage('Probe card built on the active PCB. ' +
        'Use File > Save As to save it as a .PcbDoc.');
End;
""")

    with open(path, "w", newline="\n") as f:
        f.write("\n".join(o))


def write_wiring_map(path, L):
    """Write the CSV mapping every land to its die pad and coordinates."""
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["land_id", "edge", "row", "signal", "die_pad",
                    "land_x_mm", "land_y_mm", "die_x_um", "die_y_um"])
        for pr in L["probes"]:
            d = pr["die"]
            w.writerow([pr["id"], pr["edge"], pr["row"], d["signal"], d["name"],
                        f"{pr['x']/1000:.3f}", f"{pr['y']/1000:.3f}",
                        f"{d['x']:.3f}", f"{d['y']:.3f}"])


def _edge_order(pr):
    """Sort key ordering lands T, B, L, R then by land number."""
    return ("TBLR".index(pr["edge"]), int(pr["id"][1:]))


def group_by_netclass(probes):
    """Bucket lands by net class and merge differential pairs."""
    classes = {}
    for pr in probes:
        classes.setdefault(pr["die"]["netclass"], []).append(pr)
    for members in classes.values():
        members.sort(key=lambda pr: (pr["die"]["signal"], _edge_order(pr)))
    classes = merge_diff_pairs(classes)
    return dict(sorted(classes.items(), key=lambda kv: (-len(kv[1]), kv[0])))


def _interleave_pairs(p_members, n_members):
    """Order a merged block so each P land sits above its N partner."""
    pool = list(n_members)
    ordered = []
    for p in p_members:
        sig = p["die"]["signal"]
        best_i, best_d = None, None
        for i, q in enumerate(pool):
            s = q["die"]["signal"]
            if len(s) != len(sig):
                continue
            d = sum(a != b for a, b in zip(sig, s))
            if best_d is None or d < best_d:
                best_i, best_d = i, d
        ordered.append(p)
        if best_i is not None and best_d == 1:
            ordered.append(pool.pop(best_i))
    ordered.extend(pool)
    return ordered


def merge_diff_pairs(classes):
    """Combine net classes differing only by a trailing P or N."""
    out = {}
    used = set()
    for name in classes:
        if name in used:
            continue
        base = None
        if name.endswith(" P") and (name[:-2] + " N") in classes:
            base, p_name, n_name = name[:-2], name, name[:-2] + " N"
        elif name.endswith(" N") and (name[:-2] + " P") in classes:
            base, p_name, n_name = name[:-2], name[:-2] + " P", name
        if base is not None:
            out[base.strip()] = _interleave_pairs(classes[p_name], classes[n_name])
            used.add(p_name)
            used.add(n_name)
        else:
            out[name] = classes[name]
    return out


def write_pcb_library_script(path, L):
    """Write the PcbLib script: one footprint per net class."""
    classes = group_by_netclass(L["probes"])
    cx, cy = L["cx"], L["cy"]

    o = []
    w = o.append
    w(f"""// Auto-generated by auto_probe_pcb.py -- Altium PCB-library DelphiScript.
// Open a PCB library FIRST: File > New > Library > PCB Library, make that
// .PcbLib the ACTIVE document, then File > Run Script... and run
// GenerateFootprint. Builds one footprint per net class; each holds that
// class's probe-card lands (pad designators = land ids) at their real board
// positions relative to the die center. Place them all at the same origin to
// reconstruct the complete land pattern.

Var
    Lib : IPCB_Library;
    FP  : IPCB_LibComponent;

{_pad_proc_pas("AddFPPad", "FP.AddPCBObject(Pad)")}

// --- start a named footprint. Drops any footprint of the same name from a
// --- previous run first, so re-running replaces instead of colliding.
// --- CreateNewComponent is the library-level call used by working scripts;
// --- it is equivalent to PCBServer.CreatePCBLibComp.
Procedure StartFP(Nm : String);
Var Old;
Begin
    Old := Lib.GetComponentByName(Nm);
    If Old <> Nil Then
    Begin
        Lib.RemoveComponent(Old);
        Lib.DeRegisterComponent(Old);
    End;
    FP := Lib.CreateNewComponent;
    FP.Name := Nm;
End;

Procedure EndFP;
Begin
    Lib.RegisterComponent(FP);
    // must SetState_CurrentComponent (not a plain property assign) or the
    // footprint origin / bounding box come out wrong.
    Lib.SetState_CurrentComponent(FP);
End;
""")

    build_calls = []
    for ci, (netclass, members) in enumerate(classes.items()):
        proc = f"BuildFP{ci}"
        build_calls.append(proc)
        w(f"\n// {netclass} -- {len(members)} lands")
        w(f"Procedure {proc};\nBegin")
        w(f"    StartFP({_pas_str(_part_name(netclass))});")
        for pr in members:
            w(f"    AddFPPad({mm((pr['x']-cx)/1000)}, {mm((pr['y']-cy)/1000)}, "
              f"{_pas_str(pr['id'])});")
        w("    EndFP;\nEnd;")

    w("""
Procedure GenerateFootprint;
Begin
    If PCBServer = Nil Then
    Begin
        ShowMessage('PCBServer is nil -- the PCB editor is not loaded.');
        Exit;
    End;
    Lib := PCBServer.GetCurrentPCBLibrary;
    If Lib = Nil Then
    Begin
        ShowMessage('No PCB library open. Open the .PcbLib, click its tab so it ' +
            'is the ACTIVE document, then run this again.');
        Exit;
    End;

    PCBServer.PreProcess;""")
    for proc in build_calls:
        w(f"    {proc};")
    w(f"""    PCBServer.PostProcess;

    Lib.Board.ViewManager_FullUpdate;
    ShowMessage('Built {len(build_calls)} probe-card footprints ' +
        '({len(L["probes"])} pads total). Save the .PcbLib.');
End;""")

    with open(path, "w", newline="\n") as f:
        f.write("\n".join(o))


def write_sch_library_script(path, L):
    """Write the SchLib script: one component per net class."""
    classes = group_by_netclass(L["probes"])

    o = []
    w = o.append
    w("""// Auto-generated by auto_probe_pcb.py -- Altium Sch-library DelphiScript.
// Open a schematic library FIRST: File > New > Library > Schematic Library,
// make that .SchLib the ACTIVE document, then File > Run Script... and run
// GenerateSymbol. Builds one single-part component per net class, each
// carrying a footprint-model reference to the PcbLib footprint of the same
// name. Place all of them to cover the whole probe card.

Var
    Lib  : ISch_Lib;
    Comp : ISch_Component;

Procedure AddRect(Xmil, Ymil, Wmil, Hmil : Integer);
Var R;
Begin
    R := SchServer.SchObjectFactory(eRectangle, eCreate_GlobalCopy);
    R.OwnerPartId := 1;
    R.OwnerPartDisplayMode := 0;
    R.Location := Point(MilsToCoord(Xmil), MilsToCoord(Ymil - Hmil));
    R.Corner := Point(MilsToCoord(Xmil + Wmil), MilsToCoord(Ymil));
    R.LineWidth := eSmall;
    R.IsSolid := True;
    R.AreaColor := $00E7FFFF;
    R.Color := $00000080;
    Comp.AddSchObject(R);
End;

Procedure AddPin(Xmil, Ymil : Integer; Desig, Nm : String);
Var Pin;
Begin
    Pin := SchServer.SchObjectFactory(ePin, eCreate_GlobalCopy);
    Pin.OwnerPartId := 1;
    Pin.OwnerPartDisplayMode := 0;
    Pin.Orientation := eRotate0;
    Pin.Location := Point(MilsToCoord(Xmil - """ + str(SCH_PIN_LENGTH) + """), MilsToCoord(Ymil));
    Pin.PinLength := MilsToCoord(""" + str(SCH_PIN_LENGTH) + """);
    Pin.Designator := Desig;
    Pin.Name := Nm;
    Pin.ShowName := True;
    Pin.ShowDesignator := True;
    Pin.Electrical := eElectricPassive;
    Comp.AddSchObject(Pin);
End;

Procedure StartComp(LibRef, Descr : String);
Begin
    Comp := SchServer.SchObjectFactory(eSchComponent, eCreate_GlobalCopy);
    Comp.LibReference := LibRef;
    Comp.ComponentDescription := Descr;
    Comp.Designator.Text := 'PC?';
    Comp.PartCount := 1;
    Comp.CurrentPartID := 1;
    Comp.DisplayMode := 0;
End;

// --- attach the matching PcbLib footprint and file the component away.
Procedure EndComp(FpName : String);
Var Impl;
Begin
    Impl := Comp.AddSchImplementation;
    Impl.ModelName := FpName;
    Impl.ModelType := 'PCBLIB';
    Impl.IsCurrent := True;
    Lib.AddSchComponent(Comp);
    Lib.CurrentSchComponent := Comp;
    Comp.GraphicallyInvalidate;
End;
""")

    build_calls = []
    for ci, (netclass, members) in enumerate(classes.items()):
        height = (len(members) + 1) * SCH_PIN_PITCH
        nm = _part_name(netclass)
        proc = f"BuildComp{ci}"
        build_calls.append(proc)
        w(f"\n// {netclass} -- {len(members)} lands")
        w(f"Procedure {proc};\nBegin")
        w(f"    StartComp({_pas_str(nm)}, "
          f"{_pas_str('Probe card lands -- ' + netclass)});")
        w(f"    AddRect(0, 0, {SCH_BLOCK_WIDTH}, {height});")
        for k, pr in enumerate(members):
            py = -(k + 1) * SCH_PIN_PITCH
            sig = pr["die"]["signal"] or pr["id"]
            w(f"    AddPin(0, {py}, {_pas_str(pr['id'])}, {_pas_str(sig)});")
        w(f"    EndComp({_pas_str(nm)});\nEnd;")

    w("""
Procedure GenerateSymbol;
Begin
    If SchServer = Nil Then
    Begin
        ShowMessage('SchServer is nil -- the schematic editor is not loaded.');
        Exit;
    End;
    Lib := SchServer.GetCurrentSchDocument;
    If Lib = Nil Then
    Begin
        ShowMessage('No schematic library open. Create one first: ' +
            'File > New > Library > Schematic Library, make it active, then run again.');
        Exit;
    End;""")
    for proc in build_calls:
        w(f"    {proc};")
    w(f"""    ShowMessage('Built {len(build_calls)} probe-card components ' +
        '({len(L["probes"])} pins total). Save the .SchLib.');
End;""")

    with open(path, "w", newline="\n") as f:
        f.write("\n".join(o))
    return classes


def find_input_csv():
    """Return the single CSV in the inputs folder, or stop with an error."""
    csvs = sorted(INPUT_DIR.glob("*.csv"))
    if not csvs:
        raise SystemExit(f"No .csv found in {INPUT_DIR} -- put your pinout there.")
    if len(csvs) > 1:
        names = ", ".join(p.name for p in csvs)
        raise SystemExit(f"Multiple .csv files in {INPUT_DIR} ({names}); keep just one.")
    return csvs[0]


def main():
    """Read the pinout, compute the layout, write the four output files."""
    csv_path = find_input_csv()
    pads = read_pads(csv_path)
    L = compute_layout(pads)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    counts = {e: len(v) for e, v in L["edges"].items()}
    print(f"Die {DIE_X} x {DIE_Y} um, center ({L['cx']:.1f}, {L['cy']:.1f})")
    print(f"Per-edge pad counts: {counts}")
    print(f"Land ring = keep-out x{L['scale']:.2f} -> "
          f"{L['scale']*KEEP_OUT_WIDTH/1000:.1f} x "
          f"{L['scale']*KEEP_OUT_HEIGHT/1000:.1f} mm")
    MW, MH = marker_dims()
    print(f"Keep-out {MW/1000:.1f} x {MH/1000:.1f} mm; inner-row land clears it by "
          f"X {L['scale']*KEEP_OUT_WIDTH/2 - MW/2:.0f} um, "
          f"Y {L['scale']*KEEP_OUT_HEIGHT/2 - MH/2:.0f} um")
    apr = aperture_radius()
    if min(MW, MH) / 2.0 < apr:
        print(f"  WARNING: keep-out half-extent {min(MW, MH)/2:.0f} um is inside the "
              f"{apr:.0f} um aperture radius -- the marker crosses the cutout.")
    check_fit(L)

    map_path = OUTPUT_DIR / f"{csv_path.stem}_wiring_map.csv"
    write_wiring_map(map_path, L)
    print(f"Wrote wiring map:   {map_path}")

    pas_path = OUTPUT_DIR / f"{csv_path.stem}_probe_card.pas"
    write_altium_script(pas_path, L)
    print(f"Wrote Altium script:{pas_path}")

    fp_path = OUTPUT_DIR / f"{csv_path.stem}_probecard_footprint.pas"
    write_pcb_library_script(fp_path, L)
    print(f"Wrote PCB-lib footprint: {fp_path}")

    sym_path = OUTPUT_DIR / f"{csv_path.stem}_probecard_symbol.pas"
    classes = write_sch_library_script(sym_path, L)
    print(f"Wrote Sch-lib symbol:    {sym_path}")
    cc = {k: len(v) for k, v in classes.items()}
    print(f"Net-class parts ({len(cc)}): {cc}")

    print(f"{len(pads)} die pads / {len(L['probes'])} probe lands "
          f"({VIA_DRILL} um dia drill, {PROBE_PAD_SIDE} um land).")


if __name__ == "__main__":
    main()
