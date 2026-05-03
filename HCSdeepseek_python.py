"""
HCS Design Engine (Pure Python, No UI)
======================================
Reference: ACI/PCI CODE-319-25, PCI Design Handbook 8th Ed.
Units: SI (mm, kN, MPa)

This file combines all calculation logic from the HCS Design App.
It can be used as a reference for Copilot in Excel to generate
spreadsheet formulas or VBA code.
"""

import math
import numpy as np

# =============================================================================
# 1. CONSTANTS & MATERIAL DATA
# =============================================================================
WIRE_PROPS = {
    5.0: {"area_mm2": 19.6, "fpu_MPa": 1618, "fpy_MPa": 1432, "Eps_MPa": 199050},
    7.0: {"area_mm2": 38.5, "fpu_MPa": 1515, "fpy_MPa": 1324, "Eps_MPa": 199990},
}

STRAND_PROPS = {
    "3/8 in  (d=8.4mm)":  {"d_mm": 8.4,  "area_mm2": 54.9,  "fpu_MPa": 1862, "fpy_MPa": 1675, "Eps_MPa": 196500},
    "7/16 in (d=9.7mm)":  {"d_mm": 9.7,  "area_mm2": 74.2,  "fpu_MPa": 1860, "fpy_MPa": 1674, "Eps_MPa": 196500},
    "1/2 in  (d=11.2mm)": {"d_mm": 11.2, "area_mm2": 98.7,  "fpu_MPa": 1862, "fpy_MPa": 1675, "Eps_MPa": 196500},
    "3/5 in  (d=13.4mm)": {"d_mm": 13.4, "area_mm2": 140.0, "fpu_MPa": 1860, "fpy_MPa": 1674, "Eps_MPa": 196500},
}

PRESET_TABLE = {
    "HCS 120mm — Circular core": {
        "h": 120, "b_bottom": 1197, "b_top": 1185,
        "tf_top": 40, "tf_bot": 30,
        "core_shape": "Circular", "d_core": 60, "n_core": 9,
        "gap_side": 64, "gap_between": 70,
        "h_straight": 40, "h_taper": 20,
    },
    "HCS 150mm — Teardrop core": {
        "h": 150, "b_bottom": 1197, "b_top": 1185,
        "tf_top": 40, "tf_bot": 40,
        "core_shape": "Teardrop", "d_core": 80, "n_core": 9,
        "gap_side": 66, "gap_between": 53,
        "h_straight": 40, "h_taper": 10,
    },
    "HCS 200mm — Teardrop core": {
        "h": 200, "b_bottom": 1199, "b_top": 1187,
        "tf_top": 52, "tf_bot": 50,
        "core_shape": "Teardrop", "d_core": 80, "n_core": 9,
        "gap_side": 67, "gap_between": 52,
        "h_straight": 40, "h_taper": 40,
    },
    "HCS 250mm — Capsule core": {
        "h": 250, "b_bottom": 1199, "b_top": 1187,
        "tf_top": 52, "tf_bot": 50,
        "core_shape": "Capsule", "d_core": 80, "n_core": 9,
        "gap_side": 67, "gap_between": 52,
        "h_straight": 100, "h_taper": 20,
    },
}

# =============================================================================
# 2. GEOMETRY CALCULATIONS
# =============================================================================
def calc_core_area(core_shape: str, d_core: float, h_straight: float, h_taper: float) -> float:
    """Area of one core void (mm²)."""
    if core_shape == "Circular":
        return (math.pi / 4) * d_core ** 2
    elif core_shape == "Capsule":
        return (math.pi / 4) * d_core ** 2 + d_core * h_straight
    else:  # Teardrop
        return (math.pi / 4) * d_core ** 2 + 0.65 * d_core * h_taper

def calc_h_core(core_shape: str, d_core: float, h_straight: float, h_taper: float) -> float:
    """Total height of one core void (mm)."""
    if core_shape == "Circular":
        return d_core
    elif core_shape == "Capsule":
        return d_core + h_straight
    else:
        return d_core + h_taper

def calc_modular_ratio(wc: float, f_c: float, wc_top: float, f_c_top: float) -> tuple:
    """Calculate Ec_hcs, Ec_top, n_mod (MPa, MPa, -)."""
    wc_kgm3 = wc * 1000 / 9.81
    wc_top_kgm3 = wc_top * 1000 / 9.81
    Ec_hcs = 0.043 * (wc_kgm3 ** 1.5) * math.sqrt(f_c)
    Ec_top = 0.043 * (wc_top_kgm3 ** 1.5) * math.sqrt(f_c_top)
    n_mod = Ec_top / Ec_hcs if Ec_hcs > 0 else 1.0
    return Ec_hcs, Ec_top, n_mod

# =============================================================================
# 3. SECTION PROPERTIES (Gross, Net, Composite, Eccentricity)
# =============================================================================
def _I_void_own(core_shape: str, d_core: float, h_straight: float, h_taper: float) -> float:
    """Moment of inertia of one core void about its own centroid."""
    if core_shape == "Circular":
        return math.pi / 64.0 * d_core ** 4
    elif core_shape == "Capsule":
        return (math.pi / 64.0 * d_core ** 4) + (d_core * h_straight ** 3 / 12.0)
    else:  # Teardrop
        return (math.pi / 128.0 * d_core ** 4) + (d_core * h_taper ** 3 / 36.0)

def calc_net_section(b_top: float, h: float, tf_bot: float,
                     core_shape: str, d_core: float, n_core: int,
                     h_straight: float, h_taper: float,
                     A_core_1: float, A_voids_total: float,
                     h_core: float) -> dict:
    """Return gross & net section properties (mm, mm², mm⁴)."""
    Ag = b_top * h
    yb_g = h / 2.0
    Ig = b_top * h ** 3 / 12.0
    y_void_c = tf_bot + h_core / 2.0
    An = Ag - A_voids_total
    yb = (Ag * (h / 2.0) - A_voids_total * y_void_c) / An if An > 0 else 0
    yt = h - yb
    I_gross_shifted = Ig + Ag * (h / 2.0 - yb) ** 2
    I_v1 = _I_void_own(core_shape, d_core, h_straight, h_taper)
    d_void = y_void_c - yb
    I_voids = n_core * (I_v1 + A_core_1 * d_void ** 2)
    In = I_gross_shifted - I_voids
    Sb = In / yb if yb > 0 else 0
    St = In / yt if yt > 0 else 0
    kb = In / (An * yb) if An * yb > 0 else 0
    kt = In / (An * yt) if An * yt > 0 else 0
    r2 = In / An if An > 0 else 0
    return {
        "Ag": Ag, "yb_g": yb_g, "yt_g": yt, "Ig": Ig, "Sb_g": Ig / yb_g, "St_g": Ig / yt,
        "An": An, "yb": yb, "yt": yt, "In": In, "Sb": Sb, "St": St,
        "kb": kb, "kt": kt, "r2": r2, "y_void_c": y_void_c
    }

def calc_composite_section(net_props: dict, b_top: float, h: float,
                           t_topping: float, n_mod: float, hcs_type: str) -> dict:
    """Composite section properties (HCS + topping). Returns dict with all net + composite keys."""
    An = net_props["An"]
    yb = net_props["yb"]
    In = net_props["In"]
    if t_topping <= 0:
        comp = dict(net_props)
        comp.update({"A_comp": An, "yb_comp": yb, "yt_comp": net_props["yt"],
                     "I_comp": In, "Sb_comp": net_props["Sb"], "St_comp": net_props["St"],
                     "St_hcs": net_props["St"], "St_top_tr": net_props["St"],
                     "h_total": h, "A_top_tr": 0, "b_top_tr": 0})
        return comp
    if hcs_type == "Half Slab (Open Top)":
        h_total = h + t_topping
        A_comp = b_top * h_total
        yb_comp = h_total / 2.0
        yt_comp = h_total - yb_comp
        I_comp = b_top * h_total ** 3 / 12.0
        Sb_comp = I_comp / yb_comp
        St_comp = I_comp / yt_comp
        St_hcs = I_comp / abs(h - yb_comp) if abs(h - yb_comp) > 1e-3 else 1e9
        St_top_tr = St_comp * n_mod
        comp = dict(net_props)
        comp.update({"A_comp": A_comp, "yb_comp": yb_comp, "yt_comp": yt_comp,
                     "I_comp": I_comp, "Sb_comp": Sb_comp, "St_comp": St_comp,
                     "St_hcs": St_hcs, "St_top_tr": St_top_tr,
                     "h_total": h_total, "A_top_tr": b_top * t_topping, "b_top_tr": b_top})
        return comp
    # Full HCS + transformed topping
    b_top_tr = b_top * n_mod
    A_top_tr = b_top_tr * t_topping
    y_top_c = h + t_topping / 2.0
    A_comp = An + A_top_tr
    yb_comp = (An * yb + A_top_tr * y_top_c) / A_comp if A_comp > 0 else 0
    yt_comp = h + t_topping - yb_comp
    d_net = yb_comp - yb
    d_top = y_top_c - yb_comp
    I_top_tr = b_top_tr * t_topping ** 3 / 12.0
    I_comp = In + An * d_net ** 2 + I_top_tr + A_top_tr * d_top ** 2
    Sb_comp = I_comp / yb_comp if yb_comp > 0 else 0
    St_comp = I_comp / yt_comp if yt_comp > 0 else 0
    St_hcs = I_comp / abs(h - yb_comp) if abs(h - yb_comp) > 1e-3 else 1e9
    St_top_tr = St_comp * n_mod
    comp = dict(net_props)
    comp.update({"A_comp": A_comp, "yb_comp": yb_comp, "yt_comp": yt_comp,
                 "I_comp": I_comp, "Sb_comp": Sb_comp, "St_comp": St_comp,
                 "St_hcs": St_hcs, "St_top_tr": St_top_tr,
                 "h_total": h + t_topping, "A_top_tr": A_top_tr, "b_top_tr": b_top_tr})
    return comp

def calc_ps_eccentricity(yb: float, dp_bot: float, dp_top: float,
                         n_top: int, Aps_bot: float, Aps_top: float) -> dict:
    """Eccentricity of prestressing steel from centroid."""
    e_bot = dp_bot - yb
    e_top = dp_top - yb if n_top > 0 else 0.0
    Aps_total = Aps_bot + Aps_top
    e_net = (Aps_bot * e_bot + Aps_top * e_top) / Aps_total if Aps_total > 0 else 0.0
    return {"e_bot": e_bot, "e_top": e_top, "e_net": e_net}

def get_all_section_props(inp: dict) -> dict:
    """Master function: compute all section properties from input dict."""
    b_top = inp["b_top"]
    h = inp["h"]
    tf_bot = inp["tf_bot"]
    core_shape = inp["core_shape"]
    d_core = inp["d_core"]
    n_core = inp["n_core"]
    h_straight = inp.get("h_straight", 0)
    h_taper = inp.get("h_taper", 0)
    A_core_1 = inp["A_core_1"]
    A_voids_total = inp["A_voids_total"]
    h_core = inp["h_core"]
    has_topping = inp.get("has_topping", False)
    t_topping = inp.get("t_topping", 0) if has_topping else 0
    n_mod = inp.get("n_mod", 1.0)
    hcs_type = inp["hcs_type"]
    cover_bot = inp.get("cover_bot", 35)
    cover_top = inp.get("cover_top", 30)
    n_bot = inp.get("n_bot", 0)
    n_top = inp.get("n_top", 0)
    ps_area = inp.get("ps_area", 19.6)
    Aps_bot = inp.get("Aps_bot", n_bot * ps_area)
    Aps_top = inp.get("Aps_top", n_top * ps_area)
    dp_bot = h - cover_bot
    dp_top = cover_top

    net = calc_net_section(b_top, h, tf_bot, core_shape, d_core, n_core,
                           h_straight, h_taper, A_core_1, A_voids_total, h_core)
    comp = calc_composite_section(net, b_top, h, t_topping, n_mod, hcs_type)
    ecc = calc_ps_eccentricity(net["yb"], dp_bot, dp_top, n_top, Aps_bot, Aps_top)
    result = dict(comp)
    result.update(ecc)
    return result

# =============================================================================
# 4. SELF-WEIGHT FUNCTIONS
# =============================================================================
def calc_SW_HCS(wc: float, b_bottom: float, h: float, A_voids_total: float, hcs_type: str) -> float:
    """Self-weight of HCS slab (kN/m²)."""
    A_conc = b_bottom * h - A_voids_total
    return wc * A_conc / b_bottom / 1000.0

def calc_SW_topping(wc_top: float, t_topping: float, has_topping: bool) -> float:
    """Topping self-weight (kN/m²)."""
    if has_topping and t_topping > 0:
        return wc_top * t_topping / 1000.0
    return 0.0

# =============================================================================
# 5. SPAN & LOADS (Transfer length, SFD/BMD)
# =============================================================================
def calc_transfer_development_length(ps_type: str, d_ps: float, fpu: float,
                                     fpi: float, fpy: float,
                                     assumed_loss_pct: float = 20.0) -> dict:
    fse_est = fpi * (1 - assumed_loss_pct / 100.0)
    fps_est = min(fpu, fpy + 70.0)
    if ps_type == "PC Wire (plain/indented)":
        l_t = 50.0 * d_ps
        method = "Wire: 50d"
    else:
        l_t_60 = 60.0 * d_ps
        l_t_aci = (fse_est / 20.7) * d_ps
        l_t = max(l_t_60, l_t_aci)
        method = f"Strand: max(60d, fse/20.7*d)"
    l_d = l_t + (fps_est - fse_est) * d_ps / 20.7
    return {
        "l_t": l_t, "l_d": l_d, "fse_est": fse_est, "fps_est": fps_est,
        "method_lt": method,
        "loss_note": f"Assumed loss = {assumed_loss_pct}%"
    }

def check_prestress_development(L_an: float, l_d: float) -> dict:
    if L_an >= 1.5 * l_d:
        return {"status": "FULL", "is_prestressed": True,
                "message": "Full prestress development assumed. OK."}
    elif L_an >= l_d:
        return {"status": "PARTIAL", "is_prestressed": "partial",
                "message": f"L_an/l_d = {L_an/l_d:.2f}. Partial development."}
    else:
        return {"status": "NON-PRESTRESSED", "is_prestressed": False,
                "message": f"CRITICAL: L_an/l_d = {L_an/l_d:.2f} < 1."}

def calc_factored_loads_and_diagrams(L_an, b_bottom, t_topping,
                                     wc, wc_top, has_topping,
                                     SW_HCS, SW_topping, SDL, LL,
                                     has_point_load, P1_DL, P1_LL, x_P1,
                                     P2_DL, P2_LL, x_P2, slab_position,
                                     N=200,
                                     lf_DL=1.2, lf_LL=1.6, lf_SDL=1.2,
                                     lf_P1DL=1.2, lf_P1LL=1.6,
                                     lf_P2DL=1.2, lf_P2LL=1.6,
                                     w_line_DL=0.0, w_line_LL=0.0,
                                     x_line_start=0.0, x_line_end=-1.0,
                                     lf_line_DL=1.2, lf_line_LL=1.6) -> dict:
    if x_line_end < 0:
        x_line_end = float(L_an)
    x_line_start = float(max(0.0, x_line_start))
    x_line_end = float(min(L_an, x_line_end))
    wu_area = (lf_DL * (SW_HCS + SW_topping) + lf_SDL * SDL + lf_LL * LL)
    ws_area = SW_HCS + SW_topping + SDL + LL
    wu_line = wu_area * b_bottom / 1e6
    ws_line = ws_area * b_bottom / 1e6

    # Line load
    wu_lineld = (lf_line_DL * w_line_DL + lf_line_LL * w_line_LL) / 1000.0
    ws_lineld = (w_line_DL + w_line_LL) / 1000.0
    line_length = max(0.0, x_line_end - x_line_start)

    # Point loads
    eff_w = 0.50 * L_an if slab_position == "Interior slab" else 0.25 * L_an
    rf = min(b_bottom / max(eff_w, 1.0), 1.0)
    if has_point_load:
        Pu1 = (lf_P1DL * P1_DL + lf_P1LL * P1_LL) * rf
        Ps1 = (P1_DL + P1_LL) * rf
        P2_active = (P2_DL + P2_LL) > 0
        Pu2 = (lf_P2DL * P2_DL + lf_P2LL * P2_LL) * rf if P2_active else 0.0
        Ps2 = (P2_DL + P2_LL) * rf if P2_active else 0.0
        x_P1f = float(x_P1)
        x_P2f = float(x_P2) if P2_active else L_an * 2
    else:
        Pu1 = Ps1 = Pu2 = Ps2 = 0.0
        x_P1f = L_an * 2
        x_P2f = L_an * 2

    # Line load reactions
    if wu_lineld > 0 and line_length > 0:
        a, b, Lx = x_line_start, x_line_end, float(L_an)
        Ra_u_line = wu_lineld * (b - a) * (Lx - (a + b) / 2.0) / Lx
        Rb_u_line = wu_lineld * (b - a) * ((a + b) / 2.0) / Lx
        Ra_s_line = ws_lineld * (b - a) * (Lx - (a + b) / 2.0) / Lx
    else:
        Ra_u_line = Rb_u_line = Ra_s_line = 0.0

    # Reactions
    Ra_u = (wu_line * L_an / 2 + Pu1 * (L_an - x_P1f) / L_an +
            Pu2 * (L_an - x_P2f) / L_an + Ra_u_line)
    Ra_s = (ws_line * L_an / 2 + Ps1 * (L_an - x_P1f) / L_an +
            Ps2 * (L_an - x_P2f) / L_an + Ra_s_line)

    x = np.linspace(0.0, L_an, N)
    step1u = np.where(x > x_P1f, Pu1, 0.0)
    step2u = np.where(x > x_P2f, Pu2, 0.0)
    step1s = np.where(x > x_P1f, Ps1, 0.0)
    step2s = np.where(x > x_P2f, Ps2, 0.0)

    # Line load contribution
    Q_u = wu_lineld * np.maximum(0.0, np.minimum(x, x_line_end) - x_line_start)
    Q_s = ws_lineld * np.maximum(0.0, np.minimum(x, x_line_end) - x_line_start)

    def line_moment(w, a, b, xarr):
        xeff = np.minimum(xarr, b)
        M = w * (xeff - a) * (xarr - 0.5 * (xeff + a))
        return np.where(xarr > a, M, 0.0)

    M_line_u = line_moment(wu_lineld, x_line_start, x_line_end, x)
    M_line_s = line_moment(ws_lineld, x_line_start, x_line_end, x)

    Vu = Ra_u - wu_line * x - step1u - step2u - Q_u
    Mu = (Ra_u * x - wu_line * x**2 / 2.0 - step1u * (x - x_P1f) -
          step2u * (x - x_P2f) - M_line_u)
    Vs = Ra_s - ws_line * x - step1s - step2s - Q_s
    Ms = (Ra_s * x - ws_line * x**2 / 2.0 - step1s * (x - x_P1f) -
          step2s * (x - x_P2f) - M_line_s)

    return {
        "x_arr": x, "Vu_arr": Vu, "Mu_arr": Mu, "Vs_arr": Vs, "Ms_arr": Ms,
        "Ra_u": float(Ra_u), "Rb_u": float(Ra_u),  # Rb not returned separately
        "wu_area": wu_area, "ws_area": ws_area,
        "Vu_max": float(np.max(np.abs(Vu))),
        "Mu_max": float(np.max(Mu)), "Mu_max_x": float(x[np.argmax(Mu)]),
        "Pu1_red": Pu1, "Pu2_red": Pu2, "x_P1_use": x_P1f, "x_P2_use": x_P2f,
    }

# =============================================================================
# 6. PRESTRESS LOSSES
# =============================================================================
def loss_elastic_shortening(f_ci, Ec_ci, Eps, Pi, e, An, In, Mg_sw):
    Pi_N = Pi * 1000
    Mg_sw_Nmm = Mg_sw * 1000
    fcgp = (Pi_N / An) + (Pi_N * e * e / In) - (Mg_sw_Nmm * e / In)
    return max((Eps / Ec_ci) * fcgp, 0.0)

def loss_creep(fcgp, fcdp, Eps, Ec, Kcr=2.0):
    return max(Kcr * (Eps / Ec) * (fcgp - fcdp), 0.0)

def loss_shrinkage(Eps, RH=75.0, V_S=38.0, Ksh=1.0):
    return max(8.2e-6 * Ksh * Eps * (1 - 0.06 * V_S) * (100 - RH), 0.0)

def loss_relaxation(ps_type, fpi, fpu, sum_other_losses):
    Kre, J = 34.5, 0.04
    if "Wire" in ps_type:
        C = 1.0
    else:
        ratio = fpi / fpu
        C = max(1.45 - 0.3 * ratio, 0.5)
    return max((Kre - J * sum_other_losses) * C, 0.0)

def get_prestress_losses(inp: dict) -> dict:
    RH = inp.get("RH", 75.0)
    V_S = inp.get("V_S", 38.0)
    wc = inp["wc"]
    f_ci = inp["f_ci"]
    fpi = inp["fpi"]
    fpu = inp["fpu"]
    Eps = inp["Eps"]
    Pi = inp["Pi"]
    Aps_bot = inp["Aps_bot"]
    Aps_top = inp.get("Aps_top", 0.0)
    Aps_total = Aps_bot + Aps_top
    e_net = inp.get("sp_e_net", 0.0)
    An = inp["sp_An"]
    In = inp["sp_In"]
    L = inp["L_an"]
    b_bottom = inp["b_bottom"]
    SW_HCS = inp.get("SW_HCS", 0.0)   # already calculated
    has_topping = inp.get("has_topping", False)
    SW_topping = inp.get("SW_topping", 0.0)
    SDL = inp.get("SDL", 0.0)
    Ec_ci = 0.043 * (wc * 1000 / 9.81) ** 1.5 * math.sqrt(f_ci)
    w_sw = SW_HCS * b_bottom / 1e6   # kN/mm
    Mg_sw = (w_sw * L * L) / 8.0     # kN·mm
    ES = loss_elastic_shortening(f_ci, Ec_ci, Eps, Pi, e_net, An, In, Mg_sw)
    # fcgp
    Pi_N = Pi * 1000
    Mg_sw_Nmm = Mg_sw * 1000
    fcgp = (Pi_N / An) + (Pi_N * e_net * e_net / In) - (Mg_sw_Nmm * e_net / In)
    # fcdp
    if has_topping:
        w_sdl_line = (SW_topping + SDL) * b_bottom / 1e6
        M_sdl = (w_sdl_line * L * L) / 8.0
        I_comp = inp.get("sp_I_comp", In)
        fcdp = (M_sdl * 1000 * e_net) / I_comp if I_comp > 0 else 0.0
    else:
        fcdp = 0.0
    Ec = inp.get("Ec_hcs", Ec_ci)
    CR = loss_creep(fcgp, fcdp, Eps, Ec)
    SH = loss_shrinkage(Eps, RH, V_S)
    RE = loss_relaxation(inp["ps_type"], fpi, fpu, ES + CR + SH)
    total = ES + CR + SH + RE
    total_pct = (total / fpi) * 100 if fpi > 0 else 0
    fse = fpi - total
    Pe = fse * Aps_total / 1000.0
    return {
        "pl_ES": ES, "pl_CR": CR, "pl_SH": SH, "pl_RE": RE,
        "pl_total_MPa": total, "pl_total_pct": total_pct,
        "pl_fse": fse, "pl_Pe": Pe,
        "pl_fse_MPa": fse, "pl_fpi_MPa": fpi
    }

# =============================================================================
# 7. STRESS CHECKS
# =============================================================================
LIMIT_COMP_TRANSFER = 0.60
LIMIT_TENS_TRANSFER = 0.25
LIMIT_COMP_TOTAL = 0.60
LIMIT_TENS_CLASS_T = 0.50
LIMIT_TENS_CLASS_U = 0.0

def stress_at_transfer(Pi, e_net, An, In, yb, h, Mg_sw, f_ci):
    Pi_N = Pi * 1000
    yt = h - yb
    f_bot = -Pi_N/An + Pi_N * e_net * yb / In - Mg_sw * 1000 * yb / In
    f_top = -Pi_N/An - Pi_N * e_net * yt / In + Mg_sw * 1000 * yt / In
    comp_limit = LIMIT_COMP_TRANSFER * f_ci
    tens_limit = LIMIT_TENS_TRANSFER * math.sqrt(f_ci)
    ok = (abs(f_top) <= comp_limit and abs(f_bot) <= comp_limit)
    return {"f_top": f_top, "f_bot": f_bot, "limit_comp": comp_limit,
            "limit_tens": tens_limit, "status": "OK" if ok else "NG"}

def stress_at_lifting(Pe, e_net, An, In, yb, h, Mg_sw, f_ci):
    Pe_N = Pe * 1000
    yt = h - yb
    f_bot = -Pe_N/An + Pe_N * e_net * yb / In - Mg_sw * 1000 * yb / In
    f_top = -Pe_N/An - Pe_N * e_net * yt / In + Mg_sw * 1000 * yt / In
    comp_limit = LIMIT_COMP_TRANSFER * f_ci
    tens_limit = LIMIT_TENS_TRANSFER * math.sqrt(f_ci)
    ok = (abs(f_top) <= comp_limit and abs(f_bot) <= comp_limit)
    return {"f_top": f_top, "f_bot": f_bot, "limit_comp": comp_limit,
            "limit_tens": tens_limit, "status": "OK" if ok else "NG"}

def stress_at_construction(Pe, e_net, An, In, yb, h, Mg_sw, Mg_sdl, Mg_topping, f_c):
    Pe_N = Pe * 1000
    yt = h - yb
    M_total = (Mg_sw + Mg_sdl + Mg_topping) * 1000
    f_bot = -Pe_N/An + Pe_N * e_net * yb / In - M_total * yb / In
    f_top = -Pe_N/An - Pe_N * e_net * yt / In + M_total * yt / In
    comp_limit = LIMIT_COMP_TOTAL * f_c
    tens_limit = LIMIT_TENS_CLASS_T * math.sqrt(f_c)
    ok = (abs(f_top) <= comp_limit and abs(f_bot) <= comp_limit)
    return {"f_top": f_top, "f_bot": f_bot, "limit_comp": comp_limit,
            "limit_tens": tens_limit, "status": "OK" if ok else "NG"}

def stress_at_service(Pe, e_net, An, In, yb, h,
                      comp_dict, M_DL, M_SDL, M_LL, f_c, section_class="T"):
    Pe_N = Pe * 1000
    yt = h - yb
    f_bot_nc = -Pe_N/An - Pe_N * e_net * yb / In + M_DL * 1000 * yb / In
    f_top_nc = -Pe_N/An + Pe_N * e_net * yt / In - M_DL * 1000 * yt / In
    I_comp = comp_dict["I_comp"]
    yb_comp = comp_dict["yb_comp"]
    yt_comp = comp_dict["yt_comp"]
    M_super = (M_SDL + M_LL) * 1000
    f_bot_super = M_super * yb_comp / I_comp
    f_top_super = -M_super * yt_comp / I_comp
    f_bot = f_bot_nc + f_bot_super
    f_top = f_top_nc + f_top_super
    sqrt_fc = math.sqrt(f_c)
    tens_limit = LIMIT_TENS_CLASS_T * sqrt_fc if section_class.upper() == 'T' else LIMIT_TENS_CLASS_U * sqrt_fc
    comp_limit = LIMIT_COMP_TOTAL * f_c
    ok = (abs(f_top) <= comp_limit and abs(f_bot) <= comp_limit)
    return {"f_bot": f_bot, "f_top": f_top, "limit_comp": comp_limit,
            "limit_tens": tens_limit, "status": "OK" if ok else "NG"}

def get_all_stress_checks(inp: dict) -> dict:
    Pi = inp["Pi"]
    Pe = inp.get("pl_Pe", Pi)
    e_net = inp.get("sp_e_net", 0.0)
    An = inp["sp_An"]
    In = inp["sp_In"]
    yb = inp["sp_yb"]
    h = inp["h"]
    f_ci = inp["f_ci"]
    f_c = inp["f_c"]
    L = inp["L_an"]
    b_bottom = inp["b_bottom"]
    SW_HCS = inp.get("SW_HCS", 0.0)
    SW_topping = inp.get("SW_topping", 0.0)
    SDL = inp.get("SDL", 0.0)
    LL = inp.get("LL", 0.0)
    has_topping = inp.get("has_topping", False)
    # Moments
    w_sw = SW_HCS * b_bottom / 1e6
    Mg_sw = w_sw * L * L / 8.0
    w_top = SW_topping * b_bottom / 1e6 if has_topping else 0.0
    Mg_top = w_top * L * L / 8.0
    w_sdl = SDL * b_bottom / 1e6
    M_sdl = w_sdl * L * L / 8.0
    w_ll = LL * b_bottom / 1e6
    M_ll = w_ll * L * L / 8.0

    transfer = stress_at_transfer(Pi, e_net, An, In, yb, h, Mg_sw, f_ci)
    lifting = stress_at_lifting(Pe, e_net, An, In, yb, h, Mg_sw, f_ci)
    construction = stress_at_construction(Pe, e_net, An, In, yb, h, Mg_sw, M_sdl, Mg_top, f_c)
    comp_dict = {
        "I_comp": inp.get("sp_I_comp", In),
        "yb_comp": inp.get("sp_yb_comp", yb),
        "yt_comp": inp.get("sp_yt_comp", h + inp.get("t_topping", 0) - yb),
    }
    service = stress_at_service(Pe, e_net, An, In, yb, h, comp_dict,
                                Mg_sw, M_sdl, M_ll, f_c, "T")
    return {
        "sc_transfer": transfer, "sc_lifting": lifting,
        "sc_construction": construction, "sc_service": service,
        "sc_service_class": "T"
    }

# =============================================================================
# 8. CAPACITY (Flexure & Shear)
# =============================================================================
def calc_fps(fpu, fpy, rho_p, f_c, beta1, gamma_p=0.28):
    term = (rho_p * fpu) / (beta1 * f_c) * gamma_p
    fps = fpu * (1 - term)
    return min(fps, fpy)

def calc_moment_capacity(Aps, fps, dp, f_c, b, phi=0.9):
    a = Aps * fps / (0.85 * f_c * b)
    Mn = (Aps * fps * (dp - a/2)) / 1e6
    return Mn, phi * Mn, a

def calc_Vci(f_c, bw, dp, fpe, Vu, Mu, Mg, lambda_factor=1.0):
    sqrt_fc = math.sqrt(f_c)
    Vc_min = 0.05 * lambda_factor * sqrt_fc * bw * dp / 1000.0
    return Vc_min  # simplified lower bound

def calc_Vcw(f_c, bw, dp, fpc, Vp=0, lambda_factor=1.0):
    sqrt_fc = math.sqrt(f_c)
    Vcw = (0.29 * lambda_factor * sqrt_fc + 0.3 * fpc) * bw * dp / 1000.0 + Vp
    return Vcw

def get_capacity_results(inp: dict) -> dict:
    f_c = inp["f_c"]
    fpu = inp["fpu"]
    fpy = inp["fpy"]
    Eps = inp["Eps"]
    Aps = inp["Aps_bot"]
    dp = inp["dp_bot"]
    b = inp["b_top"]
    Pe = inp.get("pl_Pe", inp["Pi"])
    # Beta1
    if f_c <= 28:
        beta1 = 0.85
    elif f_c < 55:
        beta1 = 0.85 - 0.05 * (f_c - 28) / 7
    else:
        beta1 = 0.65
    rho_p = Aps / (b * dp) if b * dp > 0 else 0
    fps = calc_fps(fpu, fpy, rho_p, f_c, beta1)
    Mn, phi_Mn, a = calc_moment_capacity(Aps, fps, dp, f_c, b)
    Mu_max = inp.get("lb_Mu_max", 0) / 1e6
    DCR_M = Mu_max / phi_Mn if phi_Mn > 0 else 999.0

    bw = inp.get("bw_shear", inp["b_bottom"] - inp["n_core"] * inp["d_core"])
    fpc = Pe * 1000 / inp["sp_An"]
    # Use Vu/Mu arrays if available
    x_arr = inp.get("lb_x_arr", np.linspace(0, inp["L_an"], 100))
    Vu_arr = inp.get("lb_Vu_arr", np.zeros_like(x_arr))
    Mu_load = inp.get("lb_Mu_arr", np.zeros_like(x_arr)) / 1e6
    Vci_arr = np.zeros_like(Vu_arr)
    Vcw_arr = np.zeros_like(Vu_arr)
    phi_Vn_arr = np.zeros_like(Vu_arr)
    for i, x in enumerate(x_arr):
        Vu = abs(Vu_arr[i])
        Vci = calc_Vci(f_c, bw, dp, fpc, Vu, Mu_load[i], 0.0)
        Vcw = calc_Vcw(f_c, bw, dp, fpc)
        Vn = min(Vci, Vcw)
        phi_Vn_arr[i] = 0.75 * Vn
        Vci_arr[i] = Vci
        Vcw_arr[i] = Vcw
    phi_Vn_min = np.min(phi_Vn_arr) if len(phi_Vn_arr) > 0 else 0.0
    Vu_max_val = inp.get("lb_Vu_max", 0)
    DCR_V = Vu_max_val / phi_Vn_min if phi_Vn_min > 0 else 999.0

    # Minimum shear reinforcement check
    h = inp["h"]
    has_topping = inp.get("has_topping", False)
    needs_Av_min = False
    if h > 317 and not has_topping:
        idx = np.argmax(np.abs(Vu_arr))
        if phi_Vn_arr[idx] > 0 and abs(Vu_arr[idx]) > 0.5 * phi_Vn_arr[idx]:
            needs_Av_min = True

    return {
        "cap_fps": fps, "cap_Mn": Mn, "cap_phi_Mn": phi_Mn, "cap_a": a,
        "cap_DCR_M": DCR_M,
        "cap_Vci_arr": Vci_arr, "cap_Vcw_arr": Vcw_arr,
        "cap_phi_Vn_arr": phi_Vn_arr, "cap_phi_Vn_min": phi_Vn_min,
        "cap_DCR_V": DCR_V, "cap_needs_Av_min": needs_Av_min
    }

# =============================================================================
# 9. DEFLECTION & VIBRATION
# =============================================================================
def camber_prestress(Pe, e, L, Ec, I):
    Pe_N = Pe * 1000
    if Ec <= 0 or I <= 0: return 0.0
    return Pe_N * e * L * L / (8.0 * Ec * I)

def deflection_selfweight(w_sw, L, Ec, I):
    w_Nmm = w_sw * 1000
    if Ec <= 0 or I <= 0: return 0.0
    return -(5.0 * w_Nmm * L**4) / (384.0 * Ec * I)

def deflection_uniform_load(w, L, Ec, I):
    return - (5.0 * w * 1000 * L**4) / (384.0 * Ec * I)

def deflection_point_load(P, a, L, Ec, I):
    if P == 0 or Ec <= 0 or I <= 0: return 0.0
    P_N = P * 1000
    a = min(a, L - a)
    return - P_N * a * (3.0 * L**2 - 4.0 * a**2) / (48.0 * Ec * I)

def check_deflection_limits(net_deflection, span_mm,
                            limit_ll_fraction=360, limit_total_fraction=240):
    limit_ll = span_mm / limit_ll_fraction
    limit_total = span_mm / limit_total_fraction
    return {
        "limit_ll_mm": limit_ll, "limit_total_mm": limit_total,
        "deflection_mm": net_deflection,
        "status_ll": "OK" if abs(net_deflection) <= limit_ll else "NG",
        "status_total": "OK" if abs(net_deflection) <= limit_total else "NG"
    }

def calc_thermal_camber(alpha_T, delta_T, L_an, h):
    if h <= 0: return 0.0
    return alpha_T * delta_T * L_an**2 / (8.0 * h)

def calc_vibration(SW_HCS, SW_topping, SDL, LL, b_bottom, L_an, Ec_hcs, I_comp,
                   damping_ratio=3.0, vibration_mode="Walking / Occupancy"):
    limits_fn = {"Walking / Occupancy": 8.0, "Machine / Equipment": 4.0, "Sensitive (Lab/Hospital)": 12.0}
    limits_ag = {"Walking / Occupancy": 0.005, "Machine / Equipment": 0.015, "Sensitive (Lab/Hospital)": 0.004}
    fn_limit = limits_fn.get(vibration_mode, 8.0)
    ag_limit = limits_ag.get(vibration_mode, 0.005)
    w_total_kNm2 = SW_HCS + SW_topping + SDL + 0.1 * LL
    m_per_mm = (w_total_kNm2 * b_bottom / 1e6 * 1000.0) / 9810.0
    if m_per_mm <= 0: m_per_mm = 1e-6
    m_per_m = m_per_mm * 1000.0
    EI_SI = Ec_hcs * I_comp * 1e-6
    L_m = L_an / 1000.0
    fn = (math.pi**2 / (2.0 * L_m**2)) * math.sqrt(EI_SI / m_per_m)
    fn_ok = fn >= fn_limit
    W_eff_kN = w_total_kNm2 * (b_bottom / 1000.0) * L_m
    beta = damping_ratio / 100.0
    ag = 0.29 * math.exp(-2.0 * math.pi * beta) / W_eff_kN if W_eff_kN > 0 else 0
    ag_ok = ag <= ag_limit
    return {"fn": fn, "fn_limit": fn_limit, "fn_ok": fn_ok,
            "ag": ag, "ag_limit": ag_limit, "ag_ok": ag_ok,
            "W_eff": W_eff_kN}

def get_deflection_results(inp: dict) -> dict:
    L = inp["L_an"]
    Pe = inp.get("pl_Pe", inp["Pi"])
    e = inp.get("sp_e_net", inp.get("sp_e_bot", 0.0))
    Ec = inp.get("Ec_hcs", 33000.0)
    I_net = inp["sp_In"]
    has_topping = inp.get("has_topping", False)
    I_comp = inp.get("sp_I_comp", I_net)
    b = inp["b_bottom"]
    h = inp["h"]
    SW_HCS = inp.get("SW_HCS", 0.0)
    SW_topping = inp.get("SW_topping", 0.0)
    SDL = inp.get("SDL", 0.0)
    LL = inp.get("LL", 0.0)

    w_sw = SW_HCS * b / 1e6
    w_topping = SW_topping * b / 1e6 if has_topping else 0.0
    w_sdl = SDL * b / 1e6
    w_ll = LL * b / 1e6

    delta_ps_initial = camber_prestress(Pe, e, L, Ec, I_net)
    delta_sw = deflection_selfweight(w_sw, L, Ec, I_net)
    delta_topping = deflection_uniform_load(w_topping, L, Ec, I_net) if has_topping else 0.0
    delta_sdl = deflection_uniform_load(w_sdl, L, Ec, I_comp)
    delta_ll = deflection_uniform_load(w_ll, L, Ec, I_comp)

    net_release = delta_ps_initial + delta_sw

    # PCI multipliers (defaults from PCI Table 4.8.3, can be customized via inputs)
    mult = inp.get("pci_multipliers", None)
    if mult is None:
        mult = {"final_camber": 2.70 if not has_topping else 2.40,
                "final_sw": 2.40, "final_sdl": 3.0, "final_ll": 1.0}
    final_camber = delta_ps_initial * mult.get("final_camber", 2.70)
    final_sw = delta_sw * mult.get("final_sw", 2.40)
    final_sdl = delta_sdl * mult.get("final_sdl", 3.0)
    final_ll = delta_ll   # LL transient, factor 1.0

    total_deflection = final_camber + final_sw + final_sdl + final_ll

    limit_ll_frac = inp.get("limit_LL_fraction", 360)
    limit_tot_frac = inp.get("limit_total_fraction", 240)
    limit_check = check_deflection_limits(total_deflection, L,
                                          limit_ll_frac, limit_tot_frac)

    # Thermal
    delta_thermal = 0.0
    if inp.get("has_thermal", False):
        alpha_T = inp.get("alpha_T", 10e-6)
        delta_T = inp.get("delta_T", 0.0)
        delta_thermal = calc_thermal_camber(alpha_T, delta_T, L, h)

    # Vibration
    vib = calc_vibration(SW_HCS, SW_topping, SDL, LL, b, L, Ec, I_comp,
                         inp.get("damping_ratio", 3.0),
                         inp.get("vibration_mode", "Walking / Occupancy"))

    return {
        "def_delta_ps_initial": delta_ps_initial,
        "def_delta_sw": delta_sw,
        "def_delta_topping": delta_topping,
        "def_delta_sdl": delta_sdl,
        "def_delta_ll": delta_ll,
        "def_net_release": net_release,
        "def_total_longterm": total_deflection,
        "def_limit_ll_mm": limit_check["limit_ll_mm"],
        "def_limit_total_mm": limit_check["limit_total_mm"],
        "def_status_ll": limit_check["status_ll"],
        "def_status_total": limit_check["status_total"],
        "def_final_camber": final_camber,
        "def_final_sw": final_sw,
        "def_final_sdl": final_sdl,
        "def_final_ll": final_ll,
        "def_delta_thermal": delta_thermal,
        "def_vib_fn": vib["fn"],
        "def_vib_fn_limit": vib["fn_limit"],
        "def_vib_fn_ok": vib["fn_ok"],
        "def_vib_ag": vib["ag"],
        "def_vib_ag_limit": vib["ag_limit"],
        "def_vib_ag_ok": vib["ag_ok"],
        "def_vib_W_eff": vib["W_eff"],
    }

# =============================================================================
# 10. MASTER DESIGN FUNCTION
# =============================================================================
def calculate_hcs_design(inputs: dict) -> dict:
    """
    Run the full HCS design chain.

    Required keys in inputs:
        h, b_bottom, b_top, tf_bot, core_shape, d_core, n_core,
        h_straight, h_taper, A_core_1, A_voids_total, h_core,
        has_topping, t_topping (if topping), hcs_type,
        wc, f_c, f_ci, wc_top, f_c_top (if topping),
        ps_type, wire_dia/strand_size, n_bot, n_top, cover_bot, cover_top,
        fpi_pct, fpu, fpy, Eps, ps_area,
        L_cc, bw_beam_L, bw_beam_R, b_bear_L, b_bear_R,
        SDL, LL, has_point_load (optional), P1_DL, P1_LL, x_P1, etc.,
        slab_position, has_line_load (optional), w_line_DL, w_line_LL, etc.,
        RH, V_S, limit_LL_fraction, limit_total_fraction, etc.

    Returns dict with all results (sp_*, pl_*, sc_*, cap_*, def_*, lb_*, etc.)
    """
    out = inputs.copy()  # we will build upon inputs

    # --- 1. Derived geometry ---
    out["A_core_1"] = calc_core_area(out["core_shape"], out["d_core"],
                                     out.get("h_straight",0), out.get("h_taper",0))
    out["A_voids_total"] = out["n_core"] * out["A_core_1"]
    out["h_core"] = calc_h_core(out["core_shape"], out["d_core"],
                                out.get("h_straight",0), out.get("h_taper",0))
    out["bw_shear"] = out["b_bottom"] - out["n_core"] * out["d_core"]
    # Elastic moduli
    has_top = out.get("has_topping", False)
    Ec_hcs, Ec_top, n_mod = calc_modular_ratio(out["wc"], out["f_c"],
                                               out.get("wc_top", out["wc"]),
                                               out.get("f_c_top", out["f_c"]))
    out["Ec_hcs"] = Ec_hcs
    out["Ec_top"] = Ec_top
    out["n_mod"] = n_mod
    # Self-weight
    out["SW_HCS"] = calc_SW_HCS(out["wc"], out["b_bottom"], out["h"],
                                out["A_voids_total"], out["hcs_type"])
    out["SW_topping"] = calc_SW_topping(out.get("wc_top", out["wc"]),
                                        out.get("t_topping", 0), has_top)
    # V/S auto
    out["V_S"] = (out["b_bottom"] * out["h"] - out["A_voids_total"]) / (2*(out["b_bottom"]+out["h"]))
    # Span
    L_clear = out["L_cc"] - out["bw_beam_L"]/2 - out["bw_beam_R"]/2
    L_an = L_clear + out["b_bear_L"]/2 + out["b_bear_R"]/2
    out["L_clear"] = L_clear
    out["L_an"] = L_an
    out["bear_min"] = max(L_clear/180.0, 50.8)
    # Prestress derived
    fpi = out["fpi_pct"]/100.0 * out["fpu"]
    out["fpi"] = fpi
    d_ps = out.get("wire_dia", 5.0) if "Wire" in out["ps_type"] else STRAND_PROPS[out["strand_size"]]["d_mm"]
    td = calc_transfer_development_length(out["ps_type"], d_ps, out["fpu"], fpi, out["fpy"], 20.0)
    out.update({f"lb_{k}": v for k, v in td.items()})
    dev = check_prestress_development(L_an, td["l_d"])
    out["lb_ps_status"] = dev["status"]
    out["lb_ps_is_ps"] = dev["is_prestressed"]
    out["lb_ps_message"] = dev["message"]
    # Prestress area & Pi
    Aps_bot = out["n_bot"] * out["ps_area"]
    Aps_top = out["n_top"] * out["ps_area"]
    Pi = (Aps_bot + Aps_top) * fpi / 1000.0
    out["Aps_bot"] = Aps_bot
    out["Aps_top"] = Aps_top
    out["Pi"] = Pi
    out["dp_bot"] = out["h"] - out["cover_bot"]
    out["dp_top"] = out["cover_top"]

    # Section properties
    sp = get_all_section_props(out)
    out.update({f"sp_{k}": v for k, v in sp.items()})

    # Factored load diagrams (use default load factors if not given)
    lf_DL = out.get("lf_DL", 1.2)
    lf_LL = out.get("lf_LL", 1.6)
    lf_SDL = out.get("lf_SDL", 1.2)
    ld = calc_factored_loads_and_diagrams(
        L_an, out["b_bottom"], out.get("t_topping",0),
        out["wc"], out.get("wc_top", out["wc"]), has_top,
        out["SW_HCS"], out["SW_topping"],
        out["SDL"], out["LL"],
        out.get("has_point_load", False),
        out.get("P1_DL",0), out.get("P1_LL",0), out.get("x_P1",0),
        out.get("P2_DL",0), out.get("P2_LL",0), out.get("x_P2",0),
        out.get("slab_position", "Interior slab"),
        200, lf_DL, lf_LL, lf_SDL,
        out.get("lf_P1DL",1.2), out.get("lf_P1LL",1.6),
        out.get("lf_P2DL",1.2), out.get("lf_P2LL",1.6),
        out.get("w_line_DL",0), out.get("w_line_LL",0),
        out.get("x_line_start",0), out.get("x_line_end", L_an),
        out.get("lf_line_DL",1.2), out.get("lf_line_LL",1.6)
    )
    out.update({f"lb_{k}": v for k, v in ld.items()})

    # Prestress losses
    losses = get_prestress_losses(out)
    out.update(losses)

    # Stress checks
    stress = get_all_stress_checks(out)
    out.update(stress)

    # Capacity
    cap = get_capacity_results(out)
    out.update(cap)

    # Deflection
    defl = get_deflection_results(out)
    out.update(defl)

    return out

# --- Example usage (comment out if not needed) ---
if __name__ == "__main__":
    sample_input = {
        # geometry & materials
        "h": 200, "b_bottom": 1199, "b_top": 1187, "tf_bot": 50, "tf_top": 52,
        "core_shape": "Teardrop", "d_core": 80, "n_core": 9,
        "h_straight": 40, "h_taper": 40,
        "hcs_type": "Full HCS (Hollow Core)", "has_topping": True, "t_topping": 60,
        "wc": 24.0, "f_c": 50.0, "f_ci": 35.0, "wc_top": 24.0, "f_c_top": 30.0,
        "ps_type": "PC Wire (plain/indented)", "wire_dia": 5.0,
        "n_bot": 10, "n_top": 0, "cover_bot": 35, "cover_top": 30,
        "fpi_pct": 75.0, "fpu": 1618.0, "fpy": 1432.0, "Eps": 199050.0, "ps_area": 19.6,
        "L_cc": 6000, "bw_beam_L": 300, "bw_beam_R": 300,
        "b_bear_L": 150, "b_bear_R": 150,
        "SDL": 1.5, "LL": 2.0,
        "slab_position": "Interior slab",
        "RH": 75.0,
        "limit_LL_fraction": 360, "limit_total_fraction": 240,
    }
    result = calculate_hcs_design(sample_input)
    print("Design completed. Keys:", list(result.keys()))