# -*- coding: utf-8 -*-
__title__ = "Cotas Automaticas de Familias"
__version__ = "2.6"
__doc__ = (
    "Insere cotas automaticas nas paredes da vista ativa (Elevacao ou Corte). "
    "v2.5: get_face_refs_up pega o MAIOR z entre todas as faces +Up (topo real). "
    "      Tolerancia de deduplicacao vertical reduzida para 0.5 cm. "
    "v2.6: Antes de criar cada cota, ordena as refs pela sua posicao projetada "
    "      no eixo da cota e remove qualquer ref cuja distancia para a anterior "
    "      seja <= 1 cm. Isso elimina segmentos com valor 0 ou quase 0 que "
    "      apareciam na visualizacao (ex: '100' em vez de '10')."
)

# ============================================================
# IMPORTS
# ============================================================
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Options,
    LocationCurve, Solid, PlanarFace, ReferenceArray, Line, XYZ,
    Transaction, DimensionType, UnitUtils,
)

try:
    from Autodesk.Revit.DB import UnitTypeId
    def to_ft(cm):
        return UnitUtils.ConvertToInternalUnits(cm, UnitTypeId.Centimeters)
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType
    def to_ft(cm):
        return UnitUtils.ConvertToInternalUnits(cm, DisplayUnitType.DUT_CENTIMETERS)

from pyrevit import revit, forms, script
import math

# ============================================================
# Globals
# ============================================================
doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()

# Distancia minima entre duas refs consecutivas na cota (em ft).
# Segmentos menores que isso geram valor "0" e sao removidos.
TOL_DIM_ZERO = to_ft(1.0)   # 1 cm

# ============================================================
# ETAPA 1 - Vista ativa
# ============================================================
view = doc.ActiveView
output.print_md("## Cotas Automaticas v2.6 - **{}**".format(view.Name))

if view.ViewType.ToString() not in ("Elevation", "Section"):
    forms.alert(
        "Esta ferramenta funciona em vistas de Elevacao ou Corte.\n"
        "Vista atual: {} ({})".format(view.Name, view.ViewType),
        exitscript=True,
    )

right    = view.RightDirection
up       = view.UpDirection
view_dir = view.ViewDirection

output.print_md(
    "Right: `{:.2f},{:.2f},{:.2f}` | Up: `{:.2f},{:.2f},{:.2f}`".format(
        right.X, right.Y, right.Z, up.X, up.Y, up.Z
    )
)

# ============================================================
# ETAPA 2 - Paredes na vista
# ============================================================
walls = list(
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_Walls)
    .WhereElementIsNotElementType()
    .ToElements()
)

output.print_md("**Paredes na vista:** {}".format(len(walls)))

if not walls:
    forms.alert("Nenhuma parede encontrada na vista ativa.", exitscript=True)

# ============================================================
# HELPERS
# ============================================================

def dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z


def pts_close(p1, p2, tol_ft=None):
    tol_ft = tol_ft or to_ft(2.0)
    dx = p1.X - p2.X
    dy = p1.Y - p2.Y
    dz = p1.Z - p2.Z
    return math.sqrt(dx * dx + dy * dy + dz * dz) < tol_ft


def get_wall_dir(wall):
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return None, None, None
    c  = loc.Curve
    p0 = c.GetEndPoint(0)
    p1 = c.GetEndPoint(1)
    raw = p1 - p0
    if raw.GetLength() < 1e-9:
        return None, None, None
    return raw.Normalize(), p0, p1


def get_face_refs_right(wall):
    """Retorna (ref_neg, ref_pos, r_neg, r_pos) para faces com normal em +/-Right."""
    opt = Options()
    opt.ComputeReferences = True
    opt.View = view
    geom = wall.get_Geometry(opt)

    ref_neg = None
    ref_pos = None
    r_neg   = None
    r_pos   = None

    for g in geom:
        if not isinstance(g, Solid):
            continue
        for face in g.Faces:
            if face.Reference is None or not isinstance(face, PlanarFace):
                continue
            n   = face.FaceNormal
            d_r = dot(n, right)
            if d_r > 0.8 and ref_pos is None:
                ref_pos = face.Reference
                r_pos   = dot(face.Origin, right)
            elif d_r < -0.8 and ref_neg is None:
                ref_neg = face.Reference
                r_neg   = dot(face.Origin, right)

    return ref_neg, ref_pos, r_neg, r_pos


def get_face_refs_up(wall):
    """
    Retorna (ref_bottom, ref_top, z_bottom, z_top).
    Itera TODAS as faces +Up/-Up e retorna a de maior/menor Z respectivamente,
    garantindo captura do topo real (nao o vao de porta/janela).
    """
    opt = Options()
    opt.ComputeReferences = True
    opt.View = view
    geom = wall.get_Geometry(opt)

    candidates_top = []
    candidates_bot = []

    for g in geom:
        if not isinstance(g, Solid):
            continue
        for face in g.Faces:
            if face.Reference is None or not isinstance(face, PlanarFace):
                continue
            n   = face.FaceNormal
            d_u = dot(n, up)
            z_f = dot(face.Origin, up)

            if d_u > 0.8:
                candidates_top.append((z_f, face.Reference))
            elif d_u < -0.8:
                candidates_bot.append((z_f, face.Reference))

    ref_top    = None
    z_top      = None
    ref_bottom = None
    z_bottom   = None

    if candidates_top:
        z_top, ref_top = max(candidates_top, key=lambda x: x[0])

    if candidates_bot:
        z_bottom, ref_bottom = min(candidates_bot, key=lambda x: x[0])

    return ref_bottom, ref_top, z_bottom, z_top


def classify_wall(wall):
    d, p0, p1 = get_wall_dir(wall)
    if d is None:
        return None, None, None, None

    d_right = abs(dot(d, right))
    d_up    = abs(dot(d, up))
    d_depth = abs(dot(d, view_dir))

    if d_right >= d_up and d_right >= d_depth:
        return "H", d, p0, p1
    elif d_up >= d_right and d_up >= d_depth:
        return "V", d, p0, p1

    return "D", d, p0, p1


# ============================================================
# FILTRO ANTI-ZERO
# Recebe lista de (valor_escalar, ref) ordenada pelo valor,
# remove refs cuja distancia para a anterior e <= TOL_DIM_ZERO.
# ============================================================

def filter_zero_segments(pairs):
    """
    pairs: lista de (valor_ft, Reference) ja ordenada por valor_ft.
    Retorna lista filtrada onde distancia entre consecutivos > TOL_DIM_ZERO.
    Estrategia: percorre em ordem; se distancia para o ultimo aceito for
    pequena demais, descarta o atual (mantem o ja registrado).
    """
    if not pairs:
        return []

    accepted = [pairs[0]]
    for val, ref in pairs[1:]:
        last_val = accepted[-1][0]
        if abs(val - last_val) > TOL_DIM_ZERO:
            accepted.append((val, ref))
        else:
            output.print_md(
                "  [FILTRO] Segmento descartado: distancia {:.4f} ft ({:.2f} cm) <= 1 cm".format(
                    abs(val - last_val),
                    abs(val - last_val) * 30.48
                )
            )

    return accepted


# ============================================================
# PAREDES - refs_h (horizontal)
# Coleta (r_val, ref) para ordenar e filtrar antes de criar cota
# ============================================================
refs_h_pairs = []   # lista de (r_val_ft, Reference)
pts_h        = []   # pontos para calcular linha de cota

for wall in walls:
    orient, _, p0, p1 = classify_wall(wall)
    if orient is None:
        continue

    ref_neg, ref_pos, r_neg, r_pos = get_face_refs_right(wall)

    pt_left    = p0 if dot(p0, right) <= dot(p1, right) else p1
    pt_right_w = p1 if pt_left == p0 else p0

    if ref_neg is not None and r_neg is not None:
        # verifica se ja existe ponto muito proximo
        if not any(pts_close(pt_left, ep) for ep in pts_h):
            refs_h_pairs.append((r_neg, ref_neg))
            pts_h.append(pt_left)

    if ref_pos is not None and r_pos is not None:
        if not any(pts_close(pt_right_w, ep) for ep in pts_h):
            refs_h_pairs.append((r_pos, ref_pos))
            pts_h.append(pt_right_w)

# Ordena por posicao no eixo Right e filtra segmentos zero
refs_h_pairs.sort(key=lambda x: x[0])
output.print_md("**Refs horizontais antes do filtro:** {}".format(len(refs_h_pairs)))
refs_h_pairs = filter_zero_segments(refs_h_pairs)
refs_h = [r for _, r in refs_h_pairs]
output.print_md("**Refs horizontais apos filtro:** {}".format(len(refs_h)))


# ============================================================
# PAREDES - refs_v (vertical)
# ============================================================
refs_v_pairs = []   # lista de (z_val_ft, Reference)
z_vals_seen  = []   # para deduplicacao

TOL_Z = to_ft(0.5)  # tolerancia de deduplicacao (0.5 cm)


def z_already(z_val):
    return any(abs(z_val - zv) < TOL_Z for zv in z_vals_seen)


for wall in walls:
    orient, _, p0, p1 = classify_wall(wall)
    if orient is None:
        continue

    ref_bottom, ref_top, z_bottom, z_top = get_face_refs_up(wall)

    if ref_top is not None and z_top is not None and not z_already(z_top):
        refs_v_pairs.append((z_top, ref_top))
        z_vals_seen.append(z_top)

    if ref_bottom is not None and z_bottom is not None and not z_already(z_bottom):
        refs_v_pairs.append((z_bottom, ref_bottom))
        z_vals_seen.append(z_bottom)

# Ordena por posicao no eixo Up e filtra segmentos zero
refs_v_pairs.sort(key=lambda x: x[0])
output.print_md("**Refs verticais antes do filtro:** {}".format(len(refs_v_pairs)))
refs_v_pairs = filter_zero_segments(refs_v_pairs)
refs_v = [r for _, r in refs_v_pairs]
output.print_md("**Refs verticais apos filtro:** {}".format(len(refs_v)))


# ============================================================
# LINHAS DE COTA
# ============================================================

def build_dim_line_h(pts, offset_cm=60.0):
    if not pts:
        return None
    offset_ft = to_ft(offset_cm)

    r_vals = [dot(p, right) for p in pts]
    r_min  = min(r_vals) - to_ft(30)
    r_max  = max(r_vals) + to_ft(30)

    y_avg = sum(p.Y for p in pts) / len(pts)
    z_avg = sum(p.Z for p in pts) / len(pts)

    def make_pt(r_val):
        return XYZ(
            right.X * r_val + up.X * (z_avg + offset_ft),
            right.Y * r_val + up.Y * (z_avg + offset_ft) + y_avg * (1 - abs(right.Y)),
            right.Z * r_val + up.Z * (z_avg + offset_ft),
        )

    pt1 = make_pt(r_min)
    pt2 = make_pt(r_max)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)


def build_dim_line_v(z_vals, offset_cm=60.0):
    if not z_vals:
        return None
    offset_ft = to_ft(offset_cm)

    u_min = min(z_vals) - to_ft(30)
    u_max = max(z_vals) + to_ft(30)

    all_bb_pts = []
    for w in walls:
        bb = w.get_BoundingBox(None)
        if bb:
            all_bb_pts.append(
                XYZ((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, 0)
            )

    if not all_bb_pts:
        return None

    cx     = sum(p.X for p in all_bb_pts) / len(all_bb_pts)
    cy     = sum(p.Y for p in all_bb_pts) / len(all_bb_pts)
    base_r = dot(XYZ(cx, cy, 0), right) + offset_ft

    def make_pt(u_val):
        return XYZ(
            right.X * base_r + up.X * u_val,
            right.Y * base_r + up.Y * u_val,
            right.Z * base_r + up.Z * u_val,
        )

    pt1 = make_pt(u_min)
    pt2 = make_pt(u_max)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)


# ============================================================
# ABERTURAS (Windows / Doors)
# ============================================================

def collect_openings_in_view():
    cats = [BuiltInCategory.OST_Windows, BuiltInCategory.OST_Doors]
    openings = []
    for cat in cats:
        openings.extend(
            list(
                FilteredElementCollector(doc, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        )
    return openings


def opening_refs_lr_bt(opening):
    opt = Options()
    opt.ComputeReferences = True
    opt.View = view
    geom = opening.get_Geometry(opt)

    lr_pos = []
    lr_neg = []
    bt_pos = []
    bt_neg = []

    for g in geom:
        if not isinstance(g, Solid):
            continue
        for face in g.Faces:
            if face.Reference is None or not isinstance(face, PlanarFace):
                continue
            n   = face.FaceNormal
            d_r = dot(n, right)
            d_u = dot(n, up)

            if abs(d_r) > 0.8:
                if d_r > 0:
                    lr_pos.append((d_r, face.Reference))
                else:
                    lr_neg.append((d_r, face.Reference))

            if abs(d_u) > 0.8:
                z_face = dot(face.Origin, up)
                if d_u > 0:
                    bt_pos.append((z_face, face.Reference))
                else:
                    bt_neg.append((z_face, face.Reference))

    ref_left  = min(lr_neg, key=lambda x: x[0])[1] if lr_neg else None
    ref_right = max(lr_pos, key=lambda x: x[0])[1] if lr_pos else None

    ref_bottom = None
    ref_top    = None
    all_bt = [(z, r) for (z, r) in bt_pos] + [(z, r) for (z, r) in bt_neg]
    if all_bt:
        all_bt_sorted = sorted(all_bt, key=lambda x: x[0])
        ref_bottom = all_bt_sorted[0][1]
        ref_top    = all_bt_sorted[-1][1]

    return ref_left, ref_right, ref_bottom, ref_top


def face_origin_value_for_refs(opening, ref_a, ref_b, axis_dir):
    opt = Options()
    opt.ComputeReferences = True
    opt.View = view
    geom = opening.get_Geometry(opt)

    va = None
    vb = None

    for g in geom:
        if not isinstance(g, Solid):
            continue
        for face in g.Faces:
            if face.Reference is None or not isinstance(face, PlanarFace):
                continue
            if face.Reference != ref_a and face.Reference != ref_b:
                continue
            v = dot(face.Origin, axis_dir)
            if face.Reference == ref_a:
                va = v
            elif face.Reference == ref_b:
                vb = v

    if va is None or vb is None:
        return None
    return va, vb


def opening_range_on_axis(opening, axis_dir):
    bb = opening.get_BoundingBox(view) or opening.get_BoundingBox(None)
    if not bb:
        return None

    corners = [
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    ]

    vals = [dot(c, axis_dir) for c in corners]
    return min(vals), max(vals)


def build_opening_dim_line_h(opening, offset_cm=40.0):
    bb = opening.get_BoundingBox(view) or opening.get_BoundingBox(None)
    if not bb:
        return None

    offset_ft = to_ft(offset_cm)
    corners = [
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    ]

    r_vals = [dot(c, right) for c in corners]
    r_min  = min(r_vals) - to_ft(10)
    r_max  = max(r_vals) + to_ft(10)
    u_vals = [dot(c, up) for c in corners]
    u_avg  = sum(u_vals) / len(u_vals)
    y_avg  = (bb.Min.Y + bb.Max.Y) / 2.0

    def make_pt(r_val):
        return XYZ(
            right.X * r_val + up.X * (u_avg + offset_ft),
            right.Y * r_val + up.Y * (u_avg + offset_ft) + y_avg * (1 - abs(right.Y)),
            right.Z * r_val + up.Z * (u_avg + offset_ft),
        )

    pt1 = make_pt(r_min)
    pt2 = make_pt(r_max)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)


def build_opening_dim_line_v(opening, offset_cm=40.0):
    bb = opening.get_BoundingBox(view) or opening.get_BoundingBox(None)
    if not bb:
        return None

    offset_ft = to_ft(offset_cm)
    corners = [
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z), XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    ]

    u_vals = [dot(c, up) for c in corners]
    u_min  = min(u_vals) - to_ft(10)
    u_max  = max(u_vals) + to_ft(10)
    mid    = XYZ(
        (bb.Min.X + bb.Max.X) / 2,
        (bb.Min.Y + bb.Max.Y) / 2,
        (bb.Min.Z + bb.Max.Z) / 2,
    )
    base_r = dot(mid, right) + offset_ft

    def make_pt(u_val):
        return XYZ(
            right.X * base_r + up.X * u_val,
            right.Y * base_r + up.Y * u_val,
            right.Z * base_r + up.Z * u_val,
        )

    pt1 = make_pt(u_min)
    pt2 = make_pt(u_max)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)


def safe_add_opening_dimension(opening, ref_a, ref_b, axis_type):
    if ref_a is None or ref_b is None:
        return False

    axis_dir = right if axis_type == "h" else up
    pair     = face_origin_value_for_refs(opening, ref_a, ref_b, axis_dir)

    if pair is not None:
        dim_len = abs(pair[1] - pair[0])
    else:
        rr      = opening_range_on_axis(opening, axis_dir)
        dim_len = abs(rr[1] - rr[0]) if rr else 0.0

    if dim_len <= TOL_DIM_ZERO:
        return False

    dim_line = (
        build_opening_dim_line_h(opening, offset_cm=40.0)
        if axis_type == "h"
        else build_opening_dim_line_v(opening, offset_cm=40.0)
    )

    if dim_line is None:
        return False

    ref_array = ReferenceArray()
    ref_array.Append(ref_a)
    ref_array.Append(ref_b)

    new_dim = doc.Create.NewDimension(view, dim_line, ref_array)
    if dim_type:
        try:
            new_dim.DimensionType = dim_type
        except Exception:
            pass
    return True


openings = collect_openings_in_view()
output.print_md("**Aberturas na vista:** {}".format(len(openings)))

# ============================================================
# Tipo de cota
# ============================================================
dim_type = None
try:
    dtypes = list(FilteredElementCollector(doc).OfClass(DimensionType).ToElements())
    if dtypes:
        dim_type = dtypes[0]
except Exception:
    pass

# ============================================================
# CRIAR COTAS
# ============================================================
dims_created = 0
errors       = 0

with Transaction(doc, "Cotas Automaticas v2.6") as t:
    t.Start()
    try:

        # ── Horizontal (paredes) ──────────────────────────────
        if len(refs_h) >= 2:
            dim_line_h = build_dim_line_h(pts_h, offset_cm=60.0)
            if dim_line_h:
                ref_array_h = ReferenceArray()
                for r in refs_h:
                    ref_array_h.Append(r)
                try:
                    new_dim = doc.Create.NewDimension(view, dim_line_h, ref_array_h)
                    if dim_type:
                        try:
                            new_dim.DimensionType = dim_type
                        except Exception:
                            pass
                    dims_created += 1
                    output.print_md(
                        "Cota horizontal criada com **{}** referencias".format(
                            ref_array_h.Size
                        )
                    )
                except Exception as e:
                    errors += 1
                    output.print_md("Erro cota horizontal: `{}`".format(str(e)))
            else:
                output.print_md("Linha horizontal invalida.")
        else:
            output.print_md("Menos de 2 referencias horizontais.")

        # ── Vertical (paredes) ────────────────────────────────
        if len(refs_v) >= 2:
            z_vals_for_line = [v for v, _ in refs_v_pairs]
            dim_line_v = build_dim_line_v(z_vals_for_line, offset_cm=60.0)
            if dim_line_v:
                ref_array_v = ReferenceArray()
                for r in refs_v:
                    ref_array_v.Append(r)
                try:
                    new_dim = doc.Create.NewDimension(view, dim_line_v, ref_array_v)
                    if dim_type:
                        try:
                            new_dim.DimensionType = dim_type
                        except Exception:
                            pass
                    dims_created += 1
                    output.print_md(
                        "Cota vertical criada com **{}** referencias".format(
                            ref_array_v.Size
                        )
                    )
                except Exception as e:
                    errors += 1
                    output.print_md("Erro cota vertical: `{}`".format(str(e)))
            else:
                output.print_md("Linha vertical invalida.")
        else:
            output.print_md(
                "Menos de 2 referencias verticais. z_vals: {}".format(
                    [round(v * 30.48, 2) for v in z_vals_seen]
                )
            )

        # ── Vaos (largura + altura) ───────────────────────────
        created_openings = 0
        for op in openings:
            try:
                ref_left, ref_right_op, ref_bottom, ref_top = opening_refs_lr_bt(op)

                if ref_left is not None and ref_right_op is not None:
                    if safe_add_opening_dimension(op, ref_left, ref_right_op, "h"):
                        created_openings += 1

                if ref_bottom is not None and ref_top is not None:
                    if safe_add_opening_dimension(op, ref_bottom, ref_top, "v"):
                        created_openings += 1

            except Exception as e:
                errors += 1
                output.print_md("Erro abertura: `{}`".format(str(e)))

        if created_openings:
            dims_created += created_openings
            output.print_md("Cotas de vaos criadas: **{}**".format(created_openings))
        else:
            output.print_md("Nenhuma cota de vao criada.")

        t.Commit()

    except Exception as e:
        t.RollBack()
        forms.alert("Erro na transacao:\n{}".format(str(e)))

# ============================================================
# RESUMO
# ============================================================
output.print_md("---")
output.print_md("## Concluido! (v2.6)")
output.print_md("- **Cotas criadas:** {}".format(dims_created))
output.print_md("- **Erros:** {}".format(errors))
output.print_md("- **Paredes processadas:** {}".format(len(walls)))
