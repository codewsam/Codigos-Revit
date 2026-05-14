# -*- coding: utf-8 -*-
__title__ = "Cotas Automáticas de Famílias"
__version__ = "2.1"
__doc__ = (
    "Insere cotas automáticas nas paredes da vista ativa (Elevação ou Corte). "
    "Detecta as direções da vista (Right/Up), coleta referências de faces laterais "
    "de todas as paredes visíveis e cria cotas lineares agrupadas por eixo. "
    "Compatível com o projeto AURORA MAIS VIVER."
)

# ============================================================
# IMPORTS
# ============================================================
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Options,
    LocationCurve, Solid, PlanarFace, ReferenceArray, Line, XYZ,
    Transaction, DimensionType, UnitUtils
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

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()

# ============================================================
# ETAPA 1 — Vista ativa
# ============================================================
view = doc.ActiveView

output.print_md("## Cotas Automaticas — **{}**".format(view.Name))

if view.ViewType.ToString() not in ("Elevation", "Section"):
    forms.alert(
        "Esta ferramenta funciona em vistas de Elevacao ou Corte.\n"
        "Vista atual: {} ({})".format(view.Name, view.ViewType),
        exitscript=True
    )

# Direcoes da vista
right    = view.RightDirection
up       = view.UpDirection
view_dir = view.ViewDirection

output.print_md("Right: `{:.2f},{:.2f},{:.2f}` | Up: `{:.2f},{:.2f},{:.2f}`".format(
    right.X, right.Y, right.Z, up.X, up.Y, up.Z))

# ============================================================
# ETAPA 2 — Buscar paredes na vista ativa
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
    return math.sqrt(dx*dx + dy*dy + dz*dz) < tol_ft

def get_wall_dir(wall):
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return None, None, None
    c = loc.Curve
    p0 = c.GetEndPoint(0)
    p1 = c.GetEndPoint(1)
    raw = p1 - p0
    if raw.GetLength() < 1e-9:
        return None, None, None
    return raw.Normalize(), p0, p1

# ============================================================
# ETAPA 8 — Obter referencias de faces laterais (normal Right)
# ============================================================
def get_face_refs_right(wall):
    opt = Options()
    opt.ComputeReferences = True
    opt.View = view
    geom = wall.get_Geometry(opt)
    ref_neg = None
    ref_pos = None
    for g in geom:
        if isinstance(g, Solid):
            for face in g.Faces:
                if face.Reference is None:
                    continue
                # CylindricalFace e outras faces curvas nao tem FaceNormal
                if not isinstance(face, PlanarFace):
                    continue
                n = face.FaceNormal
                d_r = dot(n, right)
                if d_r > 0.8 and ref_pos is None:
                    ref_pos = face.Reference
                elif d_r < -0.8 and ref_neg is None:
                    ref_neg = face.Reference
    return ref_neg, ref_pos

# ============================================================
# ETAPA 11 — Classificar paredes
# ============================================================
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
    else:
        return "D", d, p0, p1

# ============================================================
# ETAPA 9 — Coletar refs horizontais (faces com normal Right)
# de todas as paredes, filtrando coincidentes
# ============================================================
refs_h = []
pts_h  = []

for wall in walls:
    orient, d, p0, p1 = classify_wall(wall)
    if orient is None:
        continue

    ref_neg, ref_pos = get_face_refs_right(wall)

    pt_left  = p0 if dot(p0, right) <= dot(p1, right) else p1
    pt_right_w = p1 if pt_left == p0 else p0

    if ref_neg is not None:
        if not any(pts_close(pt_left, ep) for ep in pts_h):
            refs_h.append(ref_neg)
            pts_h.append(pt_left)

    if ref_pos is not None:
        if not any(pts_close(pt_right_w, ep) for ep in pts_h):
            refs_h.append(ref_pos)
            pts_h.append(pt_right_w)

output.print_md("**Referencias horizontais:** {}".format(len(refs_h)))

# ============================================================
# ETAPA 8b — Obter referencias de faces laterais (normal Up)
# ============================================================
def get_face_refs_up(wall):
    opt = Options()
    opt.ComputeReferences = True
    opt.View = view
    geom = wall.get_Geometry(opt)
    ref_neg = None
    ref_pos = None
    for g in geom:
        if isinstance(g, Solid):
            for face in g.Faces:
                if face.Reference is None:
                    continue
                # CylindricalFace e outras faces curvas nao tem FaceNormal
                if not isinstance(face, PlanarFace):
                    continue
                n = face.FaceNormal
                d_u = dot(n, up)
                if d_u > 0.8 and ref_pos is None:
                    ref_pos = face.Reference
                elif d_u < -0.8 and ref_neg is None:
                    ref_neg = face.Reference
    return ref_neg, ref_pos

# ============================================================
# ETAPA 9b — Coletar refs verticais (faces com normal Up)
# de todas as paredes, filtrando coincidentes
# ============================================================
refs_v = []
pts_v  = []

for wall in walls:
    orient, d, p0, p1 = classify_wall(wall)
    if orient is None:
        continue

    # Para cotas verticais, o agrupamento final é feito pelas faces alinhadas ao vetor Up.
    ref_neg, ref_pos = get_face_refs_up(wall)

    # Deduplicacao no eixo Right (na "linha" da cota vertical)
    pt_low   = p0 if dot(p0, right) <= dot(p1, right) else p1
    pt_high  = p1 if pt_low == p0 else p0

    if ref_neg is not None:
        if not any(pts_close(pt_low, ep) for ep in pts_v):
            refs_v.append(ref_neg)
            pts_v.append(pt_low)

    if ref_pos is not None:
        if not any(pts_close(pt_high, ep) for ep in pts_v):
            refs_v.append(ref_pos)
            pts_v.append(pt_high)

output.print_md("**Referencias verticais:** {}".format(len(refs_v)))

# ============================================================
# ETAPA 15/16 — Construir linha de cota horizontal
# Linha ao longo de Right, posicionada no Z medio das paredes
# ============================================================
def build_dim_line_h(pts, offset_cm=60.0):
    if not pts:
        return None
    offset_ft = to_ft(offset_cm)
    r_vals = [dot(p, right) for p in pts]
    r_min  = min(r_vals) - to_ft(30)
    r_max  = max(r_vals) + to_ft(30)

    # Posicao base: media dos pontos
    y_avg = sum(p.Y for p in pts) / len(pts)
    z_avg = sum(p.Z for p in pts) / len(pts)

    # Linha paralela ao Right, deslocada +offset na direcao Up
    def make_pt(r_val):
        return XYZ(
            right.X * r_val + up.X * (z_avg + offset_ft),
            right.Y * r_val + up.Y * (z_avg + offset_ft) + y_avg * (1 - abs(right.Y)),
            right.Z * r_val + up.Z * (z_avg + offset_ft)
        )

    pt1 = make_pt(r_min)
    pt2 = make_pt(r_max)

    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)

# ============================================================
# ETAPA 15c/16c — Construir linha de cota vertical
# Linha ao longo de Up, posicionada no Y medio das paredes
# deslocada +offset na direcao Right
# ============================================================
def build_dim_line_v(pts, offset_cm=60.0):
    if not pts:
        return None
    offset_ft = to_ft(offset_cm)

    u_vals = [dot(p, up) for p in pts]
    u_min  = min(u_vals) - to_ft(30)
    u_max  = max(u_vals) + to_ft(30)

    # Posicao base: media dos pontos
    x_avg = sum(p.X for p in pts) / len(pts)
    y_avg = sum(p.Y for p in pts) / len(pts)
    z_avg = sum(p.Z for p in pts) / len(pts)

    # Linha paralela ao Up, deslocada +offset na direcao Right
    def make_pt(u_val):
        return XYZ(
            right.X * (dot(XYZ(x_avg, y_avg, z_avg), right) + offset_ft) + up.X * u_val,
            right.Y * (dot(XYZ(x_avg, y_avg, z_avg), right) + offset_ft) + up.Y * u_val,
            right.Z * (dot(XYZ(x_avg, y_avg, z_avg), right) + offset_ft) + up.Z * u_val
        )

    # Como o calculo acima pode introduzir pequenas inconsistencias,
    # fazemos um fallback mais simples caso a linha fique degenerada.
    pt1 = make_pt(u_min)
    pt2 = make_pt(u_max)

    if pt1.DistanceTo(pt2) < 1e-6:
        # fallback: deslocar pelo vetor Right diretamente usando um ponto medio
        mid = XYZ(x_avg, y_avg, z_avg)
        mid_off = XYZ(mid.X + right.X * offset_ft, mid.Y + right.Y * offset_ft, mid.Z + right.Z * offset_ft)
        # reconstruir pt1/pt2 no plano usando mid_off como origem
        base_u = dot(mid_off, up)
        pt1 = XYZ(mid_off.X + up.X * (u_min - base_u), mid_off.Y + up.Y * (u_min - base_u), mid_off.Z + up.Z * (u_min - base_u))
        pt2 = XYZ(mid_off.X + up.X * (u_max - base_u), mid_off.Y + up.Y * (u_max - base_u), mid_off.Z + up.Z * (u_max - base_u))

    if pt1.DistanceTo(pt2) < 1e-6:
        return None

    return Line.CreateBound(pt1, pt2)

# ============================================================
# ETAPA 18 — Criar cotas com NewDimension
# ============================================================
dim_type = None
try:
    dtypes = list(FilteredElementCollector(doc).OfClass(DimensionType).ToElements())
    if dtypes:
        dim_type = dtypes[0]
except:
    pass

dims_created = 0
errors       = 0

with Transaction(doc, "Cotas Automaticas de Paredes") as t:
    t.Start()
    try:
        # -------------------------
        # Cotas horizontais (Right)
        # -------------------------
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
                        except:
                            pass
                    dims_created += 1
                    output.print_md("Cota horizontal criada com {} referencias".format(ref_array_h.Size))
                except Exception as e:
                    errors += 1
                    output.print_md("Erro cota horizontal: {}".format(str(e)))
            else:
                output.print_md("Linha horizontal invalida.")
        else:
            output.print_md("Menos de 2 referencias horizontais, pulando.")

        # -------------------------
        # Cotas verticais (Up)
        # -------------------------
        if len(refs_v) >= 2:
            dim_line_v = build_dim_line_v(pts_v, offset_cm=60.0)
            if dim_line_v:
                ref_array_v = ReferenceArray()
                for r in refs_v:
                    ref_array_v.Append(r)
                try:
                    new_dim = doc.Create.NewDimension(view, dim_line_v, ref_array_v)
                    if dim_type:
                        try:
                            new_dim.DimensionType = dim_type
                        except:
                            pass
                    dims_created += 1
                    output.print_md("Cota vertical criada com {} referencias".format(ref_array_v.Size))
                except Exception as e:
                    errors += 1
                    output.print_md("Erro cota vertical: {}".format(str(e)))
            else:
                output.print_md("Linha vertical invalida.")
        else:
            output.print_md("Menos de 2 referencias verticais, pulando.")

        t.Commit()

    except Exception as e:
        t.RollBack()
        forms.alert("Erro na transacao:\n{}".format(str(e)))

# ============================================================
# RESUMO
# ============================================================
output.print_md("---")
output.print_md("## Concluido!")
output.print_md("- **Cotas criadas:** {}".format(dims_created))
output.print_md("- **Erros:** {}".format(errors))
output.print_md("- **Paredes processadas:** {}".format(len(walls)))
