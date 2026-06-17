#!/usr/bin/env python3
"""
diagnostico_zip_pf2025.py  v3
Ejecutar DENTRO de PowerFactory: Tools → Python → Execute Script
Resultado: Output Window de PF  +  diagnostico_zip_resultado.txt
"""

import powerfactory
app = powerfactory.GetApplication()

import os
_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    "diagnostico_zip_resultado.txt")
_lineas = []

def p(txt=""):
    app.PrintPlain(str(txt))
    _lineas.append(str(txt))

def guardar():
    with open(_OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(_lineas))
    app.PrintPlain(f"  >> Guardado en: {_OUT}")

def probar_attrs(obj, nombres, prefijo="  "):
    """Prueba lista de nombres de atributo en obj. Devuelve dict de los que existen."""
    encontrados = {}
    for attr in nombres:
        try:
            val = obj.GetAttribute(attr)
            encontrados[attr] = val
            p(f"{prefijo}{attr:<32} = {val!r}   [EXISTE]")
        except Exception:
            pass  # no existe, ignorar
    return encontrados

# ─────────────────────────────────────────────────────────────────────────────
lods = app.GetCalcRelevantObjects("*.ElmLod") or []
p("=" * 66)
p("  DIAGNOSTICO ZIP — ElmLod / TypLod — PF 2025  (v3)")
p("=" * 66)

if not lods:
    p("  No se encontraron cargas ElmLod. Verifica proyecto/caso activo.")
    guardar()
else:
    ld = lods[0]
    p(f"  Carga  : {ld.loc_name}  ({len(lods)} cargas totales)")
    try:
        p(f"  Clase  : {ld.GetClassName()}")
    except Exception:
        pass
    p()

    # ── A. Candidatos ElmLod ──────────────────────────────────────────────────
    p("  [A] Atributos ZIP candidatos en ElmLod:")
    ELM_CANDS = [
        # ZIP PF 2024-2025
        "pzload","pbcload","qzload","qbcload",
        # PF arrays
        "cv_p(0)","cv_p(1)","cv_p(2)","cv_q(0)","cv_q(1)","cv_q(2)",
        # PF clasico
        "aP","bP","cP","aQ","bQ","cQ",
        "kpu_low","kqu_low","kpu","kqu","kpf","kqf",
        # flags / modo
        "iZip","izip","iZIP","i_zip","i_mode","mode_zip",
        "iload","loa_dep_typ","ilodtyp",
        # otro
        "pload_a","pload_b","pload_c","qload_a","qload_b","qload_c",
        "Pnorm(0)","Pnorm(1)","Pnorm(2)","Qnorm(0)","Qnorm(1)","Qnorm(2)",
        "elod_a","elod_b","elod_c",
        "zip_aP","zip_bP","zip_cP",
        # modo entrada
        "mode_inp","imode",
        # PF 2025 posibles
        "czip","bzip","azip",
        "dP(0)","dP(1)","dP(2)","dQ(0)","dQ(1)","dQ(2)",
    ]
    elm_ok = probar_attrs(ld, ELM_CANDS)
    if not elm_ok:
        p("    (ninguno encontrado en ElmLod)")
    p()

    # ── B. Candidatos TypLod ──────────────────────────────────────────────────
    p("  [B] Atributos ZIP candidatos en TypLod:")
    typ = None
    for atr in ("typ_id","TypLoad","pTypLoad"):
        try:
            t = ld.GetAttribute(atr)
            if t is not None:
                typ = t
                break
        except Exception:
            pass

    if typ is None:
        p("    Sin TypLod asignado.")
    else:
        try:
            p(f"    TypLod : {typ.loc_name}")
            p(f"    Clase  : {typ.GetClassName()}")
        except Exception:
            p(f"    TypLod : {typ.loc_name}")
        p()

        TYP_CANDS = [
            # PF clasico (0-100 %)
            "aP","bP","cP","aQ","bQ","cQ",
            # PF con guion bajo
            "a_P","b_P","c_P","a_Q","b_Q","c_Q",
            # PF 2024-2025 posibles
            "pzload","pbcload","qzload","qbcload",
            "kp_zip","kq_zip","kpu","kqu","kpf","kqf",
            "iZip","izip","i_zip",
            "cv_p(0)","cv_p(1)","cv_p(2)","cv_q(0)","cv_q(1)","cv_q(2)",
            "Pnorm(0)","Pnorm(1)","Pnorm(2)","Qnorm(0)","Qnorm(1)","Qnorm(2)",
            "dP(0)","dP(1)","dP(2)","dQ(0)","dQ(1)","dQ(2)",
            "elod_a","elod_b","elod_c",
            "czip","bzip","azip",
            # genericos voltage-dependent
            "kpu_low","kqu_low","kpu_high","kqu_high",
            "V0pu","Vmin","Vmax",
            # otros campos comunes de TypLod
            "cosn","tanfi","pf_recap",
            "systp","iZip","frnom",
        ]
        typ_ok = probar_attrs(typ, TYP_CANDS, prefijo="    ")
        if not typ_ok:
            p("    (ninguno encontrado en TypLod)")

        # ── C. Lectura bruta de TypLod (DataObject attributes) ─────────────
        p()
        p("  [C] Lectura bruta — campos TypLod conocidos PF (GetAttr directo):")
        # En PF los DataObject exponen atributos como propiedades Python directas
        PROP_NAMES = [
            "aP","bP","cP","aQ","bQ","cQ",
            "kpu","kqu","kpf","kqf",
            "iZip","frnom","cosn","systp",
        ]
        for pn in PROP_NAMES:
            try:
                val = getattr(typ, pn, "__NO_ATTR__")
                if val != "__NO_ATTR__":
                    p(f"    typ.{pn:<28} = {val!r}   [PROPIEDAD PYTHON]")
            except Exception:
                pass

    # ── D. Muestra TODOS los atributos del ElmLod (metodo alternativo) ───────
    p()
    p("  [D] Contenido dir(ElmLod) — propiedades Python (filtradas):")
    KEYWORDS_D = ["zip","zload","bcload","pload","qload","kpu","kqu",
                  "kpf","kqf","pnorm","qnorm","cv_p","cv_q","izip",
                  "elod","dP","dQ","mode_zip","ilodtyp"]
    try:
        attrs_dir = [a for a in dir(ld)
                     if not a.startswith("_")
                     and any(k in a.lower() for k in KEYWORDS_D)]
        if attrs_dir:
            for a in sorted(attrs_dir):
                try:
                    v = getattr(ld, a)
                    if not callable(v):
                        p(f"  {a:<32} = {v!r}")
                except Exception:
                    pass
        else:
            p("  (ninguno con esos keywords en dir())")
    except Exception as e:
        p(f"  Error en dir(): {e}")

    # ── E. RESUMEN ────────────────────────────────────────────────────────────
    p()
    p("  [E] RESUMEN:")
    all_ok = dict(**elm_ok)
    if typ is not None:
        try:
            all_ok.update(typ_ok)
        except Exception:
            pass
    if all_ok:
        p(f"  Atributos encontrados: {list(all_ok.keys())}")
        nums = [v for v in all_ok.values() if isinstance(v,(int,float))]
        if nums:
            if any(v > 1.5 for v in nums):
                p("  Escala probable: PORCENTAJE 0-100")
            else:
                p("  Escala probable: FRACCION 0-1")
    else:
        p("  No se encontraron atributos ZIP en ElmLod ni TypLod.")
        p("  -> ZIP posiblemente gestionado via ComLdf 'Voltage Dependency'")
        p("     o via ElmGenLoad / TypGenLoad en lugar de ElmLod/TypLod.")

    p()
    guardar()
