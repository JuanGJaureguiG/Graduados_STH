"""
build_data.py — Tablero de Graduados · UNIMINUTO Sede Tolima-Huila
Genera data.json (agregado, sin identificadores personales) a partir del
export de SAP "Graduados_SAP_<fecha>.xlsx", hoja "Listado Estudiantes".

Uso:  python build_data.py Graduados_SAP_11-5-2026.xlsx 2026-05-11
"""
import json, os, sys
import pandas as pd

SRC   = sys.argv[1] if len(sys.argv) > 1 else "Graduados_SAP_11-5-2026.xlsx"
CORTE = sys.argv[2] if len(sys.argv) > 2 else "2026-05-11"
SHEET = "Listado Estudiantes"

# Centros Universitarios fuera del alcance de la sede (regla de todos los tableros)
EXCLUIDOS = {"Cu Fresno", "Ct Puerto Boyacá", "Cu Líbano", "Cu Cajamarca", "Cu Mariquita",
             "Cu Florencia", "Cu Mocoa", "Cu Planadas"}  # Planadas no registra graduados
CU_NOMBRE = {"Cu Ibagué": "Ibagué", "Cu Neiva": "Neiva", "Cu Pitalito": "Pitalito",
             "Cu Garzón": "Garzón", "Cu Lérida": "Lérida", "Cu La Dorada": "La Dorada"}

# Unificación de denominaciones de programa (mayúsculas / tildes / variantes de SAP)
PROG_FIX = {
    "Trabajo social": "Trabajo Social",
    "Especialización En Gerencia De Proyectos": "Especialización en Gerencia de Proyectos",
    "ESPECIALIZACIÓN EN GERENCIA EDUCATIVA": "Especialización en Gerencia Educativa",
    "Especialización en gerencia de riesgos laborales seguridad y salud en el trabajo":
        "Especialización en Gerencia en Riesgos Laborales, Seguridad y Salud en el Trabajo",
    "Especialización en Auditoría Integral y Revisoría  Fiscal":
        "Especialización en Auditoría Integral y Revisoría Fiscal",
    "Especialización Psicología Organizacional": "Especialización en Psicología Organizacional",
    "Especialización en Auditoria Forense": "Especialización en Auditoría Forense",
    "Licenciatura En Educación Artística": "Licenciatura en Educación Artística",
    "Licenciatura En Educación Básica Con Énfasis En Ciencias Naturales Y Educación Ambiental":
        "Licenciatura en Educación Básica con Énfasis en Ciencias Naturales y Educación Ambiental",
}
NIVEL_FIX = {"Especializacion": "Especialización", "Tecnico Profesional": "Técnico Profesional"}
MOD_FIX = {"Hibrida": "Híbrida"}

df = pd.read_excel(SRC, sheet_name=SHEET)
df.columns = ["genero", "edad", "estrato", "cu", "modalidad", "programa", "area", "nivel", "semestre", "anio"]
n_total = len(df)

df = df.dropna(subset=["cu", "anio", "semestre"])
n_vacios = n_total - len(df)
excl = df[~df["cu"].isin(CU_NOMBRE)]
excl_detalle = excl["cu"].value_counts().to_dict()
df = df[df["cu"].isin(CU_NOMBRE)].copy()

df["cu"] = df["cu"].map(CU_NOMBRE)
df["programa"] = df["programa"].fillna("Sin programa registrado").str.strip().replace(PROG_FIX)
df["nivel"] = df["nivel"].replace(NIVEL_FIX).fillna("Sin dato")
df["modalidad"] = df["modalidad"].replace(MOD_FIX).fillna("Sin dato")
df["area"] = df["area"].fillna("Sin dato")
df["genero"] = df["genero"].fillna("Sin dato")
df["edad"] = df["edad"].fillna("Sin dato")
df["estrato"] = df["estrato"].map(lambda v: "Sin dato" if pd.isna(v) else f"Estrato {int(v)}")
df["anio"] = df["anio"].astype(int)
df["sem"] = df["semestre"].str.extract(r"(\d)").astype(int)

DIMS = ["cu", "modalidad", "programa", "area", "nivel", "genero", "edad", "estrato"]
ORDEN = {
    "edad": ["18 a 30", "31 a 35", "36 a 40", "41 a 45", "Más de 45", "Sin dato"],
    "estrato": [f"Estrato {i}" for i in range(1, 7)] + ["Sin dato"],
    "nivel": ["Pregrado", "Especialización", "Técnico Profesional", "Sin dato"],
}
dims = {}
for d in DIMS:
    vals = df[d].value_counts().index.tolist()          # por frecuencia
    if d in ORDEN:
        vals = [v for v in ORDEN[d] if v in vals] + [v for v in vals if v not in ORDEN[d]]
    dims[d] = vals

g = df.groupby(["anio", "sem"] + DIMS).size().reset_index(name="n")
rows = [[int(r.anio), int(r.sem)] + [dims[d].index(getattr(r, d)) for d in DIMS] + [int(r.n)]
        for r in g.itertuples(index=False)]

# programa -> nivel y área (fijo por programa, para agrupar en la tabla)
prog_meta = (df.groupby("programa")[["nivel", "area"]].agg(lambda s: s.mode().iat[0]).to_dict("index"))

anios = sorted(df["anio"].unique().tolist())
periodos = sorted({(int(a), int(s)) for a, s in df[["anio", "sem"]].itertuples(index=False)})
ultimo = periodos[-1]
out = {
    "corte": CORTE,
    "fuente": os.path.basename(SRC),
    "anios": anios,
    "periodos": [f"{a}-{s}" for a, s in periodos],
    "anio_parcial": [ultimo[0]] if ultimo[1] == 1 else [],
    "dims": dims,
    "cols": ["anio", "sem"] + DIMS + ["n"],
    "rows": rows,
    "prog_meta": prog_meta,
    "calidad": {
        "registros_archivo": n_total,
        "registros_vacios": n_vacios,
        "excluidos_por_centro": {k.replace("Cu ", "").replace("Ct ", ""): int(v) for k, v in excl_detalle.items()},
        "registros_tablero": int(df.shape[0]),
    },
}
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
print(f"data.json · {len(rows)} filas agregadas · {df.shape[0]} graduados · periodos {out['periodos'][0]}–{out['periodos'][-1]}")
print("Excluidos:", out["calidad"]["excluidos_por_centro"], "· vacíos:", n_vacios)
