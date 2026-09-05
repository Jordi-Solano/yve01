"""FASE E · La fila hotelera del DRR: tres estados y medias PONDERADAS.

Cuatro hoteles elegidos para que una media plana y una ponderada den
resultados claramente distintos. Si fueran parecidos, el fallo de ponderacion
no se veria y la prueba pasaria sin demostrar nada.

  Costa Azul   grande  · 12.000 noches · ocupacion 60% · GOP medido
  Plaza Mayor  pequeño ·    600 noches · ocupacion 95% · GOP medido
  Ribera       mediano ·  3.600 noches · ocupacion 75% · SIN GOP -> N/D
  Faro         no ha subido DRR                        -> estado sin_drr

Ocupacion del grupo:
    plana      = (60 + 95 + 75) / 3               = 76.7%   <- mentira
    ponderada  = 10.470 / 16.200                  = 64.6%   <- lo correcto

Y los dos denominadores son distintos a proposito: 3 hoteles tienen DRR, pero
solo 2 traen un GOP agregable. Un solo "sobre X de N" para todo seria falso.
"""
import json
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

TENANT = "fase-e-test"
os.environ["YVE_TENANT"] = TENANT
os.environ.pop("YVE_HOTEL", None)

import tenant_dirs                                     # noqa: E402

# Los DRR de prueba se generan aqui mismo (antes venian de /home/claude/drr,
# una carpeta de otro sandbox: el test no podia correr en ningun otro sitio).
import tempfile as _tf
FIXTURES = _tf.mkdtemp(prefix="drr_fase_e_")
os.environ["YVE_FIXTURES_DRR"] = FIXTURES
import fixtures_drr_fase_e as _fx                       # noqa: E402
_fx.main()
HOTELES = [
    ("HCOSTA1", "Hotel Costa Azul",  "DRR-COSTA-AZUL.xlsm"),
    ("HPLAZA2", "Hotel Plaza Mayor", "DRR-PLAZA-MAYOR.xlsm"),
    ("HRIBER3", "Hotel Ribera",      "DRR-RIBERA.xlsm"),
    ("HFARO4",  "Hotel Faro",        None),             # no sube nada
]


def _montar():
    base = os.path.join(BASE, "tenants", TENANT)
    shutil.rmtree(base, ignore_errors=True)
    tenant_dirs.tenant_base()
    datos = tenant_dirs.datos_dir()

    json.dump([{"id": h, "nombre": n, "activo": True} for h, n, _ in HOTELES],
              open(os.path.join(datos, "hoteles.json"), "w", encoding="utf-8"))

    # Cada DRR se procesa con SU hotel activo: el hotel viaja en el nombre del
    # fichero de salida, no en una columna. Sin censo, el sufijo sale vacio y
    # los tres se pisan — pasa de verdad, me paso montando esto.
    for hid, _nombre, fichero in HOTELES:
        if not fichero:
            continue
        entorno = dict(os.environ, YVE_TENANT=TENANT, YVE_HOTEL=hid)
        r = subprocess.run([sys.executable, os.path.join(BASE, "lector_drr.py"),
                            os.path.join(FIXTURES, fichero)],
                           capture_output=True, text=True, cwd=BASE, env=entorno, timeout=180)
        assert r.returncode == 0, f"lector_drr fallo con {fichero}: {r.stderr[-400:]}"

    generados = sorted(os.listdir(tenant_dirs.reportes_dir()))
    drrs = [f for f in generados if f.startswith("drr_procesado_")]
    assert len(drrs) == 3, f"esperaba 3 DRR, uno por hotel, y hay {len(drrs)}: {drrs}"
    return base


def main():
    base = _montar()
    fallos = []

    import agregador_grupo
    ag = agregador_grupo.agregado()
    por_nombre = {f["nombre"]: f for f in ag["hoteles"]}
    h = ag["hotelero"]

    def comprobar(etiqueta, esperado, obtenido, tol=0.15):
        if obtenido is None or abs(float(obtenido) - esperado) > tol:
            fallos.append(f"{etiqueta}: esperaba {esperado}, salio {obtenido}")

    # ── 1. Los tres estados ───────────────────────────────────────────────
    for nombre, estado in (("Hotel Costa Azul", "con_drr"),
                           ("Hotel Plaza Mayor", "con_drr"),
                           ("Hotel Ribera", "con_drr"),
                           ("Hotel Faro", "sin_drr")):
        real = por_nombre[nombre]["drr"]["estado"]
        if real != estado:
            fallos.append(f"{nombre}: estado {real!r}, esperaba {estado!r}")

    # El hotel sin DRR no vale CERO: sus metricas son None, no 0. Un cero
    # entraria en las medias y hundiria el grupo.
    faro = por_nombre["Hotel Faro"]["drr"]
    for campo in ("ocupacion_pct", "adr", "revpar", "gop"):
        if faro[campo] is not None:
            fallos.append(f"Faro sin DRR deberia tener {campo}=None y tiene {faro[campo]}")

    # ── 2. Ribera tiene DRR pero su GOP sale N/D (fase D dentro de la E) ──
    rib = por_nombre["Hotel Ribera"]["drr"]
    if rib["gop"] is not None:
        fallos.append(f"Ribera no trae GOP en el fichero y sale {rib['gop']}")
    if rib["ocupacion_pct"] is None:
        fallos.append("Ribera SI trae ocupacion y sale None")

    # ── 3. Las medias del grupo, ponderadas ──────────────────────────────
    comprobar("ocupacion ponderada (10.470/16.200)", 64.6, h["ocupacion_pct"])
    comprobar("ADR ponderado (1.415.700/10.470)",   135.2, h["adr"], tol=0.6)
    comprobar("RevPAR ponderado (1.415.700/16.200)", 87.4, h["revpar"], tol=0.6)
    comprobar("GOP% ponderado (383.010/1.309.500)",  29.2, h["gop_pct"], tol=0.2)

    # Y que NO son las planas, que es el fallo que veniamos a matar.
    for etiqueta, plana, real in (("ocupacion", 76.7, h["ocupacion_pct"]),
                                  ("ADR",      163.3, h["adr"]),
                                  ("RevPAR",   130.5, h["revpar"]),
                                  ("GOP%",      33.0, h["gop_pct"])):
        if real is not None and abs(float(real) - plana) < 1.0:
            fallos.append(f"{etiqueta} del grupo ({real}) coincide con la media PLANA "
                          f"({plana}): no se esta ponderando")

    # ── 4. Los denominadores, y que son DOS ──────────────────────────────
    if h["n_hoteles"] != 4: fallos.append(f"n_hoteles {h['n_hoteles']}, esperaba 4")
    if h["con_datos"] != 3: fallos.append(f"con_datos {h['con_datos']}, esperaba 3")
    if h["sin_drr"]   != 1: fallos.append(f"sin_drr {h['sin_drr']}, esperaba 1")
    if h["gop_sobre"] != 2:
        fallos.append(f"gop_sobre {h['gop_sobre']}, esperaba 2 (Ribera no trae GOP)")

    # ── 5. La fila financiera sigue cuadrando ────────────────────────────
    if not ag["cuadra"]:
        fallos.append("la fila financiera ha dejado de cuadrar")

    # ── Informe ──────────────────────────────────────────────────────────
    print("\n── Fase E · fila hotelera ────────────────────────────────────")
    for f in ag["hoteles"]:
        d = f["drr"]
        est = {"con_drr": "con DRR", "drr_viejo": "DRR viejo", "sin_drr": "SIN DRR"}[d["estado"]]
        print(f"  {f['nombre']:<20} {est:<10} "
              f"ocup {str(d['ocupacion_pct'] or '—'):>6}  ADR {str(d['adr'] or '—'):>7}  "
              f"GOP {str(d['gop'] or 'N/D'):>10}  ({d['gop_procedencia'] or '—'})")
    print(f"\n  GRUPO ponderado   ocup {h['ocupacion_pct']}%  ADR {h['adr']}  "
          f"RevPAR {h['revpar']}  GOP% {h['gop_pct']}%")
    print(f"  sobre {h['con_datos']} de {h['n_hoteles']} hoteles "
          f"(GOP sobre {h['gop_sobre']}) · {h['sin_drr']} sin DRR · {h['dias_oob']} dias OOB")
    print(f"  medias planas serian: ocup 76.7%  ADR 163.3  RevPAR 130.5  GOP% 33.0%")

    shutil.rmtree(base, ignore_errors=True)

    if fallos:
        print("\n FALLOS:")
        for f in fallos:
            print("   ·", f)
        return 1
    print("\n OK · tres estados, medias ponderadas y los dos denominadores\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
