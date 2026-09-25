# -*- coding: utf-8 -*-
"""b99 — los registros del AR aguantan escrituras a la vez (candados.py).

Con el comando del render.yaml (`--workers 1 --threads 8`) dos peticiones pueden
leer el mismo fichero, cambiar cada una lo suyo y escribir: la segunda borra lo de
la primera. Hasta el 25 sep produccion corria con UN hilo y no pasaba. Aqui se
MIDE con hilos de verdad, todos arrancando a la vez (Barrier), y con una espera
metida entre leer y escribir para que la carrera salga siempre (sin candado se
pierde casi todo; con candado, nada):

  A. contratos_grupo.registrar ×10 contratos distintos → los 10 en el registro
  B. contratos_grupo.decidir ×10 → las 10 decisiones se quedan
  C. peticiones_credito.nueva ×10 clientes → 10 peticiones con 10 ids DISTINTOS
  D. beo_contrato.numero_beo ×10 → 10 numeros de BEO distintos y seguidos
  E. almacen_datos.guardar_ajuste_ap ×10 facturas → los 10 ajustes
  F. tab_ar_real.alta_cliente_pendiente ×10 → 10 fichas en clientes_credito.xlsx
  G. lector_contratos_grupo.guardar ×10 → 10 facturas en reservas_credito.xlsx
  H. compensaciones.compensar ×10 parejas distintas → las 10 en el registro y en ajustes_ap
  I. DOS compensaciones de 1.210 a la vez sobre la MISMA pareja → entra una y la otra
     se rechaza ("No queda saldo"); sin candado las dos respondian OK (medido: una se
     perdia sin avisar; segun el momento podrian entrar las dos)
  J. peticiones_credito.firmar_solicitante ×10 (firma de Direccion apagada) → 10
     aprobadas y los 10 limites en clientes_credito.xlsx
  K. /api/ar_real/cobrar ×6 facturas a la vez (la ruta entera) → las 6 COBRADO (sin
     candado: las 6 respondian 200 y solo 1 quedaba cobrada)
  L. candados.registros es reentrante (una funcion protegida llama a otra) y no se
     bloquea con dos carpetas de datos distintas.

  python3.12 tests/test_candados_registros.py
  python3.12 tests/test_candados_registros.py --sabotaje   (candados que no cierran nada: tiene que fallar)
"""
import json
import os
import shutil
import sys
import tempfile
import threading
import time
from contextlib import contextmanager

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_BACKUP_HORA'] = 'off'
import pandas as pd            # noqa: E402

SABOTAJE = '--sabotaje' in sys.argv
N = 10
ESPERA = 0.03
TENANT = "test-candados"
H_A = "HCANDAD"


def carrera(fn, argumentos):
    """Lanza fn(*args) en un hilo por cada args, todos a la vez. Devuelve (resultados, errores)."""
    salida, errores = [None] * len(argumentos), []
    barrera = threading.Barrier(len(argumentos))

    def uno(i, args):
        try:
            barrera.wait(timeout=20)
            salida[i] = fn(*args)
        except Exception as e:          # noqa: BLE001
            errores.append(f"{type(e).__name__}: {e}")
    hilos = [threading.Thread(target=uno, args=(i, a)) for i, a in enumerate(argumentos)]
    for h in hilos:
        h.start()
    for h in hilos:
        h.join(timeout=120)
    return salida, errores


def lento(fn):
    """La funcion de lectura, con una espera DESPUES de leer: abre la ventana leer→escribir."""
    def envoltura(*a, **k):
        r = fn(*a, **k)
        time.sleep(ESPERA)
        return r
    return envoltura


def contrato(i, agencia="Viajes Meridiano S.L."):
    return {"es_contrato_grupo": True, "evento": {"id": f"EV-{i:02d}", "nombre": f"Congreso {i}"},
            "contrato_numero": f"CG-2026-{900 + i}",
            "cliente": {"nombre": f"Cliente {i} S.A.", "cif": "A08000001"},
            "agencia": {"nombre": agencia, "cif": "B28004554"},
            "alojamiento": {"fecha_entrada": "2026-10-05", "fecha_salida": "2026-10-08", "noches": 3,
                            "habitaciones": 10, "total_habitaciones": 5500, "iva_pct": 10},
            "fb": {"total": 0}, "salas": {"total": 0},
            "comisiones": {"modo": "porcentaje", "alojamiento_pct": 10}, "facturacion": {"pagador": None}}


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import candados as CAND
    import contratos_grupo as CG
    import peticiones_credito as PC
    import beo_contrato as B
    import almacen_datos as ALM
    import compensaciones as CMP
    import lector_contratos_grupo as L
    import tab_ar_real as TA

    if SABOTAJE:
        @contextmanager
        def _nada(carpeta):
            yield
        CAND.registros = _nada

    # la ventana leer→escribir, abierta en todos los lectores
    CG.leer = lento(CG.leer)
    PC.leer = lento(PC.leer)
    B._num_leer = lento(B._num_leer)
    ALM.ajustes_ap = lento(ALM.ajustes_ap)
    CMP.leer = lento(CMP.leer)
    _read_excel = pd.read_excel
    pd.read_excel = lento(_read_excel)

    DD = tempfile.mkdtemp(prefix='cand_')
    estado = {}

    def paso(nombre, fn):
        """Cada caso por separado: sin candado un fichero a medio escribir puede reventar la
        lectura (BadZipFile, visto): eso tambien es un fallo, no un test roto."""
        try:
            fn()
        except Exception as e:          # noqa: BLE001
            ok(False, f"{nombre}: excepcion {type(e).__name__}: {str(e)[:120]}")

    def caso_a():
        _, err = carrera(lambda i: CG.registrar(contrato(i), L.transformar(contrato(i), hotel_id=""), hotel_id="", datos_dir=DD),
                         [(i,) for i in range(N)])
        regs = CG.leer(DD)
        estado["ids"] = [c["id"] for c in regs]
        ok(len(regs) == N and not err, f"A · registrar ×{N}: {len(regs)} contratos en el registro {err[:1]}")

    def caso_b():
        ids = estado.get("ids") or []
        _, err = carrera(lambda cid: CG.decidir(cid, pagador="agencia", usuario="u", datos_dir=DD), [(c,) for c in ids])
        dec = [c for c in CG.leer(DD) if (c.get("pagador") or {}).get("origen") == "usuario"]
        ok(len(ids) == N and len(dec) == N and not err, f"B · decidir ×{len(ids)}: {len(dec)} decisiones guardadas {err[:1]}")

    def caso_c():
        _, err = carrera(lambda i: PC.nueva(f"Agencia {i} S.L.", "", "u", "U", datos_dir=DD), [(i,) for i in range(N)])
        pets = PC.leer(DD)
        ok(len(pets) == N and len({p['id'] for p in pets}) == N and not err,
           f"C · peticion nueva ×{N}: {len(pets)} peticiones, {len({p['id'] for p in pets})} ids distintos {err[:1]}")

    def caso_d():
        nums, err = carrera(lambda i: B.numero_beo(f"clave-{i}", DD), [(i,) for i in range(N)])
        ok(sorted(n for n in nums if n) == list(range(B.BEO_INICIAL, B.BEO_INICIAL + N)) and not err,
           f"D · numero de BEO ×{N}: {sorted(n for n in nums if n)} {err[:1]}")

    def caso_e():
        _, err = carrera(lambda i: ALM.guardar_ajuste_ap(f"FRA-{i}", {"dias_pago": 45}, "u", datos_dir=DD), [(i,) for i in range(N)])
        aj = ALM.ajustes_ap(DD)
        ok(len(aj) == N and not err, f"E · ajuste AP ×{N}: {len(aj)} ajustes {err[:1]}")

    def caso_f():
        _, err = carrera(lambda i: TA.alta_cliente_pendiente(f"Agencia Nueva {i}", "", "prueba", datos_dir=DD, hotel_id=""),
                         [(i,) for i in range(N)])
        cli = _read_excel(os.path.join(DD, 'clientes_credito.xlsx'))
        ok(len(cli) == N and not err, f"F · alta de cliente ×{N}: {len(cli)} fichas en clientes_credito.xlsx {err[:1]}")

    def caso_g():
        _, err = carrera(lambda i: L.guardar(L.transformar(contrato(i), hotel_id=""), DD), [(i,) for i in range(N)])
        res = _read_excel(os.path.join(DD, 'reservas_credito.xlsx'))
        ok(len(res) == N and not err, f"G · factura del grupo ×{N}: {len(res)} filas en reservas_credito.xlsx {err[:1]}")

    def caso_h():
        cs = {c["id"]: c for c in CG.leer(DD)}

        def compensar(i, cid, importe=1210.0):
            fg = {"numero_reserva": f"GRP-CG-2026-{900 + i}", "estado": "FACTURADO", "total": 6000.0, "hotel_id": ""}
            fc = {"numero_factura": f"VM-{i}", "total_factura": 1210.0, "base_imponible": 1000.0}
            return CMP.compensar(cs[cid], fg, fc, importe, usuario="u", datos_dir=DD)
        orden = sorted(cs)
        _, err = carrera(compensar, [(i, cid) for i, cid in enumerate(orden)])
        comps = CMP.leer(DD)
        aj = ALM.ajustes_ap(DD)
        refl = sum(1 for i in range(N) if (aj.get(f"VM-{i}") or {}).get("compensado") == 1210.0)
        ok(len(orden) == N and len(comps) == N and refl == N and not err,
           f"H · compensar ×{N}: {len(comps)} en el registro, {refl} reflejadas en la comision {err[:1]}")

    def caso_i():
        DD2 = tempfile.mkdtemp(prefix='cand2_')
        try:
            CG.registrar(contrato(0), L.transformar(contrato(0), hotel_id=""), hotel_id="", datos_dir=DD2)
            c0 = CG.decidir(CG.leer(DD2)[0]["id"], pagador="agencia", usuario="u", datos_dir=DD2)
            fg = {"numero_reserva": "GRP-CG-2026-900", "estado": "FACTURADO", "total": 6000.0, "hotel_id": ""}
            fc = {"numero_factura": "VM-0", "total_factura": 1210.0, "base_imponible": 1000.0}
            _, err = carrera(lambda: CMP.compensar(c0, fg, fc, 1210.0, usuario="u", datos_dir=DD2), [(), ()])
            total = round(sum(x["importe"] for x in CMP.leer(DD2)), 2)
            ok(total == 1210.0 and len(err) == 1,
               f"I · dos compensaciones de 1.210 a la vez sobre la misma comision: compensado {total} (una rechazada: {err[:1]})")
        finally:
            shutil.rmtree(DD2, ignore_errors=True)

    def caso_j():
        for p in PC.leer(DD):
            PC.guardar(p["id"], {"fiscal": {"razon_social": p["cliente"], "nif": "B28004554", "direccion": "C/ Balmes 200",
                                            "cp": "08006", "poblacion": "Barcelona", "pais": "España"},
                                 "informa": {"resultado": "riesgo bajo", "fecha": "2026-09-25"},
                                 "comercial": {"potencial": 50000, "limite": 10000, "revision": "2027-06-30"}}, "u", datos_dir=DD)
        pids = [p["id"] for p in PC.leer(DD)]
        _, err = carrera(lambda pid: PC.firmar_solicitante(pid, "u", "U", datos_dir=DD), [(pid,) for pid in pids])
        aprob = [p for p in PC.leer(DD) if p.get("estado") == "APROBADA"]
        cli = _read_excel(os.path.join(DD, 'clientes_credito.xlsx'))
        con_limite = int((pd.to_numeric(cli['credito_limite'], errors='coerce') == 10000).sum()) if 'credito_limite' in cli.columns else 0
        ok(len(pids) == N and len(aprob) == N and con_limite == N and not err,
           f"J · firmar ×{len(pids)}: {len(aprob)} aprobadas, {con_limite} limites en clientes_credito.xlsx {err[:1]}")

    try:
        for nombre, fn in (("A", caso_a), ("B", caso_b), ("C", caso_c), ("D", caso_d), ("E", caso_e), ("F", caso_f),
                           ("G", caso_g), ("H", caso_h), ("I", caso_i), ("J", caso_j)):
            paso(nombre, fn)
    finally:
        pd.read_excel = _read_excel
        shutil.rmtree(DD, ignore_errors=True)

    # K · la ruta entera: /api/ar_real/cobrar ×6 a la vez
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H_A
    import dashboard as D
    from tenant_dirs import datos_dir
    dd = str(datos_dir())
    if os.path.isdir(dd):
        shutil.rmtree(dd)
    os.makedirs(dd, exist_ok=True)
    try:
        json.dump([{"id": H_A, "nombre": "Hotel Candados", "activo": True}], open(os.path.join(dd, "hoteles.json"), "w"))
        pd.DataFrame([{"numero_reserva": f"FAC-2026-CORP-{k:04d}", "cliente": "Agencia", "estado": "FACTURADO", "total": 100.0 + k,
                       "fecha_emision": "2026-09-01", "hotel_id": H_A} for k in range(1, 7)]).to_excel(
            os.path.join(dd, "reservas_credito.xlsx"), index=False)
        pd.read_excel = lento(_read_excel)
        D.app.config["TESTING"] = True
        clientes = []
        for _ in range(6):
            c = D.app.test_client()
            c.post("/api/login", json={"username": "admin", "password": "admin123"})
            with c.session_transaction() as s:
                s["tenant_id"] = TENANT
                s["hotel_activo"] = H_A
            tok = (c.get("/api/csrf_token").get_json() or {}).get("token", "")
            clientes.append((c, tok))

        def cobrar(k):
            c, tok = clientes[k - 1]
            r = c.post("/api/ar_real/cobrar", json={"numero": f"FAC-2026-CORP-{k:04d}", "fecha_cobro": "2026-09-20"},
                       headers={"X-CSRF-Token": tok})
            return r.status_code
        def caso_k():
            codigos, err = carrera(cobrar, [(k,) for k in range(1, 7)])
            pd.read_excel = _read_excel
            df = _read_excel(os.path.join(dd, "reservas_credito.xlsx"))
            cobradas = int((df["estado"].astype(str).str.upper().isin(["COBRADO", "COBRADA"])).sum())
            ok(cobradas == 6 and all(c == 200 for c in codigos) and not err,
               f"K · /api/ar_real/cobrar ×6 a la vez: {cobradas} de 6 cobradas (codigos {codigos}) {err[:1]}")
        paso("K", caso_k)
    finally:
        pd.read_excel = _read_excel
        shutil.rmtree(dd, ignore_errors=True)

    # L · reentrante y sin bloqueos entre carpetas
    t0 = time.time()
    d1, d2 = tempfile.mkdtemp(), tempfile.mkdtemp()
    try:
        with CAND.registros(d1):
            with CAND.registros(d1):          # la misma carpeta dentro: no se bloquea
                with CAND.registros(d2):      # otra carpeta: independiente
                    pass
        hecho = []

        def otra_carpeta():
            with CAND.registros(d2):
                hecho.append(1)
        with CAND.registros(d1):
            h = threading.Thread(target=otra_carpeta)
            h.start(); h.join(timeout=5)
        ok(time.time() - t0 < 5 and hecho == [1], "L · reentrante en el mismo hilo y carpetas distintas no se esperan")
    finally:
        shutil.rmtree(d1, ignore_errors=True); shutil.rmtree(d2, ignore_errors=True)

    import subprocess
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_')], 'la zona oracle_* no se toca')
    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
