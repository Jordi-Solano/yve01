# -*- coding: utf-8 -*-
"""Disco persistente (Jordi, sep 2026): con YVE_DATA_DIR las carpetas de datos
viven en el disco y un "deploy" (repo nuevo) no borra nada.

  python3.12 tests/test_almacen_persistente.py
  python3.12 tests/test_almacen_persistente.py --sabotaje
"""
import os
import shutil
import subprocess
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

SABOTAJE = '--sabotaje' in sys.argv


def _repo_falso(raiz):
    """Un repo con sus carpetas de datos y ficheros de referencia, como el de verdad."""
    for c in ("datos-referencia", "facturas-procesadas", "reportes", "aprobaciones", "facturas-entrada", "tenants", "ar_real_data"):
        os.makedirs(os.path.join(raiz, c), exist_ok=True)
    open(os.path.join(raiz, "datos-referencia", "proveedores.xlsx"), "w").write("maestro v1")
    open(os.path.join(raiz, "datos-referencia", "usuarios.json"), "w").write("{}")
    return raiz


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import almacen_persistente as AP
    if SABOTAJE:
        AP.montar = lambda base=None, data_dir=None: None      # el disco no se monta nunca

    tmp = tempfile.mkdtemp(prefix='disco_')
    try:
        disco = os.path.join(tmp, "var_data")
        repo1 = _repo_falso(os.path.join(tmp, "repo1"))
        # sin variable: no hace nada
        ok(AP.montar(repo1, "") is None and not os.path.islink(os.path.join(repo1, "datos-referencia")), "sin YVE_DATA_DIR no toca nada")
        h = AP.montar(repo1, disco)
        ok(h and len(h["enlazadas"]) == 7 and "datos-referencia" in h["sembradas"], f"primer arranque: siembra y enlaza {h and len(h['enlazadas'])} carpetas")
        ok(os.path.islink(os.path.join(repo1, "datos-referencia")) and os.path.exists(os.path.join(disco, "datos-referencia", "proveedores.xlsx")),
           "la carpeta del repo es un enlace y el maestro esta en el disco")
        # la app escribe "por el repo" y acaba en el disco
        open(os.path.join(repo1, "facturas-procesadas", "facturas_ap_20260901.xlsx"), "w").write("subido por el usuario")
        open(os.path.join(repo1, "datos-referencia", "proveedores.xlsx"), "w").write("maestro EDITADO por el cliente")
        ok(os.path.exists(os.path.join(disco, "facturas-procesadas", "facturas_ap_20260901.xlsx")), "lo que la app guarda va al disco")
        # DEPLOY: repo nuevo desde cero (Render clona otra vez), con un fichero de referencia nuevo en el repo
        repo2 = _repo_falso(os.path.join(tmp, "repo2"))
        open(os.path.join(repo2, "datos-referencia", "plan_cuentas.xlsx"), "w").write("nuevo en esta version")
        h2 = AP.montar(repo2, disco)
        ok(h2 and not h2["sembradas"] and h2["copiados"] == 1, f"segundo deploy: no re-siembra, copia solo lo nuevo ({h2 and h2['copiados']} fichero)")
        _lee = lambda *p: open(os.path.join(*p)).read() if os.path.exists(os.path.join(*p)) else None
        ok(_lee(repo2, "facturas-procesadas", "facturas_ap_20260901.xlsx") == "subido por el usuario", "lo subido SOBREVIVE al deploy")
        ok(_lee(repo2, "datos-referencia", "proveedores.xlsx") == "maestro EDITADO por el cliente", "el repo nuevo NO pisa lo que el cliente edito")
        ok(os.path.exists(os.path.join(repo2, "datos-referencia", "plan_cuentas.xlsx")), "y el fichero nuevo del repo aparece")
        # reinicio del mismo contenedor: idempotente
        h3 = AP.montar(repo2, disco)
        ok(h3 and not h3["enlazadas"] and os.path.islink(os.path.join(repo2, "reportes")), "reiniciar sin deploy: nada que hacer, enlaces intactos")
        ok(os.path.exists(os.path.join(disco, AP.MARCA)), "marca de montaje en el disco")
        # estado() para /health
        os.environ["YVE_DATA_DIR"] = disco
        e = AP.estado(repo2)
        ok(e["persistente"] is True, f"estado(): persistente={e['persistente']}")
        os.environ.pop("YVE_DATA_DIR", None)
        ok(AP.estado(repo1)["persistente"] is False, "estado sin variable: no persistente")
        src = open(os.path.join(BASE, 'dashboard.py'), encoding='utf-8').read()
        ok(src.index("almacen_persistente as _AP_DISCO") < src.index("from tenant_dirs import"), "dashboard monta el disco ANTES de importar nada que abra ficheros")
        ok("almacenamiento_persistente" in src, "/health dice si hay disco")
        ok("YVE_DATA_DIR" in open(os.path.join(BASE, 'render.yaml')).read(), "render.yaml documenta la variable y el disco")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_') or f == 'lector_facturas_ap.py'], 'ni oracle_* ni clasificador')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
