"""candados.py — b99: exclusion para leer-modificar-escribir los registros de datos.

Con el comando del render.yaml (`--workers 1 --threads 8`) dos peticiones a la vez
pueden LEER el mismo fichero, cambiar cada una lo suyo y ESCRIBIR: la segunda
borra lo de la primera (medido en el test: sin candado se pierden altas, numeros
de BEO repetidos, peticiones de credito con el mismo id). Hasta el 25 sep
produccion corria con UN hilo y no pasaba; en cuanto se ponga el comando del
render.yaml, pasa.

UN candado por carpeta de datos (por tenant) para TODOS los registros del AR:
contratos_grupo.json, compensaciones_ar.json, peticiones_credito.json,
beo_numeracion.json, eventos_referencia.json, reservas_credito.xlsx,
clientes_credito.xlsx y ajustes_ap.json. Uno solo y no uno por fichero a
proposito: las operaciones se anidan (decidir quien paga toca el contrato, la
factura del grupo y la ficha del cliente; emitir toca la factura y el
contrato), y con candados por fichero dos peticiones que los toman en orden
distinto se bloquean en cruz. Con uno reentrante no hay orden que respetar.

- Entre hilos del mismo proceso: threading.RLock (reentrante: una funcion
  protegida puede llamar a otra protegida).
- Entre procesos (subprocesos del lote, otro worker): fcntl.flock sobre
  `<carpeta>/.registros.lock`, solo en el nivel mas externo de cada hilo.

Lo protegido es solo el tramo leer-cambiar-escribir: nada de llamadas a la IA
dentro (una lectura de contrato de 30 s no puede tener parado el AR entero).
"""
import functools
import inspect
import os
import threading
from contextlib import contextmanager

try:
    import fcntl
except ImportError:          # pragma: no cover (Windows: solo el candado de hilos)
    fcntl = None

FICHERO_CANDADO = ".registros.lock"

_mx = threading.Lock()
_rlocks = {}                 # carpeta real -> RLock
_local = threading.local()   # carpeta real -> (profundidad, fichero del flock) en ESTE hilo


def _clave(carpeta):
    return os.path.realpath(os.path.abspath(str(carpeta)))


def _rlock(clave):
    with _mx:
        lk = _rlocks.get(clave)
        if lk is None:
            lk = _rlocks[clave] = threading.RLock()
        return lk


@contextmanager
def registros(carpeta):
    """Exclusion sobre los registros de datos de `carpeta` (la de datos del tenant).
    Reentrante en el mismo hilo; sin `carpeta` no hace nada."""
    if not carpeta:
        yield
        return
    clave = _clave(carpeta)
    lk = _rlock(clave)
    lk.acquire()
    estado = getattr(_local, "estado", None)
    if estado is None:
        estado = _local.estado = {}
    prof, fh = estado.get(clave, (0, None))
    try:
        if prof == 0 and fcntl is not None:
            try:
                os.makedirs(clave, exist_ok=True)
                fh = open(os.path.join(clave, FICHERO_CANDADO), "a+")
                fcntl.flock(fh, fcntl.LOCK_EX)
            except OSError:
                fh = None            # sin disco donde dejar el candado: queda el de hilos
        estado[clave] = (prof + 1, fh)
        yield
    finally:
        prof2, fh2 = estado.get(clave, (1, fh))
        if prof2 <= 1:
            estado.pop(clave, None)
            if fh2 is not None:
                try:
                    fcntl.flock(fh2, fcntl.LOCK_UN)
                finally:
                    fh2.close()
        else:
            estado[clave] = (prof2 - 1, fh2)
        lk.release()


def protegido(resolver):
    """Decorador: la funcion corre con el candado de la carpeta de datos que resuelve
    `resolver(datos_dir)` a partir de SU argumento `datos_dir` (por nombre o posicion)."""
    def deco(fn):
        sig = inspect.signature(fn)

        @functools.wraps(fn)
        def envoltura(*a, **k):
            try:
                dd = sig.bind_partial(*a, **k).arguments.get("datos_dir")
            except TypeError:
                dd = k.get("datos_dir")
            with registros(resolver(dd)):
                return fn(*a, **k)
        envoltura.__protegido__ = True
        return envoltura
    return deco


def en_la_peticion(fn):
    """Decorador para rutas Flask que leen, deciden y escriben registros del AR: la
    peticion ENTERA con el candado de la carpeta de datos del tenant. Solo para rutas
    rapidas (nada de IA ni correo dentro: esas protegen solo su tramo de escritura)."""
    @functools.wraps(fn)
    def envoltura(*a, **k):
        try:
            from tenant_dirs import datos_dir as _d
            carpeta = str(_d())
        except Exception:
            carpeta = None
        with registros(carpeta):
            return fn(*a, **k)
    return envoltura


def _tmp(ruta, sufijo):
    return f"{ruta}.{os.getpid()}-{threading.get_ident()}{sufijo}"


def escribir_json(obj, ruta):
    """Escritura ATOMICA (temporal propio + os.replace): quien lea a la vez ve el fichero
    de antes o el de despues, nunca uno a medias."""
    import json
    os.makedirs(os.path.dirname(str(ruta)) or ".", exist_ok=True)
    tmp = _tmp(str(ruta), ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(obj, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, str(ruta))
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def escribir_excel(df, ruta, **kw):
    """Como escribir_json, para los .xlsx (antes `df.to_excel(ruta)` a pelo: un lector a la
    vez podia abrir el fichero a medio escribir -> BadZipFile, medido en el test)."""
    kw.setdefault("index", False)
    os.makedirs(os.path.dirname(str(ruta)) or ".", exist_ok=True)
    tmp = _tmp(str(ruta), ".tmp.xlsx")
    try:
        df.to_excel(tmp, **kw)
        os.replace(tmp, str(ruta))
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
