# -*- coding: utf-8 -*-
"""copia_seguridad.py — Yve.01 · copia diaria de los datos FUERA de Render (b81).

Que se copia: las carpetas de datos (`almacen_persistente.CARPETAS`: datos-referencia,
facturas-entrada, facturas-procesadas, reportes, aprobaciones, tenants, ar_real_data),
esten en el disco persistente (YVE_DATA_DIR) o en el propio checkout. Todo va a UN zip
con un MANIFIESTO.json dentro (fecha, carpetas, numero de ficheros).

A donde: a cualquier almacen compatible con S3 (Cloudflare R2, Backblaze B2, AWS S3,
Wasabi…) y/o a una carpeta local (para pruebas, o un segundo disco). Variables de
entorno, todas opcionales — sin ninguna, el modulo no hace nada:

  YVE_BACKUP_S3_ENDPOINT   https://<cuenta>.r2.cloudflarestorage.com  (R2, B2, Wasabi…; vacio = AWS)
  YVE_BACKUP_S3_BUCKET     nombre del bucket
  YVE_BACKUP_S3_KEY        access key id
  YVE_BACKUP_S3_SECRET     secret access key
  YVE_BACKUP_S3_REGION     'auto' (R2) · 'us-west-004' (B2) · 'eu-west-1' (AWS)… por defecto 'auto'
  YVE_BACKUP_S3_PREFIX     carpeta dentro del bucket (por defecto 'yve01/')
  YVE_BACKUP_DIR           carpeta local donde dejar tambien la copia
  YVE_BACKUP_HORA          hora UTC de la copia diaria, 'HH:MM' (por defecto 03:30); 'off' = sin planificador
  YVE_BACKUP_RETENCION     dias que se guardan las copias (por defecto 30; siempre quedan las 3 ultimas)
  YVE_BACKUP_EMAIL         a quien avisar si la copia FALLA (usa Brevo, como las notificaciones)

Restaurar: `restaurar(nombre)` baja el zip, comprueba que se abre entero, guarda antes una
copia de lo que hay ahora ("antes_de_restaurar_…", por si acaso), vacia las carpetas y
extrae. Los modulos abren los ficheros por ruta en cada peticion, asi que la app ve los
datos restaurados sin reiniciar (las caches se invalidan por huella del fichero).

Linea de comandos:  python copia_seguridad.py ahora | listar | restaurar <nombre> | estado
"""
import json
import os
import shutil
import tempfile
import threading
import time
import zipfile
from datetime import datetime, timedelta, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ESTADO_FICHERO = "copias_estado.json"
PREFIJO = "yve01_"
PREFIJO_PREVIA = "antes_de_restaurar_"
_lock = threading.Lock()


# ── que se copia ─────────────────────────────────────────────────────────────
def raiz_datos():
    """Donde viven los datos: el disco persistente si lo hay, si no el checkout."""
    d = str(os.environ.get("YVE_DATA_DIR", "") or "").strip()
    return d if d and os.path.isdir(d) else BASE_DIR


def carpetas_datos():
    try:
        from almacen_persistente import CARPETAS
    except Exception:
        CARPETAS = ["datos-referencia", "facturas-entrada", "facturas-procesadas", "reportes",
                    "aprobaciones", "tenants", "ar_real_data"]
    return list(CARPETAS)


def _ruta_estado():
    return os.path.join(raiz_datos(), ESTADO_FICHERO)


def estado():
    try:
        return json.load(open(_ruta_estado(), encoding="utf-8"))
    except Exception:
        return {}


def _guardar_estado(d):
    try:
        tmp = _ruta_estado() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, _ruta_estado())
    except Exception:
        pass


def crear_zip(destino, raiz=None, carpetas=None):
    """Mete las carpetas de datos en `destino` (zip). Devuelve {ficheros, bytes, carpetas}."""
    raiz = raiz or raiz_datos()
    carpetas = carpetas or carpetas_datos()
    n = 0
    incluidas = []
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as z:
        for nombre in carpetas:
            src = os.path.join(raiz, nombre)
            if not os.path.isdir(src):
                continue
            incluidas.append(nombre)
            for dirpath, dirs, files in os.walk(src):
                dirs[:] = [d for d in dirs if d != "__pycache__"]
                for f in files:
                    if f.endswith((".tmp", ".lock")) or f.startswith("~$"):
                        continue
                    p = os.path.join(dirpath, f)
                    arc = os.path.relpath(p, raiz)
                    try:
                        z.write(p, arc)
                        n += 1
                    except OSError:
                        continue
        manifiesto = {"aplicacion": "Yve.01", "fecha": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                      "raiz": raiz, "carpetas": incluidas, "ficheros": n, "formato": 1}
        z.writestr("MANIFIESTO.json", json.dumps(manifiesto, ensure_ascii=False, indent=2))
    return {"ficheros": n, "bytes": os.path.getsize(destino), "carpetas": incluidas}


# ── destinos ─────────────────────────────────────────────────────────────────
class DestinoLocal:
    tipo = "local"

    def __init__(self, carpeta):
        self.carpeta = carpeta
        os.makedirs(carpeta, exist_ok=True)

    def nombre(self):
        return f"carpeta {self.carpeta}"

    def subir(self, ruta, nombre):
        shutil.copy2(ruta, os.path.join(self.carpeta, nombre))

    def listar(self):
        out = []
        for f in os.listdir(self.carpeta):
            if f.endswith(".zip") and (f.startswith(PREFIJO) or f.startswith(PREFIJO_PREVIA)):
                p = os.path.join(self.carpeta, f)
                out.append({"nombre": f, "bytes": os.path.getsize(p),
                            "fecha": datetime.fromtimestamp(os.path.getmtime(p), timezone.utc).isoformat(timespec="seconds")})
        return sorted(out, key=lambda x: x["nombre"], reverse=True)

    def descargar(self, nombre, destino):
        shutil.copy2(os.path.join(self.carpeta, nombre), destino)

    def borrar(self, nombre):
        try:
            os.remove(os.path.join(self.carpeta, nombre))
        except FileNotFoundError:
            pass


class DestinoS3:
    tipo = "s3"

    def __init__(self, bucket, key, secret, endpoint="", region="auto", prefijo="yve01/", cliente=None):
        self.bucket = bucket
        self.prefijo = (prefijo or "").strip("/")
        self.prefijo = self.prefijo + "/" if self.prefijo else ""
        self.endpoint = endpoint or ""
        if cliente is not None:
            self.cli = cliente
        else:
            import boto3
            from botocore.config import Config
            kw = {"aws_access_key_id": key, "aws_secret_access_key": secret,
                  "config": Config(signature_version="s3v4", retries={"max_attempts": 4, "mode": "standard"})}
            if endpoint:
                kw["endpoint_url"] = endpoint
            if region:
                kw["region_name"] = region
            self.cli = boto3.client("s3", **kw)

    def nombre(self):
        return f"bucket {self.bucket}" + (f" ({self.endpoint})" if self.endpoint else " (AWS)")

    def _k(self, nombre):
        return self.prefijo + nombre

    def subir(self, ruta, nombre):
        self.cli.upload_file(ruta, self.bucket, self._k(nombre))

    def listar(self):
        out = []
        tok = None
        while True:
            kw = {"Bucket": self.bucket, "Prefix": self.prefijo}
            if tok:
                kw["ContinuationToken"] = tok
            r = self.cli.list_objects_v2(**kw)
            for o in r.get("Contents", []) or []:
                n = o["Key"][len(self.prefijo):]
                if "/" in n or not n.endswith(".zip"):
                    continue
                if not (n.startswith(PREFIJO) or n.startswith(PREFIJO_PREVIA)):
                    continue
                lm = o.get("LastModified")
                out.append({"nombre": n, "bytes": int(o.get("Size", 0)),
                            "fecha": lm.isoformat(timespec="seconds") if hasattr(lm, "isoformat") else str(lm or "")})
            if not r.get("IsTruncated"):
                break
            tok = r.get("NextContinuationToken")
        return sorted(out, key=lambda x: x["nombre"], reverse=True)

    def descargar(self, nombre, destino):
        self.cli.download_file(self.bucket, self._k(nombre), destino)

    def borrar(self, nombre):
        self.cli.delete_object(Bucket=self.bucket, Key=self._k(nombre))


def destinos(env=None):
    """Los destinos configurados (0, 1 o 2). Sin ninguno, la copia no hace nada."""
    env = env if env is not None else os.environ
    out = []
    bucket = (env.get("YVE_BACKUP_S3_BUCKET") or "").strip()
    if bucket:
        out.append(DestinoS3(bucket, env.get("YVE_BACKUP_S3_KEY", ""), env.get("YVE_BACKUP_S3_SECRET", ""),
                             endpoint=(env.get("YVE_BACKUP_S3_ENDPOINT") or "").strip(),
                             region=(env.get("YVE_BACKUP_S3_REGION") or "auto").strip(),
                             prefijo=env.get("YVE_BACKUP_S3_PREFIX", "yve01/")))
    carpeta = (env.get("YVE_BACKUP_DIR") or "").strip()
    if carpeta:
        out.append(DestinoLocal(carpeta))
    return out


def configurado(env=None):
    return bool(destinos(env))


def retencion_dias(env=None):
    env = env if env is not None else os.environ
    try:
        return max(1, int(env.get("YVE_BACKUP_RETENCION", "30")))
    except ValueError:
        return 30


# ── copiar ───────────────────────────────────────────────────────────────────
def _fecha_de(nombre):
    """yve01_20260912_033000.zip -> datetime UTC (o None)."""
    try:
        base = nombre[len(PREFIJO):] if nombre.startswith(PREFIJO) else nombre[len(PREFIJO_PREVIA):]
        return datetime.strptime(base[:15], "%Y%m%d_%H%M%S").replace(tzinfo=timezone.utc)
    except Exception:
        return None


def aplicar_retencion(destino, dias, ahora=None, minimo=3):
    """Borra las copias diarias mas viejas que `dias`, dejando siempre las `minimo` ultimas.
    Las copias 'antes_de_restaurar_' siguen la misma regla. Devuelve los nombres borrados."""
    ahora = ahora or datetime.now(timezone.utc)
    limite = ahora - timedelta(days=dias)
    borradas = []
    lista = destino.listar()
    diarias = [c for c in lista if c["nombre"].startswith(PREFIJO)]
    for c in diarias[minimo:]:
        f = _fecha_de(c["nombre"])
        if f and f < limite:
            destino.borrar(c["nombre"]); borradas.append(c["nombre"])
    for c in [c for c in lista if c["nombre"].startswith(PREFIJO_PREVIA)]:
        f = _fecha_de(c["nombre"])
        if f and f < limite:
            destino.borrar(c["nombre"]); borradas.append(c["nombre"])
    return borradas


def _avisar_fallo(mensaje, env=None):
    env = env if env is not None else os.environ
    dest = (env.get("YVE_BACKUP_EMAIL") or env.get("YVE_ALERTAS_EMAIL") or "").strip()
    if not dest:
        return False
    try:
        from notificaciones import enviar_email
        return bool(enviar_email(dest, "⚠ Yve.01: la copia de seguridad ha FALLADO",
                                 f"<p>La copia de seguridad diaria de Yve.01 no se ha podido hacer.</p><pre>{mensaje}</pre>"
                                 f"<p>Revisa las variables YVE_BACKUP_* en Render y el bucket.</p>", tipo="copia_seguridad"))
    except Exception:
        return False


def hacer_copia(env=None, nombre=None, prefijo=PREFIJO, motivo="diaria"):
    """Crea el zip y lo sube a todos los destinos. Devuelve el resultado (y lo apunta en el estado)."""
    env = env if env is not None else os.environ
    with _lock:
        dests = destinos(env)
        res = {"ok": False, "fecha": datetime.now(timezone.utc).isoformat(timespec="seconds"), "motivo": motivo,
               "destinos": [], "errores": []}
        if not dests:
            res["errores"].append("sin destino configurado (YVE_BACKUP_S3_BUCKET o YVE_BACKUP_DIR)")
            return res
        nombre = nombre or f"{prefijo}{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.zip"
        res["nombre"] = nombre
        tmpdir = tempfile.mkdtemp(prefix="yve_copia_")
        try:
            ruta = os.path.join(tmpdir, nombre)
            info = crear_zip(ruta)
            res.update(info)
            for d in dests:
                try:
                    d.subir(ruta, nombre)
                    borradas = aplicar_retencion(d, retencion_dias(env)) if prefijo == PREFIJO else []
                    res["destinos"].append({"tipo": d.tipo, "donde": d.nombre(), "ok": True, "borradas": borradas})
                except Exception as e:
                    res["destinos"].append({"tipo": d.tipo, "donde": d.nombre(), "ok": False, "error": str(e)[:300]})
                    res["errores"].append(f"{d.nombre()}: {str(e)[:300]}")
            res["ok"] = any(x["ok"] for x in res["destinos"]) and not res["errores"]
        except Exception as e:
            res["errores"].append(str(e)[:300])
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        est = estado()
        if prefijo == PREFIJO:
            est["ultima"] = res
            if res["ok"]:
                est["ultima_ok"] = res
        est.setdefault("historial", [])
        est["historial"] = ([{"fecha": res["fecha"], "ok": res["ok"], "nombre": res.get("nombre"), "motivo": motivo,
                              "bytes": res.get("bytes"), "errores": res["errores"][:2]}] + est["historial"])[:30]
        _guardar_estado(est)
        if not res["ok"]:
            _avisar_fallo("\n".join(res["errores"]) or "error desconocido", env)
        return res


def listar(env=None):
    """Todas las copias de todos los destinos: [{nombre, bytes, fecha, destino}]."""
    out = []
    for d in destinos(env):
        try:
            for c in d.listar():
                out.append({**c, "destino": d.tipo, "donde": d.nombre()})
        except Exception as e:
            out.append({"nombre": "", "error": str(e)[:200], "destino": d.tipo, "donde": d.nombre()})
    return sorted(out, key=lambda x: (x.get("nombre") or "", x.get("destino") or ""), reverse=True)


# ── restaurar ────────────────────────────────────────────────────────────────
def comprobar_zip(ruta):
    """Que el zip se abre entero y es de Yve. Devuelve el manifiesto."""
    with zipfile.ZipFile(ruta) as z:
        malo = z.testzip()
        if malo:
            raise ValueError(f"zip corrupto en {malo}")
        try:
            man = json.loads(z.read("MANIFIESTO.json").decode("utf-8"))
        except KeyError:
            raise ValueError("no es una copia de Yve.01 (sin MANIFIESTO.json)")
        if not man.get("carpetas") or not int(man.get("ficheros") or 0):
            raise ValueError("la copia esta vacia (0 ficheros): no se restaura nada con ella")
        for n in z.namelist():
            if n.startswith("/") or ".." in n.split("/"):
                raise ValueError(f"ruta peligrosa dentro del zip: {n}")
    return man


def descargar_copia(nombre, destino_ruta, env=None, tipo=None):
    """Baja la copia `nombre` del primer destino que la tenga (o del `tipo` pedido)."""
    ultimo = None
    for d in destinos(env):
        if tipo and d.tipo != tipo:
            continue
        try:
            d.descargar(nombre, destino_ruta)
            return d
        except Exception as e:
            ultimo = e
    raise FileNotFoundError(f"no se encontro la copia {nombre}: {ultimo}")


def restaurar(nombre, env=None, raiz=None, tipo=None, copia_previa=True):
    """Deja los datos como estaban en la copia `nombre`. Devuelve un dict con lo hecho."""
    env = env if env is not None else os.environ
    raiz = raiz or raiz_datos()
    with _lock:
        tmpdir = tempfile.mkdtemp(prefix="yve_restaurar_")
        res = {"ok": False, "nombre": nombre, "fecha": datetime.now(timezone.utc).isoformat(timespec="seconds")}
        try:
            ruta = os.path.join(tmpdir, nombre)
            d = descargar_copia(nombre, ruta, env, tipo)
            res["desde"] = d.nombre()
            man = comprobar_zip(ruta)
            res["manifiesto"] = man
            if copia_previa:
                previa = f"{PREFIJO_PREVIA}{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.zip"
                ruta_prev = os.path.join(tmpdir, previa)
                crear_zip(ruta_prev, raiz=raiz)
                for dd in destinos(env):
                    try:
                        dd.subir(ruta_prev, previa)
                    except Exception as e:
                        res.setdefault("avisos", []).append(f"copia previa no subida a {dd.nombre()}: {str(e)[:120]}")
                res["copia_previa"] = previa
            # vaciar y extraer SOLO las carpetas que trae la copia
            carpetas = man.get("carpetas") or carpetas_datos()
            for nombre_c in carpetas:
                dst = os.path.join(raiz, nombre_c)
                real = os.path.realpath(dst)
                if os.path.islink(dst) or os.path.isdir(dst):
                    for hijo in os.listdir(real):
                        p = os.path.join(real, hijo)
                        shutil.rmtree(p, ignore_errors=True) if os.path.isdir(p) and not os.path.islink(p) else os.remove(p)
                else:
                    os.makedirs(dst, exist_ok=True)
            n = 0
            with zipfile.ZipFile(ruta) as z:
                for info in z.infolist():
                    if info.filename == "MANIFIESTO.json" or info.is_dir():
                        continue
                    partes = info.filename.split("/")
                    if partes[0] not in carpetas:
                        continue
                    destino_f = os.path.join(os.path.realpath(os.path.join(raiz, partes[0])), *partes[1:])
                    os.makedirs(os.path.dirname(destino_f), exist_ok=True)
                    with z.open(info) as src, open(destino_f, "wb") as dst_f:
                        shutil.copyfileobj(src, dst_f)
                    n += 1
            res["ficheros"] = n
            res["carpetas"] = carpetas
            res["ok"] = True
        except Exception as e:
            res["error"] = str(e)[:300]
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        est = estado()
        est["ultima_restauracion"] = res
        _guardar_estado(est)
        try:
            import almacen_datos
            if hasattr(almacen_datos, "_invalidate"):
                almacen_datos._invalidate()
        except Exception:
            pass
        return res


# ── planificador (hilo dentro de la app) ─────────────────────────────────────
def _proxima(hora_hhmm, ahora=None):
    ahora = ahora or datetime.now(timezone.utc)
    try:
        hh, mm = [int(x) for x in hora_hhmm.split(":")]
    except Exception:
        hh, mm = 3, 30
    p = ahora.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if p <= ahora:
        p += timedelta(days=1)
    return p


_hilo = None


def arrancar_planificador(env=None):
    """Un hilo que hace la copia cada dia a YVE_BACKUP_HORA (UTC). Solo si hay destino."""
    global _hilo
    env = env if env is not None else os.environ
    hora = (env.get("YVE_BACKUP_HORA") or "03:30").strip().lower()
    if hora == "off" or not configurado(env) or _hilo is not None:
        return None

    def bucle():
        while True:
            espera = (_proxima(hora) - datetime.now(timezone.utc)).total_seconds()
            time.sleep(max(30, espera))
            try:
                r = hacer_copia(env, motivo="diaria")
                print(f"[copia_seguridad] {'OK' if r['ok'] else 'FALLO'} {r.get('nombre')} {r.get('bytes')} bytes {r['errores']}")
            except Exception as e:
                print(f"[copia_seguridad] error inesperado: {e}")

    _hilo = threading.Thread(target=bucle, name="copia_seguridad", daemon=True)
    _hilo.start()
    print(f"[copia_seguridad] planificada cada dia a las {hora} UTC → {', '.join(d.nombre() for d in destinos(env))}")
    return _hilo


def resumen(env=None):
    """Para el panel y /health: configurado, destinos, ultima copia."""
    env = env if env is not None else os.environ
    est = estado()
    u = est.get("ultima_ok") or {}
    dias = None
    if u.get("fecha"):
        try:
            dias = (datetime.now(timezone.utc) - datetime.fromisoformat(u["fecha"])).total_seconds() / 86400.0
        except Exception:
            dias = None
    return {"configurado": configurado(env), "hora": (env.get("YVE_BACKUP_HORA") or "03:30"),
            "retencion_dias": retencion_dias(env), "destinos": [d.nombre() for d in destinos(env)],
            "ultima": est.get("ultima"), "ultima_ok": u, "dias_desde_ultima_ok": round(dias, 2) if dias is not None else None,
            "ultima_restauracion": est.get("ultima_restauracion"),
            "al_dia": bool(u) and dias is not None and dias < 1.5}


if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else "estado"
    if cmd == "ahora":
        print(json.dumps(hacer_copia(motivo="manual"), ensure_ascii=False, indent=2))
    elif cmd == "listar":
        for c in listar():
            print(f"{c.get('nombre', ''):40} {c.get('bytes', 0):>12} {c.get('fecha', '')} {c.get('destino')} {c.get('error', '')}")
    elif cmd == "restaurar" and len(sys.argv) > 2:
        print(json.dumps(restaurar(sys.argv[2]), ensure_ascii=False, indent=2))
    else:
        print(json.dumps(resumen(), ensure_ascii=False, indent=2))
