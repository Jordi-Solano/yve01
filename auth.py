"""
auth.py — Yve.01
Módulo de autenticación y gestión de usuarios.
Usuarios almacenados en datos-referencia/usuarios.json.
"""

import os, json
from pathlib import Path
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, UserMixin

BASE_DIR       = Path(__file__).parent
USUARIOS_PATH  = BASE_DIR / "datos-referencia" / "usuarios.json"

ROLES_VALIDOS = {"admin", "financial_controller", "income_auditor", "fb_manager", "jefe_otras"}

# ── Modelo de usuario ─────────────────────────────────────────────────────

class Usuario(UserMixin):
    def __init__(self, data: dict):
        self.id       = data["username"]
        self.username  = data["username"]
        self.nombre    = data.get("nombre", "")
        self.email     = data.get("email", "")
        self.rol       = data.get("rol", "income_auditor")
        self.activo    = data.get("activo", True)

    @property
    def is_active(self):
        return self.activo

    def to_dict(self):
        return {"username": self.username, "nombre": self.nombre,
                "email": self.email, "rol": self.rol, "activo": self.activo}


# ── Almacén de usuarios ──────────────────────────────────────────────────

def _load_all():
    if USUARIOS_PATH.exists():
        try:
            return json.loads(USUARIOS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []

def _save_all(users):
    USUARIOS_PATH.parent.mkdir(exist_ok=True)
    USUARIOS_PATH.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")

def _find(username):
    for u in _load_all():
        if u["username"] == username:
            return u
    return None


# ── API pública ───────────────────────────────────────────────────────────

def inicializar_usuarios():
    """Crea usuarios por defecto si usuarios.json no existe."""
    if USUARIOS_PATH.exists() and _load_all():
        return
    defaults = [
        {"username": "admin",     "password": "admin123",  "nombre": "Administrador",     "email": "admin@hotel.com",    "rol": "admin"},
        {"username": "fc_user",   "password": "hotel2024", "nombre": "Financial Controller","email": "fc@hotel.com",      "rol": "financial_controller"},
        {"username": "auditor",   "password": "hotel2024", "nombre": "Income Auditor",     "email": "auditor@hotel.com",  "rol": "income_auditor"},
        {"username": "fbmanager", "password": "hotel2024", "nombre": "F&B Manager",        "email": "fb@hotel.com",       "rol": "fb_manager"},
    ]
    users = []
    for d in defaults:
        users.append({
            "username":      d["username"],
            "password_hash": generate_password_hash(d["password"]),
            "nombre":        d["nombre"],
            "email":         d["email"],
            "rol":           d["rol"],
            "activo":        True,
        })
    _save_all(users)


def login(username, password):
    """Verifica credenciales. Devuelve Usuario o None."""
    u = _find(username)
    if not u:
        return None
    if not u.get("activo", True):
        return None
    if check_password_hash(u["password_hash"], password):
        return Usuario(u)
    return None


def get_usuario(username):
    """Carga un usuario por username. Para flask-login user_loader."""
    u = _find(username)
    if u and u.get("activo", True):
        return Usuario(u)
    return None


def listar_usuarios():
    """Devuelve lista de usuarios (sin password_hash)."""
    return [
        {k: v for k, v in u.items() if k != "password_hash"}
        for u in _load_all()
    ]


def crear_usuario(username, password, nombre, email, rol):
    """Crea un nuevo usuario. Devuelve True/error string."""
    if not username or not password:
        return "Username y password requeridos"
    if rol not in ROLES_VALIDOS:
        return f"Rol inválido: {rol}"
    if _find(username):
        return "El usuario ya existe"
    users = _load_all()
    users.append({
        "username":      username,
        "password_hash": generate_password_hash(password),
        "nombre":        nombre,
        "email":         email,
        "rol":           rol,
        "activo":        True,
    })
    _save_all(users)
    return True


def cambiar_password(username, nueva_password):
    """Cambia la contraseña de un usuario."""
    users = _load_all()
    for u in users:
        if u["username"] == username:
            u["password_hash"] = generate_password_hash(nueva_password)
            _save_all(users)
            return True
    return "Usuario no encontrado"


def toggle_activo(username):
    """Activa/desactiva un usuario."""
    users = _load_all()
    for u in users:
        if u["username"] == username:
            u["activo"] = not u.get("activo", True)
            _save_all(users)
            return u["activo"]
    return None


def verificar_rol(usuario, roles_permitidos):
    """Verifica si el usuario tiene uno de los roles permitidos."""
    if isinstance(roles_permitidos, str):
        roles_permitidos = {roles_permitidos}
    return usuario.rol in roles_permitidos


# ── Flask-Login setup ─────────────────────────────────────────────────────

def init_login(app):
    """Configura Flask-Login en la app Flask."""
    app.secret_key = app.secret_key or os.urandom(24).hex()
    lm = LoginManager()
    lm.login_view = "login_page"
    lm.init_app(app)

    @lm.user_loader
    def load_user(user_id):
        return get_usuario(user_id)

    return lm
