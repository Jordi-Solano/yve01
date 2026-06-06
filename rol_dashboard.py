"""
Personalización de dashboard por rol
Cada rol ve solo lo que necesita
"""

ROLE_PERMISSIONS = {
    "admin": {
        "nombre": "Administrador",
        "descripcion": "Acceso completo",
        "tabs": ["ar", "ap", "drr", "banco", "notificaciones", "fb", "ar_real", "multi_hotel", "calipolis"],
        "acciones": ["aprobar", "rechazar", "editar", "exportar", "ver_usuarios"],
        "visible_en_kpis": ["revenue", "ocupacion", "gop", "alertas", "facturas_pendientes"]
    },
    "financial_controller": {
        "nombre": "Controller Financiero",
        "descripcion": "Gestión de finanzas e integración Oracle",
        "tabs": ["ar", "ap", "drr", "banco", "notificaciones", "ar_real", "calipolis"],
        "acciones": ["aprobar", "exportar", "ver_reportes"],
        "visible_en_kpis": ["revenue", "gop", "alertas", "facturas_pendientes", "oracle_status"]
    },
    "income_auditor": {
        "nombre": "Income Auditor",
        "descripcion": "Auditoría de ingresos y OTAs",
        "tabs": ["ar", "drr", "banco", "notificaciones"],
        "acciones": ["aprobar_ar", "exportar"],
        "visible_en_kpis": ["revenue", "ocupacion", "alertas_di_cert", "facturas_ar"]
    },
    "fb_manager": {
        "nombre": "Jefe de F&B",
        "descripcion": "Gestión de costos y compras F&B",
        "tabs": ["ap", "fb", "notificaciones", "calipolis"],
        "acciones": ["aprobar_ap_fb", "exportar"],
        "visible_en_kpis": ["ingresos_fb", "food_cost_pct", "alertas_fb", "facturas_ap_fb"]
    },
    "jefe_otras": {
        "nombre": "Jefe de Servicios",
        "descripcion": "Gestión de AP no-F&B",
        "tabs": ["ap", "notificaciones"],
        "acciones": ["aprobar_ap_otras"],
        "visible_en_kpis": ["facturas_ap_otras", "alertas"]
    }
}

def get_dashboard_config(role):
    """Retorna configuración de dashboard para un rol"""
    config = ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS["income_auditor"])
    
    return {
        "role": role,
        "nombre": config["nombre"],
        "descripcion": config["descripcion"],
        "tabs_disponibles": config["tabs"],
        "tab_default": config["tabs"][0] if config["tabs"] else "ar",
        "acciones_permitidas": config["acciones"],
        "kpis_visibles": config["visible_en_kpis"]
    }

def get_html_tabs(role):
    """Genera HTML de tabs según rol"""
    config = get_dashboard_config(role)
    
    tab_icons = {
        "ar": "📥 AR — OTAs",
        "ap": "📦 AP — Proveedores",
        "drr": "📊 DRR",
        "banco": "🏦 Banco",
        "notificaciones": "🔔 Notificaciones",
        "fb": "🍽️ F&B Cost",
        "ar_real": "🏢 AR Real",
        "multi_hotel": "🏨 Multi-Hotel",
        "calipolis": "🏩 Calipolis"
    }
    
    html = ""
    for tab in config["tabs_disponibles"]:
        icon_text = tab_icons.get(tab, tab)
        html += f'    <button class="tab" onclick="switchTab(\'{tab}\',this)">{icon_text}</button>\n'
    
    return html

def get_kpis_config(role):
    """Retorna qué KPIs mostrar según rol"""
    config = get_dashboard_config(role)
    
    kpis_map = {
        "revenue": {"label": "Revenue MTD", "color": "#1db954"},
        "ocupacion": {"label": "Ocupación", "color": "#1a73e8"},
        "gop": {"label": "GOP", "color": "#ff9800"},
        "alertas": {"label": "Alertas Activas", "color": "#e05252"},
        "facturas_pendientes": {"label": "Facturas Pendientes", "color": "#ff9800"},
        "oracle_status": {"label": "Oracle Status", "color": "#1a73e8"},
        "ingresos_fb": {"label": "Ingresos F&B", "color": "#1db954"},
        "food_cost_pct": {"label": "F&B Cost %", "color": "#ff9800"},
        "alertas_di_cert": {"label": "DI Certs Pendientes", "color": "#e05252"},
        "facturas_ar": {"label": "Facturas AR", "color": "#1a73e8"},
        "alertas_fb": {"label": "Alertas F&B", "color": "#e05252"},
        "facturas_ap_fb": {"label": "Facturas AP (F&B)", "color": "#ff9800"},
        "facturas_ap_otras": {"label": "Facturas AP (Otras)", "color": "#ff9800"}
    }
    
    return {kpi: kpis_map[kpi] for kpi in config["kpis_visibles"] if kpi in kpis_map}

if __name__ == "__main__":
    print("CONFIGURACIÓN DE ROLES")
    print("=" * 70)
    
    for role in ROLE_PERMISSIONS.keys():
        config = get_dashboard_config(role)
        print(f"\n{role.upper()}: {config['nombre']}")
        print(f"  Descripción: {config['descripcion']}")
        print(f"  Tabs: {', '.join(config['tabs_disponibles'])}")
        print(f"  KPIs: {', '.join(config['kpis_visibles'])}")
