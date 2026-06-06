"""
Simulador de Demo — Genera cambios realistas en tiempo real
Simula: facturas procesadas, ocupación, alertas, actividad
"""
import json
import random
from datetime import datetime, timedelta
from demo_mode import generar_datos_demo

class DemoSimulador:
    def __init__(self):
        self.data = generar_datos_demo()
        self.iteration = 0
        self.eventos = []
    
    def simular_paso(self):
        """Simula un "paso" en el tiempo (cambios realistas)"""
        self.iteration += 1
        evento = {}
        
        # 30% chance: Nueva factura AR
        if random.random() < 0.3:
            otas = ["Booking.com", "Expedia", "Hotels.com", "Despegar"]
            nueva_factura = {
                "ota": random.choice(otas),
                "factura": f"BKG-2025-06-{random.randint(5000, 9999)}",
                "importe": round(random.uniform(1500, 4500), 2),
                "fecha": (datetime.now() - timedelta(hours=random.randint(0, 24))).strftime("%Y-%m-%d"),
                "estado": random.choice(["Procesada", "Pendiente Aprobación"]),
                "di_cert": random.choice(["SI", "NO - Solicitar"]),
                "hotel": random.choice([h["id"] for h in self.data["hoteles"]])
            }
            self.data["facturas_ar"].append(nueva_factura)
            evento["tipo"] = "nueva_factura_ar"
            evento["detalle"] = f"{nueva_factura['ota']} - €{nueva_factura['importe']}"
        
        # 20% chance: Nueva factura AP
        if random.random() < 0.2:
            proveedores = ["Food Supply Co", "Cleaning Services", "Maintenance Pro", "Utilities"]
            nueva_factura = {
                "proveedor": random.choice(proveedores),
                "factura": f"FAC-{random.randint(1000, 9999)}-2025",
                "importe": round(random.uniform(800, 3500), 2),
                "tipo": random.choice(["F&B", "Servicios", "Mantenimiento"]),
                "estado": random.choice(["3-Way OK", "Pendiente albarán"]),
                "aprobada": random.choice(["SI", "NO"]),
                "hotel": random.choice([h["id"] for h in self.data["hoteles"]])
            }
            self.data["facturas_ap"].append(nueva_factura)
            evento["tipo"] = "nueva_factura_ap"
            evento["detalle"] = f"{nueva_factura['proveedor']} - €{nueva_factura['importe']}"
        
        # Cambios en ocupación (pequeños cambios realistas)
        for hotel in self.data["hoteles"]:
            cambio_occ = random.uniform(-2, 3)
            hotel["ocupacion"] = max(40, min(100, hotel["ocupacion"] + cambio_occ))
            hotel["ocupacion"] = round(hotel["ocupacion"], 1)
            
            # Revenue sube/baja con ocupación
            cambio_rev = cambio_occ * 50 * hotel["habitaciones"]
            hotel["revenue_mtd"] = max(0, hotel["revenue_mtd"] + cambio_rev)
        
        # 15% chance: Nueva alerta
        if random.random() < 0.15:
            tipos_alerta = [
                "Ocupación baja",
                "F&B Cost alto",
                "Missing DI Cert",
                "AP Discrepancia",
                "Out of Balance"
            ]
            nueva_alerta = {
                "hotel": random.choice([h["id"] for h in self.data["hoteles"]]),
                "tipo": random.choice(tipos_alerta),
                "severidad": random.choice(["info", "warning", "critical"]),
                "mensaje": f"Alerta simulada - {datetime.now().strftime('%H:%M:%S')}"
            }
            self.data["alertas"].append(nueva_alerta)
            evento["tipo"] = "nueva_alerta"
            evento["detalle"] = nueva_alerta["tipo"]
        
        # 10% chance: Resolver alerta
        if random.random() < 0.1 and len(self.data["alertas"]) > 3:
            self.data["alertas"].pop(0)
            evento["tipo"] = "alerta_resuelta"
            evento["detalle"] = "1 alerta marcada como resuelta"
        
        # Actualizar consolidado
        self._actualizar_consolidado()
        
        evento["timestamp"] = datetime.now().isoformat()
        evento["iteracion"] = self.iteration
        self.eventos.append(evento)
        
        return evento
    
    def _actualizar_consolidado(self):
        """Recalcula las métricas consolidadas"""
        self.data["consolidado"]["total_revenue_mtd"] = sum(h["revenue_mtd"] for h in self.data["hoteles"])
        self.data["consolidado"]["avg_ocupacion"] = round(sum(h["ocupacion"] for h in self.data["hoteles"]) / len(self.data["hoteles"]), 1)
        self.data["consolidado"]["facturas_ar_procesadas"] = len([f for f in self.data["facturas_ar"] if f["estado"] == "Procesada"])
        self.data["consolidado"]["alertas_activas"] = len(self.data["alertas"])
    
    def get_data(self):
        """Retorna datos actuales"""
        return self.data
    
    def get_eventos_recientes(self, limit=10):
        """Retorna últimos eventos"""
        return self.eventos[-limit:]
    
    def reset(self):
        """Reinicia la simulación"""
        self.data = generar_datos_demo()
        self.iteration = 0
        self.eventos = []

# Instancia global
_simulador = None

def get_simulador():
    global _simulador
    if _simulador is None:
        _simulador = DemoSimulador()
    return _simulador

if __name__ == "__main__":
    print("SIMULADOR DEMO")
    print("=" * 70)
    
    sim = DemoSimulador()
    
    print("\nSimulando 5 pasos...")
    for i in range(5):
        evento = sim.simular_paso()
        print(f"\nPaso {i+1}:")
        print(f"  Tipo: {evento.get('tipo', 'cambios_base')}")
        print(f"  Detalle: {evento.get('detalle', 'Cambios en ocupación y métricas')}")
        print(f"  Revenue total: €{sim.data['consolidado']['total_revenue_mtd']:,.0f}")
        print(f"  Ocupación avg: {sim.data['consolidado']['avg_ocupacion']}%")
        print(f"  Alertas: {sim.data['consolidado']['alertas_activas']}")
