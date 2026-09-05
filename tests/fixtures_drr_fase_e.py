"""DRR de prueba para la fase E — tres hoteles con perfiles DISTINTOS a proposito.

Los numeros estan elegidos para que una media PLANA y una PONDERADA den
resultados claramente diferentes. Si los tres hoteles fueran parecidos, el
error de ponderacion no se veria y la prueba pasaria sin demostrar nada — que
es justo lo que pasaba con los datos del demo.

  Hotel Costa Azul   grande, 400 hab · ocupacion 60% · con GOP medido
  Hotel Plaza Mayor  pequeño,  20 hab · ocupacion 95% · con GOP medido
  Hotel Ribera       mediano           · SIN GOP en el fichero -> N/D

  (y un cuarto hotel del censo se queda SIN subir nada -> estado "sin DRR")

Ocupacion del grupo:
    plana      = (60 + 95) / 2                = 77.5%   ← mentira
    ponderada  = (7200 + 570) / (12000 + 600) = 61.7%   ← lo correcto
El hotel de 20 habitaciones no puede pesar lo mismo que el de 400.

Nombres genericos: esto se ve en la web.
"""
import os
import sys
sys.path.insert(0, os.path.dirname(__file__))
from crear_drr import construir as _construir_drr       # el generador que vive en tests/
from datetime import datetime


def construir(nombre, fecha, metricas, dias, carpeta=None):
    """Adaptador al `crear_drr.construir` del repo (antes se importaba un
    `crear.py` de una carpeta externa del sandbox, /home/claude/drr, que ya no
    existe: el test estaba atado a una maquina). Devuelve la ruta escrita."""
    carpeta = carpeta or os.environ.get("YVE_FIXTURES_DRR") or os.path.join(os.path.dirname(__file__), "_drr_fase_e")
    hotel = nombre.replace("DRR-", "").replace(".xlsm", "").replace("-", " ").title()
    dias_tb = []
    for n, f, cuentas, descuadre in dias:
        rev_hab, rev_rest, rev_bar, g1, g2, g3 = cuentas
        dias_tb.append((n, f.strftime("%Y-%m-%d"), [
            ("INCOME", "Room Revenue", 0.0, rev_hab), ("INCOME", "Restaurant Revenue", 0.0, rev_rest),
            ("INCOME", "Bar Revenue", 0.0, rev_bar), ("EXPENSES", "Salaries", g1, 0.0),
            ("EXPENSES", "Supplies", g2, 0.0), ("EXPENSES", "Energy", g3, 0.0),
            ("ASSETS", "Cash", rev_hab + rev_rest + rev_bar - g1 - g2 - g3, 0.0)], descuadre))
    return _construir_drr(os.path.join(carpeta, nombre), metricas, dias_tb, hotel=hotel,
                          fecha_informe=fecha.strftime("%Y-%m-%d"))

FECHA = datetime(2026, 7, 29)


def dias_normales(rev_hab, rev_rest, rev_bar, descuadre_dia3=0.0):
    """Tres dias de trial balance; el tercero puede ir descuadrado."""
    return [
        (1, datetime(2026, 7, 1), (rev_hab, rev_rest, rev_bar, 4100.0, 1750.0, 620.0), 0.0),
        (2, datetime(2026, 7, 2), (rev_hab, rev_rest, rev_bar, 4050.0, 1690.0, 600.0), 0.0),
        (3, datetime(2026, 7, 3), (rev_hab, rev_rest, rev_bar, 4200.0, 1800.0, 640.0), descuadre_dia3),
    ]


# ── 1 · Hotel Costa Azul — GRANDE (400 hab), ocupacion baja, GOP medido ────
# 400 hab x 30 dias = 12.000 noches disponibles; al 60% -> 7.200 ocupadas.
costa = {
    "Occupancy %":       (0.60, 0.60, 0.62, 0.65),
    "Rooms Occupied":    (240, 7200, 7440, 7800),
    "ADR":               (120.0, 120.0, 121.0, 125.0),
    "Revenue PAR":       (72.0, 72.0, 75.0, 81.2),
    "Rooms Revenue":     (28800.0, 864000.0, 900240.0, 975000.0),
    "Food Revenue":      (6100.0, 183000.0, 189000.0, 195000.0),
    "Beverage Revenue":  (2400.0, 72000.0, 74000.0, 76000.0),
    "F&B Other":         (500.0, 15000.0, 15500.0, 16000.0),
    "F&B Revenue Total": (9000.0, 270000.0, 278500.0, 287000.0),
    "Telephone / Other": (400.0, 12000.0, 12400.0, 12800.0),
    "Total Revenue":     (38200.0, 1146000.0, 1191140.0, 1274800.0),
    "Spend PAR":         (52.0, 52.0, 53.0, 54.0),
    "GOP":               (10696.0, 320880.0, 333519.0, 369692.0),
    "GOP %":             (0.28, 0.28, 0.28, 0.29),
}

# ── 2 · Hotel Plaza Mayor — PEQUEÑO (20 hab), ocupacion alta, GOP medido ───
# 20 hab x 30 dias = 600 noches disponibles; al 95% -> 570 ocupadas.
plaza = {
    "Occupancy %":       (0.95, 0.95, 0.93, 0.90),
    "Rooms Occupied":    (19, 570, 558, 540),
    "ADR":               (210.0, 210.0, 208.0, 205.0),
    "Revenue PAR":       (199.5, 199.5, 193.4, 184.5),
    "Rooms Revenue":     (3990.0, 119700.0, 116064.0, 110700.0),
    "Food Revenue":      (900.0, 27000.0, 26500.0, 26000.0),
    "Beverage Revenue":  (420.0, 12600.0, 12300.0, 12000.0),
    "F&B Other":         (80.0, 2400.0, 2350.0, 2300.0),
    "F&B Revenue Total": (1400.0, 42000.0, 41150.0, 40300.0),
    "Telephone / Other": (60.0, 1800.0, 1760.0, 1700.0),
    "Total Revenue":     (5450.0, 163500.0, 158974.0, 152700.0),
    "Spend PAR":         (120.0, 120.0, 118.0, 115.0),
    "GOP":               (2071.0, 62130.0, 60011.0, 57226.0),
    "GOP %":             (0.38, 0.38, 0.3775, 0.3748),
}

# ── 3 · Hotel Ribera — mediano, y SIN GOP en el fichero ───────────────────
# Es el caso de la fase D dentro de la fase E: tiene DRR y cuenta para la
# ocupacion del grupo, pero su GOP sale N/D y NO entra en el GOP del grupo.
# Por eso los dos denominadores salen distintos: 3 hoteles con DRR, 2 con GOP.
ribera = {
    "Occupancy %":       (0.75, 0.75, 0.76, 0.78),
    "Rooms Occupied":    (90, 2700, 2736, 2808),
    "ADR":               (160.0, 160.0, 162.0, 165.0),
    "Revenue PAR":       (120.0, 120.0, 123.1, 128.7),
    "Rooms Revenue":     (14400.0, 432000.0, 443232.0, 463320.0),
    "Food Revenue":      (3200.0, 96000.0, 97000.0, 99000.0),
    "Beverage Revenue":  (1300.0, 39000.0, 39500.0, 40000.0),
    "F&B Other":         (280.0, 8400.0, 8500.0, 8600.0),
    "F&B Revenue Total": (4780.0, 143400.0, 145000.0, 147600.0),
    "Telephone / Other": (210.0, 6300.0, 6400.0, 6500.0),
    "Total Revenue":     (19390.0, 581700.0, 594632.0, 617420.0),
    "Spend PAR":         (88.0, 88.0, 89.0, 90.0),
    # Sin GOP ni GOP %: el generador los deja vacios.
    "GOP":               (None, None, None, None),
    "GOP %":             (None, None, None, None),
}


def main():
    construir("DRR-COSTA-AZUL.xlsm",  FECHA, costa,
              dias_normales(28800.0, 6100.0, 2400.0))
    construir("DRR-PLAZA-MAYOR.xlsm", FECHA, plaza,
              dias_normales(3990.0, 900.0, 420.0))
    # Ribera con un dia descuadrado, para que ademas se vea un OOB en el grupo.
    construir("DRR-RIBERA.xlsm",      FECHA, ribera,
              dias_normales(14400.0, 3200.0, 1300.0, descuadre_dia3=312.75))
    print("generados: DRR-COSTA-AZUL.xlsm · DRR-PLAZA-MAYOR.xlsm · DRR-RIBERA.xlsm")


if __name__ == "__main__":
    main()
