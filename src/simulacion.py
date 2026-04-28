import random
import heapq
import configparser
from enum import Enum
from pathlib import Path

import numpy as np
from scipy.stats import truncnorm, beta

from excel import generar_excel
from texto import generar_txt

# -------------------------
# CONFIGURACIÓN
# -------------------------

config = configparser.ConfigParser(inline_comment_prefixes=(';', '#'))
config.read(Path(__file__).parent.parent / "simulacion.conf")

GENERAL = config["GENERAL"]

SIM_TIME = int(GENERAL["SIM_TIME"])
SEED_RAW = GENERAL.get("SEED", "").strip()
SEED = int(SEED_RAW) if SEED_RAW else None

N_REPLICAS = int(GENERAL.get("N_REPLICAS", 100))

CAPACIDAD_BATERIA_KWH = float(GENERAL["CAPACIDAD_BATERIA_KWH"])
POTENCIA_CR_KW = float(GENERAL["POTENCIA_CR_KW"])
POTENCIA_CSR_KW = float(GENERAL["POTENCIA_CSR_KW"])

ARRIBO_SHAPE = float(GENERAL["ARRIBO_GAMMA_SHAPE"])
ARRIBO_SCALE = float(GENERAL["ARRIBO_GAMMA_SCALE"])

BATERIA_MU = float(GENERAL["BATERIA_MU"])
BATERIA_SIGMA = float(GENERAL["BATERIA_SIGMA"])

BETA_ALPHA = float(GENERAL["BETA_ALPHA"])
BETA_BETA = float(GENERAL["BETA_BETA"])

ESCENARIOS = ["ESCENARIO_REAL", "ESCENARIO_EFICIENTE", "ESCENARIO_NO_EFICIENTE"]
NOMBRES_ESCENARIOS = {
    "ESCENARIO_REAL": "Real",
    "ESCENARIO_EFICIENTE": "Eficiente",
    "ESCENARIO_NO_EFICIENTE": "No eficiente"
}

# -------------------------
# ESTRUCTURAS
# -------------------------

class Cargador:
    def __init__(self):
        self.ocupado_hasta = 0
        self.tiempo_ocioso = 0

class TipoCargador(Enum):
    CR = "CR"
    CSR = "CSR"

# -------------------------
# FUNCIONES AUXILIARES
# -------------------------

def generar_intervalo_arribo():
    return np.random.gamma(ARRIBO_SHAPE, ARRIBO_SCALE)

def generar_bateria_inicial():
    a, b = (0 - BATERIA_MU) / BATERIA_SIGMA, (100 - BATERIA_MU) / BATERIA_SIGMA
    dist = truncnorm(a, b, loc=BATERIA_MU, scale=BATERIA_SIGMA)
    return dist.rvs()

def generar_bateria_final(b_ini):
    while True:
        b_fin = beta.rvs(BETA_ALPHA, BETA_BETA) * 100
        if b_fin >= b_ini:
            return b_fin

def elegir_tipo(PROB_CR, CCR, CCSR):
    if CCR == 0:
        return TipoCargador.CSR
    if CCSR == 0:
        return TipoCargador.CR
    return TipoCargador.CR if random.random() < PROB_CR else TipoCargador.CSR

def tiempo_carga(soc_inicial, soc_final, potencia_cargador_kw):
    energia = ((soc_final - soc_inicial) / 100) * CAPACIDAD_BATERIA_KWH
    return energia / potencia_cargador_kw

def generar_tiempo_carga(tipo, b_ini, b_fin):
    if tipo == TipoCargador.CR:
        return tiempo_carga(b_ini, b_fin, POTENCIA_CR_KW) * 60
    return tiempo_carga(b_ini, b_fin, POTENCIA_CSR_KW) * 60

def tolera_espera(bateria, espera):
    if espera == 0:
        return True
    elif bateria <= 10:
        return random.random() < 0.9
    elif bateria <= 40 and espera > 30:
        return random.random() < 0.4
    elif bateria > 40 and espera > 10:
        return random.random() < 0.2
    else:
        return True

# -------------------------
# SIMULACIÓN
# -------------------------

def correr_simulacion(escenario, seed_offset=0):
    if SEED is not None:
        random.seed(SEED + seed_offset)
        np.random.seed(SEED + seed_offset)

    CCR = int(config[escenario]["CCR"])
    CCSR = config[escenario].getint("CCSR", fallback=1)
    PROB_CR = float(config[escenario]["PROB_CR"])

    reloj = 0
    eventos = []
    heapq.heappush(eventos, (0, "llegada", None))

    cargadores_CR = [Cargador() for _ in range(CCR)]
    cargadores_CSR = [Cargador() for _ in range(CCSR)]

    atendidos = {TipoCargador.CR: 0, TipoCargador.CSR: 0}
    llegadas = {TipoCargador.CR: 0, TipoCargador.CSR: 0}
    arrepentidos = {TipoCargador.CR: 0, TipoCargador.CSR: 0}
    t_espera = {TipoCargador.CR: [], TipoCargador.CSR: []}
    t_sistema = {TipoCargador.CR: [], TipoCargador.CSR: []}

    while eventos:
        reloj, evento, data = heapq.heappop(eventos)

        if evento == "llegada":
            if reloj > SIM_TIME:
                continue

            heapq.heappush(eventos, (reloj + generar_intervalo_arribo(), "llegada", None))

            b_ini = generar_bateria_inicial()
            b_fin = generar_bateria_final(b_ini)
            tipo = elegir_tipo(PROB_CR, CCR, CCSR)

            lista = cargadores_CR if tipo == TipoCargador.CR else cargadores_CSR
            if not lista:
                continue

            llegadas[tipo] += 1

            min_oc = min(c.ocupado_hasta for c in lista)
            candidatos = [c for c in lista if c.ocupado_hasta == min_oc]
            cargador = random.choice(candidatos)

            espera = max(0, cargador.ocupado_hasta - reloj)

            if not tolera_espera(b_ini, espera):
                arrepentidos[tipo] += 1
                continue

            t_serv = generar_tiempo_carga(tipo, b_ini, b_fin)

            inicio = max(reloj, cargador.ocupado_hasta)
            fin = inicio + t_serv

            if cargador.ocupado_hasta < reloj:
                cargador.tiempo_ocioso += reloj - cargador.ocupado_hasta

            cargador.ocupado_hasta = fin

            heapq.heappush(eventos, (fin, "fin", {
                "tipo": tipo,
                "inicio": inicio,
                "llegada": reloj
            }))

        elif evento == "fin":
            tipo = data["tipo"]
            atendidos[tipo] += 1
            t_espera[tipo].append(data["inicio"] - data["llegada"])
            t_sistema[tipo].append(reloj - data["llegada"])

    for c in cargadores_CR + cargadores_CSR:
        if c.ocupado_hasta < SIM_TIME:
            c.tiempo_ocioso += SIM_TIME - c.ocupado_hasta

    def stats(tipo, cargadores):
        l = llegadas[tipo]
        a = arrepentidos[tipo]

        return {
            "llegadas": l,
            "atendidos": atendidos[tipo],
            "arrepentidos": a,
            "pct_arrepentidos": (a / l * 100) if l > 0 else 0,
            "tiempo_espera_prom": np.mean(t_espera[tipo]) if t_espera[tipo] else 0,
            "tiempo_sistema_prom": np.mean(t_sistema[tipo]) if t_sistema[tipo] else 0,
            "ociosidad_por_cargador": [
                c.tiempo_ocioso / max(c.ocupado_hasta, SIM_TIME) * 100
                for c in cargadores
            ],
            "n_cargadores": len(cargadores),
        }

    return {
        "escenario": escenario,
        "nombre": NOMBRES_ESCENARIOS[escenario],
        "CCR": CCR,
        "CCSR": CCSR,
        "PROB_CR": PROB_CR,
        "CR": stats(TipoCargador.CR, cargadores_CR),
        "CSR": stats(TipoCargador.CSR, cargadores_CSR),
    }

# -------------------------
# PROMEDIO
# -------------------------

def promediar_resultados(lista):
    def avg(x): return sum(x) / len(x) if x else 0

    base = lista[0].copy()

    for tipo in ["CR", "CSR"]:
        for clave in ["llegadas", "atendidos", "arrepentidos",
                      "pct_arrepentidos", "tiempo_espera_prom", "tiempo_sistema_prom"]:
            base[tipo][clave] = avg([r[tipo][clave] for r in lista])

        n = len(lista[0][tipo]["ociosidad_por_cargador"])
        base[tipo]["ociosidad_por_cargador"] = [
            avg([r[tipo]["ociosidad_por_cargador"][i] for r in lista])
            for i in range(n)
        ]

    return base

# -------------------------
# MÉTRICAS GLOBALES Y EFICIENCIA
# -------------------------

def calcular_metricas_globales(r):
    total_llegadas = r["CR"]["llegadas"] + r["CSR"]["llegadas"]
    total_arrep = r["CR"]["arrepentidos"] + r["CSR"]["arrepentidos"]
    total_atendidos = r["CR"]["atendidos"] + r["CSR"]["atendidos"]

    # Tiempo de espera promedio ponderado
    W = (
        r["CR"]["tiempo_espera_prom"] * r["CR"]["atendidos"] +
        r["CSR"]["tiempo_espera_prom"] * r["CSR"]["atendidos"]
    ) / max(total_atendidos, 1)

    # Ociosidad promedio total
    ociosidades = r["CR"]["ociosidad_por_cargador"] + r["CSR"]["ociosidad_por_cargador"]
    O = sum(ociosidades) / len(ociosidades) if ociosidades else 0

    # % de abandono
    A = (total_arrep / total_llegadas * 100) if total_llegadas > 0 else 0

    # Usuarios atendidos
    S = total_atendidos

    return {"W": W, "O": O, "A": A, "S": S}


def calcular_eficiencia(metricas):
    W_max = max(m["W"] for m in metricas) if metricas else 0

    eficiencias = []
    for m in metricas:
        W = m["W"]
        O = m["O"]
        W_ratio = (W / W_max) if W_max != 0 else 0
        E = 1 - (0.5 * O + 0.5 * W_ratio)
        eficiencias.append(E)
    return eficiencias

# -------------------------
# MAIN
# -------------------------

if __name__ == "__main__":
    print(f"\nSimulación con {N_REPLICAS} réplicas por escenario\n")

    resultados = []

    for esc in ESCENARIOS:
        print(f"Corriendo escenario: {NOMBRES_ESCENARIOS[esc]}")

        replicas = []
        for i in range(N_REPLICAS):
            replicas.append(correr_simulacion(esc, i))

        r = promediar_resultados(replicas)
        resultados.append(r)

    metricas = [calcular_metricas_globales(r) for r in resultados]
    eficiencias = calcular_eficiencia(metricas)

    for i, r in enumerate(resultados):
        r["eficiencia_global"] = eficiencias[i]
        r["metricas_globales"] = metricas[i]  # opcional

    output_dir = Path(__file__).resolve().parent.parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    generar_txt(
        resultados,
        output_dir / "reporte_simulacion.txt",
        SIM_TIME,
        SEED,
        N_REPLICAS
    )

    generar_excel(
        resultados,
        output_dir / "analisis_sensibilidad.xlsx",
        SIM_TIME,
        SEED,
        N_REPLICAS
    )

    print("\n✓ Simulación finalizada con promedios")