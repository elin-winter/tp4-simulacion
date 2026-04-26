import random
import heapq
import numpy as np
import configparser
from scipy.stats import truncnorm, beta
from enum import Enum
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
import datetime

from excel import generar_excel
from texto import generar_txt

# -------------------------
# CONFIGURACIÓN
# -------------------------

config = configparser.ConfigParser(inline_comment_prefixes=(';', '#'))
config.read("simulacion.conf")

SIM_TIME = int(config["GENERAL"]["SIM_TIME"])
SEED = int(config["GENERAL"]["SEED"])

CAPACIDAD_BATERIA_KWH = 50
POTENCIA_CR_KW = 150
POTENCIA_CSR_KW = 22

ESCENARIOS = ["ESCENARIO_REAL"]
NOMBRES_ESCENARIOS = {
    "ESCENARIO_REAL": "Real",
}
bat_10 = 0
bat_40 = 0
bat_90 = 0

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
    return np.random.gamma(3, 12)

def generar_bateria_inicial():
    mu, sigma = 32, 15
    a, b = (0 - mu) / sigma, (100 - mu) / sigma
    dist = truncnorm(a, b, loc=mu, scale=sigma)
    return dist.rvs()  
         
def generar_bateria_final(b_ini):
    alpha, beta_param = 20, 2
    
    while True:
        # Beta.rvs() genera valores en [0, 1], multiplicamos por 100
        b_fin = beta.rvs(alpha, beta_param) * 100
        
        if b_fin >= b_ini:
            return b_fin

def elegir_tipo(PROB_CR):
    return TipoCargador.CR if random.random() < PROB_CR else TipoCargador.CSR

def tiempo_carga(soc_inicial, soc_final, potencia_cargador_kw):
    energia_necesaria = ((soc_final - soc_inicial) / 100) * CAPACIDAD_BATERIA_KWH
    return energia_necesaria / potencia_cargador_kw

def generar_tiempo_carga(tipo, b_ini, b_fin):
    if tipo == TipoCargador.CR:
        return tiempo_carga(b_ini, b_fin, POTENCIA_CR_KW) * 60
    return tiempo_carga(b_ini, b_fin, POTENCIA_CSR_KW) * 60

def tolera_espera(bateria, espera):
    global bat_10, bat_40, bat_90
    if espera == 0:
        return True
    elif bateria <= 10:
        status = random.random() < 0.9
        if (not status):
            bat_10 += 1
        return status
    elif bateria <= 40 and espera > 60:
        status = random.random() < 0.4
        if (not status):
            bat_40 += 1
        return status
    elif espera > 30:
        status =  random.random() < 0.2
        if (not status):
            bat_90 += 1
        return status
    else:
        return True

# -------------------------
# SIMULACIÓN POR ESCENARIO
# -------------------------

def correr_simulacion(escenario):
    random.seed(SEED)
    np.random.seed(SEED)

    CCR = int(config[escenario]["CCR"])
    CCSR = int(config[escenario]["CCSR"])
    PROB_CR = float(config[escenario]["PROB_CR"])

    reloj = 0
    eventos = []
    heapq.heappush(eventos, (0, "llegada", None))

    cargadores_CR = [Cargador() for _ in range(CCR)]
    cargadores_CSR = [Cargador() for _ in range(CCSR)]

    atendidos = {TipoCargador.CR: 0, TipoCargador.CSR: 0}
    llegadas_por_tipo = {TipoCargador.CR: 0, TipoCargador.CSR: 0}
    arrepentidos_por_tipo = {TipoCargador.CR: 0, TipoCargador.CSR: 0}
    tiempos_espera = {TipoCargador.CR: [], TipoCargador.CSR: []}
    tiempos_sistema = {TipoCargador.CR: [], TipoCargador.CSR: []}

    while eventos and reloj < SIM_TIME:
        reloj, evento, data = heapq.heappop(eventos)
        print("Eventos :", eventos)

        if evento == "llegada":
            ia = generar_intervalo_arribo()
            heapq.heappush(eventos, (reloj + ia, "llegada", None)) 

            b_ini = generar_bateria_inicial()
            b_fin = generar_bateria_final(b_ini)
            tipo_cargador = elegir_tipo(PROB_CR)

            llegadas_por_tipo[tipo_cargador] += 1

            lista = cargadores_CR if tipo_cargador == TipoCargador.CR else cargadores_CSR

            min_ocupado = min(c.ocupado_hasta for c in lista)
            candidatos = [c for c in lista if c.ocupado_hasta == min_ocupado]
            cargador = random.choice(candidatos)
            indice = lista.index(cargador)

            print ("cargador ocupado: ", cargador.ocupado_hasta)
            print ("reloj : ", reloj)
            
            espera = max(0, cargador.ocupado_hasta - reloj)

            print (" Espera: ", espera)
            print("--------------------\n")
            if not tolera_espera(b_ini, espera):
                arrepentidos_por_tipo[tipo_cargador] += 1
                continue
            
            
            t_serv = generar_tiempo_carga(tipo_cargador, b_ini, b_fin)
            inicio = max(reloj, cargador.ocupado_hasta)
            fin = inicio + t_serv

            if cargador.ocupado_hasta < reloj:
                cargador.tiempo_ocioso += reloj - cargador.ocupado_hasta

            cargador.ocupado_hasta = fin

            heapq.heappush(eventos, (fin, "fin", {
                "tipo": tipo_cargador,
                "indice": indice,
                "inicio": inicio,
                "llegada": reloj
            }))


        elif evento == "fin":
            tipo_cargador = data["tipo"]
            atendidos[tipo_cargador] += 1
            tiempos_espera[tipo_cargador].append(data["inicio"] - data["llegada"])
            tiempos_sistema[tipo_cargador].append(reloj - data["llegada"])

    for c in cargadores_CR + cargadores_CSR:
        if c.ocupado_hasta < SIM_TIME:
            c.tiempo_ocioso += SIM_TIME - c.ocupado_hasta

    def stats(tipo, cargadores):
        llegadas = llegadas_por_tipo[tipo]
        arrep = arrepentidos_por_tipo[tipo]
        te = tiempos_espera[tipo]
        ts = tiempos_sistema[tipo]
        n = len(cargadores)
        ociosidad_total = sum(c.tiempo_ocioso for c in cargadores)
        return {
            "llegadas": llegadas,
            "atendidos": atendidos[tipo],
            "arrepentidos": arrep,
            "pct_arrepentidos": (arrep / llegadas * 100) if llegadas > 0 else 0,
            "tiempo_espera_prom": np.mean(te) if te else 0,
            "tiempo_sistema_prom": np.mean(ts) if ts else 0,
            "ociosidad_prom_pct": ociosidad_total / (n * SIM_TIME) * 100,
            "ociosidad_por_cargador": [c.tiempo_ocioso / SIM_TIME * 100 for c in cargadores],
            "n_cargadores": n,
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
# MAIN
# -------------------------

if __name__ == "__main__":
    print(f"\nIniciando simulación — {len(ESCENARIOS)} escenarios | SIM_TIME={SIM_TIME} min\n")
    resultados = []
    for esc in ESCENARIOS:
        print(f"  Corriendo: {NOMBRES_ESCENARIOS[esc]}...")
        r = correr_simulacion(esc)
        resultados.append(r)

        for tipo_k, label in [("CR", "CR"), ("CSR", "CSR")]:
            s = r[tipo_k]
            print(f"    [{label}] Llegadas:{s['llegadas']} | Atendidos:{s['atendidos']} | "
                  f"Arrep:{s['arrepentidos']} ({s['pct_arrepentidos']:.1f}%) | "
                  f"EspProm:{s['tiempo_espera_prom']:.1f}min | Ocioso:{s['ociosidad_prom_pct']:.1f}%")

    txt_path = "C:/Users/alesc/Documents/Github/tp4-simulacion/output/reporte_simulacion.txt"
    xlsx_path = "C:/Users/alesc/Documents/Github/tp4-simulacion/output/analisis_sensibilidad.xlsx"

    generar_txt(resultados, txt_path, SIM_TIME, SEED)
    generar_excel(resultados, xlsx_path, SIM_TIME, SEED)
    print(f"Arrepentimiento 10: {bat_10}\n")
    print(f"Arrepentimiento 40: {bat_40}\n")
    print(f"Arrepentimiento 90: {bat_90}\n")
    print("\n✓ Todos los archivos generados exitosamente.")