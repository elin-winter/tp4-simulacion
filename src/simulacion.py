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
config.read("simulacion.conf")

SIM_TIME = int(config["GENERAL"]["SIM_TIME"])
SEED = int(config["GENERAL"]["SEED"])

CAPACIDAD_BATERIA_KWH = 50
POTENCIA_CR_KW = 50
POTENCIA_CSR_KW = 22

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
# SIMULACIÓN POR ESCENARIO
# -------------------------

def correr_simulacion(escenario):
    random.seed(SEED)
    np.random.seed(SEED)

    CCR = int(config[escenario]["CCR"]) 
    CCSR = config[escenario].getint("CCSR", fallback=1) 
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

    while eventos:
        # Tomar el siguiente evento de la lista (cola de prioridad)
        reloj, evento, data = heapq.heappop(eventos)

        if evento == "llegada":
            if reloj > SIM_TIME:
                continue
             
            # Generar próximo arribo y agregarlo a la lista de eventos
            ia = generar_intervalo_arribo()
            heapq.heappush(eventos, (reloj + ia, "llegada", None)) 
            

            # Generar estado inicial y final de batería, y tipo de cargador
            b_ini = generar_bateria_inicial()
            b_fin = generar_bateria_final(b_ini)
            tipo_cargador = elegir_tipo(PROB_CR)

            llegadas_por_tipo[tipo_cargador] += 1

            # Seleccionar lista de cargadores según tipo
            lista = cargadores_CR if tipo_cargador == TipoCargador.CR else cargadores_CSR

            
            # Buscar el cargador que antes se desocupa
            min_ocupado = min(c.ocupado_hasta for c in lista)
            candidatos = [c for c in lista if c.ocupado_hasta == min_ocupado]
            cargador = random.choice(candidatos)
            indice = lista.index(cargador)

            # Calcular tiempo de espera (si el cargador está ocupado)
            espera = max(0, cargador.ocupado_hasta - reloj)
            print("Espera: ", espera, "Cargador: ", tipo_cargador)
            # Verificar si el usuario tolera la espera
            if not tolera_espera(b_ini, espera):
                arrepentidos_por_tipo[tipo_cargador] += 1
                continue
            
            # Calcular tiempo de servicio y programar evento de fin de carga
            t_serv = generar_tiempo_carga(tipo_cargador, b_ini, b_fin)

            inicio = max(reloj, cargador.ocupado_hasta)
            fin = inicio + t_serv

            # Acumular tiempo ocioso si corresponde
            if cargador.ocupado_hasta < reloj:
                cargador.tiempo_ocioso += reloj - cargador.ocupado_hasta

            cargador.ocupado_hasta = fin

            heapq.heappush(eventos, (fin, "fin", {
                "tipo": tipo_cargador,
                "indice": indice,
                "inicio": inicio,
                "llegada": reloj
            }))

        # Evento de fin de carga: actualizar estadísticas
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
        return {
            "llegadas": llegadas,
            "atendidos": atendidos[tipo],
            "arrepentidos": arrep,
            "pct_arrepentidos": (arrep / llegadas * 100) if llegadas > 0 else 0,
            "tiempo_espera_prom": np.mean(te) if te else 0,
            "tiempo_sistema_prom": np.mean(ts) if ts else 0,
            "ociosidad_por_cargador": [c.tiempo_ocioso / (max(c.ocupado_hasta, SIM_TIME)) * 100 for c in cargadores],
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
                  f"EspProm:{s['tiempo_espera_prom']:.1f}min \n")
    
            for i, val in enumerate(s["ociosidad_por_cargador"]):
                print(f"  Cargador {label}-{i}: Ocioso {val:.1f}%")

    output_dir = Path(__file__).resolve().parent.parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    txt_path = output_dir / "reporte_simulacion.txt"
    xlsx_path = output_dir / "analisis_sensibilidad.xlsx"

    generar_txt(resultados, txt_path, SIM_TIME, SEED)
    generar_excel(resultados, xlsx_path, SIM_TIME, SEED)
    print("\n✓ Todos los archivos generados exitosamente.")