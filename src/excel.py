from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
import datetime


AZUL_HEADER = "1F4E79"
AZUL_CLARO = "BDD7EE"
VERDE = "375623"
VERDE_CLARO = "E2EFDA"
NARANJA = "833C00"
NARANJA_CLARO = "FCE4D6"
GRIS = "F2F2F2"
BLANCO = "FFFFFF"
ESCENARIO_COLORES = {
    "Real": ("2E75B6", "DEEAF1"),
    "Eficiente": ("375623", "E2EFDA"),
    "No eficiente": ("C00000", "FCE4D6"),
}

def estilar_celda(ws, celda, valor=None, bold=False, fondo=None, fuente_color="000000",
                  alineacion="center", borde=False, formato=None, tamaño=11):
    c = ws[celda]
    if valor is not None:
        c.value = valor
    c.font = Font(bold=bold, color=fuente_color, size=tamaño, name="Arial")
    if fondo:
        c.fill = PatternFill("solid", fgColor=fondo)
    c.alignment = Alignment(horizontal=alineacion, vertical="center", wrap_text=True)
    if borde:
        thin = Side(style="thin", color="AAAAAA")
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    if formato:
        c.number_format = formato


def generar_excel(resultados, path, sim_time, seed):
    wb = Workbook()

    # ---- Hoja 1: Resumen comparativo ----
    ws = wb.active
    ws.title = "Resumen Comparativo"
    ws.sheet_view.showGridLines = False

    # Título
    ws.merge_cells("A1:J1")
    estilar_celda(ws, "A1", "ANÁLISIS DE SENSIBILIDAD — ESTACIÓN DE CARGA VE",
                  bold=True, fondo=AZUL_HEADER, fuente_color=BLANCO, tamaño=14)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:J2")
    estilar_celda(ws, "A2", f"Simulación de {sim_time} minutos | Seed: {seed} | {datetime.datetime.now().strftime('%Y-%m-%d')}",
                  fondo="D6E4F0", fuente_color="1F4E79", tamaño=10)
    ws.row_dimensions[2].height = 16

    # Encabezados tabla
    headers = ["Métrica", "Unidad",
               "Real (CR)", "Real (CSR)",
               "Eficiente (CR)", "Eficiente (CSR)",
               "No eficiente (CR)", "No eficiente (CSR)"]
    cols_w = [28, 10, 13, 13, 13, 13, 13, 13]
    for i, (h, w) in enumerate(zip(headers, cols_w)):
        col = get_column_letter(i + 1)
        ws.column_dimensions[col].width = w
        estilar_celda(ws, f"{col}4", h, bold=True, fondo=AZUL_HEADER,
                      fuente_color=BLANCO, borde=True, tamaño=10)
    ws.row_dimensions[4].height = 28

    metricas = [
        ("N° Cargadores", "#", "n_cargadores", None),
        ("Total Llegadas", "#", "llegadas", None),
        ("Total Atendidos", "#", "atendidos", None),
        ("Arrepentidos", "#", "arrepentidos", None),
        ("% Arrepentidos", "%", "pct_arrepentidos", "0.0%"),
        ("Tiempo Prom. Espera", "min", "tiempo_espera_prom", "0.00"),
        ("Tiempo Prom. Sistema", "min", "tiempo_sistema_prom", "0.00"),
        ("Ociosidad Prom. por Cargador", "%", "ociosidad_por_cargador", "0.0%"),  # ← nuevo
    ]

    fila_inicio_datos = 5
    for i, (label, unidad, key, fmt) in enumerate(metricas):
        row = fila_inicio_datos + i
        fondo_row = GRIS if i % 2 == 0 else BLANCO
        estilar_celda(ws, f"A{row}", label, bold=True, fondo=fondo_row,
                      alineacion="left", borde=True, tamaño=10)
        estilar_celda(ws, f"B{row}", unidad, fondo=fondo_row, borde=True, tamaño=10)

        col = 3
        for r in resultados:
            for tipo_k in ["CR", "CSR"]:
                val = r[tipo_k][key]
                if key == "pct_arrepentidos":
                    val = val / 100
                elif key == "ociosidad_por_cargador":
                    val = (sum(val) / len(val) / 100) if len(val) > 0 else 0
                cell = f"{get_column_letter(col)}{row}"
                estilar_celda(ws, cell, val, fondo=fondo_row, borde=True, tamaño=10,
                              formato=fmt if fmt else "General")
                col += 1
        ws.row_dimensions[row].height = 18

    # Colorear columnas por escenario
    escenario_col_ranges = [
        (3, 4, "Real"),
        (5, 6, "Eficiente"),
        (7, 8, "No eficiente"),
    ]
    for col_s, col_e, nombre in escenario_col_ranges:
        color_h, color_l = ESCENARIO_COLORES[nombre]
        for col in range(col_s, col_e + 1):
            cl = get_column_letter(col)
            ws[f"{cl}4"].fill = PatternFill("solid", fgColor=color_h)
            for i in range(len(metricas)):
                row = fila_inicio_datos + i
                c = ws[f"{cl}{row}"]
                c.fill = PatternFill("solid", fgColor=color_l if i % 2 == 0 else BLANCO)

    # ---- Hoja 2: Detalle por Escenario ----
    ws2 = wb.create_sheet("Detalle por Escenario")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 12

    ws2.merge_cells("A1:C1")
    estilar_celda(ws2, "A1", "DETALLE POR ESCENARIO Y TIPO DE CARGADOR",
                  bold=True, fondo=AZUL_HEADER, fuente_color=BLANCO, tamaño=13)
    ws2.row_dimensions[1].height = 28

    fila = 3
    for r in resultados:
        nombre = r["nombre"]
        color_h, color_l = ESCENARIO_COLORES[nombre]

        # Escenario header
        ws2.merge_cells(f"A{fila}:C{fila}")
        estilar_celda(ws2, f"A{fila}", f"ESCENARIO: {nombre.upper()} — CR:{r['CCR']} | CSR:{r['CCSR']} | P(CR)={r['PROB_CR']:.0%}",
                      bold=True, fondo=color_h, fuente_color=BLANCO, tamaño=11)
        ws2.row_dimensions[fila].height = 20
        fila += 1

        for tipo_k, label in [("CR", "Cargadores Rápidos (CR)"), ("CSR", "Cargadores Semi-Rápidos (CSR)")]:
            s = r[tipo_k]
            # Sub-header
            ws2.merge_cells(f"A{fila}:C{fila}")
            estilar_celda(ws2, f"A{fila}", f"  {label}", bold=True, fondo=color_l,
                          fuente_color=color_h, alineacion="left", tamaño=10)
            ws2.row_dimensions[fila].height = 18
            fila += 1

            filas_det = [
                ("Total llegadas", s["llegadas"], "General"),
                ("Total atendidos", s["atendidos"], "General"),
                ("Arrepentidos", s["arrepentidos"], "General"),
                ("% Arrepentidos", s["pct_arrepentidos"] / 100, "0.0%"),
                ("Tiempo prom. espera (min)", s["tiempo_espera_prom"], "0.00"),
                ("Tiempo prom. sistema (min)", s["tiempo_sistema_prom"], "0.00"),
                #("Ociosidad promedio", s["ociosidad_prom_pct"] / 100, "0.0%"),
            ]
            for j, (met, val, fmt) in enumerate(filas_det):
                fondo_f = GRIS if j % 2 == 0 else BLANCO
                estilar_celda(ws2, f"A{fila}", met, fondo=fondo_f, alineacion="left",
                              borde=True, tamaño=10)
                estilar_celda(ws2, f"B{fila}", val, fondo=fondo_f, borde=True,
                              formato=fmt, tamaño=10)
                ws2.merge_cells(f"C{fila}:C{fila}")
                ws2.row_dimensions[fila].height = 16
                fila += 1

            # Ociosidad por cargador
            estilar_celda(ws2, f"A{fila}", "Ociosidad por cargador:",
                          bold=True, fondo=color_l, alineacion="left", tamaño=10)
            ws2.merge_cells(f"B{fila}:C{fila}")
            ws2.row_dimensions[fila].height = 16
            fila += 1
            for idx_c, oc in enumerate(s["ociosidad_por_cargador"]):
                estilar_celda(ws2, f"A{fila}", f"    {tipo_k}[{idx_c}]", fondo=BLANCO,
                              alineacion="left", borde=True, tamaño=10)
                estilar_celda(ws2, f"B{fila}", oc / 100, fondo=BLANCO, borde=True,
                              formato="0.0%", tamaño=10)
                ws2.row_dimensions[fila].height = 15
                fila += 1

        fila += 1

    # ---- Hoja 3: Datos para gráficos ----
    ws3 = wb.create_sheet("Datos Gráficos")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:G1")
    estilar_celda(ws3, "A1", "DATOS PARA GRÁFICOS COMPARATIVOS",
                  bold=True, fondo=AZUL_HEADER, fuente_color=BLANCO, tamaño=12)
    ws3.row_dimensions[1].height = 24

    escenarios_nombres = [r["nombre"] for r in resultados]
    tipos = ["CR", "CSR"]

    datasets = [
        ("% Arrepentidos", "pct_arrepentidos", "0.0%"),
        ("Tiempo Espera Prom (min)", "tiempo_espera_prom", "0.00"),
        ("Tiempo Sistema Prom (min)", "tiempo_sistema_prom", "0.00"),
        ("Ociosidad Prom por Cargador (%)", "ociosidad_por_cargador", "0.0%"),
    ]

    fila3 = 3
    chart_data_refs = []

    for ds_label, ds_key, ds_fmt in datasets:
        ws3.merge_cells(f"A{fila3}:G{fila3}")
        estilar_celda(ws3, f"A{fila3}", ds_label, bold=True, fondo=AZUL_CLARO,
                      fuente_color=AZUL_HEADER, tamaño=11)
        ws3.row_dimensions[fila3].height = 20
        fila_header = fila3 + 1
        fila3 += 1

        # Headers
        estilar_celda(ws3, f"A{fila3}", "Escenario", bold=True, fondo=AZUL_HEADER,
                      fuente_color=BLANCO, borde=True, tamaño=10)
        ws3.column_dimensions["A"].width = 16
        for j, t in enumerate(tipos):
            col_l = get_column_letter(j + 2)
            ws3.column_dimensions[col_l].width = 14
            estilar_celda(ws3, f"{col_l}{fila3}", t, bold=True, fondo=AZUL_HEADER,
                          fuente_color=BLANCO, borde=True, tamaño=10)
        ws3.row_dimensions[fila3].height = 18
        fila3 += 1

        data_start = fila3
        for r in resultados:
            estilar_celda(ws3, f"A{fila3}", r["nombre"], fondo=GRIS, alineacion="left",
                          borde=True, tamaño=10)
            for j, t in enumerate(tipos):
                val = r[t][ds_key]
                if ds_key == "ociosidad_por_cargador":
                    val = (sum(val) / len(val) / 100) if len(val) > 0 else 0
                elif ds_fmt == "0.0%":
                    val = val / 100
                col_l = get_column_letter(j + 2)
                estilar_celda(ws3, f"{col_l}{fila3}", val, fondo=BLANCO, borde=True,
                              formato=ds_fmt, tamaño=10)
            ws3.row_dimensions[fila3].height = 16
            fila3 += 1

        chart_data_refs.append((ds_label, data_start, fila3 - 1, fila_header))
        fila3 += 2

    # ---- Hoja 4: Gráficos ----
    ws4 = wb.create_sheet("Gráficos")
    ws4.sheet_view.showGridLines = False
    ws4.merge_cells("A1:N1")
    estilar_celda(ws4, "A1", "GRÁFICOS COMPARATIVOS POR ESCENARIO",
                  bold=True, fondo=AZUL_HEADER, fuente_color=BLANCO, tamaño=13)
    ws4.row_dimensions[1].height = 28

    chart_positions = ["A3", "H3", "A22", "H22"]
    for idx, (ds_label, row_s, row_e, row_h) in enumerate(chart_data_refs):
        chart = BarChart()
        chart.type = "col"
        chart.title = ds_label
        chart.style = 10
        chart.grouping = "clustered"
        chart.width = 14
        chart.height = 10

        cats = Reference(ws3, min_col=1, min_row=row_s, max_row=row_e)
        for j in range(len(tipos)):
            data_ref = Reference(ws3, min_col=j + 2, min_row=row_h, max_row=row_e)
            chart.add_data(data_ref, titles_from_data=True)

        chart.set_categories(cats)
        chart.shape = 4
        ws4.add_chart(chart, chart_positions[idx])

    wb.save(path)