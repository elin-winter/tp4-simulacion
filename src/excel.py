from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image
import datetime

# ── Paleta ────────────────────────────────────────────────────────────────────
AZUL_OSCURO   = "0D2137"   # cabeceras principales
AZUL_MED      = "1A4A7A"   # cabeceras secundarias
AZUL_CLARO    = "2E86C1"   # acento
AZUL_SUAVE    = "D6EAF8"   # fondo alterno filas
GRIS_CLARO    = "F2F3F4"   # fondo filas pares
VERDE_OK      = "1E8449"   # texto "eficiente"
NARANJA_WARN  = "CA6F1E"   # texto "moderado"
ROJO_BAD      = "C0392B"   # texto "alto abandono"
BLANCO        = "FFFFFF"
GOLD          = "F4D03F"   # acento dorado para KPIs

ESCENARIOS_LABELS = ["Real", "Eficiente", "No eficiente"]

# ── Helpers ───────────────────────────────────────────────────────────────────

def _thin(color="BBBBBB"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _thick_bottom(color=AZUL_OSCURO):
    thick = Side(style="medium", color=color)
    thin  = Side(style="thin",   color="BBBBBB")
    return Border(left=thin, right=thin, top=thin, bottom=thick)

def cel(ws, ref, value=None, bold=False, size=10, color="000000",
        bg=None, align="center", border=False, fmt=None, italic=False,
        wrap=True):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = Font(name="Calibri", bold=bold, italic=italic,
                  size=size, color=color)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align, vertical="center",
                             wrap_text=wrap)
    if border:
        c.border = _thin()
    if fmt:
        c.number_format = fmt

def set_col_widths(ws, widths: dict):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

def set_row_height(ws, row, height):
    ws.row_dimensions[row].height = height

# ── Hoja 1: Resumen Comparativo ───────────────────────────────────────────────

def _hoja_resumen(wb, resultados, sim_time, seed, n_replicas):
    ws = wb.active
    ws.title = "Resumen"
    ws.sheet_view.showGridLines = False

    # ---- Título ----
    ws.merge_cells("A1:J1")
    cel(ws, "A1",
        "ANÁLISIS DE SENSIBILIDAD — ESTACIÓN DE CARGA VE",
        bold=True, size=15, color=BLANCO, bg=AZUL_OSCURO)
    set_row_height(ws, 1, 36)

    # ---- Subtítulo ----
    ws.merge_cells("A2:J2")
    cel(ws, "A2",
        f"  {n_replicas} réplicas  ·  {sim_time} min c/u  ·  Seed: {seed or 'aleatorio'}  ·  "
        f"Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}",
        italic=True, size=9, color=BLANCO, bg=AZUL_MED, align="left")
    set_row_height(ws, 2, 18)

    # ---- Banda vacía ----
    ws.merge_cells("A3:J3")
    ws["A3"].fill = PatternFill("solid", fgColor=AZUL_CLARO)
    set_row_height(ws, 3, 4)

    # ---- Encabezados de columnas ----
    HEADERS = [
        "Métrica", "Unidad",
        "Real\nCR", "Real\nCSR",
        "Eficiente\nCR", "Eficiente\nCSR",
        "No efic.\nCR", "No efic.\nCSR",
    ]
    ROW_H = 4
    set_row_height(ws, ROW_H, 38)
    for i, h in enumerate(HEADERS):
        ref = f"{get_column_letter(i+1)}{ROW_H}"
        cel(ws, ref, h, bold=True, size=10, color=BLANCO,
            bg=AZUL_MED, border=True)

    # ---- Datos ----
    METRICAS = [
        ("N° Cargadores",       "#",   "n_cargadores",        None,    False),
        ("Llegadas promedio",   "#",   "llegadas",            "0.0",   False),
        ("Atendidos promedio",  "#",   "atendidos",           "0.0",   False),
        ("Arrepentidos prom.",  "#",   "arrepentidos",        "0.0",   False),
        ("% Abandono",          "%",   "pct_arrepentidos",    "0.0%",  True),
        ("Espera promedio",     "min", "tiempo_espera_prom",  "0.00",  False),
        ("Tiempo en sistema",   "min", "tiempo_sistema_prom", "0.00",  False),
        ("Ociosidad cargadores","%",   "ociosidad_por_cargador","0.0%",True),
    ]

    fila = ROW_H + 1
    for idx, (label, unidad, key, fmt, is_pct) in enumerate(METRICAS):
        set_row_height(ws, fila, 20)
        row_bg = GRIS_CLARO if idx % 2 == 0 else BLANCO

        cel(ws, f"A{fila}", label, bold=True, size=10, bg=row_bg,
            align="left", border=True)
        cel(ws, f"B{fila}", unidad, size=9, color="666666",
            bg=row_bg, border=True)

        col = 3
        for r in resultados:
            for tipo in ["CR", "CSR"]:
                s = r[tipo]
                if s["n_cargadores"] == 0:
                    val = "–"
                else:
                    val = s[key]
                    if key == "pct_arrepentidos":
                        val /= 100
                    elif key == "ociosidad_por_cargador":
                        val = (sum(val) / len(val) / 100) if val else 0

                # color semáforo para % abandono
                txt_color = "000000"
                if key == "pct_arrepentidos" and val != "–":
                    if val > 0.20:
                        txt_color = ROJO_BAD
                    elif val > 0.10:
                        txt_color = NARANJA_WARN
                    else:
                        txt_color = VERDE_OK

                ref = f"{get_column_letter(col)}{fila}"
                cel(ws, ref, val, size=10, color=txt_color,
                    bg=row_bg, border=True,
                    fmt=fmt if val != "–" else None)
                col += 1
        fila += 1

    # ---- Fila: Abandono total ----
    set_row_height(ws, fila, 20)
    cel(ws, f"A{fila}", "Abandono total", bold=True, size=10,
        bg=AZUL_SUAVE, border=True)
    cel(ws, f"B{fila}", "%", size=9, color="666666",
        bg=AZUL_SUAVE, border=True)
    col = 3
    for r in resultados:
        total_l = r["CR"]["llegadas"] + r["CSR"]["llegadas"]
        total_a = r["CR"]["arrepentidos"] + r["CSR"]["arrepentidos"]
        val = (total_a / total_l) if total_l > 0 else 0
        for _ in ("CR", "CSR"):
            txt_color = ROJO_BAD if val > 0.20 else (NARANJA_WARN if val > 0.10 else VERDE_OK)
            cel(ws, f"{get_column_letter(col)}{fila}", val,
                size=10, bold=True, color=txt_color,
                bg=AZUL_SUAVE, border=True, fmt="0.0%")
            col += 1
    fila += 1

    # ---- Fila: Eficiencia global ----
    set_row_height(ws, fila, 22)
    cel(ws, f"A{fila}", "Eficiencia global (↓ mejor)", bold=True, size=10,
        bg=AZUL_OSCURO, color=BLANCO, border=True)
    cel(ws, f"B{fila}", "score", size=9, color=GOLD,
        bg=AZUL_OSCURO, border=True)
    col = 3
    for r in resultados:
        # CR column → score
        cel(ws, f"{get_column_letter(col)}{fila}",
            r["eficiencia_global"],
            size=11, bold=True, color=GOLD,
            bg=AZUL_OSCURO, border=True, fmt="0.000")
        # CSR column → vacío con mismo fondo
        cel(ws, f"{get_column_letter(col+1)}{fila}", "—",
            size=10, color="888888",
            bg=AZUL_OSCURO, border=True)
        col += 2

    # ---- Anchos de columnas ----
    set_col_widths(ws, {
        "A": 26, "B": 8,
        "C": 11, "D": 11,
        "E": 11, "F": 11,
        "G": 11, "H": 11,
        "I": 2,  "J": 2,
    })


# ── Hoja 2: KPIs visuales ─────────────────────────────────────────────────────

def _hoja_kpis(wb, resultados):
    ws = wb.create_sheet("KPIs")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:F1")
    cel(ws, "A1", "PANEL DE KPIs POR ESCENARIO",
        bold=True, size=14, color=BLANCO, bg=AZUL_OSCURO)
    set_row_height(ws, 1, 32)

    ws.merge_cells("A2:F2")
    cel(ws, "A2", "Comparación rápida entre escenarios simulados",
        italic=True, size=9, color=BLANCO, bg=AZUL_MED)
    set_row_height(ws, 2, 16)

    # Tarjetas por escenario (una cada 2 cols)
    kpi_keys = [
        ("Abandono total", "pct_arrepentidos", "0.0%", True),
        ("Espera prom. CR", "tiempo_espera_prom", "0.00", False),
        ("Ociosidad CR",    "ociosidad_por_cargador", "0.0%", True),
        ("Atendidos CR",    "atendidos",        "0.0", False),
    ]

    start_row = 4
    for esc_idx, r in enumerate(resultados):
        base_col = 1 + esc_idx * 2

        # Nombre escenario
        ws.merge_cells(
            start_row=start_row, start_column=base_col,
            end_row=start_row,   end_column=base_col + 1)
        ref = f"{get_column_letter(base_col)}{start_row}"
        cel(ws, ref, r["nombre"], bold=True, size=12,
            color=BLANCO, bg=AZUL_CLARO)
        set_row_height(ws, start_row, 26)

        for k_idx, (klabel, kkey, kfmt, is_pct) in enumerate(kpi_keys):
            row = start_row + 1 + k_idx * 3

            # Label
            ws.merge_cells(start_row=row, start_column=base_col,
                           end_row=row, end_column=base_col + 1)
            lref = f"{get_column_letter(base_col)}{row}"
            cel(ws, lref, klabel, size=9, color="555555",
                bg=GRIS_CLARO, align="left")
            set_row_height(ws, row, 16)

            # Valor
            ws.merge_cells(start_row=row+1, start_column=base_col,
                           end_row=row+1, end_column=base_col + 1)
            vref = f"{get_column_letter(base_col)}{row+1}"
            cr = r["CR"]
            if kkey == "pct_arrepentidos":
                val = cr[kkey] / 100
            elif kkey == "ociosidad_por_cargador":
                oc = cr[kkey]
                val = (sum(oc) / len(oc) / 100) if oc else 0
            else:
                val = cr.get(kkey, 0)

            txt_color = AZUL_CLARO
            if kkey == "pct_arrepentidos":
                txt_color = ROJO_BAD if val > 0.20 else (NARANJA_WARN if val > 0.10 else VERDE_OK)

            cel(ws, vref, val, bold=True, size=18, color=txt_color,
                bg=BLANCO, fmt=kfmt)
            set_row_height(ws, row+1, 30)

            # Separador
            set_row_height(ws, row+2, 5)
            for c in range(base_col, base_col + 2):
                ws.cell(row=row+2, column=c).fill = PatternFill(
                    "solid", fgColor=AZUL_SUAVE)

    set_col_widths(ws, {
        "A": 14, "B": 14,
        "C": 14, "D": 14,
        "E": 14, "F": 14,
    })


# ── Hoja 3: Insights ──────────────────────────────────────────────────────────

def _hoja_insights(wb, resultados):
    ws = wb.create_sheet("Insights")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:E1")
    cel(ws, "A1", "INSIGHTS AUTOMÁTICOS",
        bold=True, size=14, color=BLANCO, bg=AZUL_OSCURO)
    set_row_height(ws, 1, 32)

    row = 3
    for r in resultados:
        # Encabezado escenario
        ws.merge_cells(f"A{row}:E{row}")
        cel(ws, f"A{row}", f"📊  Escenario: {r['nombre']}",
            bold=True, size=12, color=BLANCO, bg=AZUL_MED, align="left")
        set_row_height(ws, row, 24)
        row += 1

        cr = r["CR"]; csr = r["CSR"]
        total_l = cr["llegadas"] + csr["llegadas"]
        total_a = cr["arrepentidos"] + csr["arrepentidos"]
        abandono = total_a / max(1, total_l)

        lines = [
            ("Abandono total:",        f"{abandono:.1%}",
             ROJO_BAD if abandono > 0.20 else (NARANJA_WARN if abandono > 0.10 else VERDE_OK)),
            ("Atendidos CR / CSR:",    f"{cr['atendidos']:.0f} / {csr['atendidos']:.0f}", AZUL_CLARO),
            ("Espera prom. CR:",       f"{cr['tiempo_espera_prom']:.1f} min", "333333"),
            ("Espera prom. CSR:",      f"{csr['tiempo_espera_prom']:.1f} min", "333333"),
            ("Eficiencia global:",     f"{r['eficiencia_global']:.3f}", AZUL_OSCURO),
        ]

        for label, value, vc in lines:
            set_row_height(ws, row, 18)
            cel(ws, f"A{row}", label,  bold=True, size=10,
                bg=GRIS_CLARO, align="left", border=True)
            ws.merge_cells(f"B{row}:C{row}")
            cel(ws, f"B{row}", value, bold=True, size=10,
                color=vc, bg=BLANCO, align="left", border=True)
            row += 1

        # Recomendación
        set_row_height(ws, row, 20)
        ws.merge_cells(f"A{row}:E{row}")
        if abandono > 0.20:
            rec = "⚠️  Alta pérdida de demanda — considerar más cargadores."
            bg_rec = "FADBD8"
        elif abandono > 0.10:
            rec = "⚡  Pérdida moderada — monitorear horas pico."
            bg_rec = "FDEBD0"
        else:
            rec = "✅  Sistema eficiente — demanda bien gestionada."
            bg_rec = "D5F5E3"
        cel(ws, f"A{row}", rec, italic=True, size=10,
            bg=bg_rec, align="left")
        row += 2

    set_col_widths(ws, {"A": 26, "B": 16, "C": 14, "D": 14, "E": 14})


# ── Hoja 4: Datos para gráficos ───────────────────────────────────────────────

def _hoja_graficos(wb, resultados):
    ws = wb.create_sheet("Gráficos")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:N1")
    cel(ws, "A1", "GRÁFICOS COMPARATIVOS",
        bold=True, size=14, color=BLANCO, bg=AZUL_OSCURO)
    set_row_height(ws, 1, 32)

    # ── Datos tabla para gráficos ──────────────────────────────────────────
    headers_g = ["Escenario", "% Abandono CR", "Espera CR (min)",
                 "Ociosidad CR (%)", "Atendidos CR"]
    for i, h in enumerate(headers_g):
        ref = f"{get_column_letter(i+1)}3"
        cel(ws, ref, h, bold=True, size=9, color=BLANCO, bg=AZUL_MED, border=True)
    set_row_height(ws, 3, 18)

    for idx, r in enumerate(resultados):
        row = 4 + idx
        cr = r["CR"]
        oc = cr["ociosidad_por_cargador"]
        ocio = (sum(oc) / len(oc) / 100) if oc else 0
        vals = [
            r["nombre"],
            cr["pct_arrepentidos"] / 100,
            cr["tiempo_espera_prom"],
            ocio,
            cr["atendidos"],
        ]
        bg = GRIS_CLARO if idx % 2 == 0 else BLANCO
        for i, v in enumerate(vals):
            ref = f"{get_column_letter(i+1)}{row}"
            fmt_map = [None, "0.0%", "0.00", "0.0%", "0.0"]
            cel(ws, ref, v, size=10, bg=bg, border=True, fmt=fmt_map[i])
        set_row_height(ws, row, 18)

    n = len(resultados)

    # ── Gráfico 1: Abandono ───────────────────────────────────────────────
    ch1 = BarChart()
    ch1.type = "col"
    ch1.title = "% Abandono por Escenario (CR)"
    ch1.style = 10
    ch1.grouping = "clustered"
    ch1.width = 16; ch1.height = 10
    ch1.y_axis.numFmt = "0%"
    ch1.y_axis.title = "Tasa de abandono"
    data1 = Reference(ws, min_col=2, min_row=3, max_row=3 + n)
    cats1 = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    ch1.add_data(data1, titles_from_data=True)
    ch1.set_categories(cats1)
    ch1.series[0].graphicalProperties.solidFill = AZUL_CLARO
    ws.add_chart(ch1, "G3")

    # ── Gráfico 2: Espera ─────────────────────────────────────────────────
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "Espera Promedio CR (min)"
    ch2.style = 10
    ch2.width = 16; ch2.height = 10
    ch2.y_axis.title = "Minutos"
    data2 = Reference(ws, min_col=3, min_row=3, max_row=3 + n)
    cats2 = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    ch2.add_data(data2, titles_from_data=True)
    ch2.set_categories(cats2)
    ch2.series[0].graphicalProperties.solidFill = "2E86C1"
    ws.add_chart(ch2, "G22")

    # ── Gráfico 3: Ociosidad ──────────────────────────────────────────────
    ch3 = BarChart()
    ch3.type = "col"
    ch3.title = "Ociosidad Promedio CR (%)"
    ch3.style = 10
    ch3.width = 16; ch3.height = 10
    ch3.y_axis.numFmt = "0%"
    ch3.y_axis.title = "Ociosidad"
    data3 = Reference(ws, min_col=4, min_row=3, max_row=3 + n)
    cats3 = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    ch3.add_data(data3, titles_from_data=True)
    ch3.set_categories(cats3)
    ch3.series[0].graphicalProperties.solidFill = "117A65"
    ws.add_chart(ch3, "G41")

    set_col_widths(ws, {
        "A": 18, "B": 14, "C": 14, "D": 14, "E": 14, "F": 2
    })


# ── Hoja 5: Eficiencia ────────────────────────────────────────────────────────

def _hoja_eficiencia(wb, resultados):
    ws = wb.create_sheet("Eficiencia")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D1")
    cel(ws, "A1", "RANKING DE EFICIENCIA GLOBAL",
        bold=True, size=14, color=BLANCO, bg=AZUL_OSCURO)
    set_row_height(ws, 1, 32)

    ws.merge_cells("A2:D2")
    cel(ws, "A2",
        "Score compuesto: 0.4×Espera + 0.3×Ociosidad + 0.2×Abandono − 0.1×Atendidos  (↓ mejor)",
        italic=True, size=8, color=BLANCO, bg=AZUL_MED, align="left")
    set_row_height(ws, 2, 16)

    hdrs = ["Pos.", "Escenario", "Score", "Evaluación"]
    for i, h in enumerate(hdrs):
        cel(ws, f"{get_column_letter(i+1)}4", h,
            bold=True, size=10, color=BLANCO, bg=AZUL_MED, border=True)
    set_row_height(ws, 4, 20)

    ranked = sorted(resultados, key=lambda r: r["eficiencia_global"])
    medals = ["🥇", "🥈", "🥉"]

    for rank, r in enumerate(ranked):
        row = 5 + rank
        set_row_height(ws, row, 22)
        bg = [BLANCO, GRIS_CLARO, AZUL_SUAVE][rank % 3]
        score = r["eficiencia_global"]
        eval_txt = "Óptimo" if rank == 0 else ("Aceptable" if rank == 1 else "Mejorable")
        eval_col = [VERDE_OK, NARANJA_WARN, ROJO_BAD][rank]

        cel(ws, f"A{row}", f"{medals[rank]} #{rank+1}",
            bold=True, size=12, bg=bg, border=True)
        cel(ws, f"B{row}", r["nombre"],
            bold=True, size=11, bg=bg, align="left", border=True)
        cel(ws, f"C{row}", score, bold=True, size=12,
            color=AZUL_OSCURO, bg=bg, border=True, fmt="0.000")
        cel(ws, f"D{row}", eval_txt, bold=True, size=11,
            color=eval_col, bg=bg, border=True)

    # gráfico horizontal
    n = len(resultados)
    # tabla auxiliar
    aux_row = 5 + n + 2
    ws[f"A{aux_row}"] = "Escenario"
    ws[f"B{aux_row}"] = "Score"
    for i, r in enumerate(ranked):
        ws[f"A{aux_row+1+i}"] = r["nombre"]
        ws[f"B{aux_row+1+i}"] = r["eficiencia_global"]

    ch = BarChart()
    ch.type = "bar"          # barras horizontales
    ch.title = "Score de Eficiencia (↓ mejor)"
    ch.style = 10
    ch.width = 18; ch.height = 10
    data = Reference(ws, min_col=2, min_row=aux_row, max_row=aux_row + n)
    cats = Reference(ws, min_col=1, min_row=aux_row + 1, max_row=aux_row + n)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.solidFill = AZUL_CLARO
    ws.add_chart(ch, "F4")

    set_col_widths(ws, {"A": 10, "B": 18, "C": 12, "D": 14})


# ── Entry point ───────────────────────────────────────────────────────────────

def generar_excel(resultados, path, sim_time, seed, n_replicas):
    wb = Workbook()

    _hoja_resumen(wb, resultados, sim_time, seed, n_replicas)
    _hoja_kpis(wb, resultados)
    _hoja_insights(wb, resultados)
    _hoja_graficos(wb, resultados)
    _hoja_eficiencia(wb, resultados)

    wb.save(path)
    print(f"✓ Excel guardado en: {path}")