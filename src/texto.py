import datetime


# ── Utilidades de formato ─────────────────────────────────────────────────────

W = 68  # ancho total del reporte

def line(char="─", w=W):
    return char * w

def box_top(w=W):    return f"╔{'═' * (w-2)}╗"
def box_mid(w=W):    return f"╠{'═' * (w-2)}╣"
def box_bot(w=W):    return f"╚{'═' * (w-2)}╝"
def box_row(text, w=W):
    inner = w - 4
    return f"║  {text:<{inner}}║"

def kv(label, value, w=W, unit=""):
    dots = w - 4 - len(label) - len(str(value)) - len(unit) - 2
    return f"{label} {'·' * max(dots, 1)} {value}{' ' + unit if unit else ''}"

def bar(pct, width=30, char="█", empty="░"):
    filled = round(pct / 100 * width)
    return char * filled + empty * (width - filled)

def semaforo(pct):
    if pct > 20:  return "🔴"
    if pct > 10:  return "🟡"
    return "🟢"

def ocio_semaforo(pct):
    if pct > 40: return "⚠  sobrecapacidad"
    if pct < 10: return "⚡  alta utilización"
    return "✓  equilibrado"

def espera_tag(min_):
    if min_ > 30: return "⚠  elevada"
    if min_ > 10: return "~  moderada"
    return "✓  baja"

def pct_bar_line(label, pct, w=W):
    b = bar(min(pct, 100))
    tag = f"{pct:5.1f}%"
    return f"  {label:<6} {b} {tag}"

def center(text, w=W):
    return text.center(w)

def blank():
    return ""


# ── Bloque de tipo de cargador ────────────────────────────────────────────────

def bloque_cargador(lines, s, tipo_key, label):
    lines.append(blank())
    lines.append(f"  ┌─ {label} {'─'*(W-7-len(label))}┐")

    if s["n_cargadores"] == 0:
        lines.append(f"  │  Sin cargadores en este escenario.")
        lines.append(f"  └{'─'*(W-4)}┘")
        return

    oc_list = s["ociosidad_por_cargador"]
    oc_prom = sum(oc_list) / len(oc_list) if oc_list else 0
    arrep_pct = s["pct_arrepentidos"]
    espera    = s["tiempo_espera_prom"]

    lines.append(kv("  Cargadores",       str(s["n_cargadores"])))
    lines.append(kv("  Llegadas prom.",    f"{s['llegadas']:.1f}"))
    lines.append(kv("  Atendidos prom.",   f"{s['atendidos']:.1f}"))
    lines.append(kv("  Arrepentidos prom.",f"{s['arrepentidos']:.1f}"))
    lines.append(blank())

    lines.append(f"  Abandono   {semaforo(arrep_pct)}")
    lines.append(pct_bar_line(tipo_key, arrep_pct))

    oc_bar_pct = min(oc_prom, 100)
    lines.append(f"  Ociosidad  {ocio_semaforo(oc_prom)}")
    lines.append(pct_bar_line(tipo_key, oc_bar_pct))
    lines.append(blank())

    lines.append(kv("  Espera promedio",   f"{espera:.2f}", unit=f"min  {espera_tag(espera)}"))
    lines.append(kv("  Tiempo en sistema", f"{s['tiempo_sistema_prom']:.2f}", unit="min"))
    lines.append(blank())

    lines.append(f"  Ociosidad por cargador:")
    for i, oc in enumerate(oc_list):
        b = bar(min(oc, 100), width=20)
        lines.append(f"    [{tipo_key}{i}]  {b}  {oc:.1f}%")

    lines.append(f"  └{'─'*(W-4)}┘")


# ── Insights del escenario ────────────────────────────────────────────────────

def insights_escenario(lines, r):
    cr  = r["CR"]
    csr = r["CSR"]

    bullets = []

    for s, tipo in [(cr, "CR"), (csr, "CSR")]:
        if s["n_cargadores"] == 0:
            continue
        p = s["pct_arrepentidos"]
        if p > 20:
            bullets.append(f"🔴  {tipo}: alta pérdida de demanda ({p:.1f}% abandono) — considerar más cargadores.")
        elif p > 10:
            bullets.append(f"🟡  {tipo}: pérdida moderada ({p:.1f}%) — monitorear horas pico.")
        else:
            bullets.append(f"🟢  {tipo}: demanda bien gestionada ({p:.1f}% abandono).")

        e = s["tiempo_espera_prom"]
        if e > 30:
            bullets.append(f"⚠   {tipo}: tiempos de espera elevados ({e:.1f} min).")
        elif e > 10:
            bullets.append(f"~   {tipo}: tiempos de espera moderados ({e:.1f} min).")

        oc = s["ociosidad_por_cargador"]
        if oc:
            prom = sum(oc) / len(oc)
            if prom > 40:
                bullets.append(f"⚠   {tipo}: alta ociosidad ({prom:.1f}%) — posible sobrecapacidad.")
            elif prom < 10:
                bullets.append(f"⚡  {tipo}: cargadores muy utilizados ({prom:.1f}% ociosidad).")

    lines.append(blank())
    lines.append(f"  ╌╌╌  INSIGHTS  {'╌'*(W-18)}╌")
    for b in bullets:
        lines.append(f"  {b}")


# ── Comparación global ────────────────────────────────────────────────────────

def tabla_comparacion(lines, resultados):
    lines.append(blank())
    lines.append(line("═"))
    lines.append(center("COMPARACIÓN GLOBAL DE ESCENARIOS"))
    lines.append(line("═"))
    lines.append(blank())

    col_esc = 18
    col_val = 12
    hdr = (f"  {'Escenario':<{col_esc}}"
           f"{'Abandono':>{col_val}}"
           f"{'Espera CR':>{col_val}}"
           f"{'Ociosidad':>{col_val}}"
           f"{'Eficiencia':>{col_val}}")
    lines.append(hdr)
    lines.append(f"  {'─'*col_esc}{'─'*col_val}{'─'*col_val}{'─'*col_val}{'─'*col_val}")

    ranked = sorted(resultados, key=lambda r: r["eficiencia_global"])
    medals = ["🥇", "🥈", "🥉"]

    for r in resultados:
        cr  = r["CR"]
        csr = r["CSR"]
        total_l = cr["llegadas"] + csr["llegadas"]
        total_a = cr["arrepentidos"] + csr["arrepentidos"]
        pct_ab  = (total_a / total_l * 100) if total_l > 0 else 0

        oc = cr["ociosidad_por_cargador"]
        oc_prom = sum(oc) / len(oc) if oc else 0

        rank_pos = next(i for i, x in enumerate(ranked) if x["nombre"] == r["nombre"])
        medal = medals[rank_pos] if rank_pos < 3 else "  "

        row = (f"  {medal} {r['nombre']:<{col_esc-3}}"
               f"{semaforo(pct_ab)} {pct_ab:>6.1f}%  "
               f"{cr['tiempo_espera_prom']:>7.1f} min  "
               f"{oc_prom:>6.1f}%    "
               f"{r['eficiencia_global']:>8.4f}")
        lines.append(row)

    lines.append(blank())

    lines.append(f"  Abandono total (barras comparativas):")
    for r in resultados:
        cr  = r["CR"];  csr = r["CSR"]
        total_l = cr["llegadas"] + csr["llegadas"]
        total_a = cr["arrepentidos"] + csr["arrepentidos"]
        pct = (total_a / total_l * 100) if total_l > 0 else 0
        b = bar(min(pct, 100), width=35)
        lines.append(f"    {r['nombre']:<14} {b} {pct:.1f}%")

    lines.append(blank())

    lines.append(f"  Score de eficiencia (↓ mejor):")
    scores = [r["eficiencia_global"] for r in resultados]
    max_s = max(scores) or 1
    for r in resultados:
        pct_s = r["eficiencia_global"] / max_s * 100
        b = bar(pct_s, width=35)
        lines.append(f"    {r['nombre']:<14} {b} {r['eficiencia_global']:.4f}")


# ── Conclusión ────────────────────────────────────────────────────────────────

def conclusion(lines, resultados):
    ranked = sorted(resultados, key=lambda r: r["eficiencia_global"])
    mejor  = ranked[0]

    lines.append(blank())
    lines.append(line("═"))
    lines.append(center("CONCLUSIÓN"))
    lines.append(line("═"))
    lines.append(blank())
    lines.append(f"  El escenario con mejor desempeño global es: 【 {mejor['nombre'].upper()} 】")
    lines.append(f"  (score de eficiencia: {mejor['eficiencia_global']:.4f})")
    lines.append(blank())
    lines.append("  Los resultados corresponden a promedios de múltiples réplicas")
    lines.append("  independientes, lo que reduce la variabilidad y brinda una base")
    lines.append("  sólida para decisiones sobre la cantidad y tipo de cargadores.")
    lines.append(blank())
    lines.append(line("─"))
    lines.append(center("FIN DEL REPORTE"))
    lines.append(line("─"))


# ── Entry point ───────────────────────────────────────────────────────────────

def generar_txt(resultados, path, sim_time, seed, n_replicas):
    lines = []

    lines.append(box_top())
    lines.append(box_row(""))
    lines.append(box_row(center("REPORTE DE SIMULACIÓN", W-4).strip()))
    lines.append(box_row(center("ESTACIÓN DE CARGA — VEHÍCULOS ELÉCTRICOS", W-4).strip()))
    lines.append(box_row(""))
    lines.append(box_mid())
    lines.append(box_row(f"Generado :  {datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}"))
    lines.append(box_row(f"Réplicas :  {n_replicas}  ×  {sim_time} min c/u"))
    lines.append(box_row(f"Seed base:  {seed or 'aleatorio'}"))
    lines.append(box_row(""))
    lines.append(box_row("Todos los valores son promedios de corridas independientes."))
    lines.append(box_row(""))
    lines.append(box_bot())

    for r in resultados:
        lines.append(blank())
        lines.append(blank())
        lines.append(line("▓"))
        lines.append(center(f"  ESCENARIO: {r['nombre'].upper()}  "))
        lines.append(line("▓"))
        lines.append(kv("  CR",    str(r["CCR"]),          unit="cargadores rápidos"))
        lines.append(kv("  CSR",   str(r["CCSR"]),         unit="cargadores semi-rápidos"))
        lines.append(kv("  P(CR)", f"{r['PROB_CR']:.0%}",  unit="prob. de elegir CR"))

        bloque_cargador(lines, r["CR"],  "CR",  "CARGADORES RÁPIDOS (CR)")
        bloque_cargador(lines, r["CSR"], "CSR", "CARGADORES SEMI-RÁPIDOS (CSR)")
        insights_escenario(lines, r)

    tabla_comparacion(lines, resultados)
    conclusion(lines, resultados)

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))