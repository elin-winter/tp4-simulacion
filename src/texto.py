import datetime


def generar_txt(resultados, path, sim_time, seed):
    lines = []
    sep = "=" * 55
    lines.append(sep)
    lines.append("  REPORTE DE SIMULACIÓN — ESTACIÓN DE CARGA VE")
    lines.append(f"  Generado: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"  Duración simulación: {sim_time} min | Seed: {seed}")
    lines.append(sep)

    for r in resultados:
        lines.append(f"\n{'#'*55}")
        lines.append(f"  ESCENARIO: {r['nombre'].upper()}")
        lines.append(f"  Cargadores Rápidos (CR): {r['CCR']} | Semi-Rápidos (CSR): {r['CCSR']}")
        lines.append(f"  Prob. uso CR: {r['PROB_CR']:.0%}")
        lines.append(f"{'#'*55}")

        for tipo_key, label in [("CR", "CARGADORES RÁPIDOS"), ("CSR", "CARGADORES SEMI-RÁPIDOS")]:
            s = r[tipo_key]
            lines.append(f"\n  --- {label} ---")
            lines.append(f"  Total llegadas:           {s['llegadas']}")
            lines.append(f"  Total atendidos:          {s['atendidos']}")
            lines.append(f"  Arrepentidos:             {s['arrepentidos']} ({s['pct_arrepentidos']:.1f}%)")
            lines.append(f"  Tiempo promedio espera:   {s['tiempo_espera_prom']:.2f} min")
            lines.append(f"  Tiempo promedio sistema:  {s['tiempo_sistema_prom']:.2f} min")
            lines.append(f"  Ociosidad por cargador:")
            for i, oc in enumerate(s['ociosidad_por_cargador']):
                lines.append(f"    {tipo_key}[{i}]: {oc:.1f}%")

    lines.append(f"\n{sep}")
    lines.append("  FIN DEL REPORTE")
    lines.append(sep)

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))