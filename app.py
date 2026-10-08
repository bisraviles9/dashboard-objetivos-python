from __future__ import annotations

from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard de objetivos", page_icon="🎯", layout="wide")

GOAL_COLUMNS = ["Área", "Objetivo", "Responsable", "Meta", "Actual", "Unidad", "Fecha límite", "Trimestre"]
MONTH_COLUMNS = ["Mes", "Área", "Objetivo", "Avance mensual", "Unidad"]

DEMO_GOALS = pd.DataFrame([
    {"Área": "Ventas", "Objetivo": "Aumentar ventas", "Responsable": "Ana", "Meta": 100000, "Actual": 68000, "Unidad": "USD", "Fecha límite": "2026-12-31", "Trimestre": "T4 2026"},
    {"Área": "Marketing", "Objetivo": "Generar oportunidades", "Responsable": "Luis", "Meta": 500, "Actual": 245, "Unidad": "leads", "Fecha límite": "2026-12-31", "Trimestre": "T4 2026"},
    {"Área": "Producto", "Objetivo": "Entregar mejoras", "Responsable": "Marta", "Meta": 12, "Actual": 8, "Unidad": "entregas", "Fecha límite": "2026-12-31", "Trimestre": "T4 2026"},
    {"Área": "Atención al cliente", "Objetivo": "Mejorar satisfacción", "Responsable": "Diego", "Meta": 90, "Actual": 74, "Unidad": "%", "Fecha límite": "2026-12-31", "Trimestre": "T4 2026"},
    {"Área": "Operaciones", "Objetivo": "Reducir tiempo de entrega", "Responsable": "Sofía", "Meta": 30, "Actual": 12, "Unidad": "% reducción", "Fecha límite": "2026-12-31", "Trimestre": "T4 2026"},
])

DEMO_MONTHS = pd.DataFrame([
    {"Mes": m, "Área": area, "Objetivo": objective, "Avance mensual": value, "Unidad": unit}
    for m, area, objective, unit, values in [
        ("Enero", "Ventas", "Aumentar ventas", "USD", [4000, 4500, 5200, 5600, 6000, 6500, 7000, 7200, 8000]),
        ("Enero", "Marketing", "Generar oportunidades", "leads", [22, 25, 28, 30, 32, 35, 36, 38, 42]),
        ("Enero", "Producto", "Entregar mejoras", "entregas", [1, 1, 1, 1, 1, 1, 1, 0, 0]),
        ("Enero", "Atención al cliente", "Mejorar satisfacción", "%", [70, 71, 72, 72, 73, 73, 74, 74, 74]),
    ]
    for m, value in zip(["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre"], values)
])


def clean_goals(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in GOAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[GOAL_COLUMNS]
    df["Meta"] = pd.to_numeric(df["Meta"], errors="coerce").fillna(0)
    df["Actual"] = pd.to_numeric(df["Actual"], errors="coerce").fillna(0)
    df["Fecha límite"] = df["Fecha límite"].astype(str)
    return df


def workbook_bytes(goals: pd.DataFrame, months: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        goals.to_excel(writer, index=False, sheet_name="Objetivos")
        months.to_excel(writer, index=False, sheet_name="Seguimiento mensual")
        summary = goals.copy()
        summary["Progreso"] = (summary["Actual"] / summary["Meta"].replace(0, pd.NA)).fillna(0)
        summary["Progreso"] = summary["Progreso"].clip(lower=0, upper=1)
        summary[["Área", "Objetivo", "Responsable", "Meta", "Actual", "Unidad", "Progreso", "Fecha límite", "Trimestre"]].to_excel(writer, index=False, sheet_name="Resumen")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                letter = cells[0].column_letter
                width = min(max(max(len(str(c.value or "")) for c in cells) + 2, 12), 34)
                ws.column_dimensions[letter].width = width
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="E4793D")
        summary_ws = writer.book["Resumen"]
        for row in range(2, summary_ws.max_row + 1):
            summary_ws.cell(row, 7).number_format = "0%"
    return output.getvalue()


st.title("🎯 Dashboard de objetivos")
st.caption("Revisa el avance por área, detecta objetivos atrasados y proyecta el cierre. Los datos iniciales son de ejemplo y puedes reemplazarlos.")

with st.sidebar:
    st.header("Datos")
    uploaded = st.file_uploader("Carga tu archivo Excel", type=["xlsx"], help="Se esperan las hojas 'Objetivos' y, opcionalmente, 'Seguimiento mensual'.")
    if "goals" not in st.session_state:
        st.session_state.goals = DEMO_GOALS.copy()
        st.session_state.months = DEMO_MONTHS.copy()
    if uploaded is not None:
        try:
            sheets = pd.read_excel(uploaded, sheet_name=None)
            if "Objetivos" not in sheets:
                st.error("El archivo debe incluir una hoja llamada 'Objetivos'.")
            else:
                st.session_state.goals = clean_goals(sheets["Objetivos"])
                st.session_state.months = sheets.get("Seguimiento mensual", pd.DataFrame(columns=MONTH_COLUMNS))
                st.success("Archivo cargado.")
        except Exception as exc:
            st.error(f"No pude leer ese Excel: {exc}")
    st.divider()
    st.info("Edita la tabla en la pestaña 'Objetivos' y descarga los cambios como Excel.")

goals = clean_goals(st.session_state.goals)
goals["Progreso"] = (goals["Actual"] / goals["Meta"].replace(0, pd.NA)).fillna(0).clip(lower=0, upper=1)
goals["Diferencia"] = (goals["Meta"] - goals["Actual"]).clip(lower=0)

if goals.empty:
    st.warning("Aún no hay objetivos. Agrega uno en la pestaña Objetivos.")
else:
    overall = goals["Progreso"].mean()
    behind = goals.sort_values("Progreso").iloc[0]
    finished = int((goals["Actual"] >= goals["Meta"]).sum())
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Avance promedio", f"{overall:.0%}")
    col2.metric("Objetivos", len(goals))
    col3.metric("Completados", f"{finished}/{len(goals)}")
    col4.metric("Más atrasado", str(behind["Área"]), f"{behind['Progreso']:.0%} de avance")

    tab_dashboard, tab_objectives, tab_monthly = st.tabs(["Resumen", "Objetivos", "Seguimiento mensual"])
    with tab_dashboard:
        left, right = st.columns([1.35, 1])
        with left:
            st.subheader("Progreso por área")
            chart = goals.set_index("Área")[["Progreso"]].sort_values("Progreso")
            st.bar_chart(chart, x_label="Área", y_label="Progreso (0–1)", color="#E4793D")
        with right:
            st.subheader("Cumplimiento trimestral")
            quarter = goals.groupby("Trimestre", dropna=False).agg(Total=("Objetivo", "count"), Completados=("Progreso", lambda x: int((x >= 1).sum())))
            quarter["Avance"] = goals.groupby("Trimestre", dropna=False)["Progreso"].mean()
            st.dataframe(quarter.reset_index(), hide_index=True, use_container_width=True, column_config={"Avance": st.column_config.ProgressColumn("Avance promedio", min_value=0, max_value=1, format="percent")})
        st.subheader("Proyección de cierre")
        st.caption("Estimación lineal orientativa: ritmo de avance acumulado hasta hoy frente a los meses del año. Ajusta la meta y el avance para que refleje tu periodo real.")
        today = date.today()
        elapsed = max(today.month, 1)
        projected_factor = min(12 / elapsed, 2.0)
        projection = goals[["Área", "Objetivo", "Meta", "Actual", "Unidad"]].copy()
        projection["Proyección"] = (projection["Actual"] * projected_factor).round(1)
        projection["Cierre estimado"] = (projection["Proyección"] / projection["Meta"].replace(0, pd.NA)).fillna(0).clip(lower=0, upper=2)
        st.dataframe(projection, hide_index=True, use_container_width=True, column_config={"Cierre estimado": st.column_config.ProgressColumn("Cierre estimado", min_value=0, max_value=1, format="percent")})
        st.subheader("¿Quién necesita apoyo?")
        behind_list = goals.sort_values("Progreso").head(3)[["Área", "Objetivo", "Responsable", "Progreso", "Diferencia", "Unidad"]]
        st.dataframe(behind_list, hide_index=True, use_container_width=True, column_config={"Progreso": st.column_config.ProgressColumn("Avance", min_value=0, max_value=1, format="percent")})

    with tab_objectives:
        st.subheader("Objetivos y avances")
        edited = st.data_editor(
            goals[GOAL_COLUMNS], use_container_width=True, hide_index=True, num_rows="dynamic",
            column_config={
                "Meta": st.column_config.NumberColumn("Meta", min_value=0),
                "Actual": st.column_config.NumberColumn("Actual", min_value=0),
                "Fecha límite": st.column_config.TextColumn("Fecha límite (AAAA-MM-DD)"),
                "Trimestre": st.column_config.TextColumn("Trimestre"),
            },
            key="objective_editor",
        )
        st.session_state.goals = clean_goals(edited)
        st.caption("Haz clic en una celda para cambiarla. Puedes agregar filas al final de la tabla.")
        with st.form("add_goal", clear_on_submit=True):
            st.markdown("**Agregar un objetivo**")
            a, b, c = st.columns(3)
            area = a.text_input("Área")
            objective = b.text_input("Objetivo")
            owner = c.text_input("Responsable")
            d, e, f, g = st.columns(4)
            target = d.number_input("Meta", min_value=0.0, value=1.0)
            actual = e.number_input("Avance actual", min_value=0.0, value=0.0)
            unit = f.text_input("Unidad", value="unidades")
            due = g.date_input("Fecha límite", value=date.today())
            quarter_name = st.text_input("Trimestre", value=f"T{(due.month - 1) // 3 + 1} {due.year}")
            if st.form_submit_button("Agregar objetivo"):
                if area.strip() and objective.strip():
                    new = {"Área": area.strip(), "Objetivo": objective.strip(), "Responsable": owner.strip(), "Meta": target, "Actual": actual, "Unidad": unit.strip(), "Fecha límite": due.isoformat(), "Trimestre": quarter_name.strip()}
                    st.session_state.goals = pd.concat([st.session_state.goals, pd.DataFrame([new])], ignore_index=True)
                    st.rerun()
                else:
                    st.error("Completa al menos el área y el objetivo.")

    with tab_monthly:
        st.subheader("Evolución del avance mensual")
        months = st.session_state.months.copy()
        if not months.empty and {"Mes", "Área", "Avance mensual"}.issubset(months.columns):
            month_filter = st.selectbox("Área", ["Todas"] + sorted(months["Área"].dropna().astype(str).unique().tolist()))
            if month_filter != "Todas":
                months = months[months["Área"] == month_filter]
            pivot = months.pivot_table(index="Mes", columns="Área", values="Avance mensual", aggfunc="sum").fillna(0)
            st.line_chart(pivot, y_label="Avance registrado")
            st.dataframe(months, hide_index=True, use_container_width=True)
        else:
            st.info("No hay seguimiento mensual aún. Carga un Excel con la hoja 'Seguimiento mensual' o empieza por la pestaña Objetivos.")

st.divider()
excel = workbook_bytes(clean_goals(st.session_state.goals), st.session_state.months)
st.download_button("⬇️ Descargar proyecto de objetivos en Excel", data=excel, file_name="dashboard_objetivos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
st.caption("El archivo incluye resumen, objetivos editables y seguimiento mensual. Los datos que ves al abrir por primera vez son demostrativos.")
