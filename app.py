import io
from datetime import datetime, timedelta, date
import pandas as pd
import plotly.express as px
import streamlit as st

# =============================
# Configuração da página
# =============================
st.set_page_config(
    page_title="Dashboard de Manutenção Preventiva",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================
# Constantes
# =============================
PRED_COLS = [
    "Pred 1 (15d)",
    "Pred 2 (30d)",
    "Pred 3 (45d)",
    "Pred 4 (60d)",
    "Pred 5 (75d)",
    "Pred 6 (90d)",
    "Pred 7 (105d)",
]

REQUIRED_BASE_COLUMNS = [
    "PLACA",
    "Última Revisão",
    "Data da Próxima Revisão",
    "Intervalo de Revisão",
    *PRED_COLS,
]

EXTRA_COLUMNS_DEFAULTS = {
    "Preventiva Concluída": "PENDENTE",
    "Data da Preventiva Realizada": pd.NaT,
    "Ciclos de Preventiva Realizados": 0,
    "Observações": "",
    "Status Preventiva": "EM DIA",
    "Status Geral": "EM DIA",
    "Dias p/ Próxima": 0,
    "Faixa": "",
    "Progresso": 0.0,
    "Qtd Preditivas Realizadas": 0,
    "Qtd Preditivas Previstas Hoje": 0,
    "Qtd Preditivas Em Dia": 0,
    "Qtd Preditivas Atrasadas": 0,
    "Preditivas Pendentes": 0,
    "Próxima Preditiva Prevista": "-",
    "Pode Confirmar Preventiva": "NÃO",
}

# =============================
# Funções utilitárias
# =============================
def normalizar_resposta(valor):
    if pd.isna(valor):
        return "NÃO"
    valor = str(valor).strip().upper()
    return "SIM" if valor in {"SIM", "S", "TRUE", "1", "REALIZADA", "REALIZADO"} else "NÃO"


def carregar_arquivo(uploaded_file):
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    faltantes = [c for c in REQUIRED_BASE_COLUMNS if c not in df.columns]
    if faltantes:
        st.error(f"Colunas obrigatórias ausentes: {', '.join(faltantes)}")
        return None

    for col in ["Última Revisão", "Data da Próxima Revisão", "Data da Preventiva Realizada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col, default in EXTRA_COLUMNS_DEFAULTS.items():
        if col not in df.columns:
            df[col] = default

    for col in PRED_COLS:
        df[col] = df[col].apply(normalizar_resposta)

    df = df.drop_duplicates(subset=["PLACA"], keep="first").copy()
    return recalcular_indicadores(df)


def calcular_info_linha(row):
    hoje = pd.Timestamp(date.today())
    ultima = row["Última Revisão"]
    proxima = row["Data da Próxima Revisão"]
    intervalo = int(row["Intervalo de Revisão"] or 120)

    if pd.isna(ultima):
        ultima = proxima - pd.Timedelta(days=intervalo)

    dias_proxima = int((proxima.normalize() - hoje.normalize()).days)
    dias_decorridos = max(int((hoje.normalize() - ultima.normalize()).days), 0)

    previstas = min(dias_decorridos // 15, 7)
    realizadas = sum(row[c] == "SIM" for c in PRED_COLS)
    atrasadas = max(previstas - realizadas, 0)

    if dias_proxima < 0:
        faixa = "Atrasada"
    elif dias_proxima <= 15:
        faixa = "0-15 dias"
    elif dias_proxima <= 30:
        faixa = "16-30 dias"
    elif dias_proxima <= 60:
        faixa = "31-60 dias"
    elif dias_proxima <= 120:
        faixa = "61-120 dias"
    else:
        faixa = ">120 dias"

    status_prev = "ATRASADA" if dias_proxima < 0 else "EM DIA"
    status_geral = "PREDITIVA ATRASADA" if atrasadas > 0 else status_prev

    return {
        "Dias p/ Próxima": dias_proxima,
        "Faixa": faixa,
        "Qtd Preditivas Realizadas": realizadas,
        "Qtd Preditivas Atrasadas": atrasadas,
        "Status Preventiva": status_prev,
        "Status Geral": status_geral,
    }


def recalcular_indicadores(df):
    extras = df.apply(lambda r: pd.Series(calcular_info_linha(r)), axis=1)
    for c in extras.columns:
        df[c] = extras[c]
    return df


# =============================
# Cards (KPIs)
# =============================
def mostrar_metricas(df):
    total_ativos = df["PLACA"].nunique()
    preventivas_em_dia = (df["Status Preventiva"] == "EM DIA").sum()
    preventivas_atrasadas = (df["Status Preventiva"] == "ATRASADA").sum()
    preditivas_em_dia = (df["Qtd Preditivas Atrasadas"] == 0).sum()
    pct_em_dia = (preventivas_em_dia / total_ativos * 100) if total_ativos else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de Ativos", total_ativos)
    c2.metric("Preventivas em Dia", preventivas_em_dia)
    c3.metric("Preditivas em Dia", preditivas_em_dia)
    c4.metric("Preventivas Atrasadas", preventivas_atrasadas)
    c5.metric("% Preventivas em Dia", f"{pct_em_dia:.1f}%")


# =============================
# Dashboard
# =============================
def dashboard(df):
    st.title("🚛 Dashboard de Manutenção Preventiva")
    st.caption("Gestão baseada em meta e projeção de conclusão das preventivas")

    # ===== GRÁFICO PRINCIPAL =====
    st.subheader("🎯 Projeção de conclusão das preventivas (GRÁFICO PRINCIPAL)")

    pendentes = int((df["Preventiva Concluída"] != "SIM").sum())

    if pendentes == 0:
        st.success("✅ Todas as preventivas já foram concluídas!")
    else:
        meta_por_dia = 3
        ritmo_atual = 1

        dias_meta = int((pendentes / meta_por_dia) + 0.999)
        dias_atual = int((pendentes / ritmo_atual) + 0.999)

        dias = list(range(0, max(dias_meta, dias_atual) + 1))

        curva_meta = [max(pendentes - meta_por_dia * d, 0) for d in dias]
        curva_atual = [max(pendentes - ritmo_atual * d, 0) for d in dias]

        df_proj = pd.DataFrame({
            "Dias": dias,
            "Meta (3/dia)": curva_meta,
            "Ritmo Atual": curva_atual,
        })

        fig = px.line(
            df_proj,
            x="Dias",
            y=["Meta (3/dia)", "Ritmo Atual"],
            markers=True,
        )

        fig.update_layout(
            xaxis_title="Dias a partir de hoje",
            yaxis_title="Preventivas pendentes",
        )

        st.plotly_chart(fig, use_container_width=True)

        st.info(
            f"""
            🔹 **Preventivas pendentes:** {pendentes}  
            🔹 **Com a meta (3/dia):** ~{dias_meta} dias  
            🔹 **No ritmo atual:** ~{dias_atual} dias  
            """
        )

    st.divider()

    # ===== CARDS =====
    mostrar_metricas(df)

    st.divider()

    # ===== GRÁFICO DE FAIXA =====
    st.subheader("Quantidade de ativos por faixa de vencimento")

    faixa_order = ["Atrasada", "0-15 dias", "16-30 dias", "31-60 dias", "61-120 dias", ">120 dias"]
    faixa_counts = df["Faixa"].value_counts().reset_index()
    faixa_counts.columns = ["Faixa", "Quantidade"]

    fig_faixa = px.bar(
        faixa_counts,
        x="Faixa",
        y="Quantidade",
        text="Quantidade",
        category_orders={"Faixa": faixa_order},
    )

    st.plotly_chart(fig_faixa, use_container_width=True)


# =============================
# Cadastro / Atualização
# =============================
def pagina_cadastro(df):
    st.title("🛠️ Cadastro / Atualização")

    placa = st.selectbox("Selecione a carreta", df["PLACA"].tolist())
    idx = df.index[df["PLACA"] == placa][0]
    row = df.loc[idx]

    with st.form("form_preditivas"):
        novos = {}
        for col in PRED_COLS:
            novos[col] = st.checkbox(col, value=row[col] == "SIM")
        obs = st.text_area("Observações", value=row["Observações"])
        salvar = st.form_submit_button("Salvar")

    if salvar:
        for col, val in novos.items():
            df.at[idx, col] = "SIM" if val else "NÃO"
        df.at[idx, "Observações"] = obs
        df = recalcular_indicadores(df)
        st.session_state["df"] = df
        st.success("Atualizado com sucesso")
        st.rerun()


# =============================
# Main
# =============================
def main():
    st.sidebar.title("📂 Dados")
    uploaded = st.sidebar.file_uploader("Selecione a planilha .xlsx", type=["xlsx"])

    if uploaded is None:
        st.info("Envie uma planilha para iniciar.")
        return

    if "df" not in st.session_state:
        st.session_state["df"] = carregar_arquivo(uploaded)

    df = st.session_state["df"]

    pagina = st.sidebar.radio("Navegação", ["Dashboard", "Cadastro / Atualização"])

    if pagina == "Dashboard":
        dashboard(df)
    else:
        pagina_cadastro(df)


if __name__ == "__main__":
    main()
