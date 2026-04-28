
import io
from datetime import datetime, timedelta, date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Manutenção Preventiva",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded",
)

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


def normalizar_resposta(valor):
    if pd.isna(valor):
        return "NÃO"
    valor = str(valor).strip().upper()
    if valor in {"SIM", "S", "TRUE", "1", "REALIZADA", "REALIZADO"}:
        return "SIM"
    return "NÃO"


def carregar_arquivo(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Não foi possível ler a planilha: {e}")
        return None

    faltantes = [c for c in REQUIRED_BASE_COLUMNS if c not in df.columns]
    if faltantes:
        st.error(
            "A planilha enviada não possui todas as colunas mínimas necessárias. "
            f"Colunas faltantes: {', '.join(faltantes)}"
        )
        return None

    for col in ["Última Revisão", "Data da Próxima Revisão", "Data da Preventiva Realizada"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Garante colunas extras
    for col, default in EXTRA_COLUMNS_DEFAULTS.items():
        if col not in df.columns:
            df[col] = default

    # Ajustes de tipos
    if "Intervalo de Revisão" not in df.columns:
        df["Intervalo de Revisão"] = 120
    df["Intervalo de Revisão"] = pd.to_numeric(df["Intervalo de Revisão"], errors="coerce").fillna(120).astype(int)

    if "Ciclos de Preventiva Realizados" in df.columns:
        df["Ciclos de Preventiva Realizados"] = pd.to_numeric(
            df["Ciclos de Preventiva Realizados"], errors="coerce"
        ).fillna(0).astype(int)

    for col in PRED_COLS:
        df[col] = df[col].apply(normalizar_resposta)

    if "Preventiva Concluída" in df.columns:
        df["Preventiva Concluída"] = df["Preventiva Concluída"].fillna("PENDENTE").astype(str).str.upper()
    else:
        df["Preventiva Concluída"] = "PENDENTE"

    # Remove duplicidades mantendo a primeira placa
    if "PLACA" in df.columns:
        df = df.drop_duplicates(subset=["PLACA"], keep="first").copy()

    return recalcular_indicadores(df)


def calcular_info_linha(row, hoje=None):
    if hoje is None:
        hoje = pd.Timestamp(date.today())
    ultima = row.get("Última Revisão")
    proxima = row.get("Data da Próxima Revisão")
    intervalo = row.get("Intervalo de Revisão", 120)
    try:
        intervalo = int(intervalo)
    except Exception:
        intervalo = 120

    if pd.isna(ultima) and not pd.isna(proxima):
        ultima = proxima - pd.Timedelta(days=intervalo)
    elif pd.isna(proxima) and not pd.isna(ultima):
        proxima = ultima + pd.Timedelta(days=intervalo)
    elif pd.isna(ultima) and pd.isna(proxima):
        ultima = hoje
        proxima = hoje + pd.Timedelta(days=intervalo)

    dias_proxima = int((proxima.normalize() - hoje.normalize()).days)
    dias_decorridos = int((hoje.normalize() - ultima.normalize()).days)
    dias_decorridos = max(0, dias_decorridos)

    # Quantidade de preditivas previstas até hoje (a cada 15 dias, até 105)
    previstas = min(max(dias_decorridos // 15, 0), 7)
    realizadas = sum(normalizar_resposta(row.get(c, "NÃO")) == "SIM" for c in PRED_COLS)
    em_dia = min(realizadas, previstas)
    atrasadas = max(previstas - realizadas, 0)
    pendentes_total = max(7 - realizadas, 0)
    progresso = realizadas / 7 if 7 else 0

    # Próxima preditiva pendente
    proxima_pred_desc = "Concluídas"
    marcos = [15, 30, 45, 60, 75, 90, 105]
    for idx, col in enumerate(PRED_COLS):
        if normalizar_resposta(row.get(col, "NÃO")) == "NÃO":
            data_prevista = ultima + pd.Timedelta(days=marcos[idx])
            proxima_pred_desc = f"{col} - {data_prevista.strftime('%d/%m/%Y')}"
            break

    # Faixa baseada nos dias restantes da preventiva
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

    preventiva_concluida = str(row.get("Preventiva Concluída", "PENDENTE")).upper().strip()
    todas_preds = all(normalizar_resposta(row.get(c, "NÃO")) == "SIM" for c in PRED_COLS)

    if dias_proxima < 0 and preventiva_concluida != "SIM":
        status_prev = "ATRASADA"
    elif dias_proxima <= 15 and preventiva_concluida != "SIM":
        status_prev = "PRÓXIMA DO VENCIMENTO"
    else:
        status_prev = "EM DIA"

    if preventiva_concluida == "SIM":
        status_geral = "PREVENTIVA CONCLUÍDA"
    elif atrasadas > 0:
        status_geral = "PREDITIVA ATRASADA"
    elif status_prev == "ATRASADA":
        status_geral = "PREVENTIVA ATRASADA"
    else:
        status_geral = "EM DIA"

    return {
        "Última Revisão": ultima,
        "Data da Próxima Revisão": proxima,
        "Dias p/ Próxima": dias_proxima,
        "Faixa": faixa,
        "Progresso": progresso,
        "Qtd Preditivas Realizadas": realizadas,
        "Qtd Preditivas Previstas Hoje": previstas,
        "Qtd Preditivas Em Dia": em_dia,
        "Qtd Preditivas Atrasadas": atrasadas,
        "Preditivas Pendentes": pendentes_total,
        "Próxima Preditiva Prevista": proxima_pred_desc,
        "Pode Confirmar Preventiva": "SIM" if todas_preds else "NÃO",
        "Status Preventiva": status_prev,
        "Status Geral": status_geral,
    }


def recalcular_indicadores(df):
    hoje = pd.Timestamp(date.today())
    atualizados = df.apply(lambda row: pd.Series(calcular_info_linha(row, hoje=hoje)), axis=1)
    for col in atualizados.columns:
        df[col] = atualizados[col]
    return df


def dataframe_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Manutencao")
    output.seek(0)
    return output



def mostrar_metricas(df_filtrado):
    total_ativos = int(df_filtrado["PLACA"].nunique()) if not df_filtrado.empty else 0
    total_preventivas_em_dia = int((df_filtrado["Status Preventiva"] == "EM DIA").sum())
    total_preditivas_em_dia = int((df_filtrado["Qtd Preditivas Atrasadas"] == 0).sum())
    total_preventivas_atrasadas = int((df_filtrado["Status Preventiva"] == "ATRASADA").sum())
    total_prev_em_dia = int((df_filtrado["Status Preventiva"] == "EM DIA").sum())
    pct_prev_em_dia = (total_prev_em_dia / total_ativos * 100) if total_ativos else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de Ativos", f"{total_ativos}")
    c2.metric("Preventivas em Dia", f"{total_preventivas_em_dia}")
    c3.metric("Preditivas em Dia", f"{total_preditivas_em_dia}")
    c4.metric("Preventivas Atrasadas", f"{total_preventivas_atrasadas}")
    c5.metric("% Preventivas em Dia", f"{pct_prev_em_dia:.1f}%")


def dashboard(df):
    st.title("🚛 Dashboard de Manutenção Preventiva")
    st.caption("Acompanhe as preditivas de 15 em 15 dias e a preventiva final do ciclo de 120 dias.")

    with st.sidebar:
        st.header("Filtros")
        marcas = sorted(df["MARCA"].dropna().astype(str).unique().tolist()) if "MARCA" in df.columns else []
        tipos = sorted(df["TIPO DE FROTA"].dropna().astype(str).unique().tolist()) if "TIPO DE FROTA" in df.columns else []
        status = sorted(df["Status Geral"].dropna().astype(str).unique().tolist()) if "Status Geral" in df.columns else []

        sel_marcas = st.multiselect("Marca", marcas, default=marcas)
        sel_tipos = st.multiselect("Tipo de frota", tipos, default=tipos)
        sel_status = st.multiselect("Status geral", status, default=status)

    df_f = df.copy()
    if sel_marcas and "MARCA" in df_f.columns:
        df_f = df_f[df_f["MARCA"].astype(str).isin(sel_marcas)]
    if sel_tipos and "TIPO DE FROTA" in df_f.columns:
        df_f = df_f[df_f["TIPO DE FROTA"].astype(str).isin(sel_tipos)]
    if sel_status and "Status Geral" in df_f.columns:
        df_f = df_f[df_f["Status Geral"].astype(str).isin(sel_status)]

    mostrar_metricas(df_f)
    st.divider()

    # Resumo por faixa
    faixa_order = ["Atrasada", "0-15 dias", "16-30 dias", "31-60 dias", "61-120 dias", ">120 dias"]
    faixa_counts = (
        df_f["Faixa"].value_counts().rename_axis("Faixa").reset_index(name="Quantidade")
        if not df_f.empty else pd.DataFrame({"Faixa": [], "Quantidade": []})
    )
    if not faixa_counts.empty:
        faixa_counts["ord"] = faixa_counts["Faixa"].apply(lambda x: faixa_order.index(x) if x in faixa_order else 999)
        faixa_counts = faixa_counts.sort_values("ord").drop(columns="ord")

    pred_cols = st.columns([1.2, 1])
    with pred_cols[0]:
        st.subheader("Quantidade de ativos por faixa de vencimento")
        if faixa_counts.empty:
            st.info("Sem dados para exibir.")
        else:
            fig_faixa = px.bar(
                faixa_counts,
                x="Faixa",
                y="Quantidade",
                text="Quantidade",
                color="Faixa",
                category_orders={"Faixa": faixa_order},
            )
            fig_faixa.update_layout(showlegend=False, xaxis_title="Faixa", yaxis_title="Quantidade")
            st.plotly_chart(fig_faixa, use_container_width=True)
            st.dataframe(faixa_counts, use_container_width=True, hide_index=True)

    with pred_cols[1]:
        st.subheader("Distribuição do status geral")
        status_counts = (
            df_f["Status Geral"].value_counts().rename_axis("Status").reset_index(name="Quantidade")
            if not df_f.empty else pd.DataFrame({"Status": [], "Quantidade": []})
        )
        if status_counts.empty:
            st.info("Sem dados para exibir.")
        else:
            fig_status = px.pie(status_counts, names="Status", values="Quantidade", hole=0.5)
            st.plotly_chart(fig_status, use_container_width=True)

    graficos_cols = st.columns(2)
    with graficos_cols[0]:
        st.subheader("Preditivas realizadas por etapa")
        etapa_data = pd.DataFrame({
            "Etapa": ["0-15", "16-30", "31-45", "46-60", "61-75", "76-90", "91-105"],
            "Quantidade": [
                int((df_f[PRED_COLS[0]] == "SIM").sum()),
                int((df_f[PRED_COLS[1]] == "SIM").sum()),
                int((df_f[PRED_COLS[2]] == "SIM").sum()),
                int((df_f[PRED_COLS[3]] == "SIM").sum()),
                int((df_f[PRED_COLS[4]] == "SIM").sum()),
                int((df_f[PRED_COLS[5]] == "SIM").sum()),
                int((df_f[PRED_COLS[6]] == "SIM").sum()),
            ]
        })
        fig_etapas = px.bar(etapa_data, x="Etapa", y="Quantidade", text="Quantidade", color="Etapa")
        fig_etapas.update_layout(showlegend=False, xaxis_title="Faixa (dias)", yaxis_title="Preditivas realizadas")
        st.plotly_chart(fig_etapas, use_container_width=True)

    
with graficos_cols[1]:
    st.subheader("📈 Projeção de conclusão das preventivas")

    # Quantidade de preventivas pendentes
    pendentes = int((df_f["Status Preventiva"] == "ATRASADA").sum()) if "Status Preventiva" in df_f.columns else 0

    if pendentes == 0:
        st.success("✅ Todas as preventivas já foram concluídas!")
    else:
        meta_por_dia = 3
        ritmo_atual = 1  # estimativa conservadora

        dias_meta = int((pendentes / meta_por_dia) + 0.999)
        dias_atual = int((pendentes / ritmo_atual) + 0.999)

        dias = list(range(0, max(dias_meta, dias_atual) + 1))

        curva_meta = [max(pendentes - meta_por_dia * d, 0) for d in dias]
        curva_atual = [max(pendentes - ritmo_atual * d, 0) for d in dias]

        df_proj = pd.DataFrame({
            "Dias": dias,
            "Meta (3/dia)": curva_meta,
            "Ritmo Atual": curva_atual
        })

        fig_proj = px.line(
            df_proj,
            x="Dias",
            y=["Meta (3/dia)", "Ritmo Atual"],
            markers=True,
        )

        fig_proj.update_layout(
            xaxis_title="Dias a partir de hoje",
            yaxis_title="Preventivas pendentes",
            legend_title="Cenário",
        )

        st.plotly_chart(fig_proj, use_container_width=True)

        st.info(
            f"""
            🔹 **Preventivas pendentes:** {pendentes}  
            🔹 **Com a meta (3/dia):** ~{dias_meta} dias  
            🔹 **No ritmo atual:** ~{dias_atual} dias  
            """
        )

    st.subheader("Base detalhada")
    cols_show = [
        c for c in [
            "PLACA", "MARCA", "MODELO", "TIPO DE FROTA", "Última Revisão", "Data da Próxima Revisão",
            "Qtd Preditivas Realizadas", "Qtd Preditivas Previstas Hoje", "Qtd Preditivas Atrasadas",
            "Pode Confirmar Preventiva", "Preventiva Concluída", "Ciclos de Preventiva Realizados",
            "Dias p/ Próxima", "Faixa", "Status Geral"
        ] if c in df_f.columns
    ]
    st.dataframe(df_f[cols_show], use_container_width=True, hide_index=True)


def pagina_cadastro(df):
    st.title("🛠️ Cadastro / Atualização")
    st.caption("Registre as preditivas e confirme a preventiva do ciclo de 120 dias.")

    placas = df["PLACA"].astype(str).tolist()
    placa = st.selectbox("Selecione a carreta", placas)
    idx = df.index[df["PLACA"].astype(str) == placa][0]
    row = df.loc[idx].copy()

    info1, info2, info3, info4 = st.columns(4)
    info1.metric("Última revisão", row["Última Revisão"].strftime("%d/%m/%Y") if pd.notna(row["Última Revisão"]) else "-")
    info2.metric("Próxima revisão", row["Data da Próxima Revisão"].strftime("%d/%m/%Y") if pd.notna(row["Data da Próxima Revisão"]) else "-")
    info3.metric("Dias p/ próxima", int(row["Dias p/ Próxima"]))
    info4.metric("Ciclos concluídos", int(row.get("Ciclos de Preventiva Realizados", 0)))

    detalhes = [c for c in ["MARCA", "MODELO", "TIPO DE FROTA", "CHASSI", "Status Geral", "Próxima Preditiva Prevista"] if c in df.columns]
    if detalhes:
        st.write("### Dados do ativo")
        st.json({c: (row[c].strftime("%d/%m/%Y") if isinstance(row[c], pd.Timestamp) and pd.notna(row[c]) else (None if pd.isna(row[c]) else str(row[c]))) for c in detalhes})

    st.write("### Registrar preditivas")
    with st.form("form_preditivas"):
        novos = {}
        cols = st.columns(2)
        for i, col in enumerate(PRED_COLS):
            current = normalizar_resposta(row.get(col, "NÃO")) == "SIM"
            with cols[i % 2]:
                novos[col] = st.checkbox(col, value=current)
        obs = st.text_area("Observações", value=str(row.get("Observações", "")))
        salvar_preds = st.form_submit_button("Salvar preditivas", type="primary")

    if salvar_preds:
        for col, val in novos.items():
            df.at[idx, col] = "SIM" if val else "NÃO"
        df.at[idx, "Observações"] = obs
        df = recalcular_indicadores(df)
        st.session_state["df_manutencao"] = df
        st.success("Preditivas atualizadas com sucesso.")
        st.rerun()

    st.write("### Confirmar preventiva")
    pode_confirmar = str(row.get("Pode Confirmar Preventiva", "NÃO")) == "SIM"
    if not pode_confirmar:
        st.warning("A preventiva só pode ser confirmada após todas as 7 preditivas estarem marcadas como SIM.")
    with st.form("form_preventiva"):
        realizou_prev = st.selectbox(
            "A preventiva foi realizada?",
            options=["PENDENTE", "SIM", "NÃO"],
            index=["PENDENTE", "SIM", "NÃO"].index(str(row.get("Preventiva Concluída", "PENDENTE")).upper() if str(row.get("Preventiva Concluída", "PENDENTE")).upper() in ["PENDENTE", "SIM", "NÃO"] else "PENDENTE"),
            disabled=not pode_confirmar,
        )
        data_prev = st.date_input(
            "Data da preventiva",
            value=(row.get("Data da Preventiva Realizada") if pd.notna(row.get("Data da Preventiva Realizada")) else date.today()),
            disabled=not pode_confirmar,
        )
        resetar_ciclo = st.checkbox(
            "Ao concluir preventiva, iniciar novo ciclo de 120 dias automaticamente",
            value=True,
            disabled=not pode_confirmar,
        )
        salvar_prev = st.form_submit_button("Confirmar preventiva", disabled=not pode_confirmar)

    if salvar_prev:
        df.at[idx, "Preventiva Concluída"] = realizou_prev
        if realizou_prev == "SIM":
            data_prev_ts = pd.Timestamp(data_prev)
            df.at[idx, "Data da Preventiva Realizada"] = data_prev_ts
            df.at[idx, "Ciclos de Preventiva Realizados"] = int(df.at[idx, "Ciclos de Preventiva Realizados"]) + 1

            if resetar_ciclo:
                df.at[idx, "Última Revisão"] = data_prev_ts
                intervalo = int(df.at[idx, "Intervalo de Revisão"]) if pd.notna(df.at[idx, "Intervalo de Revisão"]) else 120
                df.at[idx, "Data da Próxima Revisão"] = data_prev_ts + pd.Timedelta(days=intervalo)
                for col in PRED_COLS:
                    df.at[idx, col] = "NÃO"
                df.at[idx, "Preventiva Concluída"] = "PENDENTE"
        elif realizou_prev == "NÃO":
            df.at[idx, "Data da Preventiva Realizada"] = pd.NaT
        df = recalcular_indicadores(df)
        st.session_state["df_manutencao"] = df
        st.success("Preventiva atualizada com sucesso.")
        st.rerun()

    st.write("### Situação atual do ativo")
    st.dataframe(
        recalcular_indicadores(df.loc[[idx]].copy())[[
            c for c in [
                "PLACA", "Qtd Preditivas Realizadas", "Qtd Preditivas Previstas Hoje", "Qtd Preditivas Atrasadas",
                "Pode Confirmar Preventiva", "Preventiva Concluída", "Ciclos de Preventiva Realizados",
                "Dias p/ Próxima", "Faixa", "Status Preventiva", "Status Geral", *PRED_COLS
            ] if c in df.columns
        ]],
        use_container_width=True,
        hide_index=True,
    )


def pagina_ajuda():
    st.title("ℹ️ Como usar")
    st.markdown(
        """
        1. Envie a planilha `.xlsx` no menu lateral.
        2. Acesse **Cadastro / Atualização** para marcar as preditivas de cada carreta.
        3. Quando as 7 preditivas estiverem com **SIM**, o app libera a confirmação da preventiva.
        4. Ao confirmar a preventiva como **SIM**, o sistema:
           - soma **+1** em `Ciclos de Preventiva Realizados`;
           - grava a data da preventiva realizada;
           - opcionalmente reinicia o ciclo de 120 dias, limpa as 7 preditivas e recalcula a próxima revisão.
        5. Use o botão **Baixar planilha atualizada** para salvar a base já tratada.

        **Indicadores do dashboard**
        - **Total de Ativos**: quantidade de placas únicas.
        - **Preventivas em Dia**: ativos cuja preventiva ainda não venceu.
        - **Preditivas em Dia**: ativos sem preditivas atrasadas até a data atual.
        - **Preventivas Atrasadas**: ativos com próxima preventiva vencida e não concluída.
        - **% Preventivas Concluídas**: percentual de ativos marcados com preventiva concluída = SIM.
        """
    )


def main():
    st.sidebar.title("📂 Dados")
    st.sidebar.write("Envie sua planilha Excel de manutenção preventiva.")
    uploaded_file = st.sidebar.file_uploader("Selecione o arquivo .xlsx", type=["xlsx"])

    if uploaded_file is None:
        st.title("Dashboard de Manutenção Preventiva")
        st.info("Envie uma planilha .xlsx para começar.")
        st.stop()

    if (
        "df_manutencao" not in st.session_state
        or st.session_state.get("nome_arquivo") != uploaded_file.name
    ):
        df = carregar_arquivo(uploaded_file)
        if df is None:
            st.stop()
        st.session_state["df_manutencao"] = df
        st.session_state["nome_arquivo"] = uploaded_file.name

    df = st.session_state["df_manutencao"]

    st.sidebar.divider()
    pagina = st.sidebar.radio(
        "Navegação",
        ["Dashboard", "Cadastro / Atualização", "Ajuda"],
        index=0,
    )

    excel_bytes = dataframe_para_excel(df)
    st.sidebar.download_button(
        label="⬇️ Baixar planilha atualizada",
        data=excel_bytes,
        file_name=f"manutencao_atualizada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if pagina == "Dashboard":
        dashboard(df)
    elif pagina == "Cadastro / Atualização":
        pagina_cadastro(df)
    else:
        pagina_ajuda()


if __name__ == "__main__":
    main()
