import streamlit as st
import pandas as pd
import json
import os
from datetime import date
from io import BytesIO

# ─── CONFIGURAÇÕES ────────────────────────────────────────
DATA_FILE      = "atendimentos.xlsx"
CADASTROS_FILE = "cadastros.json"

TIPOS       = ["Comercial", "Domicílio", "Plantão", "Coleta", "Controle"]
MEIOS_PAGTO = ["PIX", "Crédito", "Dinheiro"]
RECEBIDO    = ["PJ", "PF", "—"]
COLUNAS     = ["DATA", "HOSPITAL", "CIDADE", "CLIENTE", "PET", "TIPO",
               "VALOR", "SITUAÇÃO", "MEIO PAGTO", "DATA PAGTO",
               "RECEBIDO", "OBSERVAÇÃO"]

HOSPITAIS_DEFAULT = [
    {"nome": "Animal Vet",      "cidade": "Valinhos"},
    {"nome": "Baloo",           "cidade": "Vinhedo"},
    {"nome": "Bersan",          "cidade": "Valinhos"},
    {"nome": "Carla",           "cidade": "Jundiaí"},
    {"nome": "Colônia",         "cidade": "Jundiaí"},
    {"nome": "Domicílio",       "cidade": ""},
    {"nome": "Domicílio Carla", "cidade": "São Paulo"},
    {"nome": "Domicílio Neto",  "cidade": "Jundiaí"},
    {"nome": "Getúlio",         "cidade": "Jundiaí"},
    {"nome": "Hermínia",        "cidade": "Vinhedo"},
    {"nome": "Karol Rusa",      "cidade": "Louveira"},
    {"nome": "Mariana",         "cidade": "Vinhedo"},
    {"nome": "Neto",            "cidade": "Jundiaí"},
    {"nome": "Todo Pet",        "cidade": "Jundiaí"},
    {"nome": "Um Dois Pets",    "cidade": "Jundiaí"},
    {"nome": "Urbanpet",        "cidade": "Valinhos"},
    {"nome": "VitalVet",        "cidade": "Valinhos"},
]

# ─── FUNÇÕES DE DADOS ─────────────────────────────────────
def load_cadastros():
    if os.path.exists(CADASTROS_FILE):
        with open(CADASTROS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    dados = {"hospitais": HOSPITAIS_DEFAULT}
    save_cadastros(dados)
    return dados

def save_cadastros(data):
    with open(CADASTROS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        df["DATA"]       = pd.to_datetime(df["DATA"],       errors="coerce")
        df["DATA PAGTO"] = pd.to_datetime(df["DATA PAGTO"], errors="coerce")
        if "OBSERVAÇÃO" not in df.columns:
            df["OBSERVAÇÃO"] = ""
        return df
    return pd.DataFrame(columns=COLUNAS)

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

def fmt_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "—"

def to_excel_download(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ─── PÁGINA ───────────────────────────────────────────────
st.set_page_config(page_title="Lodi Vet", page_icon="🐾", layout="wide")
st.title("🐾 Lodi Vet — Controle de Atendimentos")

menu = st.sidebar.radio("Navegação", [
    "📋 Novo Atendimento",
    "💳 Registrar Pagamento",
    "📊 Relatórios",
    "⚙️ Cadastros"
])

cadastros       = load_cadastros()
hospitais       = cadastros["hospitais"]
nomes_hosp      = sorted([h["nome"] for h in hospitais])
cidade_por_hosp = {h["nome"]: h["cidade"] for h in hospitais}

# ─────────────────────────────────────────────────────────
# 1. NOVO ATENDIMENTO
# ─────────────────────────────────────────────────────────
if menu == "📋 Novo Atendimento":
    st.header("Novo Atendimento")

    if "inp_cidade" not in st.session_state:
        primeiro = nomes_hosp[0] if nomes_hosp else ""
        st.session_state["inp_cidade"] = cidade_por_hosp.get(primeiro, "")

    def ao_mudar_hospital():
        hosp = st.session_state.get("sel_hospital", "")
        st.session_state["inp_cidade"] = cidade_por_hosp.get(hosp, "")

    col1, col2 = st.columns(2)

    with col1:
        data_atend = st.date_input("Data do Atendimento", value=date.today())
        hospital   = st.selectbox(
            "Hospital / Local", nomes_hosp,
            key="sel_hospital", on_change=ao_mudar_hospital
        )
        cidade  = st.text_input(
            "Cidade", key="inp_cidade",
            help="Preenchida automaticamente. Edite se necessário."
        )
        cliente = st.text_input("Cliente")
        pet     = st.text_input("Pet")

    with col2:
        tipo  = st.selectbox("Tipo de Atendimento", TIPOS)
        valor = st.number_input("Valor (R$)", min_value=0.0, step=10.0, format="%.2f")
        situacao = st.selectbox("Situação", ["Em aberto", "Pago"])

        meio_pagto = data_pagto = recebido_val = None
        if situacao == "Pago":
            meio_pagto   = st.selectbox("Meio de Pagamento", MEIOS_PAGTO)
            data_pagto   = st.date_input("Data do Pagamento", value=date.today())
            rec_raw      = st.selectbox("Recebido como", RECEBIDO)
            recebido_val = None if rec_raw == "—" else rec_raw

        obs = st.text_input(
            "Observação (opcional)",
            help="Ex: pago direto para o hospital, aguardando boleto, etc."
        )

    st.divider()

    if st.button("💾 Salvar Atendimento", type="primary"):
        erros = []
        if not cliente.strip():
            erros.append("Cliente")
        if not pet.strip():
            erros.append("Pet")
        if valor <= 0:
            erros.append("Valor maior que zero")

        if erros:
            st.error(f"Preencha os campos obrigatórios: {', '.join(erros)}")
        else:
            df = load_data()
            novo = {
                "DATA":       data_atend,
                "HOSPITAL":   hospital,
                "CIDADE":     cidade,
                "CLIENTE":    cliente.strip(),
                "PET":        pet.strip(),
                "TIPO":       tipo,
                "VALOR":      valor,
                "SITUAÇÃO":   situacao,
                "MEIO PAGTO": meio_pagto,
                "DATA PAGTO": data_pagto,
                "RECEBIDO":   recebido_val,
                "OBSERVAÇÃO": obs.strip() if obs else "",
            }
            df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
            save_data(df)
            st.success(f"✅ Atendimento salvo — {cliente.strip()} / {pet.strip()}")
            st.balloons()

# ─────────────────────────────────────────────────────────
# 2. REGISTRAR PAGAMENTO
# ─────────────────────────────────────────────────────────
elif menu == "💳 Registrar Pagamento":
    st.header("Registrar Pagamento")
    df = load_data()

    if df.empty:
        st.info("Nenhum atendimento cadastrado ainda.")
    else:
        em_aberto = df[df["SITUAÇÃO"] == "Em aberto"].copy()

        if em_aberto.empty:
            st.success("🎉 Nenhum atendimento em aberto!")
        else:
            c1, c2 = st.columns(2)
            c1.metric("Atendimentos em aberto", len(em_aberto))
            c2.metric("Total a receber", fmt_brl(em_aberto["VALOR"].sum()))
            st.divider()

            busca = st.text_input("🔍 Filtrar por cliente ou hospital")
            lista = em_aberto.copy()
            if busca:
                mask = (
                    lista["CLIENTE"].str.contains(busca, case=False, na=False) |
                    lista["HOSPITAL"].str.contains(busca, case=False, na=False)
                )
                lista = lista[mask]

            for idx, row in lista.iterrows():
                dt  = pd.to_datetime(row["DATA"]).strftime("%d/%m/%Y") if pd.notna(row["DATA"]) else "—"
                obs_txt = row.get("OBSERVAÇÃO", "")
                obs_lbl = f" | ⚠️ {obs_txt}" if pd.notna(obs_txt) and str(obs_txt).strip() else ""
                lbl = (
                    f"**{row['CLIENTE']}** — {row['PET']}  |  "
                    f"{row['HOSPITAL']}  |  {dt}  |  {fmt_brl(row['VALOR'])}{obs_lbl}"
                )

                with st.expander(lbl):
                    cc1, cc2, cc3 = st.columns(3)
                    meio  = cc1.selectbox("Meio de Pagamento", MEIOS_PAGTO, key=f"mp_{idx}")
                    dtpag = cc2.date_input("Data do Pagamento", value=date.today(), key=f"dp_{idx}")
                    rec_r = cc3.selectbox("Recebido como", RECEBIDO, key=f"rc_{idx}")

                    if st.button("✅ Confirmar Pagamento", key=f"btn_{idx}"):
                        df.at[idx, "SITUAÇÃO"]   = "Pago"
                        df.at[idx, "MEIO PAGTO"] = meio
                        df.at[idx, "DATA PAGTO"] = dtpag
                        df.at[idx, "RECEBIDO"]   = None if rec_r == "—" else rec_r
                        save_data(df)
                        st.success("Pagamento registrado!")
                        st.rerun()

# ─────────────────────────────────────────────────────────
# 3. RELATÓRIOS
# ─────────────────────────────────────────────────────────
elif menu == "📊 Relatórios":
    st.header("Relatórios")
    df = load_data()

    if df.empty:
        st.info("Nenhum dado disponível ainda.")
    else:
        with st.expander("🔍 Filtros", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                dmin   = df["DATA"].min().date() if pd.notna(df["DATA"].min()) else date.today()
                dmax   = df["DATA"].max().date() if pd.notna(df["DATA"].max()) else date.today()
                dt_ini = st.date_input("De",  value=dmin)
                dt_fim = st.date_input("Até", value=dmax)
            with c2:
                f_hosp = st.multiselect("Hospital",      sorted(df["HOSPITAL"].dropna().unique()))
                f_sit  = st.multiselect("Situação",      ["Pago", "Em aberto"])
            with c3:
                f_rec  = st.multiselect("Recebido como", ["PJ", "PF"])
                f_tipo = st.multiselect("Tipo",          TIPOS)

        mask = (df["DATA"].dt.date >= dt_ini) & (df["DATA"].dt.date <= dt_fim)
        dff  = df[mask].copy()
        if f_hosp: dff = dff[dff["HOSPITAL"].isin(f_hosp)]
        if f_sit:  dff = dff[dff["SITUAÇÃO"].isin(f_sit)]
        if f_rec:  dff = dff[dff["RECEBIDO"].isin(f_rec)]
        if f_tipo: dff = dff[dff["TIPO"].isin(f_tipo)]

        st.divider()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Atendimentos",  len(dff))
        c2.metric("Faturamento",   fmt_brl(dff["VALOR"].sum()))
        c3.metric("Recebido",      fmt_brl(dff[dff["SITUAÇÃO"] == "Pago"]["VALOR"].sum()))
        c4.metric("Em aberto",     fmt_brl(dff[dff["SITUAÇÃO"] == "Em aberto"]["VALOR"].sum()))

        c5, c6, c7 = st.columns(3)
        c5.metric("Total PJ", fmt_brl(dff[dff["RECEBIDO"] == "PJ"]["VALOR"].sum()))
        c6.metric("Total PF", fmt_brl(dff[dff["RECEBIDO"] == "PF"]["VALOR"].sum()))
        c7.metric("Sem info", fmt_brl(dff[dff["RECEBIDO"].isna()]["VALOR"].sum()))

        st.divider()

        disp = dff.copy()
        disp["DATA"]       = disp["DATA"].dt.strftime("%d/%m/%Y")
        disp["DATA PAGTO"] = (
            pd.to_datetime(disp["DATA PAGTO"], errors="coerce")
            .dt.strftime("%d/%m/%Y")
            .fillna("—")
        )
        disp["VALOR"] = disp["VALOR"].apply(fmt_brl)
        disp.fillna("—", inplace=True)
        st.dataframe(disp, use_container_width=True, hide_index=True)

        st.divider()
        fname = f"relatorio_{dt_ini.strftime('%d%m%Y')}_{dt_fim.strftime('%d%m%Y')}.xlsx"
        st.download_button(
            "📥 Exportar Excel",
            to_excel_download(dff),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ─────────────────────────────────────────────────────────
# 4. CADASTROS
# ─────────────────────────────────────────────────────────
elif menu == "⚙️ Cadastros":
    st.header("Gerenciar Cadastros")

    st.subheader("Hospitais e locais cadastrados")
    df_hosp = pd.DataFrame(hospitais).rename(
        columns={"nome": "Hospital / Local", "cidade": "Cidade"}
    )
    st.dataframe(df_hosp, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Adicionar novo hospital / local")
    with st.form("add_hosp", clear_on_submit=True):
        cc1, cc2  = st.columns(2)
        novo_nome = cc1.text_input("Nome")
        nova_cid  = cc2.text_input("Cidade (deixe em branco se variável)")
        if st.form_submit_button("➕ Adicionar"):
            if novo_nome.strip():
                cadastros["hospitais"].append({
                    "nome":   novo_nome.strip(),
                    "cidade": nova_cid.strip()
                })
                save_cadastros(cadastros)
                st.success(f"**{novo_nome.strip()}** adicionado!")
                st.rerun()
            else:
                st.error("Informe o nome.")

    st.divider()
    st.subheader("Remover hospital / local")
    to_rm = st.selectbox("Selecione para remover", nomes_hosp)
    if st.button("🗑️ Remover selecionado", type="secondary"):
        cadastros["hospitais"] = [
            h for h in cadastros["hospitais"] if h["nome"] != to_rm
        ]
        save_cadastros(cadastros)
        st.success(f"**{to_rm}** removido.")
        st.rerun()