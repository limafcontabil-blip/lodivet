import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import date, datetime
import json

# ── Configuração da página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Lodi Vet",
    page_icon="🐾",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ── CSS mobile-friendly ─────────────────────────────────────────────────────
st.markdown("""
<style>
    .main > div { padding: 1rem 1rem 2rem 1rem; }
    .stButton > button {
        width: 100%;
        padding: 0.75rem;
        font-size: 1rem;
        border-radius: 10px;
        font-weight: 600;
    }
    .stSelectbox, .stTextInput, .stNumberInput, .stDateInput {
        margin-bottom: 0.25rem;
    }
    h1 { font-size: 1.5rem !important; }
    h2 { font-size: 1.2rem !important; }
    h3 { font-size: 1rem !important; }
    .card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 1rem;
        margin-bottom: 1rem;
        border-left: 4px solid #2196F3;
    }
    .card-pago {
        border-left: 4px solid #4CAF50;
    }
    .metric-box {
        background: #e3f2fd;
        border-radius: 10px;
        padding: 0.75rem;
        text-align: center;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Hospitais e cidades ─────────────────────────────────────────────────────
HOSPITAIS = {
    "Animal Vet": "Valinhos",
    "Baloo": "Vinhedo",
    "Bersan": "Valinhos",
    "Colonia": "Jundiaí",
    "Domicilio": "",
    "Domicilio Carla": "São Paulo",
    "Domicilio Neto": "Jundiaí",
    "Getulio": "Jundiaí",
    "Herminia": "Vinhedo",
    "Karol Rusa": "Louveira",
    "Mariana": "Vinhedo",
    "Todo Pet": "Jundiaí",
    "Um dois Pets": "Jundiaí",
    "Urbanpet": "Valinhos",
    "VitalVet": "Valinhos",
    "Outro / Novo hospital": None,
}

TIPOS = ["Comercial", "Domicilio", "Plantão", "Coleta", "Controle"]
SITUACOES = ["Em aberto", "Pago"]
MEIOS_PAGTO = ["PIX", "Crédito", "Débito", "Dinheiro", "Transferência"]
RECEBIDO_OPTS = ["PF", "PJ", ""]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SHEET_ID = "1e4ID8nj9oybS_wMJjFfLeCI8OPme1f-_8dJMX3n-bD0"
ABA = "Página1"

# ── Conexão Google Sheets ───────────────────────────────────────────────────
@st.cache_resource
def get_client():
    import json
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=30)
def carregar_dados():
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(ABA)
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    if df.empty:
        df = pd.DataFrame(columns=[
            "DATA","HOSPITAL","CIDADE","CLIENTE","PET",
            "TIPO","VALOR","SITUAÇÃO","MEIO PAGTO","DATA PAGTO","RECEBIDO","OBSERVAÇÃO"
        ])
    else:
        df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0)
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce").dt.date
        df["DATA PAGTO"] = pd.to_datetime(df["DATA PAGTO"], errors="coerce").dt.date
    return df

def salvar_linha(linha: list):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(ABA)
    sheet.append_row(linha, value_input_option="USER_ENTERED")
    carregar_dados.clear()

def atualizar_linha(row_index: int, situacao: str, meio: str, data_pagto, recebido: str):
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(ABA)
    sheet_row = row_index + 2
    data_str = data_pagto.strftime("%d/%m/%Y") if data_pagto else ""
    sheet.update(f"H{sheet_row}", [[situacao]])
    sheet.update(f"J{sheet_row}", [[data_str]])
    sheet.update(f"I{sheet_row}", [[meio]])
    sheet.update(f"K{sheet_row}", [[recebido]])
    carregar_dados.clear()

def get_hospitais_atuais():
    df = carregar_dados()
    extras = set(df["HOSPITAL"].dropna().unique()) - set(HOSPITAIS.keys()) - {"Outro / Novo hospital"}
    base = dict(HOSPITAIS)
    for h in sorted(extras):
        base[h] = ""
    ordered = dict(sorted({k: v for k, v in base.items() if k != "Outro / Novo hospital"}.items()))
    ordered["Outro / Novo hospital"] = None
    return ordered

# ── Navegação ───────────────────────────────────────────────────────────────
if "pagina" not in st.session_state:
    st.session_state.pagina = "registrar"

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("➕ Registrar"):
        st.session_state.pagina = "registrar"
with col2:
    if st.button("💰 Pagamentos"):
        st.session_state.pagina = "pagamentos"
with col3:
    if st.button("📊 Relatório"):
        st.session_state.pagina = "relatorio"

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — REGISTRAR ATENDIMENTO
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.pagina == "registrar":
    st.markdown("## 🐾 Novo Atendimento")

    hospitais_map = get_hospitais_atuais()
    lista_hospitais = list(hospitais_map.keys())

    # Campos FORA do form para reagir imediatamente
    hospital_sel = st.selectbox("🏥 Hospital", lista_hospitais, key="hospital_sel")

    if hospital_sel == "Outro / Novo hospital":
        hospital_final = st.text_input("Nome do novo hospital", key="novo_hospital")
        cidade_final = st.text_input("Cidade do novo hospital", key="nova_cidade")
    else:
        hospital_final = hospital_sel
        cidade_auto = hospitais_map.get(hospital_sel, "")
        if cidade_auto == "":
            cidade_final = st.text_input("📍 Cidade", value="", key="cidade_livre")
        else:
            st.text_input("📍 Cidade", value=cidade_auto, disabled=True, key="cidade_auto")
            cidade_final = cidade_auto

    # Situação também FORA do form para mostrar campos de pagamento na hora
    situacao = st.selectbox("📋 Situação", SITUACOES, key="situacao_sel")

    meio_pagto = ""
    data_pagto = None
    recebido = ""
    if situacao == "Pago":
        meio_pagto = st.selectbox("💳 Meio de Pagamento", MEIOS_PAGTO, key="meio_sel")
        data_pagto = st.date_input("📅 Data do Pagamento", value=date.today(), key="data_pagto_sel")
        recebido = st.selectbox("🧾 PF / PJ", RECEBIDO_OPTS, key="recebido_sel")

    with st.form("form_atendimento", clear_on_submit=True):
        data = st.date_input("📅 Data do Atendimento", value=date.today())
        cliente = st.text_input("👤 Cliente")
        pet = st.text_input("🐶 Pet")
        tipo = st.selectbox("🔬 Tipo de Atendimento", TIPOS)
        valor = st.number_input("💵 Valor (R$)", min_value=0.0, step=10.0, format="%.2f")
        observacao = st.text_input("📝 Observação (opcional)")
        submitted = st.form_submit_button("✅ Salvar Atendimento", use_container_width=True)

    if submitted:
        if st.session_state.get("hospital_sel") == "Outro / Novo hospital":
            hospital_final = st.session_state.get("novo_hospital", "")
            cidade_final = st.session_state.get("nova_cidade", "")

        situacao = st.session_state.get("situacao_sel", "Em aberto")
        meio_pagto = st.session_state.get("meio_sel", "")
        data_pagto = st.session_state.get("data_pagto_sel", None)
        recebido = st.session_state.get("recebido_sel", "")

        if not hospital_final:
            st.error("Informe o hospital.")
        elif not pet:
            st.error("Informe o nome do pet.")
        elif valor == 0:
            st.warning("Valor zerado. Salvando mesmo assim...")

        if hospital_final and pet:
            data_str = data.strftime("%d/%m/%Y")
            data_pagto_str = data_pagto.strftime("%d/%m/%Y") if data_pagto else ""
            linha = [
                data_str, hospital_final, cidade_final, cliente, pet,
                tipo, valor, situacao, meio_pagto, data_pagto_str, recebido, observacao
            ]
            salvar_linha(linha)
            st.success(f"✅ Atendimento de {pet} registrado com sucesso!")

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — REGISTRAR PAGAMENTOS
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.pagina == "pagamentos":
    st.markdown("## 💰 Registrar Pagamentos")

    df = carregar_dados()
    em_aberto = df[df["SITUAÇÃO"] == "Em aberto"].copy()

    if em_aberto.empty:
        st.success("🎉 Nenhum atendimento em aberto!")
    else:
        st.info(f"{len(em_aberto)} atendimento(s) em aberto")

        for idx, row in em_aberto.iterrows():
            data_fmt = row["DATA"].strftime("%d/%m/%Y") if pd.notna(row["DATA"]) else "—"
            with st.expander(f"🐾 {row['PET']} — {row['HOSPITAL']} — R$ {row['VALOR']:.0f} ({data_fmt})"):
                st.write(f"**Cliente:** {row['CLIENTE']}")
                st.write(f"**Tipo:** {row['TIPO']}")

                with st.form(f"pag_{idx}"):
                    meio = st.selectbox("Meio de Pagamento", MEIOS_PAGTO, key=f"meio_{idx}")
                    data_p = st.date_input("Data do Pagamento", value=date.today(), key=f"datap_{idx}")
                    rec = st.selectbox("PF / PJ", RECEBIDO_OPTS, key=f"rec_{idx}")
                    confirmar = st.form_submit_button("✅ Confirmar Pagamento", use_container_width=True)

                if confirmar:
                    atualizar_linha(idx, "Pago", meio, data_p, rec)
                    st.success("Pagamento registrado!")
                    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 3 — RELATÓRIO
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.pagina == "relatorio":
    st.markdown("## 📊 Relatório")

    df = carregar_dados()

    if df.empty:
        st.info("Nenhum dado ainda.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            meses = sorted(df["DATA"].dropna().apply(lambda d: d.strftime("%m/%Y")).unique(), reverse=True)
            mes_sel = st.selectbox("Mês", ["Todos"] + list(meses))
        with col2:
            hospitais_lista = ["Todos"] + sorted(df["HOSPITAL"].dropna().unique().tolist())
            hosp_sel = st.selectbox("Hospital", hospitais_lista)

        df_f = df.copy()
        if mes_sel != "Todos":
            df_f = df_f[df_f["DATA"].apply(lambda d: d.strftime("%m/%Y") if pd.notna(d) else "") == mes_sel]
        if hosp_sel != "Todos":
            df_f = df_f[df_f["HOSPITAL"] == hosp_sel]

        total = df_f["VALOR"].sum()
        recebido_val = df_f[df_f["SITUAÇÃO"] == "Pago"]["VALOR"].sum()
        aberto_val = df_f[df_f["SITUAÇÃO"] == "Em aberto"]["VALOR"].sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total", f"R$ {total:,.0f}")
        c2.metric("Recebido", f"R$ {recebido_val:,.0f}")
        c3.metric("Em aberto", f"R$ {aberto_val:,.0f}")

        st.divider()

        por_hospital = df_f.groupby("HOSPITAL")["VALOR"].sum().sort_values(ascending=False)
        st.markdown("**Por Hospital**")
        st.dataframe(por_hospital.reset_index().rename(columns={"VALOR": "Total (R$)"}), use_container_width=True)

        st.divider()

        st.markdown("**Atendimentos**")
        colunas_exibir = ["DATA", "HOSPITAL", "CLIENTE", "PET", "TIPO", "VALOR", "SITUAÇÃO"]
        st.dataframe(df_f[colunas_exibir].sort_values("DATA", ascending=False), use_container_width=True)
