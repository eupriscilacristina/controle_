import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, date
import os
import io
import hashlib

# =============================================================================
# CONFIGURAÇÃO
# =============================================================================
st.set_page_config(
    page_title="GTCON - Controle de Acessos",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Controle_", "Acessos GTCON.xlsx")
SENHA_ADM = "1882"

# Mapeamento legível dos nomes das abas
SHEET_DISPLAY_NAMES = {
    "ACESSOS SERVIDOR": "Servidor",
    "VAGOS": "Vagos",
    "ENCAMINHADOS ": "Encaminhados",
    "NIT": "NIT",
    "DIRETORIA": "Diretoria",
    "DP": "DP",
    "FISCAL": "Fiscal",
    "CONTABIL": "Contábil",
    "IMPLANTA\u00c7\u00c3O": "Implantação",
    "COMPLIANCE": "Compliance",
    "LEGALIZA\u00c7\u00c3O": "Legalização",
    "TRIBUT\u00c1RIO": "Tributário",
    "CS": "Sucesso Cliente",
    "MARKETING": "Marketing",
    "COMERCIAL": "Comercial",
    "DIVERSOS": "Diversos",
    "TROCA DE E-MAILS": "Troca de E-mails",
    "INF NOTEBOOK": "Notebooks",
    "CERTIFICADO NAYRA": "Certificados Nayra",
}

DEPT_COLORS = {
    "Servidor": "#1f77b4", "DP": "#2ca02c", "Fiscal": "#ff7f0e",
    "Contábil": "#d62728", "Implantação": "#9467bd", "Compliance": "#8c564b",
    "Legalização": "#e377c2", "Tributário": "#7f7f7f", "Sucesso Cliente": "#bcbd22",
    "Marketing": "#17becf", "Comercial": "#aec7e8", "Diretoria": "#ffbb78",
    "NIT": "#98df8a", "Diversos": "#ff9896", "Vagos": "#c5b0d5",
}

# =============================================================================
# CSS CUSTOMIZADO
# =============================================================================
def inject_custom_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    .stApp {
        font-family: 'Inter', sans-serif;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 1rem;
        max-width: 100%;
    }

    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #0e1117 0%, #1a1f2e 100%);
        border: 1px solid #2a3040;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        transition: transform 0.2s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    }
    div[data-testid="stMetric"] label {
        color: #8892a4 !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #e8ecf1 !important;
        font-weight: 700 !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 10px 24px;
        font-weight: 500;
        font-size: 0.95rem;
    }

    .topbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 12px 24px;
        background: linear-gradient(135deg, #0e1117 0%, #1a1f2e 100%);
        border: 1px solid #2a3040;
        border-radius: 12px;
        margin-bottom: 20px;
    }
    .topbar-left {
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .topbar-right {
        display: flex;
        align-items: center;
        gap: 16px;
    }
    .topbar-brand {
        font-size: 1.2rem;
        font-weight: 700;
        color: #e8ecf1;
    }
    .topbar-info {
        color: #8892a4;
        font-size: 0.82rem;
        background: rgba(37,99,235,0.1);
        padding: 4px 12px;
        border-radius: 16px;
    }
    .topbar-copyright {
        color: #5a6577;
        font-size: 0.75rem;
    }
    .btn-sair {
        padding: 8px 18px;
        border: 1px solid #ef4444;
        border-radius: 8px;
        background: transparent;
        color: #ef4444;
        font-size: 0.85rem;
        font-weight: 500;
        font-family: inherit;
        cursor: pointer;
        transition: all 0.2s;
    }
    .btn-sair:hover {
        background: rgba(239,68,68,0.1);
    }

    .login-container {
        max-width: 420px;
        margin: 80px auto;
        padding: 40px;
        background: linear-gradient(145deg, #0e1117 0%, #1a1f2e 100%);
        border: 1px solid #2a3040;
        border-radius: 16px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.4);
    }
    .login-title {
        text-align: center;
        color: #e8ecf1;
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .login-subtitle {
        text-align: center;
        color: #8892a4;
        font-size: 0.9rem;
        margin-bottom: 32px;
    }

    .dept-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        color: white;
    }

    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
    }

    h1, h2, h3 {
        color: #e8ecf1 !important;
    }
    .stMarkdown p {
        color: #c5cdd8;
    }
    </style>
    """, unsafe_allow_html=True)


# =============================================================================
# AUTENTICAÇÃO
# =============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-title">🔐 GTCON</div>', unsafe_allow_html=True)
    st.markdown('<div class="login-subtitle">Sistema de Controle de Acessos</div>', unsafe_allow_html=True)

    password = st.text_input("Senha de Acesso", type="password", placeholder="Digite sua senha...", key="login_pwd")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("Entrar", use_container_width=True, type="primary"):
            if password == SENHA_ADM:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Senha incorreta. Acesso negado.")

    st.markdown('</div>', unsafe_allow_html=True)
    return False


# =============================================================================
# LEITURA / ESCRITA DE DADOS
# =============================================================================
@st.cache_data(ttl=30)
def load_data():
    if not os.path.exists(EXCEL_PATH):
        return {}, []
    xls = pd.ExcelFile(EXCEL_PATH)
    sheets = xls.sheet_names
    data = {}
    for s in sheets:
        df = pd.read_excel(xls, sheet_name=s)
        df = df.fillna("")
        if s == "ACESSOS SERVIDOR":
            drop_cols = [c for c in df.columns if c in ["Unnamed: 3", "Unnamed: 6"]]
            df = df.drop(columns=drop_cols, errors="ignore")
            rename_map = {}
            for c in df.columns:
                if "EXACT.1" in c:
                    rename_map[c] = c.replace("EXACT.1", "EXACT.ADM")
                elif "GTCON.1" in c:
                    rename_map[c] = c.replace("GTCON.1", "GTCON.ADM")
            if rename_map:
                df = df.rename(columns=rename_map)
        data[s] = df
    return data, sheets


def render_sectioned_tab(sheet_name, df, display_name):
    st.markdown(f"### 📋 {display_name}")

    section_markers = []
    for idx, val in df.iloc[:, 0].items():
        val_str = str(val).strip()
        val_upper = val_str.upper()
        other_cols_empty = df.iloc[idx, 1:].astype(str).str.strip().replace("", pd.NA).dropna().empty
        if val_str and other_cols_empty and val_upper != "COLABORADOR" and "@" not in val_upper and "GTCON" not in val_upper and len(val_str) > 1:
            section_markers.append((idx, val_str))

    if not section_markers:
        edited_df = st.data_editor(
            df, use_container_width=True,
            height=min(500, 35 * len(df) + 80),
            hide_index=True, num_rows="dynamic",
            key=f"editor_{sheet_name}",
        )
        col1, col2, _ = st.columns([1, 1, 4])
        with col1:
            if st.button("💾 Salvar Alterações", key=f"save_{sheet_name}", type="primary"):
                if save_sheet(sheet_name, edited_df):
                    st.success("✅ Alterações salvas com sucesso!")
                    st.cache_data.clear()
                    st.rerun()
        with col2:
            if st.button("🔄 Recarregar", key=f"reload_{sheet_name}"):
                st.cache_data.clear()
                st.rerun()
        return

    sections = []
    first_marker_idx = section_markers[0][0]

    if first_marker_idx > 0:
        header_row = df.iloc[0]
        data_before = df.iloc[1:first_marker_idx]
        data_before = data_before[data_before.iloc[:, 0].astype(str).str.strip() != ""]
        if not data_before.empty:
            data_before = data_before.copy()
            data_before.columns = [str(c).strip() for c in data_before.columns]
            data_before = data_before.reset_index(drop=True)
            sections.append((display_name, data_before))

    for i, (start_idx, title) in enumerate(section_markers):
        end_idx = section_markers[i + 1][0] if i + 1 < len(section_markers) else len(df)
        chunk = df.iloc[start_idx + 1:end_idx]
        chunk = chunk[chunk.iloc[:, 0].astype(str).str.strip() != ""]
        chunk = chunk[chunk.iloc[:, 0].astype(str).str.strip().str.upper() != "COLABORADOR"]
        chunk = chunk[chunk.iloc[:, 0].astype(str).str.strip().str.upper() != title.upper()]
        if not chunk.empty:
            chunk = chunk.reset_index(drop=True)
        sections.append((title, chunk))

    all_edited = []
    for title, chunk in sections:
        if title == display_name:
            clean_title = display_name
        else:
            clean_title = title.title()

        st.markdown(f"""
        <div style="text-align:center; padding:10px 0; margin:16px 0 8px;
            background:linear-gradient(135deg,#1a1f2e,#0e1117);
            border:1px solid #2a3040; border-radius:10px;">
            <span style="color:#60a5fa; font-size:1rem; font-weight:600;">📂 {clean_title}</span>
        </div>
        """, unsafe_allow_html=True)

        if chunk.empty:
            st.info("Nenhum registro nesta seção.")
            continue

        edited = st.data_editor(
            chunk, use_container_width=True,
            height=min(350, 35 * len(chunk) + 80),
            hide_index=True, num_rows="dynamic",
            key=f"editor_{sheet_name}_{clean_title}",
        )
        all_edited.append(edited)

    col1, col2, _ = st.columns([1, 1, 4])
    with col1:
        if st.button("💾 Salvar Alterações", key=f"save_{sheet_name}", type="primary"):
            combined = pd.concat(all_edited, ignore_index=True)
            if save_sheet(sheet_name, combined):
                st.success("✅ Alterações salvas com sucesso!")
                st.cache_data.clear()
                st.rerun()
    with col2:
        if st.button("🔄 Recarregar", key=f"reload_{sheet_name}"):
            st.cache_data.clear()
            st.rerun()


def save_sheet(sheet_name, df):
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False


# =============================================================================
# KPIs
# =============================================================================
def compute_kpis(data, all_sheets):
    dept_sheet_names = [
        "DP", "FISCAL", "CONTABIL", "IMPLANTA\u00c7\u00c3O", "COMPLIANCE",
        "LEGALIZA\u00c7\u00c3O", "TRIBUT\u00c1RIO", "CS", "MARKETING",
        "COMERCIAL", "NIT", "DIRETORIA"
    ]

    total_usuarios = 0
    total_depts = 0
    dept_counts = {}

    for s in all_sheets:
        df = data.get(s, pd.DataFrame())
        if df.empty:
            continue
        display = SHEET_DISPLAY_NAMES.get(s, s)

        if s == "ACESSOS SERVIDOR":
            col_user = [c for c in df.columns if "USU" in str(c).upper() and "GTCON" in str(c).upper() and "EXACT" not in str(c).upper()]
            if col_user:
                total_usuarios = df[col_user[0]].astype(str).replace("", pd.NA).dropna().shape[0]
        elif s == "VAGOS":
            vagos_count = df.iloc[:, 0].astype(str).replace("", pd.NA).dropna().shape[0] if not df.empty else 0
        elif s in dept_sheet_names:
            total_depts += 1
            collab_col = df.columns[0] if not df.empty else None
            if collab_col:
                count = df[collab_col].astype(str).replace("", pd.NA).dropna()
                count = count[count != "COLABORADOR"]
                dept_counts[display] = count.shape[0]

    acessos_servidor = data.get("ACESSOS SERVIDOR", pd.DataFrame())
    total_acessos = 0
    if not acessos_servidor.empty:
        for col in acessos_servidor.columns:
            if "USU" in str(col).upper() and "EXACT" in str(col).upper():
                vals = acessos_servidor[col].astype(str).replace("", pd.NA).dropna()
                total_acessos = max(total_acessos, vals.shape[0])

    encaminhados = data.get("ENCAMINHADOS ", pd.DataFrame())
    enc_count = len(encaminhados) if not encaminhados.empty else 0

    notebooks = data.get("INF NOTEBOOK", pd.DataFrame())
    nb_count = len(notebooks) if not notebooks.empty else 0

    certificados = data.get("CERTIFICADO NAYRA", pd.DataFrame())
    cert_count = len(certificados) if not certificados.empty else 0

    return {
        "total_usuarios_servidor": total_acessos,
        "total_departamentos": total_depts,
        "total_colaboradores_dept": sum(dept_counts.values()),
        "encaminhados": enc_count,
        "notebooks": nb_count,
        "certificados": cert_count,
        "dept_counts": dept_counts,
    }


# =============================================================================
# DASHBOARD
# =============================================================================
def render_dashboard(data, all_sheets):
    st.markdown("## 📊 Painel Analítico")

    kpis = compute_kpis(data, all_sheets)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("👤 Usuários Servidor", kpis["total_usuarios_servidor"])
    with c2:
        st.metric("🏢 Departamentos", kpis["total_departamentos"])
    with c3:
        st.metric("👥 Colaboradores (Dept)", kpis["total_colaboradores_dept"])
    with c4:
        st.metric("📋 Certificados", kpis["certificados"])

    c5, c6, c7 = st.columns(3)
    with c5:
        st.metric("📨 Encaminhados", kpis["encaminhados"])
    with c6:
        st.metric("💻 Notebooks", kpis["notebooks"])
    with c7:
        st.metric("📧 Trocas de E-mail", len(data.get("TROCA DE E-MAILS", [])))

    st.markdown("---")

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("### 🏢 Colaboradores por Departamento")
        dept_data = kpis["dept_counts"]
        if dept_data:
            dept_df = pd.DataFrame({
                "Departamento": list(dept_data.keys()),
                "Colaboradores": list(dept_data.values())
            }).sort_values("Colaboradores", ascending=True)

            fig_bar = px.bar(
                dept_df,
                x="Colaboradores",
                y="Departamento",
                orientation="h",
                color="Colaboradores",
                color_continuous_scale=["#1a1f2e", "#2563eb", "#60a5fa"],
                text="Colaboradores"
            )
            fig_bar.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#c5cdd8"),
                xaxis=dict(gridcolor="#1e2533", title="Nº de Colaboradores"),
                yaxis=dict(gridcolor="#1e2533", title=""),
                coloraxis_showscale=False,
                margin=dict(l=0, r=0, t=10, b=0),
                height=420
            )
            fig_bar.update_traces(textposition="outside", textfont=dict(color="#e8ecf1", size=12))
            st.plotly_chart(fig_bar, use_container_width=True)

    with col_right:
        st.markdown("### 📊 Distribuição por Departamento")
        if dept_data:
            fig_pie = px.pie(
                pd.DataFrame({"D": list(dept_data.keys()), "V": list(dept_data.values())}),
                names="D",
                values="V",
                hole=0.5,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig_pie.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#c5cdd8", size=11),
                showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                margin=dict(l=0, r=0, t=10, b=0),
                height=420
            )
            fig_pie.update_traces(textposition="inside", textinfo="percent+label",
                                  textfont=dict(size=10))
            st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")
    st.markdown("### 📊 Resumo Geral por Aba")
    summary_data = []
    for s in all_sheets:
        df = data.get(s, pd.DataFrame())
        display = SHEET_DISPLAY_NAMES.get(s, s)
        summary_data.append({
            "Aba": display,
            "Linhas": len(df),
            "Colunas": len(df.columns),
        })
    summary_df = pd.DataFrame(summary_data).sort_values("Linhas", ascending=False)

    fig_summary = px.bar(
        summary_df, x="Aba", y="Linhas",
        color="Linhas",
        color_continuous_scale=["#1a1f2e", "#3b82f6"],
        text="Linhas"
    )
    fig_summary.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#c5cdd8"),
        xaxis=dict(gridcolor="#1e2533", title=""),
        yaxis=dict(gridcolor="#1e2533", title="Registros"),
        coloraxis_showscale=False,
        margin=dict(l=0, r=0, t=10, b=0),
        height=350
    )
    fig_summary.update_traces(textposition="outside", textfont=dict(color="#e8ecf1"))
    st.plotly_chart(fig_summary, use_container_width=True)

    with st.expander("📋 Tabela Resumo", expanded=False):
        st.dataframe(summary_df, use_container_width=True, hide_index=True)


# =============================================================================
# ABAS DE DADOS
# =============================================================================
def render_data_tab(sheet_name, df, display_name):
    st.markdown(f"### 📋 {display_name}")

    if df.empty:
        st.info("Esta aba não contém dados.")
        return

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        height=min(500, 35 * len(df) + 80),
        hide_index=True,
        num_rows="dynamic",
        key=f"editor_{sheet_name}",
    )

    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("💾 Salvar Alterações", key=f"save_{sheet_name}", type="primary"):
            if save_sheet(sheet_name, edited_df):
                st.success("✅ Alterações salvas com sucesso!")
                st.cache_data.clear()
                st.rerun()
    with col2:
        if st.button("🔄 Recarregar", key=f"reload_{sheet_name}"):
            st.cache_data.clear()
            st.rerun()


# =============================================================================
# NOVO REGISTRO
# =============================================================================
def render_new_record(data, all_sheets):
    st.markdown("## ➕ Novo Registro de Acesso")
    st.markdown("Preencha os campos abaixo para cadastrar um novo registro em qualquer aba do sistema.")

    tabs_labels = [SHEET_DISPLAY_NAMES.get(s, s) for s in all_sheets]
    selected_tab = st.selectbox("📁 Selecione a Aba de Destino", tabs_labels, key="new_rec_tab")

    selected_sheet = None
    for s in all_sheets:
        if SHEET_DISPLAY_NAMES.get(s, s) == selected_tab:
            selected_sheet = s
            break

    if not selected_sheet:
        return

    df = data.get(selected_sheet, pd.DataFrame())

    if df.empty:
        st.info("Aba selecionada sem estrutura de colunas.")
        return

    st.markdown(f"**Colunas da aba `{selected_tab}`:**")

    with st.form(key="new_record_form", clear_on_submit=True):
        cols = st.columns(min(len(df.columns), 3))
        new_values = {}
        for i, col in enumerate(df.columns):
            c = cols[i % len(cols)]
            new_values[col] = c.text_input(f"{col}", key=f"nr_{col}")

        submitted = st.form_submit_button("💾 Cadastrar Registro", type="primary", use_container_width=True)
        if submitted:
            has_content = any(v.strip() for v in new_values.values())
            if not has_content:
                st.warning("Preencha ao menos um campo para cadastrar.")
            else:
                new_row = pd.DataFrame([new_values])
                updated_df = pd.concat([df, new_row], ignore_index=True)
                if save_sheet(selected_sheet, updated_df):
                    st.success(f"✅ Registro adicionado em **{selected_tab}** com sucesso!")
                    st.cache_data.clear()
                    st.rerun()


# =============================================================================
# HEADER
# =============================================================================
def render_header(data, all_sheets):
    pass


# =============================================================================
# MAIN
# =============================================================================
def main():
    inject_custom_css()

    if not check_password():
        return

    data, all_sheets = load_data()

    if not data:
        st.error("❌ Arquivo de dados não encontrado. Verifique o caminho: `Controle_/Acessos GTCON.xlsx`")
        return

    total_registros = sum(len(data.get(s, [])) for s in all_sheets)
    total_abas = len(all_sheets)

    col_title, col_info = st.columns([3, 1])
    with col_title:
        st.markdown("# 🔐 GTCON - Controle de Acessos")
    with col_info:
        st.markdown(f"""
        <div style="display:flex; gap:12px; align-items:center; padding-top:12px; justify-content:flex-end;">
            <span style="color:#8892a4; font-size:0.85rem;">📋 {total_registros} registros</span>
            <span style="color:#8892a4; font-size:0.85rem;">📁 {total_abas} abas</span>
        </div>
        <div style="text-align:right; color:#5a6577; font-size:0.75rem; margin-top:2px;">GTCON Brasil &copy; 2026</div>
        """, unsafe_allow_html=True)
    st.markdown("Sistema integrado de gerenciamento e visualização de acessos corporativos.")
    st.markdown("")

    tab_dashboard, tab_controle, tab_novo = st.tabs([
        "📊 Dashboards & Indicadores",
        "📁 Controle de Acessos",
        "➕ Novo Registro"
    ])

    with tab_dashboard:
        render_dashboard(data, all_sheets)

    with tab_controle:
        st.markdown("## 📁 Controle de Acessos por Aba")
        tabs_labels = [SHEET_DISPLAY_NAMES.get(s, s) for s in all_sheets]
        sub_tabs = st.tabs(tabs_labels)

        for i, (sub_tab, sheet_name) in enumerate(zip(sub_tabs, all_sheets)):
            with sub_tab:
                display = SHEET_DISPLAY_NAMES.get(sheet_name, sheet_name)
                sectioned_sheets = ["VAGOS"]
                if sheet_name.upper().startswith("IMPLANTA") or sheet_name in sectioned_sheets:
                    render_sectioned_tab(sheet_name, data.get(sheet_name, pd.DataFrame()), display)
                else:
                    render_data_tab(sheet_name, data.get(sheet_name, pd.DataFrame()), display)

    with tab_novo:
        render_new_record(data, all_sheets)


if __name__ == "__main__":
    main()
