"""
═══════════════════════════════════════════════════════════════════════
  brief-mensuel-bourse · Streamlit Web App
  ─────────────────────────────────────────────────────────────────────
  Companion web app du repo LinkedIn-Brief-Mensuel-Bourse.

  Features :
    1. Source de données : GitHub Releases latest (.xlsx fixe)
    2. Filtres dynamiques : PEA/CTO, secteur, pays, market, sliders perf/potentiel
    3. Visualisations Plotly : scatter, treemap, histogrammes, top N
    4. Détail ticker : metrics + chart historique yfinance (MA20/50/200 + regression)
    5. Analyse IA : bouton qui copie le prompt presse-papier + redirect claude.ai
    6. Theme Bloomberg-style #00d4ff (cohérent avec la vidéo MP4)
    7. Cache : xlsx TTL 6h (refresh manuel possible), histo ticker 1h

  Déploiement : Streamlit Community Cloud
    URL cible : https://brief-mensuel-bourse.streamlit.app
    Repo source : https://github.com/Romtaug/LinkedIn-Brief-Mensuel-Bourse

  Auteur : Romain Taugourdeau
═══════════════════════════════════════════════════════════════════════
"""
from __future__ import annotations

import json
import urllib.parse
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
import yfinance as yf
from numpy import polyfit, poly1d


# ═════════════════════════════════════════════════════════════════════
#  CONFIG
# ═════════════════════════════════════════════════════════════════════

REPO_OWNER = "Romtaug"
REPO_NAME  = "LinkedIn-Brief-Mensuel-Bourse"
XLSX_URL   = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/releases/latest/download/ranking_latest.xlsx"

# Couleurs Bloomberg-style (idem CSS de la vidéo)
COLORS = {
    "bg":       "#0a1628",
    "bg_panel": "#0f1d34",
    "bg_row":   "#13243f",
    "grid":     "#1e3358",
    "blue":     "#00d4ff",
    "blue_dim": "#0080a0",
    "amber":    "#ffaa00",
    "text":     "#e8eaf0",
    "text_mid": "#8b9bb4",
    "dim":      "#4a5c7a",
    "green":    "#00ff66",
    "red":      "#ff2e2e",
    "gold":     "#ffd700",
    "pea":      "#00d4ff",
    "cto":      "#7c9eff",
}

PARRAINAGE_URL = "https://bour.so/p/GB93ZfQVNVr"
CODE_PARRAIN   = "ROTA0058"
LINKEDIN_URL   = "https://www.linkedin.com/in/romain-taugourdeau/"


# ═════════════════════════════════════════════════════════════════════
#  PAGE CONFIG + CSS
# ═════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Brief Mensuel Bourse",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": f"https://github.com/{REPO_OWNER}/{REPO_NAME}",
        "About": "Brief Mensuel Bourse — Analyse automatisée de +1000 actions PEA & CTO. "
                 "Code open-source sur GitHub.",
    },
)

# CSS custom Bloomberg-style
st.markdown(f"""
<style>
    /* Police monospace partout sauf titres */
    html, body, [class*="css"] {{
        font-family: 'JetBrains Mono', 'Courier New', monospace;
    }}
    h1, h2, h3, h4 {{
        font-family: 'Inter', sans-serif !important;
        letter-spacing: -0.5px;
    }}
    /* Bar top sticky bleue */
    .bar-top {{
        background: {COLORS['blue']};
        color: {COLORS['bg']};
        padding: 12px 24px;
        font-weight: 700;
        font-size: 14px;
        letter-spacing: 1px;
        margin: -2rem -2rem 1.5rem -2rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }}
    .live-dot {{
        display: inline-block;
        width: 10px; height: 10px;
        background: {COLORS['bg']};
        border-radius: 50%;
        margin-right: 12px;
        animation: pulse 2s infinite;
    }}
    @keyframes pulse {{
        0%, 100% {{ opacity: 1; }}
        50% {{ opacity: 0.4; }}
    }}
    /* Métriques */
    [data-testid="stMetricValue"] {{
        font-family: 'Inter', sans-serif;
        color: {COLORS['blue']};
        font-weight: 800;
    }}
    /* Sidebar */
    [data-testid="stSidebar"] {{
        background: {COLORS['bg_panel']};
    }}
    /* Boutons */
    .stButton > button {{
        background: {COLORS['bg_panel']};
        border: 2px solid {COLORS['blue']};
        color: {COLORS['blue']};
        font-weight: 700;
        letter-spacing: 1px;
    }}
    .stButton > button:hover {{
        background: {COLORS['blue']};
        color: {COLORS['bg']};
    }}
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background: {COLORS['bg_panel']};
        border-radius: 0;
        color: {COLORS['text_mid']};
        font-weight: 700;
    }}
    .stTabs [aria-selected="true"] {{
        background: {COLORS['blue']};
        color: {COLORS['bg']};
    }}
    /* Liens */
    a {{ color: {COLORS['blue']} !important; }}
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════
#  DATA LOADING (cached)
# ═════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=6 * 3600, show_spinner="📡 Chargement du brief mensuel...")
def load_xlsx_from_github(url: str = XLSX_URL) -> dict[str, pd.DataFrame]:
    """
    Télécharge l'xlsx depuis GitHub Releases (cache 6h).
    Retourne un dict {sheet_name: DataFrame}.
    Si fail → retourne dict vide + erreur Streamlit.
    """
    try:
        r = requests.get(url, timeout=30, headers={"User-Agent": "BriefBourseApp/1.0"})
        r.raise_for_status()
        xlsx_bytes = BytesIO(r.content)
        sheets = pd.read_excel(xlsx_bytes, sheet_name=None)
        return sheets
    except requests.HTTPError as e:
        st.error(f"❌ Erreur HTTP {e.response.status_code} : {url}")
        return {}
    except Exception as e:
        st.error(f"❌ Erreur de chargement : {e}")
        return {}


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_ticker_history(ticker: str, period: str = "1y") -> pd.DataFrame:
    """
    Fetch yfinance historique (cache 1h par couple ticker/period).
    Returns DataFrame avec colonne Close + index DatetimeIndex.
    """
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period=period, auto_adjust=True)
        if hist.empty:
            return pd.DataFrame()
        # Garde juste Close + Volume pour léger
        return hist[["Close", "Volume"]].copy()
    except Exception as e:
        st.warning(f"⚠️ Erreur fetch historique {ticker} : {e}")
        return pd.DataFrame()


# ═════════════════════════════════════════════════════════════════════
#  HELPERS
# ═════════════════════════════════════════════════════════════════════

def get_country_from_ticker(ticker: str) -> str:
    """Drapeau pays à partir du suffixe yfinance."""
    flag_map = {
        ".PA": "🇫🇷 France", ".DE": "🇩🇪 Allemagne", ".AS": "🇳🇱 Pays-Bas",
        ".BR": "🇧🇪 Belgique", ".MI": "🇮🇹 Italie", ".MC": "🇪🇸 Espagne",
        ".LS": "🇵🇹 Portugal", ".OL": "🇳🇴 Norvège", ".ST": "🇸🇪 Suède",
        ".HE": "🇫🇮 Finlande", ".CO": "🇩🇰 Danemark", ".VI": "🇦🇹 Autriche",
        ".IR": "🇮🇪 Irlande", ".SW": "🇨🇭 Suisse", ".L": "🇬🇧 Royaume-Uni",
        ".WA": "🇵🇱 Pologne", ".AT": "🇬🇷 Grèce", ".T": "🇯🇵 Japon",
        ".TO": "🇨🇦 Canada", ".AX": "🇦🇺 Australie", ".HK": "🇭🇰 Hong Kong",
    }
    if "." not in ticker:
        return "🇺🇸 États-Unis"
    suf = "." + ticker.rsplit(".", 1)[1]
    return flag_map.get(suf, "🌍 Autre")


def build_claude_prompt(row: dict) -> str:
    """Génère le prompt Claude pour l'analyse d'un ticker."""
    pea_str = "Oui ✅" if row.get("pea") else "Non ❌"
    country = get_country_from_ticker(row["ticker"])
    div_str = f"{row.get('div_pct', 0):.2f}%" if row.get('div_pct', 0) else "0% (pas de dividende)"

    return f"""Tu es un analyste financier indépendant. Analyse l'action suivante en français, de manière concise et actionnable (~300 mots) :

Ticker : {row['ticker']}
Nom : {row.get('name', '')}
Secteur : {row.get('sector_fr', '')}
Pays / Marché : {country}
Éligibilité PEA : {pea_str}
─────
Cours actuel : {row.get('price_eur', '?')}€ (équivalent EUR, prix natif {row.get('price', '?')} {row.get('currency', '')})
Cible analystes 12m : {row.get('target_pct', 0):+.1f}% ({row.get('reco_label', '-')} · {int(row.get('analyst_count', 0))} analystes)
Dividende : {div_str}
Performance mois précédent : {row.get('perf_1m', 0):+.1f}%
Score potentiel total : {row.get('total_pct', 0):+.1f}% (cible + dividende)

Structure ta réponse en 4 sections :
1. **L'entreprise en 3 lignes** (activité, marché, position concurrentielle)
2. **Lecture des chiffres** (que disent les analystes ? le momentum ?)
3. **Thèse haussière vs baissière** (2 arguments chacun)
4. **Conclusion** (à qui s'adresse ce titre ? quels risques majeurs ?)

⚠️ Précise que ce n'est PAS un conseil en investissement (risque de perte en capital)."""


def plot_styled_layout(fig, title: str = "", height: int = 400):
    """Applique le thème Bloomberg sur un fig Plotly."""
    fig.update_layout(
        title=title,
        paper_bgcolor=COLORS["bg"],
        plot_bgcolor=COLORS["bg_panel"],
        font=dict(family="JetBrains Mono, monospace", color=COLORS["text"], size=12),
        title_font=dict(family="Inter, sans-serif", size=18, color=COLORS["text"]),
        xaxis=dict(gridcolor=COLORS["grid"], linecolor=COLORS["grid"], zeroline=False),
        yaxis=dict(gridcolor=COLORS["grid"], linecolor=COLORS["grid"], zeroline=False),
        legend=dict(bgcolor=COLORS["bg_panel"], bordercolor=COLORS["grid"], borderwidth=1),
        height=height,
        margin=dict(l=40, r=40, t=60, b=40),
    )
    return fig


# ═════════════════════════════════════════════════════════════════════
#  HEADER
# ═════════════════════════════════════════════════════════════════════

st.markdown(f"""
<div class="bar-top">
    <span><span class="live-dot"></span>EN DIRECT · BRIEF MENSUEL EQUITY</span>
    <span>{datetime.now().strftime('%d/%m/%Y · %H:%M')}</span>
</div>
""", unsafe_allow_html=True)

st.markdown(f"# 📊 Brief Mensuel <span style='color:{COLORS['blue']}'>Bourse.</span>",
            unsafe_allow_html=True)
st.markdown(f"<p style='color:{COLORS['text_mid']}; font-size:15px;'>"
            f"+1000 actions analysées · PEA & CTO · Source Yahoo Finance · "
            f"Mise à jour mensuelle automatique"
            f"</p>", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════
#  LOAD DATA
# ═════════════════════════════════════════════════════════════════════

sheets = load_xlsx_from_github()

if not sheets:
    st.warning(
        "⚠️ Aucune donnée disponible. Le brief mensuel sera publié le **premier "
        "jour ouvré de chaque mois**. Reviens début du mois prochain !"
    )
    st.info(f"📦 [Voir le code source sur GitHub](https://github.com/{REPO_OWNER}/{REPO_NAME})")
    st.stop()

# Onglet principal = "By Total Gain" (la version complète)
df = sheets.get("By Total Gain", pd.DataFrame())
if df.empty:
    st.error("❌ Onglet 'By Total Gain' vide ou absent")
    st.stop()


# ═════════════════════════════════════════════════════════════════════
#  SIDEBAR - Filtres
# ═════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### 🎯 Filtres")

    # Éligibilité PEA / CTO
    elig = st.radio(
        "Éligibilité",
        options=["Tout", "PEA uniquement", "CTO uniquement"],
        horizontal=False,
    )

    # Secteur GICS
    if "sector_fr" in df.columns:
        secteurs = sorted(df["sector_fr"].dropna().unique().tolist())
        sectors_sel = st.multiselect("Secteur GICS", options=secteurs, default=[])

    # Pays / Marché
    if "ticker" in df.columns:
        df["country"] = df["ticker"].apply(get_country_from_ticker)
        countries = sorted(df["country"].unique().tolist())
        countries_sel = st.multiselect("Pays", options=countries, default=[])

    st.markdown("---")
    st.markdown("### 📊 Métriques")

    # Slider potentiel total
    if "total_pct" in df.columns:
        min_t, max_t = float(df["total_pct"].min()), float(df["total_pct"].max())
        pot_range = st.slider(
            "Potentiel total (%)",
            min_value=min_t, max_value=max_t,
            value=(min_t, max_t),
            step=1.0, format="%.0f%%",
        )

    # Slider perf 1 mois
    if "perf_1m" in df.columns:
        df_perf = df.dropna(subset=["perf_1m"])
        if not df_perf.empty:
            min_p, max_p = float(df_perf["perf_1m"].min()), float(df_perf["perf_1m"].max())
            perf_range = st.slider(
                "Performance mois (%)",
                min_value=min_p, max_value=max_p,
                value=(min_p, max_p),
                step=1.0, format="%.0f%%",
            )
        else:
            perf_range = None

    # Nb analystes min
    if "analyst_count" in df.columns:
        max_an = int(df["analyst_count"].max()) if len(df) else 50
        n_analysts_min = st.slider(
            "Couverture analystes (min)",
            min_value=0, max_value=max_an, value=0,
        )

    st.markdown("---")
    if st.button("🔄 Rafraîchir les données"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.markdown("### 🔗 Liens")
    st.markdown(f"📦 [Code GitHub](https://github.com/{REPO_OWNER}/{REPO_NAME})")
    st.markdown(f"💼 [Mon profil LinkedIn]({LINKEDIN_URL})")
    st.markdown(f"💳 [Parrainage Boursorama (+100€)]({PARRAINAGE_URL}) `{CODE_PARRAIN}`")


# ═════════════════════════════════════════════════════════════════════
#  APPLICATION DES FILTRES
# ═════════════════════════════════════════════════════════════════════

df_f = df.copy()
if "country" not in df_f.columns:
    df_f["country"] = df_f["ticker"].apply(get_country_from_ticker)

if elig == "PEA uniquement":
    df_f = df_f[df_f["pea"] == True]
elif elig == "CTO uniquement":
    df_f = df_f[df_f["pea"] == False]

if sectors_sel:
    df_f = df_f[df_f["sector_fr"].isin(sectors_sel)]

if countries_sel:
    df_f = df_f[df_f["country"].isin(countries_sel)]

if "total_pct" in df_f.columns and pot_range:
    df_f = df_f[(df_f["total_pct"] >= pot_range[0]) & (df_f["total_pct"] <= pot_range[1])]

if "perf_1m" in df_f.columns and perf_range:
    mask = df_f["perf_1m"].between(perf_range[0], perf_range[1]) | df_f["perf_1m"].isna()
    df_f = df_f[mask]

if "analyst_count" in df_f.columns and n_analysts_min > 0:
    df_f = df_f[df_f["analyst_count"] >= n_analysts_min]


# ═════════════════════════════════════════════════════════════════════
#  MÉTRIQUES GLOBALES (sous le titre)
# ═════════════════════════════════════════════════════════════════════

n_total = len(df_f)
n_pea   = int(df_f["pea"].sum()) if "pea" in df_f.columns else 0
n_cto   = n_total - n_pea
avg_pot = df_f["total_pct"].mean() if n_total else 0
avg_perf = df_f["perf_1m"].dropna().mean() if n_total else 0

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Actions affichées", f"{n_total}", delta=f"sur {len(df)} totales")
m2.metric("PEA · Zone EEE", f"{n_pea}", delta=f"{100*n_pea/max(1,n_total):.0f}%")
m3.metric("CTO · Mondial", f"{n_cto}", delta=f"{100*n_cto/max(1,n_total):.0f}%")
m4.metric("Potentiel moyen", f"{avg_pot:+.1f}%")
m5.metric("Perf moyenne", f"{avg_perf:+.1f}%")

st.markdown("---")


# ═════════════════════════════════════════════════════════════════════
#  TABS
# ═════════════════════════════════════════════════════════════════════

tab_table, tab_charts, tab_ticker, tab_sectors, tab_about = st.tabs([
    "📋 Classement",
    "📊 Graphiques",
    "🔍 Détail Ticker",
    "📂 Par Secteur",
    "ℹ️ À propos",
])


# ─────────────────────────────────────────────────────────────────────
#  TAB 1 : CLASSEMENT (table principale)
# ─────────────────────────────────────────────────────────────────────
with tab_table:
    st.markdown(f"### Classement des {n_total} actions filtrées")

    # Choix tri
    sort_options = {
        "Potentiel total décroissant (cible + dividende)": ("total_pct", False),
        "Perf 1 mois décroissante":   ("perf_1m", False),
        "Perf 1 mois croissante":     ("perf_1m", True),
        "Cible analystes décroissante": ("target_pct", False),
        "Dividende décroissant":      ("div_pct", False),
        "Recommandation analystes (meilleure)":  ("reco_mean", True),
        "Couverture analystes (plus suivis)":    ("analyst_count", False),
        "Cours EUR croissant":         ("price_eur", True),
        "Cours EUR décroissant":       ("price_eur", False),
        "Nom alphabétique":            ("name", True),
    }
    sort_choice = st.selectbox("Trier par", options=list(sort_options.keys()))
    sort_col, sort_asc = sort_options[sort_choice]

    df_show = df_f.copy()
    if sort_col in df_show.columns:
        df_show = df_show.sort_values(sort_col, ascending=sort_asc, na_position="last")

    # Colonnes à afficher (avec mise en forme)
    cols_show = [c for c in [
        "ticker", "name", "country", "sector_fr", "pea",
        "price_eur", "price", "currency",
        "perf_1m", "target_pct", "div_pct", "total_pct",
        "reco_label", "analyst_count", "reco_mean",
        "boursorama_link", "yahoo_link",
    ] if c in df_show.columns]

    df_display = df_show[cols_show].rename(columns={
        "ticker":       "Ticker",
        "name":         "Nom",
        "country":      "Pays",
        "sector_fr":    "Secteur",
        "pea":          "PEA",
        "price_eur":    "Cours €",
        "price":        "Cours natif",
        "currency":     "Devise",
        "perf_1m":      "Perf 1M",
        "target_pct":   "Cible 12M",
        "div_pct":      "Dividende",
        "total_pct":    "Potentiel total",
        "reco_label":   "Reco analystes",
        "analyst_count": "Nb analystes",
        "reco_mean":    "Score reco (1=★★★★★)",
        "boursorama_link": "Boursorama",
        "yahoo_link":   "Yahoo Finance",
    })

    st.dataframe(
        df_display,
        height=600,
        use_container_width=True,
        column_config={
            "PEA":         st.column_config.CheckboxColumn("PEA"),
            "Cours €":     st.column_config.NumberColumn("Cours €", format="%.2f €"),
            "Cours natif": st.column_config.NumberColumn("Natif", format="%.2f"),
            "Perf 1M":     st.column_config.NumberColumn("Perf 1M", format="%+.1f %%"),
            "Cible 12M":   st.column_config.NumberColumn("Cible 12M", format="%+.1f %%"),
            "Dividende":   st.column_config.NumberColumn("Div", format="%.2f %%"),
            "Potentiel total": st.column_config.NumberColumn("Potentiel", format="%+.1f %%"),
            "Boursorama":  st.column_config.LinkColumn("BR", display_text="🏛️"),
            "Yahoo Finance": st.column_config.LinkColumn("YF", display_text="🔍"),
        },
        hide_index=True,
    )

    # Export
    csv = df_show.to_csv(index=False).encode("utf-8")
    st.download_button(
        "💾 Télécharger le classement filtré (CSV)",
        data=csv,
        file_name=f"brief_bourse_{datetime.now().strftime('%Y-%m-%d')}.csv",
        mime="text/csv",
    )


# ─────────────────────────────────────────────────────────────────────
#  TAB 2 : GRAPHIQUES AGRÉGÉS
# ─────────────────────────────────────────────────────────────────────
with tab_charts:
    st.markdown("### 📊 Vue agrégée du classement filtré")

    if n_total < 2:
        st.info("Pas assez de données pour les graphiques (relâche les filtres).")
    else:
        # Graph 1 : Scatter Target vs Perf
        st.markdown("#### 🎯 Cible analystes vs Performance mois")
        df_scatter = df_f.dropna(subset=["target_pct", "perf_1m"])
        if not df_scatter.empty:
            fig1 = px.scatter(
                df_scatter,
                x="perf_1m", y="target_pct",
                color="sector_fr",
                hover_data=["ticker", "name", "country", "total_pct"],
                size="analyst_count" if "analyst_count" in df_scatter.columns else None,
                size_max=30,
                labels={"perf_1m": "Performance mois (%)", "target_pct": "Cible analystes 12M (%)"},
            )
            fig1.add_hline(y=0, line_dash="dash", line_color=COLORS["dim"])
            fig1.add_vline(x=0, line_dash="dash", line_color=COLORS["dim"])
            fig1 = plot_styled_layout(fig1, height=500)
            st.plotly_chart(fig1, use_container_width=True)

        col_a, col_b = st.columns(2)

        with col_a:
            # Graph 2 : Histogramme potentiel total
            st.markdown("#### 📊 Distribution du potentiel total")
            fig2 = px.histogram(
                df_f, x="total_pct", nbins=40,
                color_discrete_sequence=[COLORS["blue"]],
                labels={"total_pct": "Potentiel total (%)"},
            )
            fig2 = plot_styled_layout(fig2, height=380)
            st.plotly_chart(fig2, use_container_width=True)

        with col_b:
            # Graph 3 : Boxplot perf par secteur
            st.markdown("#### 📦 Performance mois par secteur")
            df_box = df_f.dropna(subset=["perf_1m"])
            if not df_box.empty:
                fig3 = px.box(
                    df_box, y="sector_fr", x="perf_1m",
                    color="sector_fr",
                    labels={"perf_1m": "Perf 1M (%)", "sector_fr": ""},
                    orientation="h",
                )
                fig3.update_layout(showlegend=False)
                fig3 = plot_styled_layout(fig3, height=380)
                st.plotly_chart(fig3, use_container_width=True)

        # Graph 4 : Treemap secteurs (taille = nb tickers, couleur = potentiel moyen)
        st.markdown("#### 🌳 Treemap secteurs (taille = nb tickers, couleur = potentiel moyen)")
        df_tree = (df_f.groupby("sector_fr")
                       .agg(n=("ticker", "count"), avg_pot=("total_pct", "mean"))
                       .reset_index())
        fig4 = px.treemap(
            df_tree, path=["sector_fr"], values="n", color="avg_pot",
            color_continuous_scale=[[0, COLORS["red"]], [0.5, COLORS["amber"]], [1, COLORS["green"]]],
            color_continuous_midpoint=0,
            labels={"avg_pot": "Potentiel moyen (%)", "sector_fr": "Secteur"},
        )
        fig4 = plot_styled_layout(fig4, height=450)
        st.plotly_chart(fig4, use_container_width=True)

        # Graph 5 : Top 20
        st.markdown("#### 🏆 Top 20 par potentiel total")
        top20 = df_f.nlargest(20, "total_pct")
        fig5 = px.bar(
            top20.iloc[::-1],
            x="total_pct", y="ticker",
            color="total_pct",
            color_continuous_scale=[[0, COLORS["red"]], [0.5, COLORS["amber"]], [1, COLORS["green"]]],
            text="name",
            hover_data=["sector_fr", "country", "perf_1m"],
            labels={"total_pct": "Potentiel total (%)", "ticker": ""},
        )
        fig5.update_traces(textposition="inside")
        fig5 = plot_styled_layout(fig5, height=600)
        st.plotly_chart(fig5, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────
#  TAB 3 : DÉTAIL TICKER + ANALYSE CLAUDE
# ─────────────────────────────────────────────────────────────────────
with tab_ticker:
    st.markdown("### 🔍 Analyse détaillée d'une action")

    # Sélection ticker
    tickers_sorted = df_f.sort_values("ticker")["ticker"].tolist()
    if not tickers_sorted:
        st.info("Pas de tickers dans la sélection actuelle.")
    else:
        default_idx = 0
        col_sel, col_period = st.columns([3, 1])
        with col_sel:
            ticker_sel = st.selectbox(
                "Choisis une action",
                options=tickers_sorted,
                index=default_idx,
                format_func=lambda t: f"{t} · {df_f[df_f['ticker']==t]['name'].iloc[0]}",
            )
        with col_period:
            period_sel = st.selectbox(
                "Période historique",
                options=["3mo", "6mo", "1y", "2y", "5y"],
                index=2,  # 1y par défaut
                format_func=lambda p: {"3mo":"3 mois", "6mo":"6 mois", "1y":"1 an", "2y":"2 ans", "5y":"5 ans"}[p],
            )

        row = df_f[df_f["ticker"] == ticker_sel].iloc[0].to_dict()

        # ── Bloc info ────────────────────────────────────────────
        col1, col2 = st.columns([2, 1])
        with col1:
            st.markdown(f"#### {row.get('name', ticker_sel)} ({ticker_sel})")
            st.markdown(
                f"<span style='color:{COLORS['text_mid']}'>"
                f"{get_country_from_ticker(ticker_sel)} · {row.get('sector_fr', '-')} · "
                f"{'✅ Éligible PEA' if row.get('pea') else '🌍 CTO uniquement'}"
                f"</span>",
                unsafe_allow_html=True,
            )
        with col2:
            br = row.get("boursorama_link")
            yh = row.get("yahoo_link")
            if br: st.markdown(f"🏛️ [Boursorama]({br})")
            if yh: st.markdown(f"🔍 [Yahoo Finance]({yh})")

        # ── Métriques ───────────────────────────────────────────
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Cours", f"{row.get('price_eur', 0):.2f} €",
                  delta=f"natif: {row.get('price', 0):.2f} {row.get('currency', '')}")
        m2.metric("Perf 1M", f"{row.get('perf_1m', 0):+.1f}%" if pd.notna(row.get('perf_1m')) else "-")
        m3.metric("Cible 12M", f"{row.get('target_pct', 0):+.1f}%")
        m4.metric("Dividende", f"{row.get('div_pct', 0):.2f}%")
        m5.metric("Potentiel total", f"{row.get('total_pct', 0):+.1f}%")

        st.markdown(
            f"**Reco analystes** : {row.get('reco_label', '-')} "
            f"({int(row.get('analyst_count', 0))} analystes, "
            f"score {row.get('reco_mean', 0):.2f}/5 où 1=★★★★★)"
        )

        st.markdown("---")

        # ── Bouton Analyse Claude ────────────────────────────────
        prompt = build_claude_prompt(row)

        col_btn, col_info = st.columns([1, 2])
        with col_btn:
            # Composant HTML avec copy-to-clipboard + redirect vers Claude
            prompt_js = json.dumps(prompt)
            url_claude = "https://claude.ai/new"
            st.components.v1.html(f"""
            <div style="margin: 0;">
                <button id="claude-btn" style="
                    background: #00d4ff;
                    color: #0a1628;
                    border: none;
                    padding: 12px 24px;
                    font-weight: 800;
                    font-size: 14px;
                    letter-spacing: 1px;
                    cursor: pointer;
                    border-radius: 4px;
                    width: 100%;
                    font-family: 'JetBrains Mono', monospace;
                    text-transform: uppercase;
                " onclick="copyAndRedirect()">
                    🤖 Analyser avec Claude
                </button>
                <div id="status" style="margin-top: 10px; color: #00ff66; font-family: monospace; font-size: 13px;"></div>
            </div>
            <script>
            function copyAndRedirect() {{
                const prompt = {prompt_js};
                navigator.clipboard.writeText(prompt).then(() => {{
                    document.getElementById('status').innerHTML =
                        '✅ Prompt copié ! Ouverture de Claude...';
                    setTimeout(() => {{
                        window.open('{url_claude}', '_blank');
                    }}, 500);
                }}).catch(err => {{
                    document.getElementById('status').innerHTML =
                        '⚠️ Erreur copie : ' + err + '<br>Voir le prompt ci-dessous pour copier manuellement.';
                }});
            }}
            </script>
            """, height=120)

        with col_info:
            st.info(
                "🤖 **Comment ça marche :** Le clic copie automatiquement un prompt "
                "détaillé dans ton presse-papier et ouvre Claude.ai dans un nouvel onglet. "
                "Tu n'as plus qu'à coller (Ctrl+V) pour obtenir l'analyse complète "
                "de cette action."
            )

        with st.expander("📄 Voir le prompt qui sera copié"):
            st.code(prompt, language="markdown")

        st.markdown("---")

        # ── Chart historique ────────────────────────────────────
        st.markdown(f"#### 📈 Cours historique sur {period_sel}")

        with st.spinner(f"Chargement de l'historique {ticker_sel}..."):
            hist = fetch_ticker_history(ticker_sel, period=period_sel)

        if hist.empty:
            st.warning(f"⚠️ Historique indisponible pour {ticker_sel}")
        else:
            # Calcul moyennes mobiles
            hist["MA20"]  = hist["Close"].rolling(window=20).mean()
            hist["MA50"]  = hist["Close"].rolling(window=50).mean()
            hist["MA200"] = hist["Close"].rolling(window=200).mean()

            # Régression linéaire
            x_num = list(range(len(hist)))
            slope_data = polyfit(x_num, hist["Close"].values, 1)
            trend = poly1d(slope_data)(x_num)
            slope_pct_year = (slope_data[0] / hist["Close"].iloc[-1]) * 252 * 100  # ~252 jours bourse/an

            # Plot
            fig_hist = go.Figure()
            # Cours
            fig_hist.add_trace(go.Scatter(
                x=hist.index, y=hist["Close"],
                name="Cours", line=dict(color=COLORS["blue"], width=2),
            ))
            # MA20
            if hist["MA20"].notna().any():
                fig_hist.add_trace(go.Scatter(
                    x=hist.index, y=hist["MA20"],
                    name="MA 20j", line=dict(color=COLORS["amber"], width=1.5, dash="dot"),
                ))
            # MA50
            if hist["MA50"].notna().any():
                fig_hist.add_trace(go.Scatter(
                    x=hist.index, y=hist["MA50"],
                    name="MA 50j", line=dict(color=COLORS["text_mid"], width=1.5, dash="dot"),
                ))
            # MA200
            if hist["MA200"].notna().any():
                fig_hist.add_trace(go.Scatter(
                    x=hist.index, y=hist["MA200"],
                    name="MA 200j", line=dict(color=COLORS["dim"], width=1.5, dash="dot"),
                ))
            # Régression
            fig_hist.add_trace(go.Scatter(
                x=hist.index, y=trend,
                name=f"Tendance ({slope_pct_year:+.1f}%/an)",
                line=dict(color=COLORS["green"] if slope_pct_year > 0 else COLORS["red"],
                          width=2, dash="dash"),
            ))

            fig_hist = plot_styled_layout(fig_hist, height=500)
            fig_hist.update_layout(
                xaxis_title="Date",
                yaxis_title=f"Cours ({row.get('currency', '')})",
                hovermode="x unified",
            )
            st.plotly_chart(fig_hist, use_container_width=True)

            # Stats résumé
            c1, c2, c3, c4 = st.columns(4)
            close_first = hist["Close"].iloc[0]
            close_last  = hist["Close"].iloc[-1]
            perf_period = (close_last / close_first - 1) * 100
            c1.metric("Perf période", f"{perf_period:+.1f}%")
            c2.metric("Tendance régression", f"{slope_pct_year:+.1f}%/an")
            c3.metric("Cours haut", f"{hist['Close'].max():.2f}")
            c4.metric("Cours bas",  f"{hist['Close'].min():.2f}")

            # Volume
            with st.expander("📊 Volume de transactions"):
                fig_vol = px.bar(
                    hist.reset_index(), x=hist.index.name or "Date", y="Volume",
                    color_discrete_sequence=[COLORS["blue_dim"]],
                )
                fig_vol = plot_styled_layout(fig_vol, height=250)
                st.plotly_chart(fig_vol, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────
#  TAB 4 : PAR SECTEUR (meilleur PEA vs CTO)
# ─────────────────────────────────────────────────────────────────────
with tab_sectors:
    st.markdown("### 📂 Top potentiels par secteur GICS")
    st.caption("Pour chaque secteur : meilleur ticker PEA vs meilleur ticker CTO (par potentiel total)")

    # Construit la table sec_aligned côté Streamlit
    sectors_unique = df_f["sector_fr"].dropna().unique()
    rows_aligned = []
    for sec in sectors_unique:
        df_sec = df_f[df_f["sector_fr"] == sec]
        best_pea = df_sec[df_sec["pea"] == True].nlargest(1, "total_pct")
        best_cto = df_sec[df_sec["pea"] == False].nlargest(1, "total_pct")
        rows_aligned.append({
            "Secteur": sec,
            "PEA": f"{best_pea.iloc[0]['name']} · {best_pea.iloc[0]['ticker']} · {best_pea.iloc[0]['total_pct']:+.1f}%" if not best_pea.empty else "-",
            "PEA score": best_pea.iloc[0]['total_pct'] if not best_pea.empty else None,
            "CTO": f"{best_cto.iloc[0]['name']} · {best_cto.iloc[0]['ticker']} · {best_cto.iloc[0]['total_pct']:+.1f}%" if not best_cto.empty else "-",
            "CTO score": best_cto.iloc[0]['total_pct'] if not best_cto.empty else None,
        })

    df_aligned = pd.DataFrame(rows_aligned)
    # Tri par max(PEA score, CTO score)
    df_aligned["max_score"] = df_aligned[["PEA score", "CTO score"]].max(axis=1)
    df_aligned = df_aligned.sort_values("max_score", ascending=False).drop(columns=["max_score"])

    st.dataframe(
        df_aligned,
        use_container_width=True,
        column_config={
            "PEA score": st.column_config.NumberColumn("PEA score", format="%+.1f %%"),
            "CTO score": st.column_config.NumberColumn("CTO score", format="%+.1f %%"),
        },
        hide_index=True,
        height=500,
    )

    # Récap par secteur
    st.markdown("---")
    st.markdown("### 📊 Stats par secteur")
    sec_stats = (df_f.groupby("sector_fr")
                     .agg(
                         Tickers=("ticker", "count"),
                         Potentiel_moyen=("total_pct", "mean"),
                         Perf_moyenne=("perf_1m", "mean"),
                         Dividende_moyen=("div_pct", "mean"),
                     )
                     .sort_values("Potentiel_moyen", ascending=False)
                     .reset_index()
                     .rename(columns={"sector_fr": "Secteur"}))
    st.dataframe(
        sec_stats,
        use_container_width=True,
        column_config={
            "Potentiel_moyen": st.column_config.NumberColumn("Potentiel moyen", format="%+.1f %%"),
            "Perf_moyenne":   st.column_config.NumberColumn("Perf moyenne",   format="%+.1f %%"),
            "Dividende_moyen": st.column_config.NumberColumn("Div moyen",     format="%.2f %%"),
        },
        hide_index=True,
    )


# ─────────────────────────────────────────────────────────────────────
#  TAB 5 : À PROPOS
# ─────────────────────────────────────────────────────────────────────
with tab_about:
    st.markdown(f"""
### 📊 Brief Mensuel Bourse — Companion App

Cette application est le **complément** du Brief Mensuel publié chaque mois sur LinkedIn.
Elle te permet d'**explorer librement** les +1000 actions analysées chaque mois,
avec filtres dynamiques, graphiques interactifs et analyse IA.

### 🎯 La règle d'or

> **SOCLE 50-60% = 2 ETF mondiaux** (S&P 500 + STOXX 600)  
> **FUN 40-50% = stock-picking**, 1 action par secteur minimum  
>
> Sur 10 ans, ~85% des fonds gérés activement se font ratiboiser par leur indice (étude SPIVA).
> Ton SOCLE, c'est l'ETF. Ton FUN, c'est ce que tu trouves ici.

### 📡 Sources & méthode

- **Univers** : ~1500 actions (SP500 + SP400 + NASDAQ + STOXX 600 + DAX + MDAX + SBF120 + FTSE250 + Nikkei + TSX60 + ASX50 + HSI)
- **Filtre Boursorama** : seules les actions cotables sur Boursorama sont retenues
- **Données** : Yahoo Finance (consensus analystes, cibles, dividendes, recommandations)
- **Conversion EUR** : taux FX Yahoo en temps réel
- **Mise à jour** : automatique le **premier jour ouvré** de chaque mois (hors jeudis et fériés FR)

### 🤖 Analyse IA

Le bouton **"Analyser avec Claude"** copie automatiquement un prompt structuré dans ton presse-papier
et ouvre [Claude.ai](https://claude.ai) dans un nouvel onglet. Tu n'as plus qu'à coller (Ctrl+V).

### ⚠️ Disclaimer

Les données présentées sur cette application sont issues de Yahoo Finance et ne constituent
**en aucun cas un conseil en investissement**. L'investissement en bourse comporte un **risque
de perte en capital**. Les performances passées ne préjugent pas des performances futures.

Fais tes propres recherches et consulte un conseiller financier indépendant si nécessaire.

### 🔗 Liens

- 📦 Code source : [github.com/{REPO_OWNER}/{REPO_NAME}](https://github.com/{REPO_OWNER}/{REPO_NAME})
- 💼 LinkedIn : [Romain Taugourdeau]({LINKEDIN_URL})
- 💳 Parrainage Boursorama (+100€) : [{CODE_PARRAIN}]({PARRAINAGE_URL})

---

*Brief Mensuel Bourse · Open Source · MIT License*
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════
#  FOOTER
# ═════════════════════════════════════════════════════════════════════

st.markdown("---")
st.markdown(
    f"<div style='text-align:center; color:{COLORS['dim']}; font-size:12px; padding:20px 0;'>"
    f"<strong>BRIEF MENSUEL EQUITY</strong> · "
    f"« Risque de perte en capital. Ceci n'est pas un conseil en investissement. » · "
    f"Source : Yahoo Finance"
    f"</div>",
    unsafe_allow_html=True,
)
