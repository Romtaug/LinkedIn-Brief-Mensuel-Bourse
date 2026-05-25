"""
═══════════════════════════════════════════════════════════════════════
  brief_mensuel.py · BRIEF MENSUEL EQUITY - v10
  ─────────────────────────────────────────────────────────────────────
  Pipeline complet automatisé pour LinkedIn :
    1. Vérifie si on est le premier jour ouvré du mois → sinon skip
    2. Fetch benchmarks (S&P 500, CAC 40, STOXX 600) D'ABORD (anti rate-limit)
    3. Fetch yfinance (~1480 actions) - 4 workers + pause 15s entre univers
    4. Génère xlsx + post LinkedIn + vidéo MP4 portrait 1080×1350
    5. Upload vidéo sur litterbox.catbox.moe
    6. Envoie webhook Make.com avec post + (commentaire conditionnel si split)
    7. Rotation snapshots (garde 5 derniers max)

  🆕 v10 - Changelog vs v7 :
    • Rate limit yfinance corrigé : 4 workers, benchmarks d'abord, sleep 15s
    • Section secteurs renommée "TOP 10 POTENTIELS PAR SECTEUR" (10 secteurs GICS alignés PEA vs CTO côte à côte)
    • Fix bug NaN dans liens BR : helper _safe_url() filtre None/NaN/"nan" pour ne pas afficher "BR : nan"
    • Défilement ligne par ligne aussi sur la section secteurs
    • Tickers avec 2 liens (BR + YF) dans Top 5 Perf + Top 5 Pred + Secteurs
    • Auto-linkify LinkedIn cassé via Zero-Width Space après le point
    • Réactions LinkedIn : 👍 J'aime → 💡 Instructif
    • Vidéo 30s pile (Cover 4 / Perf 5 / Pred 5 / Sect 5 / CTA 11)
    • Encodage placebo CRF 12 + audio 320k (quasi-lossless)
    • Disclaimer harmonisé : « Risque de perte en capital. Ceci n'est pas un conseil en investissement. »
    • Hook : suppression "Sans hype. Juste de la data."
    • "+1000 actions analysées" figé dans le post (valeurs exactes dans la vidéo)
    • Pills d'indices en intro vidéo : 1 ligne pleine largeur + "+X AUTRES" si débordement
    • Badge "À BIENTÔT" plus long (visible 10s+)
    • Split automatique en post + commentaire si >3000 chars LinkedIn
    • Auto-download musique : utilise mp3 dans assets/ sinon download depuis IA

  📅 PLANIFICATION :
    Le workflow PROD se déclenche tous les 1-4 du mois à 7h UTC.
    Le script vérifie : si aujourd'hui = premier jour ouvré du mois → RUN.
    Sinon → skip (exit 0 propre).

  🔐 SECRETS GITHUB ACTIONS REQUIS :
    · WEBHOOK_URL       : URL webhook Make.com
    · CODE_PARRAINAGE   : (optionnel, défaut: ROTA0058)
    · PARRAINAGE_URL    : (optionnel, défaut: bour.so/p/GB93ZfQVNVr)

  🎮 VARIABLES D'ENV :
    · TEST_MODE        : 'true' (30 tickers mélangés) ou 'false' (full ~1480)
    · SEND_TO_WEBHOOK  : 'true'/'false' (par défaut true)
    · FORCE_RUN        : 'true' pour bypass check premier jour ouvré

  Auteur : Romain Taugourdeau
═══════════════════════════════════════════════════════════════════════
"""
from __future__ import annotations

import base64
import html as html_lib
import io
import json
import logging
import os
import platform
import random
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date, time as dt_time, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

PARIS_TZ = ZoneInfo("Europe/Paris")


# ═════════════════════════════════════════════════════════════════════
#  1. CONFIG - toute la config lit les env vars (GitHub Actions secrets)
# ═════════════════════════════════════════════════════════════════════

def _env_bool(key: str, default: bool = False) -> bool:
    val = os.getenv(key, "").lower().strip()
    if val in ("true", "1", "yes", "y", "on"):
        return True
    if val in ("false", "0", "no", "n", "off"):
        return False
    return default

# ── Modes & quantités ────────────────────────────────────────────────
TEST_MODE          = _env_bool("TEST_MODE", default=True)
N_TICKERS_TEST     = int(os.getenv("N_TICKERS_TEST", "50"))  # ← TOTAL en mode test (mélangé toutes zones)
N_TOP              = 4             # Top 5 Perf + Top 5 Pred (post)
N_TOP_VIDEO        = 10            # Top 10 dans la vidéo
N_SECTOR_PER_COL   = 10            # (Legacy) 10 secteurs en PEA + 10 en CTO
N_SECTORS_ALIGNED  = 11            # 11 secteurs GICS alignés PEA vs CTO côte à côte
FILTER_BOURSO_ONLY = True          # ⚠️ Si True : on skip TOUS les tickers sans lien Boursorama
                                   # (= filtre dur sur Japon, HK, Canada, Australie, Norvège, etc.)
N_WORKERS          = 4             # ← 4 = sweet spot anti rate-limit yfinance
SLEEP_BETWEEN_UNI  = 30            # 20→30s : compense le +600 tickers Russell
SLEEP_AFTER_BENCH  = 10            # secondes après les benchmarks avant fetch universe

# ── LinkedIn ─────────────────────────────────────────────────────────
LINKEDIN_POST_MAX  = 4000         # Limite officielle API LinkedIn UGC Posts
LINKEDIN_COMMENT_MAX = 1250        # Limite officielle commentaires LinkedIn
N_ACTIONS_DISPLAY  = "+1000"       # Texte figé dans le hook (peu importe la valeur réelle)

# ── Outputs & integrations ───────────────────────────────────────────
SEND_TO_WEBHOOK    = _env_bool("SEND_TO_WEBHOOK", default=True)
WEBHOOK_URL        = os.getenv("WEBHOOK_URL", "").strip()

# ── Litterbox (host temporaire vidéo) ────────────────────────────────
LITTERBOX_EXPIRATION  = "24h"
LITTERBOX_MAX_RETRIES = 3

# ── Output paths ─────────────────────────────────────────────────────
OUT_DIR        = Path("out")
SNAPSHOT_DIR   = Path("snapshots")
MAX_SNAPSHOTS  = 5

# ── Branding & links ─────────────────────────────────────────────────
SIGNATURE         = "ROMAIN TAUGOURDEAU"
RUBRIQUE          = "BRIEF MENSUEL EQUITY"
PARRAINAGE        = os.getenv("PARRAINAGE_URL", "https://bour.so/p/GB93ZfQVNVr").strip()
CODE_PARRAINAGE   = os.getenv("CODE_PARRAINAGE", "ROTA0058").strip()
ETF_SP500_URL     = "https://www.boursorama.com/bourse/trackers/cours/1rTETZ/"
ETF_STOXX_URL     = "https://www.boursorama.com/bourse/trackers/cours/1rTESE/"

# ── Vidéo : config encodage ──────────────────────────────────────────
# Format LinkedIn Feed optimal : 1080×1350 portrait 4:5
VIDEO_W, VIDEO_H = 1080, 1350
VIDEO_FPS        = 30
VIDEO_CRF        = 18
VIDEO_PRESET     = "veryslow"
# ── Vidéo : timing total = 30s pile ──────────────────────────────────
DUR_COVER        = 5.0
DUR_TOP_PERF     = 6.5
DUR_TOP_PRED     = 6.5
DUR_SECTORS      = 6.5
DUR_CTA          = 5.5            # Long pour que le badge "À BIENTÔT" soit bien visible
TOTAL_DURATION   = DUR_COVER + DUR_TOP_PERF + DUR_TOP_PRED + DUR_SECTORS + DUR_CTA  # 30s

# ── Vidéo : effets ────────────────────────────────────────────────────
VIDEO_FADE_IN    = 0.0
VIDEO_FADE_OUT   = 0.5             # Court pour ne pas masquer le badge
VIGNETTE         = True
KEN_BURNS        = True
XFADE_DURATION   = 0.4

# ── Audio : config musique ───────────────────────────────────────────
MUSIC_FILE       = Path("assets/music.mp3")  # mp3 perso prioritaire
AUDIO_BITRATE    = "320k"          # Quasi-CD quality
AUDIO_FADE_IN    = 0.3
AUDIO_FADE_OUT   = 2.0
AUDIO_VOLUME     = 0.6

# Auto-download : si music.mp3 absent → download depuis Internet Archive (CC BY-SA 4.0)
AUTO_DOWNLOAD_MUSIC = True
DEFAULT_MUSIC_URLS = [
    "https://archive.org/download/100_free_royalty_background_music_tracks/EverythingIsGonnaBeOk.mp3",
    "https://archive.org/download/100_free_royalty_background_music_tracks/FreeLife.mp3",
    "https://archive.org/download/100_free_royalty_background_music_tracks/bright.mp3",
]


# ═════════════════════════════════════════════════════════════════════
#  2. LOGGING
# ═════════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    stream=sys.stdout,
)
log = logging.getLogger("brief")

def banner(title: str, char: str = "═", width: int = 70) -> None:
    log.info("\n%s", char * width)
    log.info("  %s", title)
    log.info("%s", char * width)


# ═════════════════════════════════════════════════════════════════════
#  3. AUTO-INSTALL (local convenience - CI installe via workflow YAML)
# ═════════════════════════════════════════════════════════════════════

def _pip_install(pkg: str) -> None:
    log.info("⚙ Installation %s…", pkg)
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", pkg],
        stdout=subprocess.DEVNULL,
    )

def _ensure_pkg(pkg: str, import_name: str | None = None) -> None:
    try:
        __import__(import_name or pkg)
    except ImportError:
        _pip_install(pkg)

# Install only if missing (idempotent - no-op en CI où deps preinstalled)
for _p, _i in [
    ("yfinance",       "yfinance"),
    ("pandas",         "pandas"),
    ("requests-cache", "requests_cache"),
    ("openpyxl",       "openpyxl"),
    ("tqdm",           "tqdm"),
    ("requests",       "requests"),
    ("pillow",         "PIL"),
    ("playwright",     "playwright"),
    ("imageio-ffmpeg", "imageio_ffmpeg"),
]:
    _ensure_pkg(_p, _i)

import pandas as pd
import requests
import requests_cache
import yfinance as yf
from PIL import Image
from playwright.sync_api import sync_playwright
from tqdm import tqdm


# Chemin vers ffmpeg (rempli par ensure_chromium_and_ffmpeg ci-dessous)
FFMPEG_BIN: str = "ffmpeg"


def ensure_chromium_and_ffmpeg() -> None:
    """Idempotent: install Playwright Chromium + system deps + verify ffmpeg."""
    global FFMPEG_BIN

    log.info("⚙ Vérification Chromium pour Playwright…")
    subprocess.check_call(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        stdout=subprocess.DEVNULL,
    )
    if platform.system() == "Linux":
        try:
            subprocess.check_call(
                [sys.executable, "-m", "playwright", "install-deps", "chromium"],
                stdout=subprocess.DEVNULL,
            )
        except subprocess.CalledProcessError:
            log.warning("  ⚠️  install-deps a échoué - pas grave si déjà installé")
    log.info("  ✓ Chromium prêt")

    # 1. Essai ffmpeg système (PATH)
    system_ffmpeg = shutil.which("ffmpeg")
    if system_ffmpeg:
        FFMPEG_BIN = system_ffmpeg
        log.info("  ✓ ffmpeg système → %s", FFMPEG_BIN)
        return

    # 2. Fallback : binaire embarqué dans imageio-ffmpeg
    try:
        import imageio_ffmpeg
        FFMPEG_BIN = imageio_ffmpeg.get_ffmpeg_exe()
        log.info("  ✓ ffmpeg via imageio-ffmpeg → %s", FFMPEG_BIN)
    except (ImportError, RuntimeError) as e:
        raise RuntimeError(
            "ffmpeg introuvable. Installe-le :\n"
            "  · Linux  : sudo apt-get install ffmpeg\n"
            "  · macOS  : brew install ffmpeg\n"
            "  · Windows: winget install Gyan.FFmpeg  (puis redémarre VS Code)\n"
            f"  Détail : {e}"
        )


# requests-cache pour yfinance (réduit le rate-limit Yahoo)
requests_cache.install_cache(".yf_cache.sqlite", expire_after=6 * 3600)


# ═════════════════════════════════════════════════════════════════════
#  4. UNIVERS - ~1480 TICKERS (100% vérifiés sur yfinance)
# ═════════════════════════════════════════════════════════════════════
# ⚠️  BRK.B → BRK-B et BF.B → BF-B : yfinance utilise le tiret pour les
#     classes d'actions. Le point fait échouer le fetch.
# ⚠️  Tickers Yahoo Finance suffixes :
#       .PA Paris   .DE Frankfurt   .AS Amsterdam   .BR Brussels
#       .MI Milano  .MC Madrid      .LS Lisbonne    .HE Helsinki
#       .OL Oslo    .ST Stockholm   .CO Copenhague  .VI Vienne
#       .IR Dublin  .SW Zurich      .L  Londres     .WA Varsovie
#       .AT Athènes .T  Tokyo       .TO Toronto     .AX Sydney
#       .HK Hong Kong

# ── S&P 500 (~503 tickers - base US) ─────────────────────────────────
SP500 = [
    "MMM","AOS","ABT","ABBV","ACN","ADBE","AMD","AES","AFL","A","APD","ABNB",
    "AKAM","ALB","ARE","ALGN","ALLE","LNT","ALL","GOOGL","GOOG","MO","AMZN",
    "AMCR","AEE","AEP","AXP","AIG","AMT","AWK","AMP","AME","AMGN","APH","ADI",
    "ANSS","AON","APA","APO","AAPL","AMAT","APTV","ACGL","ADM","ANET","AJG",
    "AIZ","T","ATO","ADSK","ADP","AZO","AVB","AVY","AXON","BKR","BALL","BAC",
    "BAX","BDX","BRK-B","BBY","TECH","BIIB","BLK","BX","BK","BA","BKNG","BSX",
    "BMY","AVGO","BR","BRO","BF-B","BLDR","BG","BXP","CHRW","CDNS","CZR","CPT",
    "CPB","COF","CAH","KMX","CCL","CARR","CAT","CBOE","CBRE","CDW","CE","COR",
    "CNC","CNP","CF","CRL","SCHW","CHTR","CVX","CMG","CB","CHD","CI","CINF",
    "CTAS","CSCO","C","CFG","CLX","CME","CMS","KO","CTSH","COIN","CL","CMCSA",
    "CAG","COP","ED","STZ","CEG","COO","CPRT","GLW","CPAY","CTVA","CSGP","COST",
    "CTRA","CRWD","CCI","CSX","CMI","CVS","DHR","DRI","DVA","DAY","DECK","DE",
    "DELL","DAL","DVN","DXCM","FANG","DLR","DFS","DG","DLTR","D","DPZ","DASH",
    "DOV","DOW","DHI","DTE","DUK","DD","EMN","ETN","EBAY","ECL","EIX","EW",
    "EA","ELV","EMR","ENPH","ETR","EOG","EPAM","EQT","EFX","EQIX","EQR","ERIE",
    "ESS","EL","EG","EVRG","ES","EXC","EXE","EXPE","EXPD","EXR","XOM","FFIV",
    "FDS","FICO","FAST","FRT","FDX","FIS","FITB","FSLR","FE","FI","F","FTNT",
    "FTV","FOXA","FOX","BEN","FCX","GRMN","IT","GE","GEHC","GEV","GEN","GNRC",
    "GD","GIS","GM","GPC","GILD","GPN","GL","GDDY","GS","HAL","HIG","HAS",
    "HCA","DOC","HSIC","HSY","HES","HPE","HLT","HOLX","HD","HON","HRL","HST",
    "HWM","HPQ","HUBB","HUM","HBAN","HII","IBM","IEX","IDXX","ITW","INCY","IR",
    "PODD","INTC","ICE","IFF","IP","IPG","INTU","ISRG","IVZ","INVH","IQV","IRM",
    "JBHT","JBL","JKHY","J","JNJ","JCI","JPM","K","KVUE","KDP","KEY","KEYS",
    "KMB","KIM","KMI","KKR","KLAC","KHC","KR","LHX","LH","LRCX","LW","LVS",
    "LDOS","LEN","LII","LIN","LYV","LKQ","LMT","L","LOW","LULU","LYB","MTB",
    "MPC","MKTX","MAR","MMC","MLM","MAS","MA","MTCH","MKC","MCD","MCK","MDT",
    "MRK","META","MET","MTD","MGM","MCHP","MU","MSFT","MAA","MRNA","MHK","MOH",
    "TAP","MDLZ","MPWR","MNST","MCO","MS","MOS","MSI","MSCI","NDAQ","NTAP","NFLX",
    "NEM","NWSA","NWS","NEE","NKE","NI","NDSN","NSC","NTRS","NOC","NCLH","NRG",
    "NUE","NVDA","NVR","NXPI","ORLY","OXY","ODFL","OMC","ON","OKE","ORCL","OTIS",
    "PCAR","PKG","PLTR","PANW","PSKY","PH","PAYX","PAYC","PYPL","PNR","PEP","PFE",
    "PCG","PM","PSX","PNW","PNC","POOL","PPG","PPL","PFG","PG","PGR","PLD",
    "PRU","PEG","PTC","PSA","PHM","PWR","QCOM","DGX","RL","RJF","RTX","O",
    "REG","REGN","RF","RSG","RMD","RVTY","ROK","ROL","ROP","ROST","RCL","SPGI",
    "CRM","SBAC","SLB","STX","SRE","NOW","SHW","SPG","SWKS","SJM","SW","SNA",
    "SOLV","SO","LUV","SWK","SBUX","STT","STLD","STE","SYK","SMCI","SYF","SNPS",
    "SYY","TMUS","TROW","TTWO","TPR","TRGP","TGT","TEL","TDY","TER","TSLA","TXN",
    "TPL","TXT","TMO","TJX","TKO","TSCO","TT","TDG","TRV","TRMB","TFC","TYL",
    "TSN","USB","UBER","UDR","ULTA","UNP","UAL","UPS","URI","UNH","UHS","VLO",
    "VTR","VLTO","VRSN","VRSK","VZ","VRTX","VTRS","VICI","V","VST","VMC","WRB",
    "GWW","WAB","WBA","WMT","DIS","WBD","WM","WAT","WEC","WFC","WELL","WST",
    "WDC","WY","WSM","WMB","WTW","WDAY","WYNN","XEL","XYL","YUM","ZBRA","ZBH","ZTS",
]

# ── S&P 400 MidCap (~400 US mid caps - Boursorama compatible) ────────
SP400_MID = [
    "ACIW","ACM","ADC","AFG","AGCO","ALE","ALK","ALV","AMG","AN","APG","APLE",
    "ARW","ASGN","ATR","AUB","AVA","AVNT","AXS","AYI","BC","BCC","BCO","BCPC",
    "BERY","BFAM","BFH","BIO","BJ","BKH","BLD","BMI","BPOP","BWA","BWXT","BXMT",
    "BYD","CACI","CADE","CASY","CATY","CBSH","CBT","CDP","CFR","CHCO","CHDN",
    "CHE","CHX","CIEN","CIVI","CLF","CLH","CMC","CNH","CNM","CNXC","CMA",
    "COKE","COLM","CR","CRC","CROX","CRS","CSL","CUBE","CW","DECK","DINO",
    "DKS","DLB","DOCS","DT","DTM","DY","EAT","EEFT","EHC","ELS","EME","ENS",
    "ENV","EPC","EPR","EQH","ESAB","ESI","ESNT","ETSY","EVR","EWBC","EXEL",
    "EXLS","FAF","FBP","FCN","FFIN","FHB","FHI","FHN","FIVE","FIX","FLG",
    "FLO","FLR","FN","FNB","FOUR","FR","FRPT","FSS","FTI","G","GAP","GATX",
    "GBCI","GFF","GGG","GHC","GME","GMED","GNTX","GNW","GO","GPI","GPK","GTLS",
    "GTLB","GXO","H","HASI","HE","HELE","HGV","HOG","HOMB","HP","HQY","HRB",
    "HUBG","HXL","IART","IBKR","IBOC","IBP","ICUI","IDA","IDCC","IEX","INGR",
    "INSP","IPAR","IRDM","ITT","JEF","JLL","JWN","JXN","KAI","KBR","KD","KEX",
    "KMPR","KNF","KNSL","KRC","KSS","LAMR","LANC","LECO","LEG","LITE","LIVN",
    "LNC","LNTH","LOPE","LSCC","LSTR","M","MAN","MASI","MAT","MATX","MC","MDU",
    "MEDP","MGY","MIDD","MLI","MMS","MOG-A","MORN","MSA","MSM","MTG","MTH",
    "MTSI","MTZ","MUR","MUSA","NBIX","NBR","NCNO","NEU","NFG","NJR","NNN","NOG",
    "NOV","NSA","NSP","NTNX","NVST","NWE","OC","OGE","OGN","OGS","OHI","OII",
    "OLED","OLLI","OLN","OMF","ORI","OSCR","OSK","OVV","OZK","PB","PBF","PCH",
    "PCTY","PCVX","PEN","PII","PIPR","PNFP","POR","POST","PR","PRGO","PRI",
    "PRVA","PSN","PTEN","R","RBC","RDN","RGA","RGEN","RGLD","RH","RHI","RIVN",
    "RL","RNR","RPM","RRC","RRX","RUSHA","RYAN","SAIA","SAIC","SAM","SANM",
    "SBS","SEE","SF","SFM","SHC","SHOO","SIGI","SITE","SKX","SLAB","SLG","SLGN",
    "SLM","SMG","SNDR","SNV","SNX","SON","SR","SRPT","SSB","SSD","ST","STAG",
    "STWD","SWX","SXT","TCBI","TDC","TDS","TDW","TEX","TFX","TGNA","THC","THG",
    "THO","TKR","TMHC","TNL","TOL","TPH","TPX","TREX","TRMK","TRNO","TTC","TTEK",
    "TWLO","TWO","UAA","UA","UFPI","UGI","UHS","UNF","UNFI","UNM","USFD","UTHR",
    "VC","VFC","VLY","VMI","VNO","VNT","VOYA","VSAT","VSCO","VVV","WAFD","WAL",
    "WBS","WCC","WEN","WEX","WERN","WEX","WH","WLK","WMS","WOLF","WPC","WSC",
    "WSO","WTRG","WTS","WU","X","XPO","XRX","YELP","ZD","ZWS","ZION",
]

# ── S&P 600 SmallCap (~600 US small caps qualité - Boursorama compatible) ─
SP600_SMALL = [
    "AAOI","AAON","ABCB","ABG","ABM","ABR","ACA","ACAD","ACEL","ACLS","ADEA",
    "ADMA","ADNT","ADUS","AEIS","AEO","AGS","AGYS","AHCO","AIN","AIR","AIT",
    "AKR","AL","ALEX","ALG","ALGT","ALK","ALKS","ALRM","AMBC","AMC","AMED",
    "AMN","AMOT","AMR","AMSF","AMWD","ANDE","ANIP","AORT","APAM","APLE","APOG",
    "ARCB","ARCH","ARCT","ARI","ARLO","AROC","ASB","ASIX","ASTH","ASTL","ATEN",
    "ATGE","ATKR","ATMU","ATSG","AUR","AVAV","AVNS","AWI","AWR","AX","AXL",
    "AZTA","B","BANC","BANR","BBSI","BBW","BCRX","BDC","BFS","BGC","BGS","BHE",
    "BHLB","BJRI","BKE","BKU","BLBD","BLMN","BLX","BMRC","BNL","BOH","BOX",
    "BRC","BRT","BSIG","BSRR","BTU","BV","BXC","CABO","CAKE","CAL","CALM",
    "CALX","CARG","CARS","CASH","CATO","CBL","CBU","CBZ","CCOI","CCRN","CCS",
    "CDE","CDP","CEIX","CEVA","CFFN","CHCT","CHEF","CHUY","CIO","CIR","CIVB",
    "CLB","CLW","CMP","CMTL","CNK","CNMD","CNS","CNXC","CODI","COHU","CON",
    "CONN","COOP","CORE","CORT","COTY","CPF","CPK","CPRX","CRC","CRI","CRK",
    "CRSR","CSGS","CSR","CSTL","CSWI","CTBI","CTLP","CTOS","CTS","CUBI","CUTR",
    "CVBF","CVCO","CVI","CVLT","CWAN","CWEN","CWH","CWK","CWST","CWT","CXM",
    "CXW","DAN","DBI","DCO","DEA","DEI","DEN","DFH","DFIN","DGII","DHC","DIOD",
    "DJCO","DK","DLX","DNB","DNOW","DOMO","DORM","DRH","DRQ","DV","DXC","DXPE",
    "EAF","ECG","ECPG","EFC","EGBN","EGY","EIG","ELME","EMBC","ENR","ENSG",
    "ENTA","ENV","EPAC","EPRT","EQC","ESE","ETD","ETSY","EVH","EVRI","EVTC",
    "EXLS","EXPI","EXTR","FBNC","FBP","FBRT","FCF","FCFS","FDP","FELE","FFBC",
    "FFG","FFIC","FFWM","FIBK","FIZZ","FL","FLS","FLWS","FN","FOR","FORM",
    "FRGE","FRO","FSB","FSCO","FSP","FUL","FULT","FWRD","GBX","GCO","GDEN",
    "GDOT","GEF","GEO","GES","GFF","GIII","GMS","GNL","GNTY","GO","GOLF","GOR",
    "GPI","GPMT","GPOR","GRBK","GTY","GVA","HAE","HAFC","HAIN","HBI","HCC",
    "HCKT","HCSG","HEES","HFFG","HFWA","HGV","HI","HLF","HLIT","HLX","HMN",
    "HOFT","HOPE","HP","HRMY","HSII","HSTM","HTH","HTLD","HTLF","HUBG","HVT",
    "HZO","IBP","ICFI","ICHR","IDT","IIIN","IIPR","INDB","INN","INVA","IOSP",
    "IPAR","IRWD","ITGR","ITRI","JACK","JBI","JBLU","JBSS","JBT","JCI","JOE",
    "JOUT","JWN","KALU","KAR","KE","KELYA","KFRC","KFS","KFY","KLG","KMT",
    "KN","KOS","KRG","KRO","KRYS","KSS","KTB","KW","KWR","LAUR","LAW","LBRT",
    "LCII","LCNB","LEG","LGIH","LGND","LILA","LMAT","LMND","LMNR","LNN",
    "LNW","LPG","LQDT","LRN","LTC","LXP","LZB","MAC","MATW","MATX","MBC","MBI",
    "MBIN","MBUU","MC","MCBS","MCRI","MCS","MCW","MCY","MD","MDC","MDP","MDRX",
    "MEG","MERC","MGEE","METC","MGPI","MGRC","MGY","MHO","MLAB","MLI","MLKN",
    "MLR","MMI","MMSI","MNRO","MOD","MODG","MODV","MOG-A","MOV","MP","MPW",
    "MPX","MQ","MRCY","MSEX","MSGE","MSGS","MTH","MTRN","MTUS","MUR","MWA",
    "MYE","MYGN","MYRG","NABL","NARI","NATL","NBHC","NBR","NEO","NEOG","NESR",
    "NEU","NGVT","NJR","NMIH","NOG","NOMD","NPK","NPO","NPWR","NSP","NTB",
    "NTCT","NTGR","NUS","NUVL","NWBI","NWE","NWL","NWN","NX","NXRT","NXT",
    "NYMT","OCFC","OFG","OGS","OII","OIS","OLPX","OMCL","OMI","ONB","OPCH",
    "ORA","OSIS","OSTK","OSW","OTTR","OUT","OXM","PACK","PAG","PARR","PATK",
    "PAYO","PBI","PCRX","PCYO","PDCO","PDM","PEB","PECO","PFBC","PFC","PFS",
    "PGNY","PHAT","PI","PINC","PIPR","PIRS","PJT","PK","PLAB","PLAY","PLMR",
    "PLUS","PLXS","PMT","POWI","POWL","PPBI","PRA","PRDO","PRG","PRGS","PRK",
    "PRLB","PRO","PRTH","PRVA","PSMT","PTCT","PTGX","PUMP","PZZA","QCRH","QNST",
    "QTRX","RAMP","RBBN","RBC","RC","RCKT","RCM","RCMT","RCUS","RDFN","REPL",
    "REPX","REVG","REZI","RGNX","RGR","RH","RICK","RILY","RLI","RLJ","RM",
    "RNST","ROCC","ROG","ROIC","RPAY","RUN","RUSHA","RWT","RXO","RXST","RXT",
    "RYI","SABR","SAFE","SAFT","SAH","SAM","SANM","SASR","SATS","SBCF","SBGI",
    "SBH","SBSI","SCHL","SCL","SCSC","SCVL","SDGR","SEAS","SEM","SENEA","SFBS",
    "SFL","SHAK","SHEN","SHO","SHOO","SI","SIBN","SIGA","SIRI","SITC","SITM",
    "SJW","SKT","SKYW","SLCA","SLG","SLM","SLP","SLVM","SM","SMP","SMPL","SMTC",
    "SNCY","SNDR","SNDX","SNEX","SOI","SON","SONO","SPB","SPNE","SPNT","SPOK",
    "SPR","SPSC","SPT","SPTN","SPWH","SR","SRCE","SRDX","SSTK","STAA","STAR",
    "STBA","STC","STEL","STER","STGW","STNG","STRA","STRL","STRO","SUM","SUPN",
    "SVC","SVRA","SWBI","SWI","SWX","SXC","SXI","SYBT","SYRE","TALO","TBBK",
    "TBI","TCBK","TCMD","TCPC","TCS","TDC","TDS","TDW","TGI","TGTX","THFF",
    "THR","THRY","THS","TILE","TMP","TNC","TNDM","TPB","TPC","TPH","TR","TRC",
    "TREE","TRMK","TRN","TRNS","TRST","TRTN","TRUP","TSE","TTGT","TTI","TTMI",
    "TVTX","TWI","TWNK","TXG","TXMD","UCB","UCBI","UE","UFCS","UFI","UFPI",
    "UGI","UHT","UIS","UMBF","UNF","UNFI","UNIT","UNTY","UPLD","URBN","USCR",
    "USNA","USPH","UTL","UTMD","UVE","UVSP","VBTX","VCEL","VCYT","VECO","VERX",
    "VG","VGR","VICR","VIR","VIRT","VKTX","VPG","VRA","VRDN","VRE","VREX",
    "VRRM","VRTS","VSAT","VSEC","VSH","VYX","WABC","WAFD","WD","WDFC","WERN",
    "WGO","WHD","WK","WKC","WKME","WLY","WMC","WNC","WOR","WRBY","WRLD","WS",
    "WSC","WSFS","WSR","WT","WTS","WWW","XHR","XNCR","XPEL","XPER","YELP","YETI",
    "YOU","ZD","ZIM","ZIP","ZUO",
]

# ── STOXX Europe Mid Cap (~150 mid caps EU Bourso-compatibles) ─────────
STOXX_MID_EU = [
    # 🇫🇷 France .PA (PEA)
    "ALMDT.PA","ALTR.PA","ATARI.PA","BIG.PA","BLC.PA","BNB.PA","BNV.PA","BSD.PA",
    "BVI.PA","CGG.PA","CIV.PA","COX.PA","DSY.PA","ECP.PA","EOS.PA","FII.PA",
    "FORE.PA","FOUG.PA","HCO.PA","INF.PA","IPH.PA","KORI.PA","LACR.PA","LIN.PA",
    "LSS.PA","MAA.PA","MERY.PA","NANO.PA","ORP.PA","RBT.PA","RIN.PA","RUI.PA",
    "SESL.PA","SII.PA","SOG.PA","STF.PA","SY.PA","THEP.PA","UG.PA","VANTI.PA",
    # 🇩🇪 Allemagne .DE (PEA)
    "AOX.DE","BIO3.DE","CEC.DE","CWC.DE","DEQ.DE","DRI.DE","DUE.DE","EVO.DE",
    "G1A.DE","GFT.DE","INH.DE","KCO.DE","KU2.DE","NDA.DE","NOEJ.DE","O2D.DE",
    "PFV.DE","PSAN.DE","RHK.DE","SANT.DE","SOW.DE","TKMS.DE","TUI1.DE","WAF.DE",
    # 🇮🇹 Italie .MI (PEA)
    "ALK.MI","BPE.MI","BMED.MI","BPSO.MI","CEM.MI","CMB.MI","DAL.MI","DLG.MI",
    "ENV.MI","ERG.MI","FILA.MI","GVS.MI","IG.MI","IGD.MI","IGL.MI","IRE.MI",
    "MFEA.MI","MFEB.MI","RWAY.MI","SOL.MI","SRS.MI","TES.MI","TGYM.MI","UNIR.MI",
    # 🇪🇸 Espagne .MC (PEA)
    "ALM.MC","APAM.MC","APPS.MC","ATRY.MC","CAF.MC","CIE.MC","EBRO.MC","ECR.MC",
    "EDR.MC","ENC.MC","FAE.MC","GEST.MC","LOG.MC","ORY.MC","PRS.MC","PSG.MC",
    "TUB.MC","UBS.MC","VID.MC","VIS.MC","ZOT.MC",
    # 🇳🇱 Pays-Bas .AS (PEA)
    "ABN.AS","AALB.AS","ACOMO.AS","ARCAD.AS","ASRNL.AS","BAMNB.AS","CTPNV.AS",
    "FAGR.AS","FUR.AS","NSI.AS","PHARM.AS","SBMO.AS","TKWY.AS","VPK.AS","WHA.AS",
    # 🇧🇪 Belgique .BR (PEA)
    "BAR.BR","BPOST.BR","CFEB.BR","DIE.BR","EVS.BR","JEN.BR","LOTB.BR","TESB.BR",
    # 🇬🇧 UK .L (CTO)
    "ASHM.L","AVON.L","CAPC.L","CRDA.L","DPLM.L","ECM.L","ENT.L","HSV.L","ICP.L",
    "JD.L","MNDI.L","SAFE.L","TET.L","WTAN.L",
]

# ── NASDAQ 100 hors-SP500 (tech US Bourso-compatible) ─────────────────
NASDAQ100_EXTRA = [
    "ADP","ASML","AZN","BIDU","BIIB","BKR","CDW","CHKP","CMCSA","COIN","COST",
    "CPRT","CRWD","CSGP","CSX","CTAS","CTSH","DLTR","DOCU","DXCM","EA","EBAY",
    "EXC","FANG","FAST","FTNT","GFS","GILD","HON","ILMN","INTC","INTU","ISRG",
    "JD","KDP","KHC","KLAC","LCID","LRCX","LULU","MAR","MCHP","MDLZ","MELI",
    "META","MNST","MRVL","MSFT","MU","NXPI","ODFL","ORLY","PANW","PAYX","PCAR",
    "PDD","PEP","PYPL","REGN","ROST","SBUX","SGEN","SIRI","SNPS","TEAM","TMUS",
    "TROW","TXN","VRSK","VRSN","VRTX","WBA","WDAY","XEL","ZM","ZS",
]
# Note : beaucoup chevauchent SP500. Le set() dans run_data_pipeline dédupliquera.

# ── FTSE 250 (UK mid-caps Bourso-compatible) ──────────────────────────
FTSE250 = [
    "4IMI.L","ABDN.L","ALPH.L","APAX.L","APTD.L","AVON.L","BAB.L","BBOX.L","BME.L",
    "BOY.L","BRSC.L","BRWM.L","BVIC.L","BWY.L","CAPC.L","CAY.L","CCC.L","CCL.L",
    "CINE.L","CLI.L","CRST.L","CTY.L","DARK.L","DOM.L","DPH.L","DTY.L","ELM.L",
    "ENT.L","ESP.L","ESYS.L","EXPP.L","FCSS.L","FOUR.L","FUTR.L","GAW.L","GFRD.L",
    "GFTU.L","GNS.L","GPE.L","GRG.L","HAS.L","HFD.L","HFEL.L","HGT.L","HICL.L",
    "HMSO.L","HOC.L","HSL.L","HSV.L","HSX.L","HTG.L","HVO.L","IGG.L","INCH.L",
    "INDV.L","INPP.L","INVP.L","JE.L","JMAT.L","KIE.L","LMP.L","MAB1.L","MAB2.L",
    "MGAM.L","MGGT.L","MKS.L","MONY.L","MRO.L","MSLH.L","MTRO.L","NCC.L","NESF.L",
    "OXIG.L","PAGE.L","PCT.L","PETS.L","PIN.L","PNN.L","PSON.L","QQ.L","RAT.L",
    "RDW.L","RHL.L","RHM.L","RMV.L","RNK.L","RNO.L","RSW.L","SAVE.L","SCT.L",
    "SDP.L","SHA.L","SOG.L","SPI.L","SPT.L","SQZ.L","SXS.L","SYNC.L","TBCG.L",
    "TET.L","TIFS.L","TRN.L","TRY.L","TUI.L","TYMN.L","UKW.L","ULE.L","VCT.L",
    "VEC.L","VED.L","VLE.L","VSVS.L","VTY.L","WG.L","WIZZ.L","WKP.L","WMH.L",
    "WTAN.L","YOUG.L",
]
# ── DAX 40 (Allemagne large-cap) ─────────────────────────────────────
DAX = [
    "ADS.DE","AIR.DE","ALV.DE","BAS.DE","BAYN.DE","BMW.DE","BNR.DE","CBK.DE",
    "CON.DE","1COV.DE","DTG.DE","DBK.DE","DB1.DE","DHL.DE","DTE.DE","EOAN.DE",
    "FRE.DE","HNR1.DE","HEI.DE","HEN3.DE","IFX.DE","MBG.DE","MRK.DE","MTX.DE",
    "MUV2.DE","P911.DE","PAH3.DE","QIA.DE","RHM.DE","RWE.DE","SAP.DE","SRT3.DE",
    "SIE.DE","ENR.DE","SHL.DE","SY1.DE","VOW3.DE","VNA.DE","ZAL.DE","BEI.DE",
]

# ── MDAX 50 (Allemagne mid-cap) ──────────────────────────────────────
MDAX = [
    "HOT.DE","LHA.DE","KBX.DE","TLX.DE","NDX1.DE","AIXA.DE","TKA.DE",
    "NDA.DE","HAG.DE","LEG.DE","DHER.DE","R3NK.DE","EVK.DE","NEM.DE",
    "GBF.DE","RAA.DE","KGX.DE","EVD.DE","FNTN.DE","TUI1.DE","FTK.DE","TEG.DE",
    "PUM.DE","FRA.DE","AG1.DE","SDF.DE","BC8.DE","FPE3.DE","UTDI.DE","8TRA.DE",
    "TKMS.DE","AMV0.DE","AT1.DE","DWS.DE","WCH.DE","JEN.DE","KRN.DE","DEZ.DE",
    "SHA0.DE","BOSS.DE","HLE.DE","LXS.DE","SZG.DE","IOS.DE","JUN3.DE","RRTL.DE",
"SAX.DE","RDC.DE",
]

# ── SDAX (Allemagne small caps PEA) ──────────────────────────────────
SDAX = [
    "ADV.DE","BVB.DE","CAP.DE","COK.DE","DEX.DE","DRW3.DE","ELG.DE","ENVITEC.DE",
    "GFJ.DE","GIL.DE","GMM.DE","GYC.DE","HLAG.DE","HHFA.DE","INH.DE","IVU.DE",
    "KCO.DE","KU2.DE","KWS.DE","LEO.DE","M5Z.DE","MEO.DE","MOR.DE","MUM.DE",
    "NA9.DE","NWO.DE","PFV.DE","PNE3.DE","PSAN.DE","S92.DE","SAX.DE","SFQ.DE",
    "SGL.DE","SIX2.DE","SOW.DE","SPR.DE","STM.DE","TTK.DE","VAR1.DE","VBK.DE",
    "VOS.DE","WAC.DE","WAF.DE","WCH.DE","ZAL.DE",
]

# ── TecDAX 30 (Allemagne tech mid caps PEA) ──────────────────────────
TECDAX = [
    "1U1.DE","AIXA.DE","BC8.DE","CCC3.DE","COK.DE","DRI.DE","EVT.DE",
    "FNTN.DE","FRA.DE","HHFA.DE","IFX.DE","INH.DE","JEN.DE","KRN.DE",
    "M5Z.DE","MOR.DE","NA9.DE","NEM.DE","NEX.DE","P7C.DE","PNE3.DE",
    "QIA.DE","S92.DE","SAP.DE","SHL.DE","SOW.DE","TEG.DE","TKMS.DE",
    "UTDI.DE","WAF.DE","WCH.DE",
]

# ── SBF 120 Mid (France hors CAC 40) ─────────────────────────────────
SBF120_MID = [
    "ADP.PA","AF.PA","ATE.PA","AMUN.PA","ARG.PA","ATO.PA","AYV.PA",
    "BEN.PA","BB.PA","BIM.PA","BOL.PA","CARM.PA","CLARI.PA","COFA.PA",
    "COV.PA","AM.PA","DBG.PA","FGR.PA","ELIOR.PA","ELIS.PA","EMEIS.PA","ERA.PA",
    "ES.PA","RF.PA","ENX.PA","FDJ.PA","FRVIA.PA","GFC.PA","GET.PA","GTT.PA",
    "ICAD.PA","IDL.PA","NK.PA","ITP.PA","IPN.PA","IPS.PA","DEC.PA","LI.PA",
    "MAU.PA","MEDCL.PA","MERY.PA","MRN.PA","MMT.PA","NEOEN.PA","NEX.PA","NXI.PA",
    "OPM.PA","PLNW.PA","PLX.PA","RCO.PA","RXL.PA","CBE.PA","RUI.PA","SK.PA",
    "DIM.PA","SCR.PA","SESG.PA","SW.PA","SOI.PA","SOP.PA","SPIE.PA","TE.PA",
    "TFI.PA","TRI.PA","UBI.PA","FR.PA","VK.PA","VLA.PA","VRLA.PA","VCT.PA",
    "VIRP.PA","VIRI.PA","VU.PA","MF.PA","WLN.PA",
]

# ── CAC Mid 60 + Small caps FR qualité (PEA + Bourso-compatibles) ────
# Ajout pour rééquilibrer le ratio PEA vs CTO (était 262/884)
CAC_MID_60 = [
    # CAC Mid 60 (mid caps Euronext Paris)
    "AKW.PA","ALBLD.PA","ALD.PA","ALMDT.PA","ALTA.PA","BNG.PA","BOI.PA","BON.PA",
    "BVI.PA","CAFO.PA","COFA.PA","COV.PA","CRTO.PA","DERI.PA","DIM.PA","DG.PA",
    "ELIOR.PA","ENX.PA","ERA.PA","EXE.PA","EXN.PA","FII.PA","FNAC.PA","FORE.PA",
    "GTT.PA","HCO.PA","ICAD.PA","INF.PA","IPS.PA","IPN.PA","IPH.PA","JCQ.PA",
    "KOF.PA","LACR.PA","LI.PA","LISI.PA","LR.PA","MAU.PA","MAA.PA","MMB.PA",
    "MERY.PA","MF.PA","MMT.PA","NANO.PA","NEOEN.PA","NXI.PA","OPM.PA","ORP.PA",
    "PLNW.PA","RCO.PA","RUI.PA","RXL.PA","SCR.PA","SESL.PA","SK.PA","SOI.PA",
    "SOP.PA","SOMA.PA","SPIE.PA","TE.PA","TFI.PA","UBI.PA","VCT.PA","VIRI.PA",
    "VIRP.PA","VK.PA","VLA.PA","VRLA.PA","VTSC.PA","VU.PA","WAVE.PA","WLN.PA",
    # Small caps FR qualité (post CAC Mid)
    "ABCA.PA","BEN.PA","BB.PA","BIM.PA","CBE.PA","CRI.PA","DBG.PA","DBV.PA",
    "ELIS.PA","EMEIS.PA","ES.PA","FGR.PA","FR.PA","GFC.PA","GET.PA","GUIL.PA",
    "IDL.PA","ITP.PA","MEDCL.PA","NEX.PA","PLX.PA","SECH.PA","SESG.PA","SW.PA",
    "TRI.PA","NACON.PA","RF.PA","DEC.PA","TIPI.PA",
]

# ── CAC Small (small caps FR PEA Bourso-compatibles) ─────────────────
CAC_SMALL = [
    "ALAGR.PA","ALATA.PA","ALCJ.PA","ALMER.PA","ALORA.PA","ALTER.PA","BLEE.PA",
    "BUR.PA","CDA.PA","CGM.PA","CHSR.PA","COFA.PA","DLTA.PA","ENGI.PA","ESKER.PA",
    "FNAC.PA","GTT.PA","HCO.PA","INF.PA","JCQ.PA","LACR.PA","LOCAL.PA","MAA.PA",
    "MMB.PA","MMT.PA","NANO.PA","ORP.PA","PIG.PA","PLNW.PA","SOMA.PA","TIPI.PA",
    "TOUP.PA","TRI.PA","VANTI.PA","VIRP.PA","VTSC.PA","WAVE.PA","XFAB.PA",
]

# ── STOXX Europe 600 ventilé par pays (PEA + UK FTSE 100 + Suisse + Irlande) ─
# ── STOXX Europe 600 ventilé par pays (PEA + UK FTSE 100 + Suisse + Irlande) ─
_STOXX_NATIONAL = (
    # CAC 40 - France .PA (PEA)
    ["AC.PA","AI.PA","AIR.PA","ALO.PA","AKE.PA","BNP.PA","BVI.PA","EN.PA","CAP.PA",
     "CA.PA","ACA.PA","BN.PA","DSY.PA","EDEN.PA","ENGI.PA","EL.PA","ERF.PA",
     "RMS.PA","KER.PA","LR.PA","OR.PA","MC.PA","ML.PA","ORA.PA","RI.PA",
     "PUB.PA","RNO.PA","SAF.PA","SGO.PA","SAN.PA","SU.PA","GLE.PA","STLAP.PA",
     "STMPA.PA","TEP.PA","HO.PA","TTE.PA","VIE.PA","DG.PA","VIV.PA"]
    # FTSE 100 - UK .L (NON-PEA, CTO uniquement)
    + ["AAL.L","ABF.L","ADM.L","AHT.L","ANTO.L","AZN.L","AUTO.L","AV.L","BA.L",
       "BARC.L","BATS.L","BDEV.L","BEZ.L","BKG.L","BLND.L","BNZL.L","BP.L",
       "BRBY.L","BT-A.L","CCH.L","CNA.L","CPG.L","CRDA.L","CRH.L","CTEC.L",
       "DCC.L","DGE.L","DPLM.L","EDV.L","EXPN.L","EZJ.L","FCIT.L","FRES.L",
       "GLEN.L","GSK.L","HIK.L","HLN.L","HSBA.L","HWDN.L","IAG.L","ICG.L",
       "IHG.L","III.L","IMB.L","IMI.L","INF.L","ITRK.L","ITV.L","JD.L","KGF.L",
       "LAND.L","LGEN.L","LLOY.L","LSEG.L","MKS.L","MNDI.L","MNG.L","MRO.L",
       "NG.L","NWG.L","NXT.L","OCDO.L","PHNX.L","PRU.L","PSH.L","PSN.L","PSON.L",
       "REL.L","RIO.L","RKT.L","RR.L","RS1.L","SBRY.L","SDR.L","SGE.L","SGRO.L",
       "SHEL.L","SMIN.L","SMT.L","SN.L","SPX.L","SSE.L","STAN.L","STJ.L","SVT.L",
       "TSCO.L","TW.L","ULVR.L","UTG.L","UU.L","VOD.L","WEIR.L","WPP.L","WTB.L"]
    # IBEX 35 - Espagne .MC (PEA)
    + ["ACS.MC","ACX.MC","AENA.MC","AMS.MC","ANA.MC","ANE.MC","BBVA.MC","BKT.MC",
       "CABK.MC","CLNX.MC","COL.MC","ELE.MC","ENG.MC","FDR.MC","FER.MC","GRF.MC",
       "IBE.MC","IDR.MC","ITX.MC","LOG.MC","MAP.MC","MEL.MC","MRL.MC",
       "MTS.MC","NTGY.MC","PUIG.MC","RED.MC","REP.MC","ROVI.MC","SAB.MC","SAN.MC",
       "SCYR.MC","SLR.MC","TEF.MC","UNI.MC"]
    # AEX 25 - Pays-Bas .AS (PEA)
    + ["MT.AS","ADYEN.AS","AGN.AS","AD.AS","AKZA.AS","ASM.AS","ASML.AS","ASRNL.AS",
       "BESI.AS","DSFIR.AS","EXO.AS","GLPG.AS","HEIA.AS","IMCD.AS","INGA.AS",
       "KPN.AS","NN.AS","PHIA.AS","PRX.AS","RAND.AS","REN.AS","SHELL.AS","UNA.AS",
       "URW.AS","WKL.AS"]
    # BEL 20 - Belgique .BR (PEA)
    + ["ABI.BR","ACKB.BR","AED.BR","AGS.BR","ARGX.BR","AZE.BR","COFB.BR","ELI.BR",
       "GBLB.BR","KBC.BR","MELE.BR","PROX.BR","SOF.BR","SOLB.BR","TNET.BR","UCB.BR",
       "UMI.BR","VGP.BR","WDP.BR"]
    # FTSE MIB - Italie .MI (PEA)
    + ["A2A.MI","AMP.MI","AZM.MI","BAMI.MI","BPE.MI","BMED.MI","BMPS.MI","BPSO.MI",
       "CPR.MI","DIA.MI","ENEL.MI","ENI.MI","RACE.MI","FBK.MI","G.MI","HER.MI",
       "INW.MI","ISP.MI","INTE.MI","IG.MI","IP.MI","LDO.MI","MB.MI","MONC.MI",
       "NEXI.MI","PIRC.MI","PIA.MI","PRY.MI","PST.MI","REC.MI","SPM.MI","SRG.MI",
       "STLAM.MI","STMMI.MI","TIT.MI","TRN.MI","TEN.MI","UCG.MI","UNI.MI"]
    # OMX Stockholm - Suède .ST (PEA)
    + ["ABB.ST","ALFA.ST","ASSA-B.ST","ATCO-A.ST","ATCO-B.ST","AZN.ST","BOL.ST",
       "ELUX-B.ST","ERIC-B.ST","ESSITY-B.ST","EVO.ST","GETI-B.ST","HEXA-B.ST",
       "HM-B.ST","INVE-B.ST","KINV-B.ST","NDA-SE.ST","NIBE-B.ST","SAND.ST",
       "SCA-B.ST","SEB-A.ST","SHB-A.ST","SINCH.ST","SKF-B.ST","SWED-A.ST",
       "TEL2-B.ST","TELIA.ST","VOLV-B.ST"]
    # OMX Helsinki - Finlande .HE (PEA)
    + ["ELISA.HE","FORTUM.HE","KESKOB.HE","KNEBV.HE","METSO.HE","NESTE.HE",
       "NOKIA.HE","NDA-FI.HE","ORNBV.HE","OUT1V.HE","SAMPO.HE","STERV.HE",
       "TELIA1.HE","TYRES.HE","UPM.HE","VALMT.HE","WRT1V.HE"]
    # OMX Copenhagen - Danemark .CO (PEA)
    + ["AMBU-B.CO","BAVA.CO","CARL-B.CO","CHR.CO","COLO-B.CO","DANSKE.CO",
       "DEMANT.CO","DSV.CO","FLS.CO","GMAB.CO","GN.CO","ISS.CO","JYSK.CO",
       "MAERSK-B.CO","NDA-DK.CO","NETC.CO","NOVO-B.CO","NZYM-B.CO","ORSTED.CO",
       "PNDORA.CO","RBREW.CO","ROCK-B.CO","TRYG.CO","VWS.CO"]
    # Oslo Børs - Norvège .OL (NON-PEA, Norvège hors EEE pour PEA)
    + ["AKERBP.OL","BAKKA.OL","DNB.OL","EQNR.OL","FRO.OL","GJF.OL","MOWI.OL",
       "NHY.OL","ORK.OL","SALM.OL","SCATC.OL","SUBC.OL","TEL.OL","TGS.OL",
       "TOM.OL","YAR.OL"]
    # ATX - Autriche .VI (PEA)
    + ["ANDR.VI","BAWAG.VI","EBS.VI","IIA.VI","LNZ.VI","OMV.VI","POST.VI","RBI.VI",
       "SBO.VI","STR.VI","TKA.VI","UQA.VI","VER.VI","VIG.VI","VOE.VI","WIE.VI"]
    # SMI - Suisse .SW (NON-PEA, Suisse hors EEE)
    + ["ABBN.SW","ALC.SW","GEBN.SW","GIVN.SW","HOLN.SW","KNIN.SW","LOGN.SW","LONN.SW",
       "NESN.SW","NOVN.SW","PGHN.SW","ROG.SW","SCMN.SW","SGSN.SW","SIKA.SW","SLHN.SW",
       "SOON.SW","SREN.SW","UBSG.SW","ZURN.SW"]
    # ISEQ - Irlande .IR (PEA)
    + ["BIRG.IR","CRH.IR","FBD.IR","GLB.IR","GRP.IR","HBRN.IR","KMR.IR","KRZ.IR",
       "OIZ.IR","RYA.IR","SK3.IR"]
    # PSI 20 - Portugal .LS (PEA)
    + ["ALTR.LS","BCP.LS","COR.LS","CTT.LS","EDP.LS","EDPR.LS","GALP.LS","IBS.LS",
       "JMT.LS","MOTA.LS","NOS.LS","NVG.LS","REN.LS","RAM.LS","SEM.LS","SON.LS"]
    # WIG 20 - Pologne .WA (PEA, marché EEE)
    + ["ALE.WA","ALR.WA","BDX.WA","CDR.WA","CPS.WA","DNP.WA","JSW.WA","KGH.WA",
       "KRU.WA","KTY.WA","LPP.WA","MBK.WA","OPL.WA","PCO.WA","PEO.WA","PGE.WA",
       "PKN.WA","PKO.WA","PZU.WA","SPL.WA"]
    # ASE - Grèce .AT (PEA, marché EEE)
    + ["AEGN.AT","ALPHA.AT","ARAIG.AT","BELA.AT","CENER.AT","ELPE.AT","ETE.AT",
       "EUROB.AT","EXAE.AT","GEKTERNA.AT","HTO.AT","JUMBO.AT","LAMDA.AT",
       "METLEN.AT","MOH.AT","MYTIL.AT","ALWN.AT","OTE.AT","PPC.AT","TPEIR.AT",
       "SARANTI.AT","TENERGY.AT","VIO.AT"]
)
STOXX = sorted(set(_STOXX_NATIONAL))


# ── Nikkei 225 - Japon .T (top ~100 actions les plus liquides) ───────
NIKKEI = [
    "7203.T","6758.T","9984.T","6861.T","8035.T","7974.T","6098.T","8306.T",
    "9432.T","6501.T","4063.T","4543.T","6981.T","6594.T","6857.T","6902.T",
    "6367.T","6273.T","6326.T","4502.T","4503.T","4519.T","4523.T","4568.T",
    "4661.T","4901.T","6201.T","6301.T","6302.T","6305.T","6471.T","6503.T",
    "6504.T","6506.T","6508.T","6701.T","6702.T","6724.T","6752.T","6753.T",
    "6762.T","6770.T","6841.T","6920.T","6954.T","7011.T","7012.T","7013.T",
    "7201.T","7211.T","7261.T","7267.T","7269.T","7270.T","7272.T","7733.T",
    "7735.T","7741.T","7751.T","7752.T","7832.T","8001.T","8002.T","8015.T",
    "8031.T","8053.T","8058.T","8264.T","8267.T","8308.T","8309.T","8316.T",
    "8411.T","8591.T","8601.T","8604.T","8630.T","8725.T","8750.T","8766.T",
    "8795.T","8801.T","8802.T","8830.T","9001.T","9005.T","9007.T","9009.T",
    "9020.T","9021.T","9101.T","9104.T","9202.T","9301.T","9433.T","9501.T",
    "9502.T","9503.T","9531.T","9602.T","9613.T",
]

# ── TSX 60 - Canada .TO ──────────────────────────────────────────────
TSX60 = [
    "RY.TO","TD.TO","BNS.TO","BMO.TO","CM.TO","NA.TO","CNR.TO","CP.TO","ENB.TO",
    "TRP.TO","SU.TO","CNQ.TO","CVE.TO","IMO.TO","MFC.TO","SLF.TO","GWO.TO",
    "BCE.TO","T.TO","RCI-B.TO","SHOP.TO","ATD.TO","MG.TO","GIB-A.TO","OTEX.TO",
    "CSU.TO","TRI.TO","ABX.TO","AEM.TO","K.TO","WPM.TO","FNV.TO","FFH.TO",
    "BAM.TO","GIL.TO","NTR.TO","POW.TO","IFC.TO","FTS.TO","EMA.TO","H.TO",
    "AQN.TO","CCO.TO","NXE.TO","TECK-B.TO","FM.TO","IVN.TO","L.TO","EMP-A.TO",
    "DOL.TO","QSR.TO","RBA.TO","WCN.TO","WSP.TO","TFII.TO","CAE.TO","BIP-UN.TO",
    "BEP-UN.TO",
]

# ── ASX 50 - Australie .AX ───────────────────────────────────────────
ASX50 = [
    "BHP.AX","CSL.AX","CBA.AX","NAB.AX","ANZ.AX","WBC.AX","MQG.AX","WES.AX",
    "WOW.AX","COL.AX","RIO.AX","TLS.AX","GMG.AX","FMG.AX","TCL.AX","STO.AX",
    "ALL.AX","REA.AX","COH.AX","BXB.AX","ASX.AX","SUN.AX","QBE.AX","IAG.AX",
    "S32.AX","JBH.AX","MIN.AX","NEM.AX","ORG.AX","RMD.AX","PLS.AX","EVN.AX",
    "JHX.AX","AGL.AX","ORI.AX","AMP.AX","TPG.AX","NWS.AX","MFG.AX","TWE.AX",
    "BSL.AX","ALD.AX","AZJ.AX","NXT.AX","A2M.AX","SCG.AX","IGO.AX","SOL.AX",
    "LLC.AX","DXS.AX",
]

# ── Hang Seng 50 - Hong Kong .HK ─────────────────────────────────────
HSI = [
    "0700.HK","0941.HK","1299.HK","0939.HK","0005.HK","0388.HK","0883.HK",
    "0001.HK","0016.HK","0011.HK","0027.HK","0066.HK","0101.HK","0175.HK",
    "0267.HK","0288.HK","0291.HK","0386.HK","0688.HK","0762.HK","0823.HK",
    "0857.HK","0960.HK","0992.HK","1038.HK","1044.HK","1093.HK","1109.HK",
    "1113.HK","1177.HK","1211.HK","1378.HK","1810.HK","1928.HK","1972.HK",
    "2018.HK","2020.HK","2269.HK","2313.HK","2318.HK","2319.HK","2331.HK",
    "2382.HK","2388.HK","2628.HK","3328.HK","3690.HK","3988.HK","9618.HK",
    "9888.HK","9988.HK","9999.HK",
]


# ═════════════════════════════════════════════════════════════════════
#  5. PEA, SECTEURS, DRAPEAUX, INDICES
# ═════════════════════════════════════════════════════════════════════

# Suffixes Yahoo Finance considérés comme PEA-éligibles (EEE + Royaume-Uni hors PEA + Suisse hors PEA)
PEA_SUFFIXES = {".PA", ".DE", ".AS", ".BR", ".MI", ".MC", ".LS", ".HE",
                ".OL", ".ST", ".CO", ".VI", ".IR", ".AT", ".WA", ".PR"}

# Cas particuliers (sociétés cotées Amsterdam mais hors PEA pour raison structurelle)
PEA_OVERRIDES: dict[str, bool] = {
    "MT.AS":  True,    # ArcelorMittal (Luxembourg, mais EEE)
    "URW.AS": True,    # Unibail-Rodamco-Westfield
}

def is_pea(ticker: str) -> bool:
    """Détermine si un ticker est éligible au PEA (Plan d'Épargne en Actions)."""
    if ticker in PEA_OVERRIDES:
        return PEA_OVERRIDES[ticker]
    if "." not in ticker:
        return False  # Pas de suffixe → US → CTO
    return "." + ticker.rsplit(".", 1)[1] in PEA_SUFFIXES

# Mapping secteur Yahoo (anglais) → label FR officiel GICS
SECTOR_FR = {
    "Technology":             "Technologies de l'information",
    "Communication Services": "Services de communication",
    "Consumer Cyclical":      "Consommation discrétionnaire",
    "Consumer Defensive":     "Consommation de base",
    "Energy":                 "Énergie",
    "Financial Services":     "Services financiers",
    "Healthcare":             "Santé",
    "Industrials":            "Industrie",
    "Basic Materials":        "Matériaux",
    "Real Estate":            "Immobilier",
    "Utilities":              "Services aux collectivités",
}

# Pour affichage compact (vidéo + post) : (label_court, emoji)
SECTOR_DISPLAY: dict[str, tuple[str, str]] = {
    "Technologies de l'information": ("Tech. info.",        "💻"),
    "Services de communication":     ("Communication",      "📡"),
    "Consommation discrétionnaire":  ("Conso. discrét.",    "🛍️"),
    "Consommation de base":          ("Conso. de base",     "🛒"),
    "Énergie":                       ("Énergie",            "⚡"),
    "Services financiers":           ("Finance",            "💳"),
    "Santé":                         ("Santé",              "🏥"),
    "Industrie":                     ("Industrie",          "🏭"),
    "Matériaux":                     ("Matériaux",          "⛏️"),
    "Immobilier":                    ("Immobilier",         "🏢"),
    "Services aux collectivités":    ("Services coll.",     "💡"),
}

def get_sector_display(sector_fr: str) -> tuple[str, str]:
    return SECTOR_DISPLAY.get(sector_fr, (sector_fr, "📌"))

# Drapeau par suffixe de marché
FLAG = {
    ".PA":"🇫🇷", ".DE":"🇩🇪", ".AS":"🇳🇱", ".BR":"🇧🇪", ".MI":"🇮🇹",
    ".MC":"🇪🇸", ".LS":"🇵🇹", ".OL":"🇳🇴", ".ST":"🇸🇪", ".HE":"🇫🇮",
    ".CO":"🇩🇰", ".VI":"🇦🇹", ".IR":"🇮🇪", ".SW":"🇨🇭", ".L":"🇬🇧",
    ".WA":"🇵🇱", ".AT":"🇬🇷",
    ".T":"🇯🇵",  ".TO":"🇨🇦", ".AX":"🇦🇺", ".HK":"🇭🇰",
}

def get_flag(ticker: str) -> str:
    if "." not in ticker:
        return "🇺🇸"
    return FLAG.get("." + ticker.rsplit(".", 1)[1], "🌍")


# ── Pills d'indices affichées dans la cover (1 ligne pleine largeur) ─
# Limité à ~10 pills max pour tenir sur 1 ligne à 1080px. Le reste va dans "+N AUTRES".
INDEX_PILLS_ALL = [
    "SP 500",            # 🇺🇸 ~$50T
    "NASDAQ 100",        # 🇺🇸 ~$25T
    "STOXX 600",         # 🇪🇺 ~€14T
    "SP 400 MID",        # 🇺🇸 ~$3.5T
    "SP 600 SMALL",      # 🇺🇸 ~$1.2T
    "FTSE 100",          # 🇬🇧 ~$2.5T
    "SBF 120",           # 🇫🇷 ~$2T
    "CAC 40",            # 🇫🇷 ~$2T
    "DAX 40",            # 🇩🇪 ~€1.9T
    "AEX 25",            # 🇳🇱 ~$1T
    "IBEX 35",           # 🇪🇸 ~$800B
    "FTSE MIB 40",       # 🇮🇹 ~$700B
    "FTSE 250",          # 🇬🇧 ~$500B
    "STOXX MID 200",     # 🇪🇺 ~€450B
    "MDAX 50",           # 🇩🇪 ~€430B
    # 🌍 NON-Bourso (tri par market cap DESC)
    "NIKKEI 225",        # 🇯🇵 ~$5T
    "HANG SENG",         # 🇭🇰 ~$2.5T
    "TSX 60",            # 🇨🇦 ~$1.8T
    "ASX 50",            # 🇦🇺 ~$1.5T
    "ASE",               # 🇬🇷 ~$100B
    "WIG 20",            # 🇵🇱 ~$100B
]
INDEX_PILLS_VISIBLE_MAX = 9


# ── Benchmarks (3 indices fetched en PREMIER pour éviter rate-limit) ─
BENCHMARKS = [
    {"ticker": "^GSPC",  "label": "S&P 500",   "flag": "🇺🇸"},
    {"ticker": "^STOXX", "label": "STOXX 600", "flag": "🇪🇺"},
    {"ticker": "^FCHI",  "label": "CAC 40",    "flag": "🇫🇷"},
]


# ═════════════════════════════════════════════════════════════════════
#  6. URL BUILDERS - Boursorama (prio) + Yahoo Finance (toujours dispo)
# ═════════════════════════════════════════════════════════════════════

# Préfixes Boursorama par marché (pour construire l'URL canonique)
BOURSO_PREFIX = {
    ".PA": "1rP",  ".AS": "1rA",  ".BR": "FF11-", ".LS": "1rL",
    ".MI": "1g",   ".MC": "FF55-",".DE": "1z",    ".SW": "2a",
}

# Exchanges Google Finance par suffixe (pour fallback)
GF_EXCHANGE = {
    ".PA": "EPA",  ".AS": "AMS",  ".BR": "EBR",  ".LS": "ELI",
    ".IR": "DUB",  ".MI": "BIT",  ".MC": "BME",  ".DE": "ETR",
    ".SW": "SWX",  ".L":  "LON",  ".ST": "STO",  ".HE": "HEL",
    ".CO": "CPH",  ".OL": "OSL",  ".VI": "VIE",  ".WA": "WSE",
    ".AT": "ATH",
    ".T":  "TYO",  ".TO": "TSE",  ".AX": "ASX",  ".HK": "HKG",
}

US_EX_MAP = {
    "NMS": "NASDAQ", "NCM": "NASDAQ", "NGM": "NASDAQ",
    "NYQ": "NYSE",   "ASE": "NYSEAMERICAN", "PCX": "NYSEARCA",
}

def boursorama_url(ticker: str) -> str | None:
    """Construit l'URL Boursorama pour un ticker (None si non couvert)."""
    if "." not in ticker:
        # US : Boursorama propose une page directe via le ticker
        return f"https://www.boursorama.com/cours/{ticker.replace('-', '.')}/"
    base, suf = ticker.rsplit(".", 1)
    suffix = "." + suf
    if suffix == ".L":
        return f"https://www.boursorama.com/cours/1u{base}.L/"
    if suffix in BOURSO_PREFIX:
        return f"https://www.boursorama.com/cours/{BOURSO_PREFIX[suffix]}{base}/"
    return None  # Marchés non couverts par Boursorama (Tokyo, Sydney, HK, etc.)

def yahoo_url(ticker: str) -> str:
    """URL Yahoo Finance - fonctionne pour TOUS les tickers."""
    return f"https://finance.yahoo.com/quote/{ticker}/"

def google_finance_url(ticker: str, yf_exchange: str | None = None) -> str | None:
    if "." not in ticker:
        gf_ex = US_EX_MAP.get(yf_exchange or "", "NYSE")
        return f"https://www.google.com/finance/quote/{ticker.replace('-', '.')}:{gf_ex}"
    base, suf = ticker.rsplit(".", 1)
    suffix = "." + suf
    if suffix in GF_EXCHANGE:
        return f"https://www.google.com/finance/quote/{base}:{GF_EXCHANGE[suffix]}"
    return None


# ═════════════════════════════════════════════════════════════════════
#  7. FETCHER YFINANCE (threadé · retry · sanity check · rate limit safe)
# ═════════════════════════════════════════════════════════════════════

# Labels FR pour les recommandations analystes
RECO_LABEL_FR = {
    "strong_buy":  "Achat fort",
    "buy":         "Achat",
    "hold":        "Conserver",
    "sell":        "Vendre",
    "strong_sell": "Vente forte",
    "underperform":"Sous-performance",
}

def fetch_one(ticker: str, market: str,
              fx_rates: dict[str, float] | None = None,
              max_retries: int = 3) -> dict[str, Any] | None:
    """
    Fetch un ticker depuis yfinance + sanity checks.
    Drop si data incomplète ou outlier (split mal géré, etc.).
    
    Args:
        ticker     : symbole Yahoo (ex: AAPL, SAN.PA, BARC.L)
        market     : libellé du marché (pour logs/stats)
        fx_rates   : taux de change { devise: rate_vers_eur }
                     Si None → on n'expose pas price_eur (mais price natif OK).
        max_retries: tentatives avec backoff
    
    Returns: dict avec toutes les métriques OU None si data inutilisable.
             Contient notamment :
               - price        : prix natif (devise locale)
               - currency     : code ISO (USD, EUR, GBp, JPY, ...)
               - price_eur    : prix converti en EUR (None si fx_rates absent)
    """
    for attempt in range(max_retries):
        try:
            t = yf.Ticker(ticker)
            info = t.info

            price = (info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"))
            target = info.get("targetMedianPrice") or info.get("targetMeanPrice")
            sector   = info.get("sector")
            name     = info.get("longName") or info.get("shortName") or ticker
            isin     = info.get("isin")
            exchange = info.get("exchange")
            currency = info.get("currency") or ""    # ex: "USD", "EUR", "GBp"

            # ── Sanity checks de base ────────────────────────────────
            if not price or not target or price <= 0:
                return None
            if sector not in SECTOR_FR:
                return None

            # ── Calculs métriques ────────────────────────────────────
            div_rate   = info.get("dividendRate") or 0
            div_pct    = (div_rate / price) * 100 if div_rate else 0
            target_pct = (target - price) / price * 100

            # Exclude outliers extrêmes (target > +200% ou < -90% = data bug)
            if target_pct > 200 or target_pct < -90:
                return None

            target_high       = info.get("targetHighPrice")
            target_low        = info.get("targetLowPrice")
            target_high_pct   = (target_high - price) / price * 100 if target_high else None
            target_low_pct    = (target_low  - price) / price * 100 if target_low  else None
            target_spread_pct = (
                (target_high - target_low) / price * 100
                if (target_high and target_low) else None
            )

            reco_key   = info.get("recommendationKey")
            reco_mean  = info.get("recommendationMean")
            n_analysts = (info.get("numberOfAnalystOpinions")
                          or info.get("numberOfAnalysts") or 0)
            reco_label = RECO_LABEL_FR.get(reco_key, reco_key or "-")

            # ── Perf mois précédent (calendaire complet) ─────────────
            perf_1m = None
            try:
                today      = datetime.now().date()
                start_prev = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
                end_prev   = today.replace(day=1)
                hist = t.history(start=start_prev, end=end_prev)
                if len(hist) >= 2:
                    perf_1m = (hist["Close"].iloc[-1] / hist["Close"].iloc[0] - 1) * 100
                    # Sanity check : exclure outliers extrêmes (split mal géré)
                    if perf_1m is not None and (perf_1m > 100 or perf_1m < -80):
                        log.warning("⚠️  %s : perf_1m suspecte (%+.1f%%) → exclu",
                                    ticker, perf_1m)
                        perf_1m = None
            except Exception:
                pass

            # ── FILTRE BOURSORAMA STRICT (si activé) ─────────────────
            # Skip les tickers non couverts par Boursorama (Japon, HK, Canada, ASX, Norvège, etc.)
            # → garantit que tous les tickers du brief auront un lien BR fonctionnel
            bourso_link = boursorama_url(ticker)
            if FILTER_BOURSO_ONLY and not bourso_link:
                return None

            # ── Conversion EUR (si fx_rates fournis) ─────────────────
            price_eur = to_eur(price, currency, fx_rates) if fx_rates else None

            return {
                "ticker": ticker,
                "name": name,
                "sector": sector,
                "sector_fr": SECTOR_FR[sector],
                "market": market,
                "price": round(price, 2),
                "currency": currency,
                "price_eur": price_eur,
                "div_pct": round(div_pct, 2),
                "target_pct": round(target_pct, 2),
                "target_high_pct":   round(target_high_pct, 2)   if target_high_pct   is not None else None,
                "target_low_pct":    round(target_low_pct, 2)    if target_low_pct    is not None else None,
                "target_spread_pct": round(target_spread_pct, 2) if target_spread_pct is not None else None,
                "reco_label": reco_label,
                "reco_mean": round(reco_mean, 2) if reco_mean else None,
                "analyst_count": int(n_analysts) if n_analysts else 0,
                "total_pct": round(target_pct + div_pct, 2),
                "perf_1m": round(perf_1m, 2) if perf_1m is not None else None,
                "pea": is_pea(ticker),
                "isin": isin or "",
                "boursorama_link": bourso_link,
                "yahoo_link":      yahoo_url(ticker),
                "google_link":     google_finance_url(ticker, exchange),
            }
        except Exception:
            if attempt == max_retries - 1:
                return None
            time.sleep(1 + attempt * 2)  # Backoff exponentiel léger
    return None


def fetch_universe(tickers: list[str], market: str,
                   fx_rates: dict[str, float] | None = None) -> list[dict[str, Any]]:
    """Fetch parallèle d'un univers de tickers. N_WORKERS threads."""
    rows: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=N_WORKERS) as pool:
        futs = {pool.submit(fetch_one, t, market, fx_rates): t for t in tickers}
        for fut in tqdm(as_completed(futs), total=len(futs),
                        desc=f"  {market:<6}", ncols=70, leave=False):
            r = fut.result()
            if r:
                rows.append(r)
    return rows


def fetch_benchmarks() -> list[dict[str, Any]]:
    """
    Fetch perf 1 mois (mois calendaire précédent complet) pour les indices benchmarks.
    
    ⚠️  CRITIQUE : à appeler EN PREMIER, avant les universes,
    pour ne pas se faire rate-limit par Yahoo en fin de course.
    """
    results: list[dict[str, Any]] = []
    today      = datetime.now().date()
    start_prev = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    end_prev   = today.replace(day=1)
    for bm in BENCHMARKS:
        try:
            t = yf.Ticker(bm["ticker"])
            hist = t.history(start=start_prev, end=end_prev)
            perf = None
            if len(hist) >= 2:
                perf = (hist["Close"].iloc[-1] / hist["Close"].iloc[0] - 1) * 100
            results.append({**bm, "perf_1m": round(perf, 2) if perf is not None else None})
            time.sleep(2)  # Petit délai entre chaque benchmark (politesse Yahoo)
        except Exception as e:
            log.warning("  ⚠️  Benchmark %s : %s", bm["ticker"], e)
            results.append({**bm, "perf_1m": None})
    return results


# ═════════════════════════════════════════════════════════════════════
#  7bis. FX RATES - Conversion des prix natifs vers EUR
# ═════════════════════════════════════════════════════════════════════
# Toutes les devises rencontrées dans l'univers Bourso-compatible.
# La paire `EURXXX=X` Yahoo donne "combien d'XXX pour 1 EUR".
# Donc pour convertir XXX → EUR, on fait : price_eur = price * (1 / rate).
# Cas particulier : .L (London) cote en GBp (pence), pas en GBP (livres).
# 1 GBp = 1/100 GBP, donc rate_GBp = rate_GBP / 100.

FX_PAIRS = {
    "USD": "EURUSD=X",   # US (.US implicite = pas de suffixe)
    "GBP": "EURGBP=X",   # rare sur yfinance (.L = GBp en général)
    "GBp": "EURGBP=X",   # London → pence
    "JPY": "EURJPY=X",   # Tokyo .T
    "CAD": "EURCAD=X",   # Toronto .TO
    "CHF": "EURCHF=X",   # Suisse .SW
    "AUD": "EURAUD=X",   # Sydney .AX
    "HKD": "EURHKD=X",   # Hong Kong .HK
    "SEK": "EURSEK=X",   # Stockholm .ST
    "DKK": "EURDKK=X",   # Copenhagen .CO
    "NOK": "EURNOK=X",   # Oslo .OL
    "PLN": "EURPLN=X",   # Warsaw .WA
}

# Fallback hardcodé si yfinance fail (taux approximatifs mai 2026).
# À actualiser manuellement de temps en temps, mais sert juste de filet de sécu.
FX_FALLBACK: dict[str, float] = {
    "EUR": 1.0,
    "USD": 0.92,
    "GBP": 1.18,
    "GBp": 0.0118,
    "JPY": 0.0061,
    "CAD": 0.68,
    "CHF": 1.04,
    "AUD": 0.61,
    "HKD": 0.118,
    "SEK": 0.087,
    "DKK": 0.134,
    "NOK": 0.087,
    "PLN": 0.235,
}


def fetch_fx_rates() -> dict[str, float]:
    """
    Récupère les taux de change yfinance pour convertir chaque devise vers EUR.
    
    Returns: dict {code_devise_ISO: taux_pour_convertir_en_EUR}
             Ex: {"EUR": 1.0, "USD": 0.92, "GBp": 0.0118, "JPY": 0.0061, ...}
    
    Stratégie :
      1. EUR = 1.0 (référence)
      2. Pour chaque devise : fetch EURXXX=X (= XXX par 1 EUR), on inverse → 1/rate
      3. Cas .L : GBp = GBP/100 → on divise encore par 100
      4. Si fetch fail → fallback hardcodé (FX_FALLBACK)
    """
    rates: dict[str, float] = {"EUR": 1.0}
    log.info("\n💱  Fetch FX rates (devises → EUR)…")
    for ccy, pair in FX_PAIRS.items():
        try:
            t = yf.Ticker(pair)
            info = t.info
            # Plusieurs champs possibles selon yfinance version
            rate_eur_to_x = (info.get("regularMarketPrice")
                             or info.get("previousClose")
                             or info.get("bid")
                             or info.get("ask"))
            if rate_eur_to_x and rate_eur_to_x > 0:
                if ccy == "GBp":
                    # London cote en pence (1 GBp = GBP/100)
                    rates[ccy] = 1.0 / (rate_eur_to_x * 100.0)
                else:
                    rates[ccy] = 1.0 / rate_eur_to_x
                log.info("   %s = %.6f EUR  (%s = %.4f)",
                         ccy, rates[ccy], pair, rate_eur_to_x)
            else:
                rates[ccy] = FX_FALLBACK.get(ccy, 1.0)
                log.warning("   ⚠️  %s : taux indispo → fallback %.6f",
                            ccy, rates[ccy])
            time.sleep(0.5)  # Politesse Yahoo
        except Exception as e:
            rates[ccy] = FX_FALLBACK.get(ccy, 1.0)
            log.warning("   ⚠️  %s : %s → fallback %.6f", ccy, e, rates[ccy])
    log.info("✅  %d devises chargées", len(rates))
    return rates


def to_eur(price: float | None, currency: str | None,
           fx_rates: dict[str, float]) -> float | None:
    """
    Convertit un prix natif vers EUR. None si impossible.
    """
    if price is None or pd.isna(price):
        return None
    if not currency:
        return None
    rate = fx_rates.get(currency, FX_FALLBACK.get(currency))
    if rate is None:
        return None
    return round(float(price) * rate, 2)


# ═════════════════════════════════════════════════════════════════════
#  8. UTILITIES TEXTE / FORMATAGE
# ═════════════════════════════════════════════════════════════════════

MOIS_FR = {
    "January":"Janvier","February":"Février","March":"Mars","April":"Avril",
    "May":"Mai","June":"Juin","July":"Juillet","August":"Août",
    "September":"Septembre","October":"Octobre","November":"Novembre",
    "December":"Décembre",
}

def to_fr_period(period_str: str) -> str:
    """'MAY 2026' → 'MAI 2026'"""
    parts = period_str.strip().split()
    if not parts:
        return period_str
    mois_traduit = MOIS_FR.get(parts[0].capitalize(), parts[0])
    return " ".join([mois_traduit.upper()] + parts[1:])

def to_fr_month_year(yyyy_mm: str) -> str:
    """'2026-04' → 'Avril 2026'"""
    try:
        y, m = yyyy_mm.split("-")
        mois_en = datetime(int(y), int(m), 1).strftime("%B")
        return f"{MOIS_FR.get(mois_en, mois_en)} {y}"
    except Exception:
        return yyyy_mm

def _next_month_fr(period_fr: str) -> str:
    """'MAI 2026' → 'juin' (mois suivant en minuscules pour la phrase teasing)."""
    parts = period_fr.strip().lower().split()
    mois_to_next = {
        "janvier": "février", "février": "mars", "mars": "avril",
        "avril": "mai", "mai": "juin", "juin": "juillet",
        "juillet": "août", "août": "septembre", "septembre": "octobre",
        "octobre": "novembre", "novembre": "décembre", "décembre": "janvier",
    }
    if not parts:
        return "mois prochain"
    return mois_to_next.get(parts[0], "mois prochain")

def cap_name(name: str) -> str:
    """Force la 1ère lettre en majuscule (sans toucher au reste)."""
    if not name:
        return ""
    return name[0].upper() + name[1:]

def smart_trunc(s: str, n: int = 22) -> str:
    """Tronque sur le dernier espace avant n chars, ajoute …"""
    if not s:
        return ""
    s = str(s).strip()
    if len(s) <= n:
        return s
    cut = s[:n].rsplit(" ", 1)[0]
    return (cut if cut else s[:n]) + "…"

def safe_ticker(t: str) -> str:
    """
    Casse l'auto-linkify LinkedIn (.DE, .ST, .BR, .AT, etc.).
    Insère un Zero-Width Space (U+200B) après le point :
    - Visuellement IDENTIQUE à l'œil nu
    - LinkedIn ne reconnaît plus le pattern comme un TLD → pas de lien fantôme
    """
    return t.replace(".", ".\u200B")

def clean_reco(label: Any) -> str:
    if not label or str(label).lower() in ("none", "nan", "-", "", "-"):
        return ""
    return str(label)

def fmt_signed_pct(v: Any) -> str:
    if v is None or pd.isna(v): return "-"
    return f"{v:+.1f}%".replace(".", ",")

def fmt_price(price_eur: Any) -> str:
    """
    Formate un prix EUR pour affichage. Ex: '148.30€'.
    Retourne '-' si valeur invalide / manquante / None / NaN.
    """
    if price_eur is None:
        return "-"
    try:
        if pd.isna(price_eur):
            return "-"
    except (TypeError, ValueError):
        return "-"
    try:
        return f"{float(price_eur):.2f}€".replace(".", ",")
    except (TypeError, ValueError):
        return "-"

def perf_class(v: Any) -> str:
    """Classe CSS pour la couleur (positive/négative/neutre)."""
    if v is None or pd.isna(v): return "neut"
    if v > 0: return "pos"
    if v < 0: return "neg"
    return "neut"

def reco_color_class(reco_mean: Any) -> str:
    """Classe CSS pour la couleur des étoiles selon recommandation moyenne (1=strong buy, 5=strong sell)."""
    if reco_mean is None or pd.isna(reco_mean): return "reco-na"
    if reco_mean <= 1.8: return "reco-strong"
    if reco_mean <= 2.5: return "reco-buy"
    if reco_mean <= 3.5: return "reco-hold"
    if reco_mean <= 4.2: return "reco-sell"
    return "reco-vsell"

def reco_stars(reco_mean: Any) -> str:
    """1.0 = ★★★★★ (Strong Buy), 5.0 = ☆☆☆☆☆ (Strong Sell)"""
    if reco_mean is None or pd.isna(reco_mean): return "-"
    score = max(0, min(5, 6 - reco_mean))
    full = int(round(score))
    return "★" * full + "☆" * (5 - full)


def _safe_url(v: Any) -> str | None:
    """
    Sanitize une URL venant du DataFrame.
    
    ⚠️  Quand boursorama_url() retourne None et que ça passe par pandas → NaN.
    `bool(NaN) == True` donc `if url` ne suffit pas. Cette helper unifie le check :
    None, NaN, "", "nan" → None.
    """
    if v is None: return None
    try:
        if isinstance(v, float) and pd.isna(v): return None
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    if not s.startswith("http"):
        return None
    return s


# ═════════════════════════════════════════════════════════════════════
#  9. SNAPSHOTS (sauvegarde mensuelle pour archivage / future diff)
# ═════════════════════════════════════════════════════════════════════

def _easter_date(year: int) -> date:
    """
    Calcule la date de Pâques (dimanche) pour une année donnée.
    Algorithme de Gauss/Meeus (grégorien). Aucune dépendance externe.
    """
    a = year % 19
    b = year // 100
    c = year % 100
    d_ = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d_ - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def french_holidays(year: int) -> set[date]:
    """
    Retourne l'ensemble des jours fériés français pour une année donnée
    (métropole). Inclut les fixes et les mobiles (Pâques, Ascension,
    Pentecôte).
    """
    easter = _easter_date(year)
    return {
        date(year, 1, 1),               # Jour de l'an
        easter + timedelta(days=1),     # Lundi de Pâques
        date(year, 5, 1),               # Fête du Travail
        date(year, 5, 8),               # Victoire 1945
        easter + timedelta(days=39),    # Ascension (jeudi)
        easter + timedelta(days=50),    # Lundi de Pentecôte
        date(year, 7, 14),              # Fête Nationale
        date(year, 8, 15),              # Assomption
        date(year, 11, 1),              # Toussaint
        date(year, 11, 11),             # Armistice
        date(year, 12, 25),             # Noël
    }


def is_french_holiday(d: date) -> bool:
    """True si `d` est un jour férié français (métropole)."""
    return d in french_holidays(d.year)


def is_first_business_day_of_month(today: date | None = None) -> bool:
    """
    True si aujourd'hui est le premier jour ouvré "publiable" du mois.
    Logique : on part du 1er du mois, on avance tant que le jour est :
      - un weekend (samedi/dimanche)
      - un férié français (métropole)
      - un JEUDI (créneau déjà occupé par d'autres posts LinkedIn récurrents)
    Le jour cible doit matcher `today`.
    """
    if today is None:
        today = date.today()
    d = date(today.year, today.month, 1)
    # 3 = jeudi (skippé : autre post hebdo), 5 = samedi, 6 = dimanche
    EXCLUDED_WEEKDAYS = {3, 5, 6}
    while d.weekday() in EXCLUDED_WEEKDAYS or is_french_holiday(d):
        d += timedelta(days=1)
    return today == d


def sleep_until_paris(hour: int = 9, minute: int = 0) -> None:
    """
    Bloque l'exécution jusqu'à ce que l'heure locale Paris atteigne
    `hour:minute` du jour courant. Gère DST automatiquement via zoneinfo.

    Si l'heure cible est déjà passée → no-op (log warning).
    Bypass en TEST_MODE / FORCE_RUN pour ne pas bloquer les tests.
    """
    if TEST_MODE or _env_bool("FORCE_RUN", default=False):
        log.info("⏭️  sleep_until_paris bypass (TEST_MODE / FORCE_RUN)")
        return

    now_paris = datetime.now(PARIS_TZ)
    target = datetime.combine(
        now_paris.date(), dt_time(hour=hour, minute=minute), tzinfo=PARIS_TZ
    )
    wait_s = (target - now_paris).total_seconds()

    if wait_s <= 0:
        log.warning(
            "⚠️  Heure cible %02d:%02d Paris déjà passée (now=%s) — publication immédiate",
            hour, minute, now_paris.strftime("%H:%M:%S"),
        )
        return

    log.info(
        "⏰  Attente jusqu'à %02d:%02d Paris (now=%s, sleep=%.0f min)",
        hour, minute, now_paris.strftime("%H:%M:%S"), wait_s / 60,
    )
    time.sleep(wait_s)
    log.info("✅  Heure cible atteinte → envoi webhook")


def rotate_snapshots(max_keep: int = MAX_SNAPSHOTS) -> None:
    """Garde seulement les N snapshots les plus récents. Supprime les plus vieux."""
    if not SNAPSHOT_DIR.exists():
        return
    files = sorted(SNAPSHOT_DIR.glob("ranking_*.json"),
                   key=lambda p: p.name, reverse=True)
    to_keep = files[:max_keep]
    to_delete = files[max_keep:]
    for f in to_delete:
        try:
            f.unlink()
            log.info("🗑  Snapshot rotation : supprimé %s", f.name)
        except Exception as e:
            log.warning("  ⚠️  Impossible de supprimer %s : %s", f.name, e)
    if to_keep:
        log.info("📚  Snapshots conservés : %d (max %d)", len(to_keep), max_keep)


def save_snapshot(df: pd.DataFrame, suffix: str) -> Path:
    """Sauvegarde le DataFrame en JSON dans snapshots/."""
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    month_key = datetime.now().strftime("%Y-%m")
    path = SNAPSHOT_DIR / f"ranking_{month_key}{suffix}.json"
    df.to_json(path, orient="records")
    log.info("📸  snapshot sauvegardé → %s", path)
    rotate_snapshots()
    return path

# NB : load_prev_ranks et diff_tag sont conservés pour usage futur
# mais NON appelés dans v10 (Romain en phase de correction, pas de diff affichée).

def load_prev_ranks(suffix: str) -> tuple[dict, dict, bool]:
    """Charge le snapshot du mois précédent pour calculer la diff (non utilisé en v10)."""
    prev_month = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
    prev_path = SNAPSHOT_DIR / f"ranking_{prev_month}{suffix}.json"
    if not prev_path.exists():
        return {}, {}, False
    prev = pd.read_json(prev_path, orient="records")
    pea = (prev[prev["pea"] == True]
           .sort_values("total_pct", ascending=False).reset_index(drop=True))
    cto = (prev[prev["pea"] == False]
           .sort_values("total_pct", ascending=False).reset_index(drop=True))
    pea_ranks = {row["ticker"]: i + 1 for i, row in pea.head(N_TOP).iterrows()}
    cto_ranks = {row["ticker"]: i + 1 for i, row in cto.head(N_TOP).iterrows()}
    return pea_ranks, cto_ranks, True

def diff_tag(ticker: str, prev_ranks: dict, cur_rank: int) -> str:
    """Tag d'évolution rank vs mois précédent (non affiché en v10)."""
    if not prev_ranks:           return ""
    if ticker not in prev_ranks: return "  🆕 Nouvelle entrée"
    delta = prev_ranks[ticker] - cur_rank
    if delta > 0: return f"  🔼 +{delta} place{'s' if delta > 1 else ''} vs mois dernier"
    if delta < 0: return f"  🔽 {abs(delta)} place{'s' if abs(delta) > 1 else ''} vs mois dernier"
    return "  ↔️  Stable"


# ═════════════════════════════════════════════════════════════════════
#  10. PIPELINE PRINCIPALE - fetch benchmarks + universes + score
# ═════════════════════════════════════════════════════════════════════

def build_test_universe(all_us, all_eu, all_de, all_intl, n_total: int) -> list[tuple[str, str]]:
    """
    En mode TEST : on prend N_TICKERS_TEST tickers AU TOTAL répartis entre marchés.
    Retourne une liste [(ticker, market), ...] mélangée mais représentative.
    """
    # Répartition cible : ~30% US, 30% EU, 20% DE, 20% INTL
    n_us   = max(1, int(n_total * 0.30))
    n_eu   = max(1, int(n_total * 0.30))
    n_de   = max(1, int(n_total * 0.20))
    n_intl = n_total - n_us - n_eu - n_de
    
    selected = []
    selected.extend([(t, "SP500") for t in all_us[:n_us]])
    selected.extend([(t, "STOXX") for t in all_eu[:n_eu]])
    selected.extend([(t, "DAX")   for t in all_de[:n_de]])
    selected.extend([(t, "INTL")  for t in all_intl[:n_intl]])
    return selected


def run_data_pipeline() -> tuple[pd.DataFrame, list[dict], str, str, str]:
    """
    Fetch yfinance → DataFrame triée + benchmarks.
    
    ⚠️  Ordre CRITIQUE pour éviter le rate limit Yahoo Finance :
      1. Benchmarks D'ABORD (S&P 500, STOXX 600, CAC 40)
      2. Pause 10s
      3. SP500 → pause 15s → STOXX → pause 15s → DAX → pause 15s → INTL
    
    Returns: (df, benchmarks, snapshot, period, suffix)
    """
    snapshot = datetime.now().strftime("%Y-%m-%d")
    period   = datetime.now().strftime("%B %Y").upper()
    suffix   = "_test" if TEST_MODE else ""

    # ── Construction des univers ─────────────────────────────────────
    all_us   = list(dict.fromkeys(SP500 + SP400_MID + SP600_SMALL + NASDAQ100_EXTRA))
    all_eu   = list(dict.fromkeys(STOXX + SBF120_MID + FTSE250 + STOXX_MID_EU + CAC_MID_60 + CAC_SMALL))
    all_de   = list(dict.fromkeys(DAX + MDAX + SDAX + TECDAX))
    all_intl = list(dict.fromkeys(NIKKEI + TSX60 + ASX50 + HSI))
    
    # ⚠️ DÉDUPLICATION GLOBALE : un ticker peut être dans 2 listes (ex: TUI1.DE dans MDAX + STOXX_MID_EU)
    # On retire chaque ticker des listes suivantes pour éviter double-fetch + doublons dans le DataFrame
    _seen: set[str] = set()
    def _dedup(lst: list[str]) -> list[str]:
        out = []
        for t in lst:
            if t not in _seen:
                _seen.add(t)
                out.append(t)
        return out
    
    all_us   = _dedup(all_us)
    all_eu   = _dedup(all_eu)
    all_de   = _dedup(all_de)
    all_intl = _dedup(all_intl)

    # ── 1. Fetch benchmarks AVANT tout (anti rate-limit) ─────────────
    log.info("\n📊  Fetch benchmarks (S&P 500, CAC 40, STOXX 600) - EN PREMIER (anti rate-limit)…")
    benchmarks = fetch_benchmarks()
    for bm in benchmarks:
        perf = bm.get("perf_1m")
        perf_str = f"{perf:+.2f}%".replace(".", ",") if perf is not None else "N/A"
        log.info("   %s %s : %s", bm["flag"], bm["label"], perf_str)

    # ── 1bis. Fetch FX rates (devises natives → EUR) ─────────────────
    # ⚠️  Doit être fait AVANT les universes (pour pouvoir convertir
    #     chaque prix immédiatement) mais APRÈS les benchmarks (anti rate-limit).
    fx_rates = fetch_fx_rates()

    log.info("   💤 Pause %ds avant fetch universes…", SLEEP_AFTER_BENCH)
    time.sleep(SLEEP_AFTER_BENCH)

    # ── 2. Fetch universes ───────────────────────────────────────────
    if TEST_MODE:
        # Mode test : 30 tickers au total, mélangés
        test_selection = build_test_universe(all_us, all_eu, all_de, all_intl, N_TICKERS_TEST)
        log.info("\n📡  Fetch yfinance MODE TEST : %d tickers (mélangés)", len(test_selection))
        rows: list[dict[str, Any]] = []
        # Groupé par marché pour les logs et le rate-limit
        for market in ["SP500", "STOXX", "DAX", "INTL"]:
            sub = [t for t, m in test_selection if m == market]
            if sub:
                rows.extend(fetch_universe(sub, market, fx_rates))
                if market != "INTL":
                    time.sleep(3)  # Mini pause même en test
        n_total = len(test_selection)
        t0 = time.time()
        elapsed = time.time() - t0
    else:
        # Mode prod : tout l'univers
        n_total = len(all_us) + len(all_eu) + len(all_de) + len(all_intl)
        log.info("\n📡  Fetch yfinance MODE PROD : %d tickers (US=%d EU=%d DE=%d INTL=%d)",
                 n_total, len(all_us), len(all_eu), len(all_de), len(all_intl))
        log.info("   ⏱  Estimation : 35-45 min avec %d workers + sleep %ds entre univers",
                 N_WORKERS, SLEEP_BETWEEN_UNI)

        t0 = time.time()
        rows = []
        for tickers, market in [
            (all_us,   "SP500"),
            (all_eu,   "STOXX"),
            (all_de,   "DAX"),
            (all_intl, "INTL"),
        ]:
            rows.extend(fetch_universe(tickers, market, fx_rates))
            if market != "INTL":  # Pas de pause après le dernier
                log.info("   💤 Pause %ds anti rate-limit Yahoo (entre %s et suivant)...",
                         SLEEP_BETWEEN_UNI, market)
                time.sleep(SLEEP_BETWEEN_UNI)
        elapsed = time.time() - t0

    if not rows:
        raise RuntimeError("❌ Aucune data récupérée - vérifie connexion/yfinance/rate limit")

    # ── 3. Stats finales ─────────────────────────────────────────────
    stats = pd.Series([r["market"] for r in rows]).value_counts().to_dict()
    log.info("  → %d/%d lignes valides en %.1fs  (US=%d EU=%d DE=%d INTL=%d)",
             len(rows), n_total, elapsed,
             stats.get("SP500", 0), stats.get("STOXX", 0),
             stats.get("DAX", 0), stats.get("INTL", 0))

    df = (pd.DataFrame(rows)
            .sort_values("total_pct", ascending=False)
            .reset_index(drop=True))

    log.info("📊  PEA / CTO  : %d PEA · %d CTO",
             int(df["pea"].sum()), int((~df["pea"]).sum()))
    log.info("👥  Couverture analystes : %d/%d",
             int((df["analyst_count"] > 0).sum()), len(df))
    # Couverture conversion EUR : combien de tickers ont un prix EUR valide
    n_eur_ok = int(df["price_eur"].notna().sum())
    log.info("💱  Couverture conversion EUR : %d/%d (%.0f%%)",
             n_eur_ok, len(df), 100 * n_eur_ok / max(1, len(df)))
    # Diversité devises (debug)
    if "currency" in df.columns:
        ccy_counts = df["currency"].value_counts().to_dict()
        log.info("💱  Devises rencontrées : %s",
                 ", ".join(f"{k}={v}" for k, v in ccy_counts.items()))

    return df, benchmarks, snapshot, period, suffix


# ═════════════════════════════════════════════════════════════════════
#  11. EXCEL EXPORT (6 onglets de classements)
# ═════════════════════════════════════════════════════════════════════

def export_xlsx(df: pd.DataFrame, snapshot: str, suffix: str) -> Path:
    """Génère le fichier Excel avec 6 onglets de classements thématiques."""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"ranking_{snapshot}{suffix}.xlsx"

    best_per_sec = (df.sort_values("total_pct", ascending=False)
                      .groupby("sector_fr", sort=False)
                      .head(2).reset_index(drop=True))

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="By Total Gain", index=False)
        df.sort_values("target_pct", ascending=False).to_excel(
            w, sheet_name="By Potential", index=False)
        df[df["div_pct"] > 0].sort_values("div_pct", ascending=False).to_excel(
            w, sheet_name="By Dividend", index=False)
        df.dropna(subset=["perf_1m"]).sort_values("perf_1m", ascending=False).to_excel(
            w, sheet_name="By Perf 1m", index=False)
        df.dropna(subset=["reco_mean"]).sort_values("reco_mean").to_excel(
            w, sheet_name="By POTENTIELS", index=False)
        best_per_sec.to_excel(w, sheet_name="By Sector", index=False)

    log.info("💾  xlsx 6 onglets → %s", path)
    return path


# ═════════════════════════════════════════════════════════════════════
#  12. RANKINGS PRÉPARÉS (alimente post LinkedIn + vidéo)
# ═════════════════════════════════════════════════════════════════════

class Rankings:
    """
    Container pour les classements préparés.
    
    Tri uniformisé v10 :
      - top_perf_*  : par perf_1m desc (mois précédent)
      - top_conv_*  : par total_pct desc (target + dividende)
      - sec_*       : par total_pct desc, 1 ticker par secteur, top N_SECTOR_PER_COL
    """
    def __init__(self, df: pd.DataFrame, suffix: str,
                 benchmarks: list[dict] | None = None) -> None:
        self.df = df
        self.benchmarks = benchmarks or []
        self.df_pea = df[df["pea"] == True].copy()
        self.df_cto = df[df["pea"] == False].copy()

        # ── Top par PEA/CTO séparés (pour la VIDÉO : N_TOP_VIDEO=10) ─
        self.top_perf_pea = (self.df_pea.dropna(subset=["perf_1m"])
                             .sort_values("perf_1m", ascending=False).head(N_TOP_VIDEO))
        self.top_perf_cto = (self.df_cto.dropna(subset=["perf_1m"])
                             .sort_values("perf_1m", ascending=False).head(N_TOP_VIDEO))

        self.top_conv_pea = (self.df_pea[self.df_pea["total_pct"] > 0]
                             .sort_values("total_pct", ascending=False).head(N_TOP_VIDEO))
        self.top_conv_cto = (self.df_cto[self.df_cto["total_pct"] > 0]
                             .sort_values("total_pct", ascending=False).head(N_TOP_VIDEO))

        # ── Sectors PEA / CTO (1 par secteur, top N_SECTOR_PER_COL) ──
        # Tri par total_pct desc dans chaque secteur, puis on prend le 1er par secteur,
        # puis on garde les N_SECTOR_PER_COL meilleurs secteurs.
        self.sec_pea = self._top_sectors(self.df_pea, N_SECTOR_PER_COL)
        self.sec_cto = self._top_sectors(self.df_cto, N_SECTOR_PER_COL)

        # ── Sectors ALIGNÉS : comparaison côte à côte PEA vs CTO ─────
        # Pour chaque secteur GICS : meilleur ticker PEA + meilleur ticker CTO.
        # Tri par max(potentiel PEA, potentiel CTO) desc.
        self.sec_aligned = self._sectors_aligned()

        # ── Top MÉLANGÉS PEA+CTO (pour le POST LinkedIn - Top 5) ─────
        self.top_perf_all = (df.dropna(subset=["perf_1m"])
                             .sort_values("perf_1m", ascending=False)
                             .head(N_TOP))

        self.top_conv_all = (df[df["total_pct"] > 0]
                             .sort_values("total_pct", ascending=False)
                             .head(N_TOP))

        # ── Stats globales ──────────────────────────────────────────
        self.n_pea_total = int(df["pea"].sum())
        self.n_cto_total = int((~df["pea"]).sum())
        self.n_total     = self.n_pea_total + self.n_cto_total
        self.n_covered   = int((df["analyst_count"] > 0).sum())

        # ── Highlights ───────────────────────────────────────────────
        self.best_perf_pea = self.top_perf_pea.iloc[0] if len(self.top_perf_pea) else None
        self.best_perf_cto = self.top_perf_cto.iloc[0] if len(self.top_perf_cto) else None
        self.best_upside   = (df.sort_values("target_pct", ascending=False).iloc[0]
                              if len(df) else None)
        self.n_strong_buy  = int((df["reco_label"].apply(clean_reco)
                                    .str.lower().str.contains("achat fort", na=False)).sum())

    @staticmethod
    def _top_sectors(df_subset: pd.DataFrame, n: int) -> pd.DataFrame:
        """
        Retourne les N meilleurs secteurs (1 ticker par secteur), triés par total_pct desc.
        """
        if df_subset.empty:
            return df_subset
        return (df_subset.sort_values("total_pct", ascending=False)
                .groupby("sector_fr", sort=False)
                .head(1).reset_index(drop=True)
                .head(n))

    def _sectors_aligned(self) -> list[dict]:
        """
        Retourne la liste des secteurs GICS alignés avec meilleur ticker PEA + meilleur ticker CTO.
        
        Pour chaque secteur GICS (11 au total) :
          - Récupère le ticker PEA avec le plus gros total_pct
          - Récupère le ticker CTO avec le plus gros total_pct
          - Skip si les 2 sont absents
        Trié par max(meilleur_PEA, meilleur_CTO) desc.
        
        Returns: [
          {
            "sector_fr": "Technologies de l'information",
            "pea": {ticker, name, total_pct, ...} or None,
            "cto": {ticker, name, total_pct, ...} or None,
            "score_sort": max des potentiels (pour le tri),
          },
          ...
        ]
        """
        # Tous les secteurs GICS connus (= valeurs de SECTOR_FR)
        all_sectors = list(SECTOR_FR.values())
        result: list[dict] = []

        for sector_fr in all_sectors:
            # Best PEA dans ce secteur (si présent)
            pea_in_sec = self.df_pea[self.df_pea["sector_fr"] == sector_fr]
            best_pea = None
            if not pea_in_sec.empty:
                best_pea = (pea_in_sec.sort_values("total_pct", ascending=False)
                                       .iloc[0].to_dict())

            # Best CTO dans ce secteur (si présent)
            cto_in_sec = self.df_cto[self.df_cto["sector_fr"] == sector_fr]
            best_cto = None
            if not cto_in_sec.empty:
                best_cto = (cto_in_sec.sort_values("total_pct", ascending=False)
                                       .iloc[0].to_dict())

            # Skip si les 2 sont vides (peu probable mais sécurité)
            if best_pea is None and best_cto is None:
                continue

            # Score pour le tri = max des potentiels existants
            scores = []
            if best_pea is not None: scores.append(best_pea["total_pct"])
            if best_cto is not None: scores.append(best_cto["total_pct"])

            result.append({
                "sector_fr": sector_fr,
                "pea": best_pea,
                "cto": best_cto,
                "score_sort": max(scores) if scores else 0,
            })

        # Tri par score desc
        result.sort(key=lambda x: x["score_sort"], reverse=True)
        return result


# ═════════════════════════════════════════════════════════════════════
#  13. POST LINKEDIN - markdown-style FR avec split automatique
# ═════════════════════════════════════════════════════════════════════

def _build_links(r: dict, mode: str) -> str:
    """
    Construit la ligne de liens selon le mode :
      - "br_yf"   : 🏛️ BR + 🔍 YF (2 liens)
      - "br_only" : 🏛️ BR uniquement
      - "none"    : pas de liens
    """
    if mode == "none":
        return ""
    bourso = _safe_url(r.get("boursorama_link"))
    yahoo  = _safe_url(r.get("yahoo_link"))
    parts = []
    if bourso: parts.append(f"🏛️ {bourso}")
    if yahoo and mode == "br_yf": parts.append(f"🔍 {yahoo}")
    # Si pas de BR mais YF dispo (cas avec FILTER_BOURSO_ONLY=False), on garde YF en fallback
    if not bourso and yahoo and mode == "br_only":
        parts.append(f"🔍 {yahoo}")
    return ("\n↳ " + " · ".join(parts)) if parts else ""


def _row_perf_post(r: dict, rank: int, links_mode: str = "br_yf") -> str:
    """Formate une ligne du Top 5 PERF pour le POST."""
    flag    = get_flag(r["ticker"])
    medal   = {1: "🥇", 2: "🥈", 3: "🥉", 4: "4️⃣", 5: "5️⃣"}.get(rank, f"{rank:02d}")
    name    = smart_trunc(cap_name(r.get("name", "")), 28)
    elig    = "✅ PEA" if r.get("pea") else "🌍 CTO"
    val = fmt_signed_pct(r.get("perf_1m"))
    price_s = fmt_price(r.get("price_eur"))
    safe_t  = safe_ticker(r["ticker"])
    links   = _build_links(r, links_mode)
    return f"{medal} {name} 📈 {val} · 💵 {price_s} · {flag} {safe_t} {elig}{links}"


def _row_pot_post(r: dict, rank: int, links_mode: str = "br_yf") -> str:
    """Formate une ligne du Top 5 POTENTIEL pour le POST."""
    flag    = get_flag(r["ticker"])
    medal   = {1: "🥇", 2: "🥈", 3: "🥉", 4: "4️⃣", 5: "5️⃣"}.get(rank, f"{rank:02d}")
    name    = smart_trunc(cap_name(r.get("name", "")), 24)
    elig    = "✅ PEA" if r.get("pea") else "🌍 CTO"
    score = fmt_signed_pct(r.get("total_pct"))
    stars   = reco_stars(r.get("reco_mean"))
    stars_block = f"{stars} · " if stars != "-" else ""
    price_s = fmt_price(r.get("price_eur"))
    safe_t  = safe_ticker(r["ticker"])
    links   = _build_links(r, links_mode)
    return f"{medal} {name} 🎯 {score} · {stars_block}💵 {price_s} · {flag} {safe_t} {elig}{links}"


def _row_sec_post(r: dict, with_2_links: bool = True) -> str:
    """Formate une ligne secteur pour le POST (1 ou 2 liens selon longueur dispo)."""
    flag   = get_flag(r["ticker"])
    sec_label, emoji = get_sector_display(r["sector_fr"])
    score  = fmt_signed_pct(r.get("total_pct"))
    name   = smart_trunc(cap_name(r.get("name", "")), 22)
    elig   = "✅PEA" if r.get("pea") else "🌍CTO"
    safe_t = safe_ticker(r["ticker"])

    bourso = _safe_url(r.get("boursorama_link"))
    yahoo  = _safe_url(r.get("yahoo_link"))
    parts = []
    if bourso: parts.append(f"🏛️ {bourso}")
    if yahoo and with_2_links: parts.append(f"🔍 {yahoo}")
    if not bourso and yahoo:   parts.append(f"🔍 {yahoo}")
    links = "\n↳ " + " · ".join(parts) if parts else ""

    return f"{emoji} {sec_label} · {name} {score} · {flag} {safe_t} {elig}{links}"


def _ticker_inline(r: dict | None, with_links: bool = True, label: str = "") -> str:
    """
    Formate un ticker en 1 ligne courte pour la section secteurs alignée.
    
    Args:
        r : dict du ticker (None si pas de ticker côté concerné)
        with_links : True = avec lien BR (priorité) ou YF en fallback, False = sans liens
        label : "🇪🇺 PEA" ou "🌍 CTO"
    
    Returns:
        "🇪🇺 PEA · Sanofi +28.5% · 🇫🇷 SAN.PA\\n     ↳ https://..."
        OU "🇪🇺 PEA · -" si r est None
    """
    if r is None:
        return f"{label} · -"
    flag    = get_flag(r["ticker"])
    name    = smart_trunc(cap_name(r.get("name", "")), 22)
    score   = fmt_signed_pct(r.get("total_pct"))
    safe_t  = safe_ticker(r["ticker"])
    line1   = f"{label} · {name} {score} · {flag} {safe_t}"

    if not with_links:
        return line1

    # Priorité Bourso, fallback Yahoo si pas de Bourso (avec _safe_url anti-NaN)
    url = _safe_url(r.get("boursorama_link")) or _safe_url(r.get("yahoo_link"))
    if not url:
        return line1
    return f"{line1}\n     ↳ {url}"


def _block_sec_aligned_post(sector_data: dict, with_links: bool = True) -> str:
    """
    Formate un bloc secteur pour le POST : on prend LE MEILLEUR ticker (PEA ou CTO),
    avec drapeau, score, éligibilité PEA, et 2 liens (BR + YF).
    
    Format :
      💻 Tech. info.
      🇺🇸 Accenture plc 🎯 +42.3% · ACN 🌍CTO
      ↳ 🏛️ https://... · 🔍 https://...
    """
    sec_label, emoji = get_sector_display(sector_data["sector_fr"])
    
    pea = sector_data.get("pea")
    cto = sector_data.get("cto")
    
    if pea is None and cto is None:
        return f"{emoji} {sec_label}\n-"
    
    # Choisir le meilleur des 2 candidats (par total_pct)
    candidates = []
    if pea is not None: candidates.append(pea)
    if cto is not None: candidates.append(cto)
    best = max(candidates, key=lambda x: x.get("total_pct", 0))
    
    flag    = get_flag(best["ticker"])
    name    = smart_trunc(cap_name(best.get("name", "")), 24)
    score   = fmt_signed_pct(best.get("total_pct"))
    price_s = fmt_price(best.get("price_eur"))
    safe_t  = safe_ticker(best["ticker"])
    elig    = "✅ PEA" if best.get("pea") else "🌍 CTO"
    
    line1 = f"{name} 🎯 {score} · 💵 {price_s} · {flag} {safe_t} {elig}"
    
    if not with_links:
        return f"{emoji} {sec_label}\n{line1}"
    
    # Liens BR + YF (jamais drop, conformément à la règle "liens intouchables")
    bourso = _safe_url(best.get("boursorama_link"))
    yahoo  = _safe_url(best.get("yahoo_link"))
    parts = []
    if bourso: parts.append(f"🏛️ {bourso}")
    if yahoo:  parts.append(f"🔍 {yahoo}")
    links_line = "\n↳ " + " · ".join(parts) if parts else ""
    
    return f"{emoji} {sec_label}\n{line1}{links_line}"


def _build_post_complete(rk: Rankings, period_fr: str, prev_month_fr: str) -> tuple[str, str]:
    """
    Construit le post LinkedIn dans 1 SEUL post.
    
    ⚠️  Stratégie v11 : LIENS INTOUCHABLES (BR + YF préservés à tous les niveaux).
    On drop progressivement uniquement le contenu marketing/explicatif.
    
    Cascade de dégradation :
      N0 : Tout (hook long + CTA emojis + Épingle + sous-titres + 10 hashtags + bench)
      N1 : Drop hook long "85% des fonds..." (utilise hook_court ultra-minimal)
      N2 : + Drop bloc CTA réactions 💡👏❤️
      N3 : + Drop "📌 Épingle · 🔁 Partage" + sous-titres explicatifs Top 5
      N4 : + Réduit hashtags (10 → 5)
      N5 : + Drop bench_line "📊 Marché en..."
    """
    next_month_fr = _next_month_fr(period_fr)
    BAR_S = "━" * 16

    # ── Bench line (optionnel dès N5) ───────────────────────────────
    bench_line = ""
    if rk.benchmarks:
        bench_parts = []
        for bm in rk.benchmarks:
            perf = bm.get("perf_1m")
            if perf is not None:
                bench_parts.append(f"{bm['label']} {fmt_signed_pct(perf)}")
        if bench_parts:
            bench_line = f"📊 Marché en {prev_month_fr.lower()} : " + " · ".join(bench_parts) + "\n\n"

    # ── Hooks (long / court) ────────────────────────────────────────
    hook_normal = f"🚨 {N_ACTIONS_DISPLAY} actions analysées !"

    hook_court = f"🚨 {N_ACTIONS_DISPLAY} actions analysées !"

    # ── Sous-titres explicatifs (optionnels dès N3) ─────────────────
    perf_subtitle = "Les plus fortes hausses (PEA + CTO)."
    pot_subtitle  = "Score = cible 12m + div. ★★★★★ = consensus achat fort."

    # ── Hashtags (10 ou 5) ──────────────────────────────────────────
    hashtags_full = "#Bourse #YahooFinance #Boursorama #Consensus"
    hashtags_min  = "#Bourse #YahooFinance #Boursorama #Consensus"  

    # ── Blocs CTA optionnels ────────────────────────────────────────
    cta_emojis_block = """💬 Choisis ta réaction selon ta stratégie :
💡 Instructif = je suis 100% ETF
👏 Bravo = je fais du stock-picking
❤️ Adore = approche hybride ETF + stock-picking"""

    share_block = "📌 Épingle  ·  🔁 Partage à un débutant en bourse"

    # ── Règle d'or (TOUJOURS présente, jamais drop) ─────────────────
    rule_dor = f"""🎯 RÈGLE D'OR
SOCLE 50-60% = 2 ETF mondiaux
🇺🇸 S&P 500 {ETF_SP500_URL}
🇪🇺 STOXX 600 {ETF_STOXX_URL}
FUN 40-50% = stock-picking, 1 action/secteur minimum"""

    def _build_cta(with_emojis: bool, with_share: bool, hashtags: str) -> str:
        """Construit le bloc CTA final selon les flags."""
        parts = []
        parts.append(
            "⚠️ « Risque de perte en capital. Ceci n'est pas un conseil. »\n"
            f"💳 Parrainage Boursorama {CODE_PARRAINAGE} (+100€) : {PARRAINAGE}\n"
            f"🔔 Prochain brief début {next_month_fr} !"
        )
        parts.append(hashtags)
        return "\n\n".join(parts)

    def _build(hook: str, with_bench: bool, with_subtitles: bool,
               with_cta_emojis: bool, with_share: bool, hashtags: str) -> str:
        """Construit le post complet. Liens BR+YF TOUJOURS présents."""
        # Liens toujours en BR + YF (intouchables)
        perf_rows = "\n\n".join(_row_perf_post(r.to_dict(), i, links_mode="br_yf")
                                for i, (_, r) in enumerate(rk.top_perf_all.head(N_TOP).iterrows(), 1))
        pot_rows = "\n\n".join(_row_pot_post(r.to_dict(), i, links_mode="br_yf")
                               for i, (_, r) in enumerate(rk.top_conv_all.head(N_TOP).iterrows(), 1))
        sec_blocks = "\n\n".join(_block_sec_aligned_post(s, with_links=True)
                                 for s in rk.sec_aligned[:N_SECTORS_ALIGNED])

        # En-tête (bench optionnel)
        head = (bench_line if with_bench else "") + hook

        # Sous-titres Top 5 (optionnels)
        perf_sub = f"\n{perf_subtitle}" if with_subtitles else ""
        pot_sub  = f"\n{pot_subtitle}"  if with_subtitles else ""

        cta = _build_cta(with_emojis=with_cta_emojis, with_share=with_share, hashtags=hashtags)

        return f"""{head}
{BAR_S}
📊 BRIEF BOURSE · {period_fr}
{BAR_S}

📈 TOP {N_TOP} PERFORMANCES - {prev_month_fr.split()[0]}{perf_sub}

{perf_rows}

{BAR_S}

⭐ TOP {N_TOP} POTENTIELS (cible + dividende){pot_sub}

{pot_rows}

{BAR_S}

📂 TOP POTENTIELS PAR SECTEUR

{sec_blocks}

{BAR_S}

{rule_dor}

{cta}"""

    # ── Cascade de dégradation (liens INTOUCHABLES partout) ──────────
    levels = [
        # (description, hook, with_bench, with_subtitles, with_cta_emojis, with_share, hashtags)
        ("N0 : Tout",                                       hook_normal, True,  True,  True,  True,  hashtags_full),
        ("N1 : Drop hook long",                             hook_court,  True,  True,  True,  True,  hashtags_full),
        ("N2 : + Drop CTA emojis",                          hook_court,  True,  True,  False, True,  hashtags_full),
        ("N3 : + Drop Épingle/Partage + sous-titres",       hook_court,  True,  False, False, False, hashtags_full),
        ("N4 : + Drop 5 hashtags",                          hook_court,  True,  False, False, False, hashtags_min),
        ("N5 : + Drop bench_line",                          hook_court,  False, False, False, False, hashtags_min),
    ]

    final_post = ""
    candidate = ""
    for desc, hook, wb, ws, we, wsh, hts in levels:
        candidate = _build(hook, wb, ws, we, wsh, hts)
        if len(candidate) <= LINKEDIN_POST_MAX:
            final_post = candidate
            log.info("  ✅ Niveau retenu : %s (%d/%d chars)",
                     desc, len(candidate), LINKEDIN_POST_MAX)
            break
        else:
            log.info("  ⏭  %s : %d chars > %d → essai niveau suivant",
                     desc, len(candidate), LINKEDIN_POST_MAX)

    if not final_post:
        # Dernier recours (très rare) : tronquer brutalement
        final_post = candidate[:LINKEDIN_POST_MAX - 50] + "\n\n[Tronqué]"
        log.error("❌ Toutes les versions dépassent %d chars - tronqué à %d",
                  LINKEDIN_POST_MAX, len(final_post))

    return final_post, ""


def build_linkedin_post(rk: Rankings, period_fr: str, prev_month_fr: str,
                        snapshot: str) -> tuple[str, str]:
    """
    Construit le post LinkedIn + le commentaire optionnel.
    
    Returns: (post_text, comment_text)
      - post_text : toujours <= 3000 chars
      - comment_text : "" si tout rentre dans le post, sinon la section secteurs
    """
    post, comment = _build_post_complete(rk, period_fr, prev_month_fr)
    return post, comment


# ═════════════════════════════════════════════════════════════════════
#  14. VIDÉO MP4 - HTML/CSS rendu Playwright + ffmpeg concat demuxer
# ═════════════════════════════════════════════════════════════════════

# Fonts Google : Inter (titres) + JetBrains Mono (datas tabulaires)
_GFONTS = ("@import url('https://fonts.googleapis.com/css2?"
           "family=Inter:wght@400;500;600;700;800;900&"
           "family=JetBrains+Mono:wght@400;500;700&display=swap');")

# CSS global (palette Bloomberg-style bleu nuit + accents bleu cyan #00d4ff)
_CSS = """
* { margin:0; padding:0; box-sizing:border-box; }
:root {
  --bg:#0a1628; --bg-pan:#0f1d34; --bg-row:#13243f; --grid:#1e3358;
  --blue:#00d4ff; --blue-dim:#0080a0; --amber:#ffaa00;
  --text:#e8eaf0; --text-mid:#8b9bb4; --dim:#4a5c7a;
  --green:#00ff66; --red:#ff2e2e;
  --gold:#ffd700; --silver:#c8c8c8; --bronze:#cd7f32;
  --pea:#00d4ff; --cto:#7c9eff;
}
body {
  width:1080px; height:1350px; background:var(--bg); color:var(--text);
  font-family:'JetBrains Mono','Courier New',monospace;
  font-size:18px; line-height:1.4; overflow:hidden;
  -webkit-font-smoothing:antialiased;
}
.slide { width:1080px; height:1350px; position:relative; }
.bar-top {
  position:absolute; top:0; left:0; right:0; height:58px;
  background:var(--blue); color:#0a1628;
  display:flex; align-items:center; padding:0 32px;
  font-weight:700; font-size:15px; letter-spacing:0.6px;
}
.bar-top .live { display:inline-block; width:10px; height:10px;
  background:#0a1628; border-radius:50%; margin-right:12px; }
.bar-top .sep { color:rgba(10,22,40,0.4); margin:0 14px; }
.bar-top .right { margin-left:auto; font-variant-numeric:tabular-nums; }
.bar-bot {
  position:absolute; bottom:0; left:0; right:0; height:52px;
  display:flex; align-items:center; padding:0 32px;
  border-top:1px solid var(--grid); color:var(--dim); font-size:13px;
  background:#050d1a; letter-spacing:0.5px;
}
.bar-bot .center { flex:1; text-align:center; color:var(--blue); font-weight:600; }
.body { padding:72px 50px 58px; height:1350px; }
.pos { color:var(--green); } .neg { color:var(--red); } .neut { color:var(--text-mid); }
.tabnum { font-variant-numeric:tabular-nums; }

/* Cover */
.cover-eyebrow { color:var(--blue); font-size:18px; font-weight:700;
  letter-spacing:6px; text-transform:uppercase; margin-top:80px; }
.cover-title { font-family:'Inter',sans-serif; font-weight:900;
  font-size:160px; line-height:0.95; color:#fff;
  letter-spacing:-5px; margin-top:40px; }
.cover-title .or { color:var(--blue); }
.cover-period { font-family:'Inter',sans-serif; font-weight:900;
  font-size:96px; color:var(--blue); margin-top:20px; letter-spacing:-2px;
  line-height:0.95; }
.cover-stats { margin-top:80px; display:flex; gap:30px; }
.stat-card { border:2px solid var(--blue-dim); padding:28px 32px;
  flex:1; background:var(--bg-pan); }
.stat-num { font-family:'Inter',sans-serif; font-weight:900;
  font-size:64px; color:var(--blue); line-height:1; letter-spacing:-2px; }
.stat-label { color:var(--text-mid); font-size:13px; letter-spacing:2px;
  text-transform:uppercase; margin-top:14px; }

/* Pills d'indices - 1 ligne pleine largeur */
.cover-indices { display:flex; flex-wrap:nowrap; gap:9px; margin-top:40px;
  width:100%; align-items:center; justify-content:space-between; }
.idx-pill { border:2px solid var(--blue); padding:8px 12px;
  font-weight:700; font-size:12px; color:var(--blue); letter-spacing:0.8px;
  white-space:nowrap; flex-shrink:0; }
.idx-pill.more { border-style:dashed; color:var(--text-mid); border-color:var(--text-mid); }

/* Section benchmarks (sous les pills, 3 cartes PERF MOIS) */
.bench-section { margin-top:40px; padding-top:30px;
  border-top:2px solid var(--grid); }
.bench-title { color:var(--text-mid); font-size:13px; font-weight:700;
  letter-spacing:3px; text-transform:uppercase; margin-bottom:18px; }
.bench-grid { display:flex; gap:20px; }
.bench-card { flex:1; background:var(--bg-pan);
  border:1.5px solid var(--grid); padding:18px 20px; }
.bench-head { color:var(--text); font-family:'Inter',sans-serif;
  font-weight:700; font-size:16px; letter-spacing:-0.2px; }
.bench-perf { font-family:'Inter',sans-serif; font-weight:900;
  font-size:36px; margin-top:8px; letter-spacing:-1px; }
.bench-period { color:var(--dim); font-size:10px; letter-spacing:1.5px;
  text-transform:uppercase; margin-top:4px; }

/* Dual panel (Top Perf / Top Pred / Sectors) */
.dual-title { font-family:'Inter',sans-serif; font-weight:900;
  font-size:56px; color:#fff; letter-spacing:-1.5px; }
.dual-title .or { color:var(--blue); }
.dual-sub { color:var(--text-mid); font-size:14px; margin-top:6px; letter-spacing:0.5px; }
.dual-grid { display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-top:24px; }
.panel { background:var(--bg-pan); border:1px solid var(--grid); overflow:hidden; }
.panel-head { padding:16px 22px; display:flex; align-items:center;
  justify-content:space-between; border-bottom:1px solid var(--grid); }
.panel-head.pea { border-left:5px solid var(--pea); }
.panel-head.cto { border-left:5px solid var(--cto); }
.panel-tag { font-size:17px; font-weight:700; letter-spacing:1.5px; }
.panel-tag.pea { color:var(--pea); } .panel-tag.cto { color:var(--cto); }
.panel-info { font-size:12px; color:var(--text-mid); letter-spacing:1px; }

/* Row : 10 lignes confortables dans 1350px portrait */
.row { padding:10px 18px; border-bottom:1px solid var(--grid);
  display:grid; grid-template-columns:42px 1fr auto;
  align-items:center; gap:12px; height:96px; overflow:hidden; }
.row > * { min-width:0; }
.row-main { min-width:0; overflow:hidden; }
.row.alt { background:var(--bg-row); }
.row.hidden { visibility:hidden; }
.rk { font-family:'Inter',sans-serif; font-weight:800;
  font-size:28px; text-align:center; color:var(--blue); }
.rk.gold { color:var(--gold); } .rk.silver { color:var(--silver); } .rk.bronze { color:var(--bronze); }
.row-main .name { font-family:'Inter',sans-serif; font-weight:700;
  font-size:20px; color:var(--text); letter-spacing:-0.2px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.row-main .ticker { color:var(--blue); font-family:'JetBrains Mono';
  font-weight:600; font-size:13px; margin-top:3px; letter-spacing:0.3px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.row-main .meta { display:flex; gap:12px; margin-top:4px;
  font-size:13px; color:var(--text-mid); letter-spacing:0.3px;
  flex-wrap:nowrap; white-space:nowrap; overflow:hidden; }
.row-main .meta .k { color:var(--dim); margin-right:3px; }
.row-num { text-align:right; }
.row-num .big { font-family:'Inter',sans-serif; font-weight:800;
  font-size:26px; letter-spacing:-0.5px; }
.row-num .small { font-size:10px; color:var(--text-mid); letter-spacing:1.2px;
  margin-top:3px; text-transform:uppercase; font-weight:600; }
.row-num .small-2 { font-size:11px; color:var(--dim); letter-spacing:0.3px;
  margin-top:2px; font-family:'JetBrains Mono'; }
.stars { font-family:'JetBrains Mono'; font-size:13px; letter-spacing:1px; }
.reco-strong { color:var(--green); } .reco-buy { color:#80ff80; }
.reco-hold { color:var(--amber); } .reco-sell { color:#ff8855; }
.reco-vsell { color:var(--red); } .reco-na { color:var(--dim); }

/* Section secteurs ALIGNÉE - 11 lignes, secteur | PEA | CTO */
.sec-aligned-row { padding:4px 14px; border-bottom:1px solid var(--grid);
  display:grid; grid-template-columns:140px 1fr 1fr;
  align-items:center; gap:14px; min-height:66px; }
.sec-aligned-row.hidden { visibility:hidden; }
.sec-aligned-row.alt { background:var(--bg-row); }

/* Colonne 1 : Label du secteur */
.sec-label-col { display:flex; flex-direction:column; gap:2px; }
.sec-label-emoji { font-size:30px; line-height:1; }
.sec-label-text { color:var(--text-mid); font-size:12px;
  letter-spacing:1.2px; text-transform:uppercase; font-weight:600;
  line-height:1.2; margin-top:4px; }

/* Colonnes 2 et 3 : Ticker PEA / Ticker CTO */
.sec-cell { display:grid; grid-template-columns:auto 1fr auto;
  align-items:center; gap:8px; padding:8px 10px;
  background:rgba(255,255,255,0.02); border-left:3px solid var(--grid); }
.sec-cell.has-pea { border-left-color:var(--pea); }
.sec-cell.has-cto { border-left-color:var(--cto); }
.sec-cell.empty { opacity:0.3; border-left-color:var(--dim); }
.sec-cell-flag { font-size:18px; }
.sec-cell-info { min-width:0; }
.sec-cell-name { font-family:'Inter',sans-serif; font-weight:700;
  font-size:17px; color:var(--text); letter-spacing:-0.3px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; line-height:1.2; }
.sec-cell-tk { color:var(--blue); font-family:'JetBrains Mono';
  font-weight:600; font-size:11px; letter-spacing:0.3px; margin-top:2px; }
.sec-cell-stars { font-family:'JetBrains Mono'; font-size:11px;
  letter-spacing:0.8px; margin-top:3px; line-height:1.2; }
.sec-cell-stars .dim { color:var(--dim); font-size:10px; margin-left:2px; }
.sec-cell-detail { font-family:'JetBrains Mono'; font-size:11px;
  color:var(--dim); letter-spacing:0.3px; margin-top:3px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; line-height:1.2; }
.sec-cell-score { font-family:'Inter',sans-serif; font-weight:800;
  font-size:26px; letter-spacing:-0.5px; text-align:right; }
.sec-cell-empty { color:var(--dim); font-size:11px; font-style:italic;
  text-align:center; grid-column:1/-1; }

/* En-tête section secteurs : labels PEA / CTO sur fond bleu */
.sec-headers { display:grid; grid-template-columns:140px 1fr 1fr;
  gap:14px; padding:8px 14px; margin-bottom:14px;
  background:var(--bg-pan); border:1px solid var(--grid); }
.sec-headers .h-label { color:var(--text-mid); font-size:11px;
  font-weight:700; letter-spacing:1.5px; text-transform:uppercase; }
.sec-headers .h-pea { color:var(--pea); }
.sec-headers .h-cto { color:var(--cto); }

/* CTA (slide finale) */
.cta-eyebrow { color:var(--blue); font-size:16px; font-weight:700;
  letter-spacing:5px; text-transform:uppercase; margin-top:40px; }
.cta-title { font-family:'Inter',sans-serif; font-weight:900;
  font-size:108px; line-height:0.96; color:#fff; letter-spacing:-3px; margin-top:26px; }
.cta-title .or { color:var(--blue); }
.cta-sub { color:var(--text-mid); font-size:21px; line-height:1.55; margin-top:30px; }
.etf-row { display:flex; gap:22px; margin-top:55px; }
.etf-card { flex:1; background:var(--bg-pan);
  border:2px solid var(--blue-dim); padding:24px 28px; }
.etf-card .flag { font-size:32px; }
.etf-card .name { color:var(--blue); font-weight:700; font-size:17px;
  margin-top:10px; letter-spacing:0.5px; }
.etf-card .desc { color:var(--text-mid); font-size:13px; margin-top:8px; line-height:1.4; }
.cta-quiz { margin-top:55px; padding:34px;
  border-left:5px solid var(--blue); background:var(--bg-pan); }
.cta-quiz-q { font-family:'Inter',sans-serif; font-weight:700; font-size:26px; color:var(--text); }
.cta-quiz-hint { color:var(--text-mid); font-size:17px;
  margin-top:14px; font-style:italic; line-height:1.5; }

/* Réactions LinkedIn (style natif) - 💡 ETF, 👏 Stock-picking, ❤️ Hybride */
.cta-reactions { display:flex; justify-content:center; gap:48px;
  margin-bottom:24px; padding-bottom:24px;
  border-bottom:1px solid var(--grid); }
.reaction-item { display:flex; flex-direction:column; align-items:center; gap:4px; }
.reaction-emoji { font-size:54px; line-height:1; }
.reaction-label { color:var(--blue); font-size:16px;
  letter-spacing:1.5px; text-transform:uppercase; font-weight:800;
  margin-top:6px; }
.reaction-sub { color:var(--text-mid); font-size:11px;
  letter-spacing:1.2px; text-transform:uppercase; font-weight:600;
  font-style:italic; }

.cta-disc { color:var(--dim); font-size:13px; margin-top:35px;
  font-style:italic; line-height:1.6; }

/* Badge "prochain brief" - sobre bleu, look pro et bien visible */
.next-brief-badge {
  margin-top:30px;
  padding:18px 24px;
  border:2.5px solid var(--blue);
  background:rgba(0,212,255,0.08);
  text-align:center;
  font-family:'Inter',sans-serif;
  font-weight:800;
  font-size:22px;
  color:var(--blue);
  letter-spacing:1.5px;
}
"""


def _wrap(body_html: str, section: str, snapshot: str, period_fr: str, n_total: int) -> str:
    """Wrappe une slide dans le template HTML complet (barre haut/bas + body)."""
    # Convert ISO snapshot (YYYY-MM-DD) → français DD/MM/YYYY
    try:
        y, m, d = snapshot.split("-")
        snapshot_fr = f"{d}/{m}/{y}"
    except Exception:
        snapshot_fr = snapshot
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>{_GFONTS}{_CSS}</style></head><body><div class="slide">
<div class="bar-top">
  <span class="live"></span>EN DIRECT <span class="sep">│</span>
  BRIEF MENSUEL &gt; {section}
  <span class="right">{snapshot_fr}</span>
</div>
{body_html}
<div class="bar-bot">
  <span>{SIGNATURE}</span>
  <span class="center">BRIEF MENSUEL · {period_fr}</span>
  <span>{n_total} VALEURS</span>
</div>
</div></body></html>"""


# ── HTML : Cover slide (intro 4s avec Ken Burns + stats + pills + benchmarks) ──

def html_cover(rk: Rankings, snapshot: str, period_fr: str, frame: int = 3) -> str:
    """
    Cover slide avec Ken Burns subtil : zoom 1.00 → 1.01 sur les 3 frames.
    Affiche progressivement : période → stats → pills → benchmarks.
    
    Args:
        frame : 1 / 2 / 3 (progression du Ken Burns)
    """
    show_period = frame >= 1
    show_stats  = frame >= 2
    show_bench  = frame >= 3

    # Ken Burns SUBTIL : scale 1.00 → 1.005 → 1.010
    zoom_scale = {1: 1.000, 2: 1.005, 3: 1.010}.get(frame, 1.000) if KEN_BURNS else 1.000
    kb_style = f'style="transform:scale({zoom_scale}); transform-origin:center center; transition:transform 1s ease-out;"'

    period_html = f'<div class="cover-period">{period_fr}</div>' if show_period else ""

    # ── 3 cartes stats (valeurs exactes, pas figé "+1000") ──────────
    stats_html = ""
    if show_stats:
        stats_html = f"""
<div class="cover-stats">
  <div class="stat-card"><div class="stat-num">{rk.n_total}</div>
    <div class="stat-label">ACTIONS ANALYSÉES</div></div>
  <div class="stat-card"><div class="stat-num">{rk.n_pea_total}</div>
    <div class="stat-label">PEA · ZONE EEE</div></div>
  <div class="stat-card"><div class="stat-num">{rk.n_cto_total}</div>
    <div class="stat-label">CTO · MONDIAL</div></div>
</div>
"""
        # ── Pills d'indices : 1 ligne pleine largeur + "+N AUTRES" si débordement ──
        visible = INDEX_PILLS_ALL[:INDEX_PILLS_VISIBLE_MAX]
        n_more = len(INDEX_PILLS_ALL) - len(visible)
        pills_html = "\n".join(f'<div class="idx-pill">{p}</div>' for p in visible)
        if n_more > 0:
            pills_html += f'\n<div class="idx-pill more">+{n_more} AUTRES</div>'
        stats_html += f'<div class="cover-indices">{pills_html}</div>'

    # ── Section benchmarks (apparaît frame 3, 3 cartes PERF MOIS) ──
    bench_html = ""
    if show_bench and rk.benchmarks:
        bench_cards = ""
        for bm in rk.benchmarks:
            perf = bm.get("perf_1m")
            if perf is None:
                perf_str = "-"
                perf_cls = "neut"
            else:
                perf_str = f"{perf:+.2f}%".replace(".", ",")
                perf_cls = "pos" if perf > 0 else "neg" if perf < 0 else "neut"
            bench_cards += f"""
<div class="bench-card">
  <div class="bench-head">{bm['flag']} {bm['label']}</div>
  <div class="bench-perf tabnum {perf_cls}">{perf_str}</div>
  <div class="bench-period">PERF MOIS</div>
</div>"""

        # Le mois précédent (analyse) dans le titre "MARCHÉ EN AVRIL 2026"
        prev_month = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
        prev_month_fr_upper = to_fr_month_year(prev_month).upper()

        bench_html = f"""
<div class="bench-section">
  <div class="bench-title">MARCHÉ EN {prev_month_fr_upper}</div>
  <div class="bench-grid">{bench_cards}</div>
</div>
"""

    body = f"""<div class="body" {kb_style}>
  <div class="cover-eyebrow">// BRIEF MENSUEL EQUITY</div>
  <div class="cover-title">BRIEF<br><span class="or">BOURSE.</span></div>
  {period_html}
  {stats_html}
  {bench_html}
</div>"""
    return _wrap(body, "BRIEF", snapshot, period_fr, rk.n_total)


# ── Helpers row HTML pour les slides Perf / Pred ─────────────────────

def _row_perf_html(rank: int, r: dict) -> str:
    """HTML d'une ligne Top Perf dans la vidéo."""
    flag    = get_flag(r["ticker"])
    medal   = {1:"gold", 2:"silver", 3:"bronze"}.get(rank, "")
    perf    = r.get("perf_1m")
    perf_s  = fmt_signed_pct(perf)
    target  = fmt_signed_pct(r.get("target_pct"))
    div     = f"💰 {r['div_pct']:.1f}%".replace(".", ",") if r.get("div_pct", 0) > 0 else ""
    price_s = fmt_price(r.get("price_eur"))
    name = cap_name(r.get("name", ""))
    alt     = "alt" if rank % 2 == 0 else ""
    sec_label, sec_emoji = get_sector_display(r.get("sector_fr", ""))
    return f"""
<div class="row {alt}">
  <div class="rk {medal}">{rank:02d}</div>
  <div class="row-main">
    <div class="name">{html_lib.escape(name)}</div>
    <div class="ticker">{flag}&nbsp;{r["ticker"]} · {sec_emoji} {html_lib.escape(sec_label)}</div>
    <div class="meta">
      <span class="tabnum">💵 {price_s}</span>
      <span class="tabnum">🎯 {target}</span>
      {f'<span class="tabnum">{div}</span>' if div else ''}
    </div>
  </div>
  <div class="row-num">
    <div class="big tabnum {perf_class(perf)}">{perf_s}</div>
    <div class="small">PERF MOIS</div>
  </div>
</div>"""


def _row_conv_html(rank: int, r: dict) -> str:
    """HTML d'une ligne Top POTENTIELS dans la vidéo."""
    flag    = get_flag(r["ticker"])
    medal   = {1:"gold", 2:"silver", 3:"bronze"}.get(rank, "")
    rm      = r.get("reco_mean")
    stars   = reco_stars(rm)
    reco_lb = r.get("reco_label") or "—"
    n_an    = int(r.get("analyst_count", 0) or 0)
    target  = fmt_signed_pct(r.get("target_pct"))
    div_pct = r.get("div_pct", 0) or 0
    div_str = f" · 💰 {div_pct:.1f}%".replace(".", ",") if div_pct > 0 else ""
    price_s = fmt_price(r.get("price_eur"))
    score   = r.get("total_pct")
    score_s = fmt_signed_pct(score)
    name = cap_name(r.get("name", ""))
    alt     = "alt" if rank % 2 == 0 else ""
    sec_label, sec_emoji = get_sector_display(r.get("sector_fr", ""))
    return f"""
<div class="row {alt}">
  <div class="rk {medal}">{rank:02d}</div>
  <div class="row-main">
    <div class="name">{html_lib.escape(name)}</div>
    <div class="ticker">{flag}&nbsp;{r["ticker"]} · {sec_emoji} {html_lib.escape(sec_label)}</div>
    <div class="meta">
      {f'<span class="stars {reco_color_class(rm)}">{stars} ({n_an})</span>' if stars != "-" else ''}
    </div>
  </div>
  <div class="row-num">
    <div class="big tabnum {perf_class(score)}">{score_s}</div>
    <div class="small">POTENTIEL TOTAL</div>
    <div class="small-2">💵 {price_s} · 🎯 {target}{div_str}</div>
  </div>
</div>"""

def _sec_cell_html(row: dict | None, side: str) -> str:
    if row is None:
        return f'<div class="sec-cell empty"><div class="sec-cell-empty">-</div></div>'

    flag    = get_flag(row["ticker"])
    score   = fmt_signed_pct(row.get("total_pct"))
    score_c = perf_class(row.get("total_pct"))
    target  = fmt_signed_pct(row.get("target_pct"))
    div_pct = row.get("div_pct", 0) or 0
    div_str = f" · 💰 {div_pct:.1f}%".replace(".", ",") if div_pct > 0 else ""
    price_s = fmt_price(row.get("price_eur"))
    name    = cap_name(row.get("name", ""))
    rm      = row.get("reco_mean")
    stars   = reco_stars(rm)
    n_an    = int(row.get("analyst_count", 0) or 0)
    stars_line = (f'<div class="sec-cell-stars {reco_color_class(rm)}">{stars} '
                  f'<span class="dim">({n_an})</span></div>') if stars != "-" else ""
    return f"""<div class="sec-cell has-{side}">
  <div class="sec-cell-flag">{flag}</div>
  <div class="sec-cell-info">
    <div class="sec-cell-name">{html_lib.escape(name)}</div>
    <div class="sec-cell-tk">{row["ticker"]}</div>
    {stars_line}
    <div class="sec-cell-detail">💵 {price_s} · 🎯 {target}{div_str}</div>
  </div>
  <div class="sec-cell-score tabnum {score_c}">{score}</div>
</div>"""


def _row_sec_aligned_html(sector_data: dict, alt: bool = False) -> str:
    """
    HTML d'une ligne alignée : Label secteur | Cellule PEA | Cellule CTO.
    """
    sec_label, emoji = get_sector_display(sector_data["sector_fr"])
    pea_cell = _sec_cell_html(sector_data.get("pea"), "pea")
    cto_cell = _sec_cell_html(sector_data.get("cto"), "cto")
    alt_cls = "alt" if alt else ""
    return f"""
<div class="sec-aligned-row {alt_cls}">
  <div class="sec-label-col">
    <div class="sec-label-emoji">{emoji}</div>
    <div class="sec-label-text">{html_lib.escape(sec_label)}</div>
  </div>
  {pea_cell}
  {cto_cell}
</div>"""


def _row_sec_hidden_aligned() -> str:
    return '<div class="sec-aligned-row hidden"></div>'


def _row_hidden() -> str:
    return '<div class="row hidden"></div>'

def _sec_row_hidden() -> str:
    return '<div class="sec-row hidden"></div>'


# ── HTML : Top 10 PERFORMANCES (défilement ligne par ligne) ──────────

def html_perf(rk: Rankings, snapshot: str, period_fr: str, visible: int) -> str:
    """
    Slide Top 10 PERFORMANCES, 2 colonnes PEA / CTO.
    `visible` = nombre de lignes affichées (le reste = hidden) → permet le défilement.
    """
    pea_data = rk.top_perf_pea.head(N_TOP_VIDEO).to_dict("records")
    cto_data = rk.top_perf_cto.head(N_TOP_VIDEO).to_dict("records")
    rows_pea, rows_cto = "", ""
    for i in range(N_TOP_VIDEO):
        rows_pea += _row_perf_html(i+1, pea_data[i]) if (i < visible and i < len(pea_data)) else _row_hidden()
        rows_cto += _row_perf_html(i+1, cto_data[i]) if (i < visible and i < len(cto_data)) else _row_hidden()
    prev_month = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
    prev_month_fr = to_fr_month_year(prev_month)
    body = f"""<div class="body">
  <div class="dual-title">TOP {N_TOP_VIDEO} <span class="or">PERFORMANCES</span></div>
  <div class="dual-sub">{period_fr}  ·  Mois calendaire précédent complet ({prev_month_fr})  ·  Source : Yahoo Finance</div>
  <div class="dual-grid">
    <div class="panel"><div class="panel-head pea">
      <div class="panel-tag pea">🇪🇺 PEA · ZONE EEE</div>
      <div class="panel-info">{rk.n_pea_total} VALEURS</div></div>{rows_pea}</div>
    <div class="panel"><div class="panel-head cto">
      <div class="panel-tag cto">🌍 CTO · MONDIAL</div>
      <div class="panel-info">{rk.n_cto_total} VALEURS</div></div>{rows_cto}</div>
  </div>
</div>"""
    return _wrap(body, f"PERFORMANCES · TOP {N_TOP_VIDEO}", snapshot, period_fr, rk.n_total)


# ── HTML : Top 10 POTENTIELS (défilement ligne par ligne) ────────────

def html_conv(rk: Rankings, snapshot: str, period_fr: str, visible: int) -> str:
    """
    Slide Top 10 POTENTIELS, 2 colonnes PEA / CTO.
    Tri = total_pct desc (cible analystes + dividende).
    """
    pea_data = rk.top_conv_pea.head(N_TOP_VIDEO).to_dict("records")
    cto_data = rk.top_conv_cto.head(N_TOP_VIDEO).to_dict("records")
    rows_pea, rows_cto = "", ""
    for i in range(N_TOP_VIDEO):
        rows_pea += _row_conv_html(i+1, pea_data[i]) if (i < visible and i < len(pea_data)) else _row_hidden()
        rows_cto += _row_conv_html(i+1, cto_data[i]) if (i < visible and i < len(cto_data)) else _row_hidden()
    body = f"""<div class="body">
  <div class="dual-title">TOP {N_TOP_VIDEO} <span class="or">POTENTIELS</span></div>
  <div class="dual-sub">Consensus analystes (★★★★★ = Achat fort)  ·  Filtre : potentiel total &gt; 0  ·  Score = cible + dividende</div>
  <div class="dual-grid">
    <div class="panel"><div class="panel-head pea">
      <div class="panel-tag pea">🇪🇺 PEA · ZONE EEE</div>
      <div class="panel-info">{rk.n_pea_total} VALEURS</div></div>{rows_pea}</div>
    <div class="panel"><div class="panel-head cto">
      <div class="panel-tag cto">🌍 CTO · MONDIAL</div>
      <div class="panel-info">{rk.n_cto_total} VALEURS</div></div>{rows_cto}</div>
  </div>
</div>"""
    return _wrap(body, f"POTENTIELS · TOP {N_TOP_VIDEO}", snapshot, period_fr, rk.n_total)


# ── HTML : TOP 10 POTENTIELS PAR SECTEUR (défilement ligne par ligne) ─

def html_sectors(rk: Rankings, snapshot: str, period_fr: str, visible: int) -> str:
    """
    Slide TOP POTENTIELS PAR SECTEUR - layout aligné PEA vs CTO côte à côte.
    
    11 lignes (1 par secteur GICS) :
      Label secteur | Meilleur ticker PEA dans ce secteur | Meilleur ticker CTO dans ce secteur
    
    Défilement ligne par ligne (`visible` = nombre de lignes affichées).
    """
    aligned_data = rk.sec_aligned[:N_SECTORS_ALIGNED]
    n_total_lines = len(aligned_data)

    # En-tête avec labels PEA / CTO
    headers_html = """
<div class="sec-headers">
  <div class="h-label">SECTEUR GICS</div>
  <div class="h-label h-pea">🇪🇺 MEILLEUR PEA</div>
  <div class="h-label h-cto">🌍 MEILLEUR CTO</div>
</div>"""

    # Lignes (visible = nombre actuellement affichées, le reste hidden pour défilement)
    rows = ""
    for i in range(n_total_lines):
        if i < visible:
            rows += _row_sec_aligned_html(aligned_data[i], alt=(i % 2 == 1))
        else:
            rows += _row_sec_hidden_aligned()

    body = f"""<div class="body">
  <div class="dual-title">TOP <span class="or">POTENTIELS PAR SECTEUR</span></div>
  <div class="dual-sub">Meilleur ticker PEA vs CTO par secteur  ·  Score = potentiel cible + dividende  ·  Tri par max(PEA,CTO)</div>
  {headers_html}
  {rows}
</div>"""
    return _wrap(body, f"POTENTIELS PAR SECTEUR · TOP {N_SECTORS_ALIGNED}", snapshot, period_fr, rk.n_total)


# ── HTML : Slide CTA finale (11s, badge "À BIENTÔT" sobre bleu) ──────

def html_cta(rk: Rankings, snapshot: str, period_fr: str) -> str:
    """Slide CTA finale avec ETF + réactions LinkedIn (💡 Instructif) + badge prochain brief."""
    next_month = _next_month_fr(period_fr)
    body = f"""<div class="body">
  <div class="cta-eyebrow">// LA RÈGLE D'OR</div>
  <div class="cta-title">ETF <span class="or">D'ABORD.</span><br>
    STOCK-PICKING<br>ENSUITE.</div>
  <div class="cta-sub">
    Sur 10 ans, ~85% des fonds gérés activement<br>
    se font ratiboiser par leur indice (étude SPIVA).<br>
    Ton SOCLE, c'est l'ETF. 50-60% du portif minimum.
  </div>
  <div class="etf-row">
    <div class="etf-card">
      <div class="flag">🇺🇸</div>
      <div class="name">ETF S&amp;P 500</div>
      <div class="desc">500 leaders mondiaux US. Ton benchmark de référence.</div>
    </div>
    <div class="etf-card">
      <div class="flag">🇪🇺</div>
      <div class="name">ETF STOXX EUROPE 600</div>
      <div class="desc">Diversification Europe. Complément naturel du S&amp;P 500.</div>
    </div>
  </div>
  <div class="cta-quiz">
    <div class="cta-reactions">
      <div class="reaction-item">
        <div class="reaction-emoji">💡</div>
        <div class="reaction-label">ETF</div>
        <div class="reaction-sub">Instructif</div>
      </div>
      <div class="reaction-item">
        <div class="reaction-emoji">👏</div>
        <div class="reaction-label">Stock-picking</div>
        <div class="reaction-sub">Bravo</div>
      </div>
      <div class="reaction-item">
        <div class="reaction-emoji">❤️</div>
        <div class="reaction-label">Hybride</div>
        <div class="reaction-sub">Adore</div>
      </div>
    </div>
    <div class="cta-quiz-q">💬 Et toi, quelle est ta stratégie ?</div>
    <div class="cta-quiz-hint">Lis bien la légende sous chaque emoji 👇</div>
  </div>
  <div class="next-brief-badge">
    🚀 À BIENTÔT POUR LE BRIEF DE {next_month.upper()}
  </div>
  <div class="cta-disc">
    « Risque de perte en capital. Ceci n'est pas un conseil »
  </div>
</div>"""
    return _wrap(body, "À BIENTÔT", snapshot, period_fr, rk.n_total)


# ═════════════════════════════════════════════════════════════════════
#  15. RENDER FRAMES - Playwright orchestrator (génère tous les PNG)
# ═════════════════════════════════════════════════════════════════════

def _playwright_render_html_to_png(playwright, html: str, out_path: Path,
                                    width: int = VIDEO_W, height: int = VIDEO_H) -> None:
    """Rend un HTML en PNG via Playwright Chromium."""
    browser = playwright.chromium.launch(headless=True)
    try:
        context = browser.new_context(viewport={"width": width, "height": height},
                                      device_scale_factor=1)
        page = context.new_page()
        page.set_content(html, wait_until="networkidle", timeout=20000)
        # Petit délai pour laisser les Google Fonts charger
        page.wait_for_timeout(800)
        page.screenshot(path=str(out_path), full_page=False,
                        clip={"x": 0, "y": 0, "width": width, "height": height})
        context.close()
    finally:
        browser.close()


def _render_html_threaded(html: str, out_path: Path) -> None:
    """Wrapper qui isole Playwright dans un thread (compat Windows asyncio)."""
    exc_holder: list[Exception] = []
    def _runner():
        try:
            with sync_playwright() as p:
                _playwright_render_html_to_png(p, html, out_path)
        except Exception as e:
            exc_holder.append(e)
    th = threading.Thread(target=_runner)
    th.start()
    th.join()
    if exc_holder:
        raise exc_holder[0]


def render_frames_to_disk(rk: Rankings, snapshot: str, period_fr: str,
                          tmpdir: Path) -> list[tuple[Path, float]]:
    """
    Génère TOUTES les frames PNG de la vidéo, retourne [(path, duration), ...].
    
    Allocation 30s :
      - Cover  4s = 3 frames Ken Burns (durées 1.3 / 1.3 / 1.4s)
      - Perf   5s = 10 frames défilement (0.4s chacune) + dernière prolongée
      - Pred   5s = 10 frames défilement (idem)
      - Sect   5s = 10 frames défilement (idem)
      - CTA   11s = 1 frame (durée 11s)
    """
    tmpdir.mkdir(parents=True, exist_ok=True)
    frames: list[tuple[Path, float]] = []

    # ── COVER : 3 frames Ken Burns ──────────────────────────────────
    log.info("\n🎬 RENDU FRAMES - Cover (3 frames Ken Burns)")
    cover_durations = [DUR_COVER * 0.33, DUR_COVER * 0.33, DUR_COVER * 0.34]
    for i in range(1, 4):
        path = tmpdir / f"01_cover_{i:02d}.png"
        log.info("   ↳ frame cover %d/3", i)
        _render_html_threaded(html_cover(rk, snapshot, period_fr, frame=i), path)
        frames.append((path, cover_durations[i-1]))

    # ── TOP 10 PERFORMANCES : défilement 10 frames + hold ───────────
    log.info("🎬 RENDU FRAMES - Top Performances (défilement 10 lignes)")
    # 10 frames de défilement (apparition d'1 ligne par frame), dernière prolongée
    n_anim_frames = N_TOP_VIDEO  # 10
    anim_duration = DUR_TOP_PERF * 0.55  # 55% pour anim
    hold_duration = DUR_TOP_PERF * 0.45  # 45% = 2.93s hold final
    per_frame     = anim_duration / n_anim_frames  # 0.325s par frame anim
    for v in range(1, n_anim_frames + 1):
        path = tmpdir / f"02_perf_{v:02d}.png"
        log.info("   ↳ frame perf %d/%d", v, n_anim_frames)
        _render_html_threaded(html_perf(rk, snapshot, period_fr, visible=v), path)
        # Dernière frame = anim + hold pour laisser le temps de lire
        dur = per_frame + (hold_duration if v == n_anim_frames else 0.0)
        frames.append((path, dur))

    # ── TOP 10 POTENTIELS : défilement 10 frames + hold ─────────────
    log.info("🎬 RENDU FRAMES - Top POTENTIELS (défilement 10 lignes)")
    for v in range(1, n_anim_frames + 1):
        path = tmpdir / f"03_pred_{v:02d}.png"
        log.info("   ↳ frame pred %d/%d", v, n_anim_frames)
        _render_html_threaded(html_conv(rk, snapshot, period_fr, visible=v), path)
        dur = per_frame + (hold_duration if v == n_anim_frames else 0.0)
        frames.append((path, dur))

    # ── TOP 10 POTENTIELS PAR SECTEUR : défilement 10 frames + hold ─
    log.info("🎬 RENDU FRAMES - Sectors (défilement %d secteurs alignés PEA vs CTO)", N_SECTORS_ALIGNED)
    for v in range(1, N_SECTORS_ALIGNED + 1):
        path = tmpdir / f"04_sect_{v:02d}.png"
        log.info("   ↳ frame sect %d/%d", v, N_SECTORS_ALIGNED)
        _render_html_threaded(html_sectors(rk, snapshot, period_fr, visible=v), path)
        dur = (DUR_SECTORS * 0.65 / N_SECTORS_ALIGNED) + (DUR_SECTORS * 0.35 if v == N_SECTORS_ALIGNED else 0.0)
        frames.append((path, dur))

    # ── CTA : 1 seule frame, durée totale 11s ──────────────────────
    log.info("🎬 RENDU FRAMES - CTA (1 frame, 11s)")
    cta_path = tmpdir / "05_cta.png"
    _render_html_threaded(html_cta(rk, snapshot, period_fr), cta_path)
    frames.append((cta_path, DUR_CTA))

    # Total réel = somme des durées
    total = sum(d for _, d in frames)
    log.info("✅ %d frames générées, durée totale = %.2fs (objectif %.0fs)",
             len(frames), total, TOTAL_DURATION)
    return frames


# ═════════════════════════════════════════════════════════════════════
#  16. AUTO-DOWNLOAD MUSIQUE - fallback Internet Archive CC BY-SA 4.0
# ═════════════════════════════════════════════════════════════════════

def auto_download_music(target_path: Path) -> Path | None:
    """
    Si target_path n'existe pas, essaie de télécharger une musique
    depuis DEFAULT_MUSIC_URLS (Internet Archive).
    Retourne target_path si succès, None sinon.
    """
    if target_path.exists() and target_path.stat().st_size > 1000:
        log.info("🎵  Musique perso trouvée → %s", target_path)
        return target_path

    target_path.parent.mkdir(parents=True, exist_ok=True)

    log.info("🎵  Musique perso absente, tentative de téléchargement depuis Internet Archive…")
    for url in DEFAULT_MUSIC_URLS:
        try:
            log.info("   ↳ Essai : %s", url.split("/")[-1])
            r = requests.get(url, timeout=30, stream=True,
                             headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            with open(target_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            if target_path.stat().st_size > 50000:  # 50KB minimum
                log.info("   ✓ Musique téléchargée → %s (%.1f KB)",
                         target_path, target_path.stat().st_size / 1024)
                return target_path
            target_path.unlink(missing_ok=True)
        except Exception as e:
            log.warning("   ⚠️  Échec : %s", e)
            continue

    log.warning("🎵  Aucune musique disponible - la vidéo sera muette")
    return None


# ═════════════════════════════════════════════════════════════════════
#  17. ASSEMBLE MP4 - ffmpeg concat demuxer + musique + effects
# ═════════════════════════════════════════════════════════════════════

def assemble_mp4(frames: list[tuple[Path, float]], out_path: Path,
                 music_path: Path | None, tmpdir: Path) -> Path:
    """
    Assemble les PNGs en MP4 via ffmpeg concat demuxer.
    
    Pipeline ffmpeg :
      1. Input concat (liste de PNGs avec durations explicites)
      2. Filtre video : fade out + vignette éventuelle
      3. Input audio (musique en loop si présente, sinon anullsrc)
      4. Filtre audio : volume + fade in/out
      5. Encode : libx264 preset placebo CRF 12 + AAC 320k + faststart
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # ── Liste concat (format ffmpeg concat demuxer) ─────────────────
    concat_list = tmpdir / "concat.txt"
    with open(concat_list, "w", encoding="utf-8") as f:
        for path, duration in frames:
            f.write(f"file '{path.resolve().as_posix()}'\n")
            f.write(f"duration {duration:.4f}\n")
        # ⚠️  Ajouter la dernière image SANS duration (requis par concat demuxer)
        f.write(f"file '{frames[-1][0].resolve().as_posix()}'\n")

    total_duration = sum(d for _, d in frames)

    # ── Construction des filtres vidéo ──────────────────────────────
    vfilter_parts = [f"fps={VIDEO_FPS}"]
    if VIGNETTE:
        vfilter_parts.append("vignette=PI/5")
    if VIDEO_FADE_OUT > 0:
        fade_start = max(0, total_duration - VIDEO_FADE_OUT)
        vfilter_parts.append(f"fade=t=out:st={fade_start:.3f}:d={VIDEO_FADE_OUT}")
    vfilter = ",".join(vfilter_parts)

    # ── Commande ffmpeg ─────────────────────────────────────────────
    cmd = [
        FFMPEG_BIN, "-y",
        # Input 1 : frames concat
        "-f", "concat", "-safe", "0", "-i", str(concat_list),
    ]
    # Input 2 : musique (loop) ou silence
    if music_path and music_path.exists():
        cmd += ["-stream_loop", "-1", "-i", str(music_path)]
    else:
        cmd += ["-f", "lavfi", "-i", f"anullsrc=channel_layout=stereo:sample_rate=48000"]

    # Filtres
    cmd += [
        "-vf", vfilter,
        "-af", f"volume={AUDIO_VOLUME},afade=t=in:st=0:d={AUDIO_FADE_IN},"
               f"afade=t=out:st={max(0, total_duration - AUDIO_FADE_OUT):.3f}:d={AUDIO_FADE_OUT}",
    ]
    # Encodage vidéo : qualité maximale demandée
    cmd += [
        "-c:v", "libx264", "-preset", VIDEO_PRESET, "-crf", str(VIDEO_CRF),
        "-pix_fmt", "yuv420p", "-r", str(VIDEO_FPS),
        # Encodage audio : AAC 320k
        "-c:a", "aac", "-b:a", AUDIO_BITRATE,
        "-shortest",  # Couper sur la durée vidéo
        # Durée stricte pour éviter dérives
        "-t", f"{total_duration:.3f}",
        # Optimisations LinkedIn
        "-movflags", "+faststart",
        str(out_path),
    ]

    log.info("🎞  ffmpeg assemble (~%.0fs vidéo, preset=%s CRF=%d)…",
             total_duration, VIDEO_PRESET, VIDEO_CRF)
    log.info("   ⏱  Encodage placebo : compte 10-15 minutes pour qualité max")

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        log.error("❌ ffmpeg stderr (50 dernières lignes) :")
        for line in result.stderr.splitlines()[-50:]:
            log.error("   %s", line)
        raise RuntimeError(f"ffmpeg failed with code {result.returncode}")

    size_mb = out_path.stat().st_size / (1024 * 1024)
    log.info("✅  Vidéo MP4 → %s (%.1f MB, %.1fs)",
             out_path, size_mb, total_duration)
    return out_path


# ═════════════════════════════════════════════════════════════════════
#  18. MAKE VIDEO - orchestrateur principal
# ═════════════════════════════════════════════════════════════════════

def make_video(rk: Rankings, snapshot: str, period_fr: str) -> Path:
    """Orchestre la génération de la vidéo MP4 complète."""
    banner(f"🎬 GÉNÉRATION VIDÉO MP4 (target {TOTAL_DURATION:.0f}s)")

    # Vérifications
    ensure_chromium_and_ffmpeg()

    # Musique (perso ou auto-download)
    music_path = auto_download_music(MUSIC_FILE) if AUTO_DOWNLOAD_MUSIC else MUSIC_FILE
    if not (music_path and music_path.exists()):
        log.warning("🎵  Aucune musique : vidéo sans son")
        music_path = None

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    video_path = OUT_DIR / f"brief_{snapshot}{'_test' if TEST_MODE else ''}.mp4"

    # Génération frames dans tmpdir
    with tempfile.TemporaryDirectory(prefix="brief_frames_") as tmp:
        tmpdir = Path(tmp)
        frames = render_frames_to_disk(rk, snapshot, period_fr, tmpdir)
        assemble_mp4(frames, video_path, music_path, tmpdir)

    return video_path


# ═════════════════════════════════════════════════════════════════════
#  19. UPLOAD LITTERBOX - hébergement temporaire pour Make.com → LinkedIn
# ═════════════════════════════════════════════════════════════════════

def upload_to_litterbox(file_path: Path, expiration: str = LITTERBOX_EXPIRATION,
                        max_retries: int = LITTERBOX_MAX_RETRIES) -> str:
    """
    Upload un fichier sur litterbox.catbox.moe (24h-72h max).
    Retourne l'URL publique de téléchargement.
    
    Pourquoi litterbox plutôt que catbox ? Plus stable pour les fichiers vidéo,
    pas de limite de bande passante côté lecture.
    """
    url = "https://litterbox.catbox.moe/resources/internals/api.php"
    log.info("⬆  Upload vers litterbox (%s)…", expiration)

    for attempt in range(1, max_retries + 1):
        try:
            with open(file_path, "rb") as f:
                files = {"fileToUpload": (file_path.name, f, "video/mp4")}
                data  = {"reqtype": "fileupload", "time": expiration}
                r = requests.post(url, files=files, data=data, timeout=600)
                r.raise_for_status()
                uploaded_url = r.text.strip()
                if not uploaded_url.startswith("http"):
                    raise RuntimeError(f"Réponse litterbox inattendue : {uploaded_url[:200]}")
                log.info("   ✓ Upload OK → %s", uploaded_url)
                return uploaded_url
        except Exception as e:
            log.warning("   ⚠️  Tentative %d/%d échouée : %s", attempt, max_retries, e)
            if attempt < max_retries:
                wait = 5 * attempt
                log.info("   💤 Pause %ds avant retry…", wait)
                time.sleep(wait)
    raise RuntimeError(f"Upload litterbox échoué après {max_retries} tentatives")


# ═════════════════════════════════════════════════════════════════════
#  20. SEND WEBHOOK - Make.com (post + commentaire conditionnel)
# ═════════════════════════════════════════════════════════════════════

def send_webhook(snapshot: str, period_fr: str, post: str, comment: str,
                 video_url: str, video_path: Path, n_total: int) -> None:
    """
    Envoie le payload à Make.com.
    
    Payload :
      - post_text     : toujours présent, <= 3000 chars (validé côté code)
      - comment_text  : "" si tout rentre dans le post, sinon section secteurs
      - video_url     : URL litterbox 24h
      - video_mime    : "video/mp4"
      - video_title   : titre LinkedIn de la vidéo
      - date, period, message, filename : métadata
    
    ⚠️  Configuration Make.com requise :
      Module 1 : Créer post LinkedIn UGC avec post_text + video_url
      Module 2 : Filter (if comment_text != "") → Créer commentaire LinkedIn
                 sur le post créé, avec comment_text
    """
    if not WEBHOOK_URL:
        log.warning("⚠️  WEBHOOK_URL absent : pas d'envoi Make.com")
        return

    n_disp = N_ACTIONS_DISPLAY  # "+1000" figé

    payload = {
        "date":         snapshot,
        "period":       period_fr,
        "message":      f"Brief Mensuel Bourse - {period_fr.title()} - {n_disp} actions",
        "filename":     video_path.name,
        "video_base64": video_url,  # NB: contient la base64 (renommé pour clarté côté Make.com)
        "video_mime":   "video/mp4",
        "video_title":  f"Brief Mensuel Bourse · {period_fr.title()} · {n_disp} actions analysées",
        "post_text":    post,
        "comment_text": comment,
        "n_actions":    n_total,
        "post_chars":   len(post),
        "comment_chars": len(comment),
    }

    log.info("📤  Webhook → Make.com (%d chars post, %d chars commentaire)…",
             len(post), len(comment))
    try:
        r = requests.post(WEBHOOK_URL, json=payload, timeout=30)
        log.info("   ↳ HTTP %d - %s", r.status_code, r.text[:300])
        r.raise_for_status()
        log.info("✅  Webhook envoyé avec succès")
    except Exception as e:
        log.error("❌  Échec webhook : %s", e)
        raise


# ═════════════════════════════════════════════════════════════════════
#  21. MAIN - orchestrateur global
# ═════════════════════════════════════════════════════════════════════

def main() -> int:
    """
    Orchestrateur global du brief mensuel.
    
    Étapes :
      1. Check premier jour ouvré du mois (sauf TEST_MODE / FORCE_RUN)
      2. Validation des secrets (WEBHOOK_URL si SEND_TO_WEBHOOK)
      3. Pipeline data : fetch benchmarks → fetch univers → DataFrame
      4. Export XLSX 6 onglets
      5. Snapshot mensuel JSON
      6. Préparation Rankings (Top 5/10 + secteurs)
      7. Construction post LinkedIn (avec split conditionnel si >3000 chars)
      8. Génération vidéo MP4 (30s, placebo CRF 12, audio 320k)
      9. Upload sur litterbox
      10. Envoi webhook Make.com (post + comment + video_url)
    
    Returns: exit code (0 = succès, 1 = erreur)
    """
    banner("🚀 BRIEF MENSUEL EQUITY - v10", char="═")
    log.info("Mode      : %s", "🧪 TEST" if TEST_MODE else "🚀 PROD")
    log.info("Date      : %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log.info("Webhook   : %s", "configuré ✓" if WEBHOOK_URL else "ABSENT ⚠️")
    log.info("Parrainage: %s", CODE_PARRAINAGE)

    # ── 1. Check premier jour ouvré (sauf TEST_MODE ou FORCE_RUN) ────
    force_run = _env_bool("FORCE_RUN", default=False)
    if not TEST_MODE and not force_run:
        if not is_first_business_day_of_month():
            log.info("\nℹ️  Aujourd'hui n'est PAS le premier jour ouvré du mois.")
            log.info("    → exit 0 (skip propre, le workflow réessaie demain)")
            return 0
        log.info("✅  Premier jour ouvré du mois → on lance le brief")

    # ── 2. Validation secrets ────────────────────────────────────────
    send_webhook_flag = SEND_TO_WEBHOOK and _env_bool("SEND_TO_WEBHOOK", default=True)
    if send_webhook_flag and not WEBHOOK_URL:
        log.error("❌  SEND_TO_WEBHOOK=true mais WEBHOOK_URL absent")
        return 1

    try:
        # ── 3. Pipeline data ────────────────────────────────────────
        df, benchmarks, snapshot, period, suffix = run_data_pipeline()
        period_fr = to_fr_period(period)

        # Mois calendaire précédent (pour le hook "Marché en avril 2026 :")
        prev_month_key = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
        prev_month_fr  = to_fr_month_year(prev_month_key)

        # ── 4. Export XLSX 6 onglets ────────────────────────────────
        xlsx_path = export_xlsx(df, snapshot, suffix)

        # ── 5. Snapshot mensuel JSON ────────────────────────────────
        save_snapshot(df, suffix)

        # ── 6. Préparation Rankings ─────────────────────────────────
        rk = Rankings(df, suffix, benchmarks=benchmarks)
        log.info("📊  Highlights :")
        if rk.best_perf_pea is not None:
            log.info("   🇪🇺 Meilleure perf PEA : %s (%+.1f%%)",
                     rk.best_perf_pea["ticker"], rk.best_perf_pea["perf_1m"])
        if rk.best_perf_cto is not None:
            log.info("   🌍 Meilleure perf CTO : %s (%+.1f%%)",
                     rk.best_perf_cto["ticker"], rk.best_perf_cto["perf_1m"])
        if rk.best_upside is not None:
            log.info("   🎯 Plus gros upside   : %s (%+.1f%%)",
                     rk.best_upside["ticker"], rk.best_upside["target_pct"])
        log.info("   ⭐ Achats forts        : %d", rk.n_strong_buy)

        # ── 7. Construction post LinkedIn ───────────────────────────
        banner("📝  POST LINKEDIN")
        post, comment = build_linkedin_post(rk, period_fr, prev_month_fr, snapshot)
        log.info("📝  Post   : %d chars / %d max", len(post), LINKEDIN_POST_MAX)
        log.info("📝  Comm.  : %d chars / %d max",
                 len(comment), LINKEDIN_COMMENT_MAX)

        # Sauvegarde locale (pour debug)
        post_file = OUT_DIR / f"linkedin_post_{snapshot}{suffix}.txt"
        post_file.write_text(post + ("\n\n=== COMMENT ===\n\n" + comment if comment else ""),
                             encoding="utf-8")
        log.info("💾  Post sauvegardé → %s", post_file)

        # ── 8. Génération vidéo MP4 ─────────────────────────────────
        banner("🎬  GÉNÉRATION VIDÉO")
        video_path = make_video(rk, snapshot, period_fr)

        # ── 9. Encodage base64 + Webhook (skip litterbox) ───────────
        if send_webhook_flag:
            banner("📤  WEBHOOK BASE64 (skip litterbox)")
            # Encode vidéo en base64 (pas d'upload externe)
            video_bytes = video_path.read_bytes()
            video_b64 = base64.b64encode(video_bytes).decode("utf-8")
            size_mb = len(video_bytes) / (1024 * 1024)
            size_b64_mb = len(video_b64) / (1024 * 1024)
            log.info("📦  Vidéo MP4 : %.1f MB → base64 : %.1f MB", size_mb, size_b64_mb)
            if size_b64_mb > 5:
                log.warning("⚠️  Payload base64 (%.1f MB) > 5 MB limite Make.com Core/Free !",
                            size_b64_mb)

            # ── Sleep jusqu'à 09:00 Paris pile (best slot LinkedIn B2B) ──
            banner("⏰  ATTENTE CRÉNEAU OPTIMAL (09:00 Paris)")
            sleep_until_paris(hour=9, minute=0)

            send_webhook(snapshot, period_fr, post, comment,
                         video_b64, video_path, rk.n_total)

        # ── Banner final ─────────────────────────────────────────────
        banner("✅  BRIEF MENSUEL TERMINÉ", char="═")
        log.info("📊  %d actions analysées", rk.n_total)
        log.info("💾  Fichiers générés dans : %s", OUT_DIR.resolve())
        log.info("📸  Snapshot conservé dans : %s", SNAPSHOT_DIR.resolve())
        return 0

    except Exception as e:
        log.exception("\n❌  ÉCHEC du brief mensuel : %s", e)
        return 1


# ═════════════════════════════════════════════════════════════════════
#  Entry point
# ═════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    sys.exit(main())
