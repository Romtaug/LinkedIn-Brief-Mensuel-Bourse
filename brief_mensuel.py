"""
═══════════════════════════════════════════════════════════════════════
  brief_mensuel.py · BRIEF MENSUEL EQUITY — GITHUB ACTIONS · PREMIUM v7
  ─────────────────────────────────────────────────────────────────────
  Pipeline complet automatisé pour LinkedIn :
    1. Vérifie si on est le premier jour ouvré du mois → sinon skip
    2. Fetch yfinance (~1075 actions) + benchmarks (S&P 500, CAC 40, STOXX 600)
    3. Génère xlsx + post LinkedIn + vidéo MP4 portrait 1080×1350
    4. Upload vidéo sur litterbox.catbox.moe
    5. Envoie webhook Make.com → LinkedIn auto-post
    6. Rotation snapshots (garde 5 derniers max)
    7. Commit snapshots dans le repo (gestion historique)

  🆕 v7 — Transitions xfade entre sections (cover → perf → conv → sec → cta)
  🆕 v7 — Tri uniforme : Top Potentiel/PREDICTION par total_pct desc partout
  🆕 v7 — CSS .cta-disc fixé
  
  📅 PLANIFICATION :
    Le workflow GitHub Actions se déclenche tous les 1-4 du mois à 7h UTC.
    Le script vérifie : si aujourd'hui = premier jour ouvré du mois → RUN.
    Sinon → skip (exit 0 propre).
    
  🔐 SECRETS GITHUB ACTIONS REQUIS :
    · WEBHOOK_URL       : URL webhook Make.com
    · CODE_PARRAINAGE   : (optionnel, défaut: ROTA0058)
    · PARRAINAGE_URL    : (optionnel, défaut: bour.so/p/GB93ZfQVNVr)

  🎮 VARIABLES D'ENV :
    · TEST_MODE : 'true' (mode test 28 actions) ou 'false' (prod ~1075)
    · SEND_TO_WEBHOOK : 'true'/'false' (par défaut true)
    · FORCE_RUN : 'true' pour bypass check premier jour ouvré (tests)

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
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Any

# ═════════════════════════════════════════════════════════════════════
#  1. CONFIG  ←  lit depuis env vars (GitHub Actions secrets)
# ═════════════════════════════════════════════════════════════════════

def _env_bool(key: str, default: bool = False) -> bool:
    val = os.getenv(key, "").lower().strip()
    if val in ("true", "1", "yes", "y", "on"):
        return True
    if val in ("false", "0", "no", "n", "off"):
        return False
    return default

# ── Modes & quantities ───────────────────────────────────────────────
TEST_MODE        = _env_bool("TEST_MODE", default=True)
N_TICKERS_TEST   = 10
N_TOP            = 5
N_TOP_VIDEO      = 10
N_SECTOR         = 1
N_WORKERS        = 12

# ── Outputs & integrations ───────────────────────────────────────────
SEND_TO_WEBHOOK  = _env_bool("SEND_TO_WEBHOOK", default=True)
WEBHOOK_URL      = os.getenv("WEBHOOK_URL", "").strip()

# ── Litterbox ────────────────────────────────────────────────────────
LITTERBOX_EXPIRATION = "24h"
LITTERBOX_MAX_RETRIES = 3

OUT_DIR          = Path("out")
SNAPSHOT_DIR     = Path("snapshots")
MAX_SNAPSHOTS    = 5

# ── Branding & links ─────────────────────────────────────────────────
SIGNATURE        = "ROMAIN TAUGOURDEAU"
RUBRIQUE         = "BRIEF MENSUEL EQUITY"

PARRAINAGE       = os.getenv("PARRAINAGE_URL", "https://bour.so/p/GB93ZfQVNVr").strip()
CODE_PARRAINAGE  = os.getenv("CODE_PARRAINAGE", "ROTA0058").strip()
ETF_SP500_URL    = "https://www.boursorama.com/bourse/trackers/cours/1rTETZ/"
ETF_STOXX_URL    = "https://www.boursorama.com/bourse/trackers/cours/1rTESE/"

# ── Video config ─────────────────────────────────────────────────────
# Format LinkedIn Feed optimal : 1080×1350 portrait 4:5
# (le post prend +20% de surface visuelle vs carré, +engagement)
VIDEO_W, VIDEO_H = 1080, 1350
VIDEO_FPS        = 30
VIDEO_CRF        = 16                       # 16 = quasi master · 18 = excellent · 20 = très bon
VIDEO_PRESET     = "veryslow"               # veryslow = meilleure compression (encodage + long, mais image + nette)

# ── Audio config (PREMIUM) ───────────────────────────────────────────
MUSIC_FILE       = Path("assets/music.mp3") # Place ton MP3 royalty-free ici
AUDIO_BITRATE    = "256k"                   # AAC bitrate (256 = quasi-CD quality)
AUDIO_FADE_IN    = 0.3                      # secondes (court : musique démarre quasi instantanément)
AUDIO_FADE_OUT   = 2.0                      # secondes
AUDIO_VOLUME     = 0.6                      # 0.0-1.0 : ne pas couvrir la voix off potentielle

# Auto-download : si True et music.mp3 absent → télécharge depuis Internet Archive
# Source : Adrian Diaz · 100 Free Royalty Background Tracks · CC BY-SA 4.0
AUTO_DOWNLOAD_MUSIC = True
DEFAULT_MUSIC_URLS = [
    "https://archive.org/download/100_free_royalty_background_music_tracks/EverythingIsGonnaBeOk.mp3",
    "https://archive.org/download/100_free_royalty_background_music_tracks/FreeLife.mp3",
    "https://archive.org/download/100_free_royalty_background_music_tracks/bright.mp3",
]

# ── Effets vidéo (PREMIUM) ───────────────────────────────────────────
VIDEO_FADE_IN    = 0.0                      # 0 = pas de fondu noir au début (vidéo démarre direct)
VIDEO_FADE_OUT   = 1.0                      # secondes de fade out à la fin
VIGNETTE         = True                     # vignette cinématique légère
KEN_BURNS        = True                     # zoom subtil sur la cover
XFADE_DURATION   = 0.4                      # 🆕 fondu entre sections principales (cover→perf→conv→sec→cta)


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
#  3. AUTO-INSTALL (local convenience — CI installe via workflow)
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

# Install only if missing (idempotent — does nothing in CI where deps preinstalled)
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


# Chemin vers le binaire ffmpeg (rempli par ensure_chromium_and_ffmpeg)
FFMPEG_BIN: str = "ffmpeg"


def ensure_chromium_and_ffmpeg() -> None:
    """Idempotent: install Playwright Chromium + system deps + verify ffmpeg.
    Met à jour la variable globale FFMPEG_BIN avec le chemin du binaire."""
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
            log.warning("  ⚠️  install-deps a échoué — pas grave si déjà installé")
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


# requests cache pour yfinance (réduit le rate-limit Yahoo)
requests_cache.install_cache(".yf_cache.sqlite", expire_after=6 * 3600)


# ═════════════════════════════════════════════════════════════════════
#  4. UNIVERS — ~1075 TICKERS (100% vérifiés)
# ═════════════════════════════════════════════════════════════════════
# ⚠️  BRK.B → BRK-B et BF.B → BF-B : yfinance utilise le tiret pour les
# classes d'actions. Le point fait échouer le fetch.

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
    "PCAR","PKG","PLTR","PANW","PARA","PH","PAYX","PAYC","PYPL","PNR","PEP","PFE",
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

DAX = [
    "ADS.DE","AIR.DE","ALV.DE","BAS.DE","BAYN.DE","BMW.DE","BNR.DE","CBK.DE",
    "CON.DE","1COV.DE","DTG.DE","DBK.DE","DB1.DE","DHL.DE","DTE.DE","EOAN.DE",
    "FRE.DE","HNR1.DE","HEI.DE","HEN3.DE","IFX.DE","MBG.DE","MRK.DE","MTX.DE",
    "MUV2.DE","P911.DE","PAH3.DE","QIA.DE","RHM.DE","RWE.DE","SAP.DE","SRT3.DE",
    "SIE.DE","ENR.DE","SHL.DE","SY1.DE","VOW3.DE","VNA.DE","ZAL.DE","BEI.DE",
]

MDAX = [
    "HOT.DE","LHA.DE","KBX.DE","TLX.DE","NDX1.DE","AIXA.DE","TKA.DE",
    "NDA.DE","HAG.DE","LEG.DE","DHER.DE","R3NK.DE","EVK.DE","NEM.DE",
    "GBF.DE","RAA.DE","KGX.DE","EVD.DE","FNTN.DE","TUI1.DE","FTK.DE","TEG.DE",
    "PUM.DE","FRA.DE","AG1.DE","SDF.DE","BC8.DE","FPE3.DE","UTDI.DE","8TRA.DE",
    "TKMS.DE","AMV0.DE","AT1.DE","DWS.DE","WCH.DE","JEN.DE","KRN.DE","DEZ.DE",
    "SHA0.DE","BOSS.DE","HLE.DE","LXS.DE","SZG.DE","IOS.DE","JUN3.DE","RRTL.DE",
    "SAX.DE","RDC.DE",
]

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

_STOXX_NATIONAL = (
    # CAC 40 — France .PA (PEA)
    ["AC.PA","AI.PA","AIR.PA","ALO.PA","AKE.PA","BNP.PA","BVI.PA","EN.PA","CAP.PA",
     "CA.PA","ACA.PA","BN.PA","DSY.PA","EDEN.PA","ENGI.PA","EL.PA","ERF.PA",
     "RMS.PA","KER.PA","LR.PA","OR.PA","MC.PA","ML.PA","ORA.PA","RI.PA",
     "PUB.PA","RNO.PA","SAF.PA","SGO.PA","SAN.PA","SU.PA","GLE.PA","STLAP.PA",
     "STMPA.PA","TEP.PA","HO.PA","TTE.PA","VIE.PA","DG.PA","VIV.PA"]
    # FTSE 100 — UK .L (NON-PEA)
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
    + ["ACS.MC","ACX.MC","AENA.MC","AMS.MC","ANA.MC","ANE.MC","BBVA.MC","BKT.MC",
       "CABK.MC","CLNX.MC","COL.MC","ELE.MC","ENG.MC","FDR.MC","FER.MC","GRF.MC",
       "IBE.MC","IDR.MC","ITX.MC","LOG.MC","MAP.MC","MEL.MC","MRL.MC",
       "MTS.MC","NTGY.MC","PUIG.MC","RED.MC","REP.MC","ROVI.MC","SAB.MC","SAN.MC",
       "SCYR.MC","SLR.MC","TEF.MC","UNI.MC"]
    + ["MT.AS","ADYEN.AS","AGN.AS","AD.AS","AKZA.AS","ASM.AS","ASML.AS","ASRNL.AS",
       "BESI.AS","DSFIR.AS","EXO.AS","GLPG.AS","HEIA.AS","IMCD.AS","INGA.AS",
       "KPN.AS","NN.AS","PHIA.AS","PRX.AS","RAND.AS","REN.AS","SHELL.AS","UNA.AS",
       "URW.AS","WKL.AS"]
    + ["ABI.BR","ACKB.BR","AED.BR","AGS.BR","ARGX.BR","AZE.BR","COFB.BR","ELI.BR",
       "GBLB.BR","KBC.BR","MELE.BR","PROX.BR","SOF.BR","SOLB.BR","TNET.BR","UCB.BR",
       "UMI.BR","VGP.BR","WDP.BR"]
    + ["A2A.MI","AMP.MI","AZM.MI","BAMI.MI","BPE.MI","BMED.MI","BMPS.MI","BPSO.MI",
       "CPR.MI","DIA.MI","ENEL.MI","ENI.MI","RACE.MI","FBK.MI","G.MI","HER.MI",
       "INW.MI","ISP.MI","INTE.MI","IG.MI","IP.MI","LDO.MI","MB.MI","MONC.MI",
       "NEXI.MI","PIRC.MI","PIA.MI","PRY.MI","PST.MI","REC.MI","SPM.MI","SRG.MI",
       "STLAM.MI","STMMI.MI","TIT.MI","TRN.MI","TEN.MI","UCG.MI","UNI.MI"]
    + ["ABB.ST","ALFA.ST","ASSA-B.ST","ATCO-A.ST","ATCO-B.ST","AZN.ST","BOL.ST",
       "ELUX-B.ST","ERIC-B.ST","ESSITY-B.ST","EVO.ST","GETI-B.ST","HEXA-B.ST",
       "HM-B.ST","INVE-B.ST","KINV-B.ST","NDA-SE.ST","NIBE-B.ST","SAND.ST",
       "SCA-B.ST","SEB-A.ST","SHB-A.ST","SINCH.ST","SKF-B.ST","SWED-A.ST",
       "TEL2-B.ST","TELIA.ST","VOLV-B.ST"]
    + ["ELISA.HE","FORTUM.HE","KESKOB.HE","KNEBV.HE","METSO.HE","NESTE.HE",
       "NOKIA.HE","NDA-FI.HE","ORNBV.HE","OUT1V.HE","SAMPO.HE","STERV.HE",
       "TELIA1.HE","TYRES.HE","UPM.HE","VALMT.HE","WRT1V.HE"]
    + ["AMBU-B.CO","BAVA.CO","CARL-B.CO","CHR.CO","COLO-B.CO","DANSKE.CO",
       "DEMANT.CO","DSV.CO","FLS.CO","GMAB.CO","GN.CO","ISS.CO","JYSK.CO",
       "MAERSK-B.CO","NDA-DK.CO","NETC.CO","NOVO-B.CO","NZYM-B.CO","ORSTED.CO",
       "PNDORA.CO","RBREW.CO","ROCK-B.CO","TRYG.CO","VWS.CO"]
    + ["AKERBP.OL","BAKKA.OL","DNB.OL","EQNR.OL","FRO.OL","GJF.OL","MOWI.OL",
       "NHY.OL","ORK.OL","SALM.OL","SCATC.OL","SUBC.OL","TEL.OL","TGS.OL",
       "TOM.OL","YAR.OL"]
    + ["ANDR.VI","BAWAG.VI","EBS.VI","IIA.VI","LNZ.VI","OMV.VI","POST.VI","RBI.VI",
       "SBO.VI","STR.VI","TKA.VI","UQA.VI","VER.VI","VIG.VI","VOE.VI","WIE.VI"]
    + ["ABBN.SW","ALC.SW","GEBN.SW","GIVN.SW","HOLN.SW","KNIN.SW","LOGN.SW","LONN.SW",
       "NESN.SW","NOVN.SW","PGHN.SW","ROG.SW","SCMN.SW","SGSN.SW","SIKA.SW","SLHN.SW",
       "SOON.SW","SREN.SW","UBSG.SW","ZURN.SW"]
    + ["BIRG.IR","CRH.IR","FBD.IR","GLB.IR","GRP.IR","HBRN.IR","KMR.IR","KRZ.IR",
       "OIZ.IR","RYA.IR","SK3.IR"]
    + ["ALTR.LS","BCP.LS","COR.LS","CTT.LS","EDP.LS","EDPR.LS","GALP.LS","IBS.LS",
       "JMT.LS","MOTA.LS","NOS.LS","NVG.LS","REN.LS","RAM.LS","SEM.LS","SON.LS"]
    # WIG 20 — Pologne .WA (PEA, marché EEE)
    + ["ALE.WA","ALR.WA","BDX.WA","CDR.WA","CPS.WA","DNP.WA","JSW.WA","KGH.WA",
       "KRU.WA","KTY.WA","LPP.WA","MBK.WA","OPL.WA","PCO.WA","PEO.WA","PGE.WA",
       "PKN.WA","PKO.WA","PZU.WA","SPL.WA"]
    # ASE — Grèce .AT (PEA, marché EEE)
    + ["AEGN.AT","ALPHA.AT","ARAIG.AT","BELA.AT","CENER.AT","ELPE.AT","ETE.AT",
       "EUROB.AT","EXAE.AT","GEKTERNA.AT","HTO.AT","JUMBO.AT","LAMDA.AT",
       "METLEN.AT","MOH.AT","MYTIL.AT","OPAP.AT","OTE.AT","PPC.AT","TPEIR.AT",
       "SARANTI.AT","TENERGY.AT","VIO.AT"]
)
STOXX = sorted(set(_STOXX_NATIONAL))


# ═════════════════════════════════════════════════════════════════════
#  4bis. UNIVERS CTO INTERNATIONAL (hors US, hors UE)
# ═════════════════════════════════════════════════════════════════════
# Tickers Yahoo Finance suffixes :
#   .T   → Tokyo (Nikkei)
#   .TO  → Toronto (TSX)
#   .AX  → Sydney (ASX)
#   .HK  → Hong Kong (Hang Seng)

# Nikkei 225 — Japon (top ~100 actions les plus liquides)
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

# TSX 60 — Canada
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

# ASX 50 — Australie (top large-cap)
ASX50 = [
    "BHP.AX","CSL.AX","CBA.AX","NAB.AX","ANZ.AX","WBC.AX","MQG.AX","WES.AX",
    "WOW.AX","COL.AX","RIO.AX","TLS.AX","GMG.AX","FMG.AX","TCL.AX","STO.AX",
    "ALL.AX","REA.AX","COH.AX","BXB.AX","ASX.AX","SUN.AX","QBE.AX","IAG.AX",
    "S32.AX","JBH.AX","MIN.AX","NEM.AX","ORG.AX","RMD.AX","PLS.AX","EVN.AX",
    "JHX.AX","AGL.AX","ORI.AX","AMP.AX","TPG.AX","NWS.AX","MFG.AX","TWE.AX",
    "BSL.AX","ALD.AX","AZJ.AX","NXT.AX","A2M.AX","SCG.AX","IGO.AX","SOL.AX",
    "LLC.AX","DXS.AX",
]

# Hang Seng 50 — Hong Kong (large-cap chinoises et HK)
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
#  5. PEA, SECTEURS, DRAPEAUX
# ═════════════════════════════════════════════════════════════════════

PEA_SUFFIXES = {".PA", ".DE", ".AS", ".BR", ".MI", ".MC", ".LS", ".HE",
                ".OL", ".ST", ".CO", ".VI", ".IR", ".AT", ".WA", ".PR"}
PEA_OVERRIDES = {"MT.AS": True, "URW.AS": True}

def is_pea(ticker: str) -> bool:
    if ticker in PEA_OVERRIDES:
        return PEA_OVERRIDES[ticker]
    if "." not in ticker:
        return False
    return "." + ticker.rsplit(".", 1)[1] in PEA_SUFFIXES

# Yahoo `sector` (en) → label FR officiel GICS
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

# Pour l'affichage compact dans la vidéo et le post : (label_court, emoji)
SECTOR_DISPLAY: dict[str, tuple[str, str]] = {
    "Technologies de l'information": ("Tech. info.",        "💻"),
    "Services de communication":     ("Communication",      "📡"),
    "Consommation discrétionnaire":  ("Conso. discrét.",    "🛍️"),
    "Consommation de base":          ("Conso. de base",     "🛒"),
    "Énergie":                       ("Énergie",            "⚡"),
    "Services financiers":           ("Finance",            "🏦"),
    "Santé":                         ("Santé",              "🏥"),
    "Industrie":                     ("Industrie",          "🏭"),
    "Matériaux":                     ("Matériaux",          "⛏️"),
    "Immobilier":                    ("Immobilier",         "🏢"),
    "Services aux collectivités":    ("Services coll.",     "💡"),
}

def get_sector_display(sector_fr: str) -> tuple[str, str]:
    return SECTOR_DISPLAY.get(sector_fr, (sector_fr, "📌"))

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


# ═════════════════════════════════════════════════════════════════════
#  6. URL BUILDERS — Boursorama (direct) + Yahoo + Google Finance
# ═════════════════════════════════════════════════════════════════════

BOURSO_PREFIX = {
    ".PA": "1rP",  ".AS": "1rA",  ".BR": "FF11-", ".LS": "1rL",
    ".MI": "1g",   ".MC": "FF55-",".DE": "1z",    ".SW": "2a",
}

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
    if "." not in ticker:
        return f"https://www.boursorama.com/cours/{ticker.replace('-', '.')}/"
    base, suf = ticker.rsplit(".", 1)
    suffix = "." + suf
    if suffix == ".L":
        return f"https://www.boursorama.com/cours/1u{base}.L/"
    if suffix in BOURSO_PREFIX:
        return f"https://www.boursorama.com/cours/{BOURSO_PREFIX[suffix]}{base}/"
    return None

def yahoo_url(ticker: str) -> str:
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

def best_link(row: dict[str, Any]) -> str:
    """1 lien unique : Bourso prio, sinon Yahoo."""
    return row.get("boursorama_link") or row.get("yahoo_link") or ""


# ═════════════════════════════════════════════════════════════════════
#  7. FETCHER YFINANCE (threadé · retry · sanity check)
# ═════════════════════════════════════════════════════════════════════

RECO_LABEL_FR = {
    "strong_buy":  "Achat fort",
    "buy":         "Achat",
    "hold":        "Conserver",
    "sell":        "Vendre",
    "strong_sell": "Vente forte",
    "underperform":"Sous-performance",
}

def fetch_one(ticker: str, market: str, max_retries: int = 3) -> dict[str, Any] | None:
    """Fetch un ticker. Drop si data incomplète ou outlier."""
    for attempt in range(max_retries):
        try:
            t = yf.Ticker(ticker)
            info = t.info
            price    = info.get("currentPrice") or info.get("regularMarketPrice")
            target   = info.get("targetMeanPrice")
            sector   = info.get("sector")
            name     = info.get("longName") or info.get("shortName") or ticker
            isin     = info.get("isin")
            exchange = info.get("exchange")

            if not price or not target or price <= 0:
                return None
            if sector not in SECTOR_FR:
                return None

            div_rate   = info.get("dividendRate") or 0
            div_pct    = (div_rate / price) * 100 if div_rate else 0
            target_pct = (target - price) / price * 100

            if target_pct > 200 or target_pct < -90:
                return None

            target_high = info.get("targetHighPrice")
            target_low  = info.get("targetLowPrice")
            target_high_pct = (target_high - price) / price * 100 if target_high else None
            target_low_pct  = (target_low  - price) / price * 100 if target_low  else None
            target_spread_pct = (
                (target_high - target_low) / price * 100
                if (target_high and target_low) else None
            )

            reco_key   = info.get("recommendationKey")
            reco_mean  = info.get("recommendationMean")
            n_analysts = (info.get("numberOfAnalystOpinions")
                          or info.get("numberOfAnalysts") or 0)
            reco_label = RECO_LABEL_FR.get(reco_key, reco_key or "—")

            # Perf mois précédent (calendaire complet)
            perf_1m = None
            try:
                today      = datetime.now().date()
                start_prev = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
                end_prev   = today.replace(day=1)
                hist = t.history(start=start_prev, end=end_prev)
                if len(hist) >= 2:
                    perf_1m = (hist["Close"].iloc[-1] / hist["Close"].iloc[0] - 1) * 100
                    # Sanity check : exclure outliers extrêmes (split mal géré, etc.)
                    if perf_1m is not None and (perf_1m > 100 or perf_1m < -80):
                        log.warning("⚠️  %s : perf_1m suspecte (%+.1f%%) → exclu",
                                    ticker, perf_1m)
                        perf_1m = None
            except Exception:
                pass

            return {
                "ticker": ticker, "name": name, "sector": sector,
                "sector_fr": SECTOR_FR[sector], "market": market,
                "price": round(price, 2), "div_pct": round(div_pct, 2),
                "target_pct": round(target_pct, 2),
                "target_high_pct": round(target_high_pct, 2) if target_high_pct is not None else None,
                "target_low_pct":  round(target_low_pct, 2)  if target_low_pct  is not None else None,
                "target_spread_pct": round(target_spread_pct, 2) if target_spread_pct is not None else None,
                "reco_label": reco_label,
                "reco_mean": round(reco_mean, 2) if reco_mean else None,
                "analyst_count": int(n_analysts) if n_analysts else 0,
                "total_pct": round(target_pct + div_pct, 2),
                "perf_1m": round(perf_1m, 2) if perf_1m is not None else None,
                "pea": is_pea(ticker), "isin": isin or "",
                "boursorama_link": boursorama_url(ticker),
                "yahoo_link":      yahoo_url(ticker),
                "google_link":     google_finance_url(ticker, exchange),
            }
        except Exception:
            if attempt == max_retries - 1:
                return None
            time.sleep(1 + attempt)
    return None


def fetch_universe(tickers: list[str], market: str) -> list[dict[str, Any]]:
    rows = []
    with ThreadPoolExecutor(max_workers=N_WORKERS) as pool:
        futs = {pool.submit(fetch_one, t, market): t for t in tickers}
        for fut in tqdm(as_completed(futs), total=len(futs),
                        desc=f"  {market:<6}", ncols=70, leave=False):
            r = fut.result()
            if r:
                rows.append(r)
    return rows


# ── Benchmarks indices (S&P 500, CAC 40, STOXX 600) ──────────────────
BENCHMARKS = [
    {"ticker": "^GSPC",  "label": "S&P 500",    "flag": "🇺🇸"},
    {"ticker": "^STOXX", "label": "STOXX 600",  "flag": "🇪🇺"},
    {"ticker": "^FCHI",  "label": "CAC 40",     "flag": "🇫🇷"},
]

def fetch_benchmarks() -> list[dict[str, Any]]:
    """Fetch perf 1 mois (mois calendaire précédent complet) pour les indices benchmarks."""
    results = []
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
        except Exception as e:
            log.warning("  ⚠️  Benchmark %s : %s", bm["ticker"], e)
            results.append({**bm, "perf_1m": None})
    return results


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
    """ 'MAY 2026' → 'MAI 2026' """
    parts = period_str.strip().split()
    if not parts:
        return period_str
    mois_traduit = MOIS_FR.get(parts[0].capitalize(), parts[0])
    return " ".join([mois_traduit.upper()] + parts[1:])

def to_fr_month_year(yyyy_mm: str) -> str:
    """ '2026-04' → 'Avril 2026' """
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
    if not name:
        return ""
    return name[0].upper() + name[1:]

def smart_trunc(s: str, n: int = 22) -> str:
    """Tronque sur le dernier espace avant n, ajoute …"""
    if not s:
        return ""
    s = str(s).strip()
    if len(s) <= n:
        return s
    cut = s[:n].rsplit(" ", 1)[0]
    return (cut if cut else s[:n]) + "…"

def n_actions_display(n: int) -> str:
    if n >= 1000: return "+1 000"
    if n >= 500:  return "+500"
    if n >= 100:  return f"+{(n // 100) * 100}"
    return str(n)

def clean_reco(label: Any) -> str:
    if not label or str(label).lower() in ("none", "nan", "—", "", "-"):
        return ""
    return str(label)

def fmt_signed_pct(v: Any) -> str:
    if v is None or pd.isna(v): return "—"
    return f"{v:+.1f}%"

def perf_class(v: Any) -> str:
    if v is None or pd.isna(v): return "neut"
    if v > 0: return "pos"
    if v < 0: return "neg"
    return "neut"

def reco_color_class(reco_mean: Any) -> str:
    if reco_mean is None or pd.isna(reco_mean): return "reco-na"
    if reco_mean <= 1.8: return "reco-strong"
    if reco_mean <= 2.5: return "reco-buy"
    if reco_mean <= 3.5: return "reco-hold"
    if reco_mean <= 4.2: return "reco-sell"
    return "reco-vsell"

def reco_stars(reco_mean: Any) -> str:
    """1.0 = ★★★★★, 5.0 = ☆☆☆☆☆"""
    if reco_mean is None or pd.isna(reco_mean): return "—"
    score = max(0, min(5, 6 - reco_mean))
    full = int(round(score))
    return "★" * full + "☆" * (5 - full)


# ═════════════════════════════════════════════════════════════════════
#  9. SNAPSHOTS (diff mois précédent)
# ═════════════════════════════════════════════════════════════════════

def is_first_business_day_of_month(today: date | None = None) -> bool:
    """True si aujourd'hui est le premier jour ouvré du mois (lundi-vendredi).
    Si 1er = samedi/dimanche → reporte au lundi suivant."""
    if today is None:
        today = date.today()
    d = date(today.year, today.month, 1)
    while d.weekday() >= 5:  # 5 = samedi, 6 = dimanche
        d += timedelta(days=1)
    return today == d


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
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    month_key = datetime.now().strftime("%Y-%m")
    path = SNAPSHOT_DIR / f"ranking_{month_key}{suffix}.json"
    df.to_json(path, orient="records")
    log.info("📸  snapshot sauvegardé → %s", path)
    # Rotation auto
    rotate_snapshots()
    return path

def load_prev_ranks(suffix: str) -> tuple[dict, dict, bool]:
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
    if not prev_ranks:           return ""
    if ticker not in prev_ranks: return "  🆕 Nouvelle entrée"
    delta = prev_ranks[ticker] - cur_rank
    if delta > 0: return f"  🔼 +{delta} place{'s' if delta > 1 else ''} vs mois dernier"
    if delta < 0: return f"  🔽 {abs(delta)} place{'s' if abs(delta) > 1 else ''} vs mois dernier"
    return "  ↔️  Stable"


# ═════════════════════════════════════════════════════════════════════
#  10. PIPELINE PRINCIPALE — fetch + score
# ═════════════════════════════════════════════════════════════════════

def run_data_pipeline() -> tuple[pd.DataFrame, list[dict], str, str, str]:
    """Fetch yfinance → DataFrame triée + benchmarks. Returns df, benchmarks, snapshot, period, suffix."""
    snapshot = datetime.now().strftime("%Y-%m-%d")
    period   = datetime.now().strftime("%B %Y").upper()
    suffix   = "_test" if TEST_MODE else ""

    all_us   = list(dict.fromkeys(SP500))
    all_eu   = list(dict.fromkeys(STOXX + SBF120_MID))
    all_de   = list(dict.fromkeys(DAX + MDAX))
    all_intl = list(dict.fromkeys(NIKKEI + TSX60 + ASX50 + HSI))

    if TEST_MODE:
        sp500_lst = all_us[:N_TICKERS_TEST]
        stoxx_lst = all_eu[:N_TICKERS_TEST]
        dax_lst   = all_de[:N_TICKERS_TEST]
        intl_lst  = all_intl[:N_TICKERS_TEST]
    else:
        sp500_lst, stoxx_lst, dax_lst, intl_lst = all_us, all_eu, all_de, all_intl

    n_total = len(sp500_lst) + len(stoxx_lst) + len(dax_lst) + len(intl_lst)
    log.info("\n📡  Fetch yfinance  (%d tickers : US=%d EU=%d DE=%d INTL=%d)",
             n_total, len(sp500_lst), len(stoxx_lst), len(dax_lst), len(intl_lst))

    t0 = time.time()
    rows = (
        fetch_universe(sp500_lst, "SP500")
        + fetch_universe(stoxx_lst, "STOXX")
        + fetch_universe(dax_lst,   "DAX")
        + fetch_universe(intl_lst,  "INTL")
    )
    elapsed = time.time() - t0

    if not rows:
        raise RuntimeError("❌ Aucune data récupérée — vérifie connexion/yfinance/rate limit")

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
    log.info("👥  Couverture analystes : %d/%d", int((df["analyst_count"] > 0).sum()), len(df))

    # Fetch benchmarks indices (S&P 500, CAC 40, STOXX 600)
    log.info("\n📊  Fetch benchmarks (S&P 500, CAC 40, STOXX 600)…")
    benchmarks = fetch_benchmarks()
    for bm in benchmarks:
        perf = bm.get("perf_1m")
        perf_str = f"{perf:+.2f}%" if perf is not None else "N/A"
        log.info("   %s %s : %s", bm["flag"], bm["label"], perf_str)

    return df, benchmarks, snapshot, period, suffix


# ═════════════════════════════════════════════════════════════════════
#  11. EXCEL EXPORT (6 onglets)
# ═════════════════════════════════════════════════════════════════════

def export_xlsx(df: pd.DataFrame, snapshot: str, suffix: str) -> Path:
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
            w, sheet_name="By PREDICTION", index=False)
        best_per_sec.to_excel(w, sheet_name="By Sector", index=False)

    log.info("💾  xlsx 6 onglets → %s", path)
    return path


# ═════════════════════════════════════════════════════════════════════
#  12. RANKINGS PRÉPARÉS (alimente post LinkedIn + vidéo)
# ═════════════════════════════════════════════════════════════════════

class Rankings:
    """Container pour les classements préparés."""
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

        _pea_conv = self.df_pea.dropna(subset=["reco_mean"])
        _cto_conv = self.df_cto.dropna(subset=["reco_mean"])
        # 🆕 v7 : tri uniformisé par total_pct desc (même tri que le POST)
        self.top_conv_pea = (_pea_conv[_pea_conv["total_pct"] > 0]
                             .sort_values("total_pct", ascending=False).head(N_TOP_VIDEO))
        self.top_conv_cto = (_cto_conv[_cto_conv["total_pct"] > 0]
                             .sort_values("total_pct", ascending=False).head(N_TOP_VIDEO))

        self.sec_pea = (self.df_pea.sort_values("total_pct", ascending=False)
                        .groupby("sector_fr", sort=False)
                        .head(N_SECTOR).reset_index(drop=True))
        self.sec_cto = (self.df_cto.sort_values("total_pct", ascending=False)
                        .groupby("sector_fr", sort=False)
                        .head(N_SECTOR).reset_index(drop=True))

        # ── Top MÉLANGÉS PEA+CTO (pour le POST LinkedIn) ─────────────
        # Top 5 perf : par perf_1m desc, mélangé (TOUS marchés)
        self.top_perf_all = (df.dropna(subset=["perf_1m"])
                             .sort_values("perf_1m", ascending=False)
                             .head(N_TOP))

        # Top 5 cible + dividende : par total_pct (target_pct + div_pct) desc
        self.top_conv_all = (df[df["total_pct"] > 0]
                             .sort_values("total_pct", ascending=False)
                             .head(N_TOP))

        # Best par secteur GICS : 1 ticker par secteur, mélangé
        self.sec_all = (df.sort_values("total_pct", ascending=False)
                        .groupby("sector_fr", sort=False)
                        .head(1).reset_index(drop=True))

        # Stats globales
        self.n_pea_total = int(df["pea"].sum())
        self.n_cto_total = int((~df["pea"]).sum())
        self.n_total     = self.n_pea_total + self.n_cto_total
        self.n_covered   = int((df["analyst_count"] > 0).sum())

        # Highlights
        self.best_perf_pea = self.top_perf_pea.iloc[0] if len(self.top_perf_pea) else None
        self.best_perf_cto = self.top_perf_cto.iloc[0] if len(self.top_perf_cto) else None
        self.best_upside   = (df.sort_values("target_pct", ascending=False).iloc[0]
                              if len(df) else None)
        self.n_strong_buy  = int((df["reco_label"].apply(clean_reco)
                                    .str.lower().str.contains("achat fort", na=False)).sum())

        # Diff mois précédent
        self.prev_ranks_pea, self.prev_ranks_cto, self.has_prev = load_prev_ranks(suffix)


# ═════════════════════════════════════════════════════════════════════
#  13. POST LINKEDIN — markdown-style FR
# ═════════════════════════════════════════════════════════════════════

BAR = "━" * 52
SUB = "─" * 52


def build_linkedin_post(rk: Rankings, period_fr: str, prev_month_fr: str,
                        snapshot: str) -> str:
    """Post LinkedIn FINAL (max 4000 chars).
    - Hook accrocheur (3 premières lignes critiques pour LinkedIn)
    - Top 5 PERF (PEA+CTO mélangés, par perf 1m) → 2 liens B/Y
    - Top 5 POTENTIEL (PEA+CTO mélangés, par target_pct + div_pct) → 2 liens B/Y + étoiles
    - Best par secteur GICS (1 par secteur) → 1 lien Bourso
    - Règle d'or + diversification (sectorielle)
    - Mini mention 'à dans 1 mois pour le brief de [mois+1]'
    - Tag ✅PEA / 🌍CTO + pédagogie par section + CTA fort"""

    # Mois suivant (pour la mini mention de fin de post)
    next_month_fr = _next_month_fr(period_fr)
    # ── Format ticker avec 2 liens (Top 5 perf) ──────────────────────
    def _row_perf(r: dict, rank: int) -> str:
        flag  = get_flag(r["ticker"])
        medal = {1: "🥇", 2: "🥈", 3: "🥉", 4: "4️⃣", 5: "5️⃣"}.get(rank, f"{rank:02d}")
        name  = smart_trunc(cap_name(r.get("name", "")), 28)
        elig  = "✅PEA" if r.get("pea") else "🌍CTO"
        val   = f"{r['perf_1m']:+.1f}%" if pd.notna(r.get("perf_1m")) else "—"

        # ⭐ NOM D'ENTREPRISE EN AVANT (sans lien : tickers non cliquables)
        return f"{medal} {name}  📈 {val}  ·  {flag} {r['ticker']} {elig}"

    # ── Format ticker (Top 5 potentiel) — sans lien ──────────────────
    def _row_pot(r: dict, rank: int) -> str:
        flag  = get_flag(r["ticker"])
        medal = {1: "🥇", 2: "🥈", 3: "🥉", 4: "4️⃣", 5: "5️⃣"}.get(rank, f"{rank:02d}")
        name  = smart_trunc(cap_name(r.get("name", "")), 24)
        elig  = "✅PEA" if r.get("pea") else "🌍CTO"
        score = f"{r['total_pct']:+.1f}%"
        stars = reco_stars(r.get("reco_mean"))

        # ⭐ NOM D'ENTREPRISE EN AVANT (sans lien : tickers non cliquables)
        return f"{medal} {name}  🎯 {score}  {stars}  ·  {flag} {r['ticker']} {elig}"

    # ── Format ticker secteur (1 ligne sans lien, plus compact) ──────
    def _row_sec(r: dict) -> str:
        flag = get_flag(r["ticker"])
        sec_label, emoji = get_sector_display(r["sector_fr"])
        score = f"{r['total_pct']:+.1f}%"
        name = smart_trunc(cap_name(r.get("name", "")), 22)
        elig = "✅PEA" if r.get("pea") else "🌍CTO"
        # ⭐ NOM D'ENTREPRISE EN AVANT, pas de lien (déjà dans Top 5+5)
        return f"{emoji} {sec_label} · {name} {score} · {flag} {r['ticker']} {elig}"

    # ── Build sections ───────────────────────────────────────────────
    perf_rows = "\n\n".join(_row_perf(r.to_dict(), i)
                            for i, (_, r) in enumerate(rk.top_perf_all.head(5).iterrows(), 1))
    pot_rows  = "\n\n".join(_row_pot(r.to_dict(), i)
                            for i, (_, r) in enumerate(rk.top_conv_all.head(5).iterrows(), 1))
    sec_rows  = "\n".join(_row_sec(r.to_dict())
                          for _, r in rk.sec_all.iterrows())

    BAR_S = "━" * 32

    # ── Ligne benchmark (S&P 500, CAC 40, STOXX 600 perf 1m) ─────────
    # ⭐ EN PREMIÈRE LIGNE DU POST (avant le hook)
    bench_line = ""
    if rk.benchmarks:
        parts = []
        for bm in rk.benchmarks:
            perf = bm.get("perf_1m")
            if perf is not None:
                parts.append(f"{bm['label']} {perf:+.1f}%")
        if parts:
            bench_line = f"📊 Marché en {prev_month_fr.lower()} : " + " · ".join(parts) + "\n\n"

    # ── POST ─────────────────────────────────────────────────────────
    n_disp = n_actions_display(rk.n_total)
    post = f"""\
{bench_line}🚨 {n_disp} actions analysées ce mois-ci.

85% des fonds gérés activement se font battre par leur indice sur 10 ans.
Frais, biais, hasard : tout joue contre toi en stock-picking pur.

Solution : ETF en socle, stock-picking pour le fun.
Ce brief alimente la 2e partie. Sans hype. Juste de la data.

{BAR_S}
📊 BRIEF BOURSE · {period_fr}
{BAR_S}

📈 TOP 5 PERFORMANCES DU MOIS
Ce qui a le plus monté en {prev_month_fr} (PEA + CTO confondus).

{perf_rows}

{BAR_S}

⭐ TOP 5 POTENTIEL (cible analystes + dividende)
Score = upside 12 mois + rendement dividende.
Étoiles = consensus analystes (★★★★★ = Achat fort).

{pot_rows}

{BAR_S}

📂 BEST PAR SECTEUR
Aucun secteur ne domine 2 décennies de suite. Diversifie.

{sec_rows}

{BAR_S}

🎯 RÈGLE D'OR
SOCLE (50-60%) = 2 ETF mondiaux. Tu copies le marché.
🇺🇸 ETF S&P 500  {ETF_SP500_URL}
🇪🇺 ETF STOXX 600  {ETF_STOXX_URL}
FUN (40-50% max) = stock-picking diversifié.
🔀 Vise 1 action par secteur minimum : Tech, Finance, Santé, Industrie...
Quand un secteur baisse, un autre compense.

{BAR_S}

👍 J'aime  ·  👏 Bravo  ·  ❤️ Adore
💬 Et toi, quelle est ta stratégie d'investissement ?
ETF · Stock-picking · Hybride ? Détaille en commentaire 👇

📌 Épingle  ·  🔁 Partage à un débutant en bourse

⚠️ Brief informatif · NE constitue PAS un conseil en investissement (ni CIF ni CGP) · Risque de perte en capital.

💳 Boursorama via parrainage {CODE_PARRAINAGE} (+100€ chacun) : {PARRAINAGE}

🔔 À dans 1 mois pour le brief de {next_month_fr}.

#BriefMensuelBourse #Investissement #PEA #ETF #Bourse
#Python #DataScience #YahooFinance #Boursorama #Prediction
"""

    # Safety check
    if len(post) > 3950:
        log.warning("⚠️  Post LinkedIn = %d chars (>3950, limite 4000) — risque de rejet API",
                    len(post))
    else:
        log.info("  📏 Post LinkedIn : %d chars (limite 4000, marge %d)",
                 len(post), 4000 - len(post))
    return post


# ═════════════════════════════════════════════════════════════════════
#  14. VIDÉO MP4 — HTML/CSS rendu Playwright + ffmpeg concat demuxer
# ═════════════════════════════════════════════════════════════════════

_GFONTS = ("@import url('https://fonts.googleapis.com/css2?"
           "family=Inter:wght@400;500;600;700;800;900&"
           "family=JetBrains+Mono:wght@400;500;700&display=swap');")

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
.body { padding:85px 50px 70px; height:1350px; }
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
  font-size:72px; color:var(--blue); line-height:1; letter-spacing:-2px; }
.stat-label { color:var(--text-mid); font-size:13px; letter-spacing:2px;
  text-transform:uppercase; margin-top:14px; }
.cover-indices { display:flex; flex-wrap:wrap; gap:12px; margin-top:40px; }
.idx-pill { border:2px solid var(--blue); padding:10px 18px;
  font-weight:700; font-size:14px; color:var(--blue); letter-spacing:1px; }

/* Section benchmarks (sous les pills) */
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

/* Dual panel */
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
  align-items:center; gap:12px; min-height:84px; }
.row.alt { background:var(--bg-row); }
.row.hidden { visibility:hidden; }
.rk { font-family:'Inter',sans-serif; font-weight:800;
  font-size:28px; text-align:center; color:var(--blue); }
.rk.gold { color:var(--gold); } .rk.silver { color:var(--silver); } .rk.bronze { color:var(--bronze); }
/* ⭐ NOM D'ENTREPRISE EN AVANT (gros), ticker en petit dessous */
.row-main .name { font-family:'Inter',sans-serif; font-weight:700;
  font-size:20px; color:var(--text); letter-spacing:-0.2px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.row-main .ticker { color:var(--blue); font-family:'JetBrains Mono';
  font-weight:600; font-size:13px; margin-top:3px; letter-spacing:0.3px; }
.row-main .meta { display:flex; gap:14px; margin-top:4px;
  font-size:13px; color:var(--text-mid); letter-spacing:0.3px; }
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

/* Sectors */
.sec-grid { display:grid; grid-template-columns:1fr 1fr; gap:22px; margin-top:30px; }
.sec-panel { background:var(--bg-pan); border:1px solid var(--grid); }
.sec-list { padding:10px 0; }
.sec-row { padding:14px 22px; border-bottom:1px solid var(--grid);
  display:grid; grid-template-columns:38px 1fr auto;
  align-items:center; gap:12px; }
.sec-row:last-child { border-bottom:none; }
.sec-emoji { font-size:24px; text-align:center; }
.sec-info .sec-name { font-size:12px; color:var(--text-mid);
  letter-spacing:1.5px; text-transform:uppercase; }
.sec-info .sec-stock { font-family:'Inter',sans-serif; font-weight:700;
  font-size:17px; color:var(--text); margin-top:3px; }
.sec-info .sec-tk { color:var(--blue); font-family:'JetBrains Mono';
  font-weight:600; font-size:13px; margin-top:2px; display:block; }
.sec-num { font-family:'Inter',sans-serif; font-weight:800; font-size:22px; }

/* CTA */
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
.cta-quiz-opts { margin-top:20px; display:flex; flex-direction:column; gap:12px; }
.cta-quiz-opt { font-size:18px; color:var(--text-mid); }
.cta-quiz-opt b { color:var(--blue); margin-right:14px;
  font-family:'Inter'; font-weight:800; }

/* Réactions LinkedIn (style natif) */
.cta-reactions { display:flex; justify-content:center; gap:48px;
  margin-bottom:24px; padding-bottom:24px;
  border-bottom:1px solid var(--grid); }
.reaction-item { display:flex; flex-direction:column; align-items:center; gap:6px; }
.reaction-emoji { font-size:54px; line-height:1; }
.reaction-label { color:var(--text-mid); font-size:13px;
  letter-spacing:1.5px; text-transform:uppercase; font-weight:600; }

/* Hint sous la question */
.cta-quiz-hint { color:var(--text-mid); font-size:17px;
  margin-top:14px; font-style:italic; line-height:1.5; }

.cta-disc { color:var(--dim); font-size:13px; margin-top:35px;
  font-style:italic; line-height:1.6; }

/* Badge "prochain brief" - encadré bleu, look pro */
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
    # Convert ISO snapshot (YYYY-MM-DD) to French DD/MM/YYYY
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


def html_cover(rk: Rankings, snapshot: str, period_fr: str, frame: int = 3) -> str:
    """Cover slide avec Ken Burns subtil : zoom 1.00 → 1.04 sur les 3 frames."""
    # Le mois est visible DÈS LA 1ÈRE FRAME (changement v6 : avant frame=2)
    show_period = frame >= 1
    show_stats  = frame >= 2
    show_bench  = frame >= 3
    # Ken Burns SUBTIL : scale 1.00 → 1.005 → 1.01 (à peine perceptible)
    zoom_scale = {1: 1.000, 2: 1.005, 3: 1.010}.get(frame, 1.000) if KEN_BURNS else 1.000
    kb_style = f'style="transform:scale({zoom_scale}); transform-origin:center center; transition:transform 1s ease-out;"'
    period_html = f'<div class="cover-period">{period_fr}</div>' if show_period else ""
    stats_html  = ""
    if show_stats:
        n_disp = n_actions_display(rk.n_total)
        stats_html = f"""
<div class="cover-stats">
  <div class="stat-card"><div class="stat-num">{n_disp}</div>
    <div class="stat-label">ACTIONS ANALYSÉES</div></div>
  <div class="stat-card"><div class="stat-num">{rk.n_pea_total}</div>
    <div class="stat-label">PEA · ZONE EEE</div></div>
  <div class="stat-card"><div class="stat-num">{rk.n_cto_total}</div>
    <div class="stat-label">CTO · MONDIAL</div></div>
</div>
<div class="cover-indices">
  <div class="idx-pill">SP 500</div>
  <div class="idx-pill">CAC 40</div>
  <div class="idx-pill">DAX</div>
  <div class="idx-pill">MDAX</div>
  <div class="idx-pill">STOXX EUROPE</div>
  <div class="idx-pill">SBF 120</div>
  <div class="idx-pill">FTSE 100</div>
  <div class="idx-pill">WIG 20</div>
  <div class="idx-pill">ASE</div>
  <div class="idx-pill">NIKKEI 225</div>
  <div class="idx-pill">TSX 60</div>
  <div class="idx-pill">ASX 50</div>
  <div class="idx-pill">HANG SENG</div>
</div>
"""

    # Section benchmarks (apparaît frame 3, en bas de la cover)
    bench_html = ""
    if show_bench and rk.benchmarks:
        bench_cards = ""
        for bm in rk.benchmarks:
            perf = bm.get("perf_1m")
            if perf is None:
                perf_str = "—"
                perf_cls = "neut"
            else:
                perf_str = f"{perf:+.2f}%"
                perf_cls = "pos" if perf > 0 else "neg" if perf < 0 else "neut"
            bench_cards += f"""
<div class="bench-card">
  <div class="bench-head">{bm['flag']} {bm['label']}</div>
  <div class="bench-perf tabnum {perf_cls}">{perf_str}</div>
  <div class="bench-period">PERF MOIS</div>
</div>"""
        bench_html = f"""
<div class="bench-section">
  <div class="bench-title">MARCHÉ EN {period_fr.upper()}</div>
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


def _row_perf(rank: int, r: dict) -> str:
    flag    = get_flag(r["ticker"])
    medal   = {1:"gold",2:"silver",3:"bronze"}.get(rank, "")
    perf    = r.get("perf_1m")
    perf_s  = fmt_signed_pct(perf)
    target  = fmt_signed_pct(r.get("target_pct"))
    div     = f"💰 {r['div_pct']:.1f}%" if r.get("div_pct", 0) > 0 else ""
    name    = smart_trunc(cap_name(r.get("name", "")), 26)
    alt     = "alt" if rank % 2 == 0 else ""
    return f"""
<div class="row {alt}">
  <div class="rk {medal}">{rank:02d}</div>
  <div class="row-main">
    <div class="name">{html_lib.escape(name)}</div>
    <div class="ticker">{flag}&nbsp;{r["ticker"]}</div>
    <div class="meta">
      <span><span class="k">CIBLE</span><span class="tabnum">{target}</span></span>
      {f'<span>{div}</span>' if div else ''}
    </div>
  </div>
  <div class="row-num">
    <div class="big tabnum {perf_class(perf)}">{perf_s}</div>
    <div class="small">PERF MOIS</div>
  </div>
</div>"""

def _row_conv(rank: int, r: dict) -> str:
    flag    = get_flag(r["ticker"])
    medal   = {1:"gold",2:"silver",3:"bronze"}.get(rank, "")
    rm      = r.get("reco_mean")
    stars   = reco_stars(rm)
    reco_lb = r.get("reco_label") or "—"
    n_an    = int(r.get("analyst_count", 0) or 0)
    target  = fmt_signed_pct(r.get("target_pct"))
    score   = r.get("total_pct")
    score_s = fmt_signed_pct(score)
    name    = smart_trunc(cap_name(r.get("name", "")), 26)
    alt     = "alt" if rank % 2 == 0 else ""
    return f"""
<div class="row {alt}">
  <div class="rk {medal}">{rank:02d}</div>
  <div class="row-main">
    <div class="name">{html_lib.escape(name)}</div>
    <div class="ticker">{flag}&nbsp;{r["ticker"]}</div>
    <div class="meta">
      <span class="stars {reco_color_class(rm)}">{stars}</span>
      <span><span class="k">{html_lib.escape(reco_lb)}</span> ({n_an})</span>
    </div>
  </div>
  <div class="row-num">
    <div class="big tabnum {perf_class(score)}">{score_s}</div>
    <div class="small">POTENTIEL TOTAL</div>
    <div class="small-2">🎯 Cible 12 mois : {target}</div>
  </div>
</div>"""

def _row_hidden() -> str:
    return '<div class="row hidden"></div>'


def html_perf(rk: Rankings, snapshot: str, period_fr: str, visible: int) -> str:
    pea_data = rk.top_perf_pea.head(N_TOP_VIDEO).to_dict("records")
    cto_data = rk.top_perf_cto.head(N_TOP_VIDEO).to_dict("records")
    rows_pea, rows_cto = "", ""
    for i in range(N_TOP_VIDEO):
        rows_pea += _row_perf(i+1, pea_data[i]) if (i < visible and i < len(pea_data)) else _row_hidden()
        rows_cto += _row_perf(i+1, cto_data[i]) if (i < visible and i < len(cto_data)) else _row_hidden()
    body = f"""<div class="body">
  <div class="dual-title">TOP {N_TOP_VIDEO} <span class="or">PERFORMANCES</span></div>
  <div class="dual-sub">{period_fr}  ·  Mois calendaire précédent complet  ·  Source : Yahoo Finance</div>
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


def html_conv(rk: Rankings, snapshot: str, period_fr: str, visible: int) -> str:
    pea_data = rk.top_conv_pea.head(N_TOP_VIDEO).to_dict("records")
    cto_data = rk.top_conv_cto.head(N_TOP_VIDEO).to_dict("records")
    rows_pea, rows_cto = "", ""
    for i in range(N_TOP_VIDEO):
        rows_pea += _row_conv(i+1, pea_data[i]) if (i < visible and i < len(pea_data)) else _row_hidden()
        rows_cto += _row_conv(i+1, cto_data[i]) if (i < visible and i < len(cto_data)) else _row_hidden()
    body = f"""<div class="body">
  <div class="dual-title">TOP {N_TOP_VIDEO} <span class="or">PREDICTION</span></div>
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
    return _wrap(body, f"PREDICTION · TOP {N_TOP_VIDEO}", snapshot, period_fr, rk.n_total)


def _sec_rows(sec_df: pd.DataFrame, max_rows: int = 8, visible: int | None = None) -> str:
    """Si visible est None, affiche tout. Sinon n'affiche que les `visible` premiers."""
    out, seen, count = "", set(), 0
    for _, r in sec_df.iterrows():
        s = r["sector_fr"]
        if s in seen:
            continue
        seen.add(s)
        count += 1
        if count > max_rows:
            break
        # ⭐ Si on a un nombre visible et qu'on dépasse → row hidden (placeholder)
        if visible is not None and count > visible:
            out += '<div class="sec-row hidden"></div>'
            continue
        label, emoji = get_sector_display(s)
        score   = fmt_signed_pct(r.get("total_pct"))
        score_c = perf_class(r.get("total_pct"))
        name    = smart_trunc(cap_name(r.get("name", "")), 22)
        flag    = get_flag(r["ticker"])
        # ⭐ NOM D'ENTREPRISE EN AVANT, ticker en petit dessous
        out += f"""
<div class="sec-row">
  <div class="sec-emoji">{emoji}</div>
  <div class="sec-info">
    <div class="sec-name">{html_lib.escape(label)}</div>
    <div class="sec-stock">{html_lib.escape(name)}</div>
    <div class="sec-tk">{flag} {r["ticker"]}</div>
  </div>
  <div class="sec-num tabnum {score_c}">{score}</div>
</div>"""
    return out

def html_sectors(rk: Rankings, snapshot: str, period_fr: str,
                 visible_pea: int | None = None,
                 visible_cto: int | None = None) -> str:
    """visible_pea/cto : si fourni, défilement progressif (rows hidden au-delà)."""
    body = f"""<div class="body">
  <div class="dual-title">TOP <span class="or">PAR SECTEUR</span></div>
  <div class="dual-sub">Meilleure opportunité dans chaque secteur GICS  ·  Score = potentiel cible + dividende</div>
  <div class="sec-grid">
    <div class="sec-panel"><div class="panel-head pea">
      <div class="panel-tag pea">🇪🇺 PEA · ZONE EEE</div>
      <div class="panel-info">PAR SECTEUR</div></div>
      <div class="sec-list">{_sec_rows(rk.sec_pea, visible=visible_pea)}</div></div>
    <div class="sec-panel"><div class="panel-head cto">
      <div class="panel-tag cto">🌍 CTO · MONDIAL</div>
      <div class="panel-info">PAR SECTEUR</div></div>
      <div class="sec-list">{_sec_rows(rk.sec_cto, visible=visible_cto)}</div></div>
  </div>
</div>"""
    return _wrap(body, "SECTEURS", snapshot, period_fr, rk.n_total)


def html_cta(rk: Rankings, snapshot: str, period_fr: str) -> str:
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
        <div class="reaction-emoji">👍</div>
        <div class="reaction-label">J'aime</div>
      </div>
      <div class="reaction-item">
        <div class="reaction-emoji">👏</div>
        <div class="reaction-label">Bravo</div>
      </div>
      <div class="reaction-item">
        <div class="reaction-emoji">❤️</div>
        <div class="reaction-label">Adore</div>
      </div>
    </div>
    <div class="cta-quiz-q">💬 Réagis + commente ta stratégie</div>
    <div class="cta-quiz-hint">ETF · Stock-picking · Hybride ? Détaille en commentaire 👇</div>
  </div>
  <div class="next-brief-badge">
    🚀 À BIENTÔT POUR LE BRIEF DE {next_month.upper()}
  </div>
  <div class="cta-disc">
    Brief informatif uniquement · Ne constitue PAS un conseil en investissement<br>
    (je ne suis ni CIF ni conseiller patrimonial) · Risque de perte en capital
  </div>
</div>"""
    return _wrap(body, "À BIENTÔT", snapshot, period_fr, rk.n_total)


def render_frames_to_disk(rk: Rankings, snapshot: str, period_fr: str,
                          tmp_dir: Path) -> list[tuple[Path, float]]:
    """Render all frames to PNG files. Returns flat list [(path, duration_seconds), ...].
    Wrapped in a thread to avoid Jupyter asyncio conflict."""
    result: dict[str, Any] = {"frames": [], "error": None}

    def _work():
        # ── FIX Windows + Jupyter ─────────────────────────────────
        if platform.system() == "Windows":
            try:
                import asyncio
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                loop = asyncio.ProactorEventLoop()
                asyncio.set_event_loop(loop)
            except (AttributeError, RuntimeError):
                pass

        try:
            tmp_dir.mkdir(parents=True, exist_ok=True)
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page(viewport={"width": VIDEO_W, "height": VIDEO_H})

                idx = [0]
                def shot(html_str: str, dur_s: float) -> tuple[Path, float]:
                    idx[0] += 1
                    out_path = tmp_dir / f"frame_{idx[0]:03d}.png"
                    page.set_content(html_str, wait_until="networkidle")
                    page.screenshot(path=str(out_path), full_page=False)
                    return (out_path, dur_s)

                # ── COVER (3 frames Ken Burns) ──────────────────
                for f, d in [(1, 0.4), (2, 1.2), (3, 2.2)]:
                    result["frames"].append(shot(html_cover(rk, snapshot, period_fr, f), d))

                # ── TOP PERF (défilement 0.5s/ligne + hold 2.5s) ──
                for v in range(1, N_TOP_VIDEO + 1):
                    result["frames"].append(shot(html_perf(rk, snapshot, period_fr, v), 0.5))
                result["frames"].append(shot(html_perf(rk, snapshot, period_fr, N_TOP_VIDEO), 2.5))

                # ── TOP CONV (défilement 0.5s/ligne + hold 2.5s) ──
                for v in range(1, N_TOP_VIDEO + 1):
                    result["frames"].append(shot(html_conv(rk, snapshot, period_fr, v), 0.5))
                result["frames"].append(shot(html_conv(rk, snapshot, period_fr, N_TOP_VIDEO), 2.5))

                # ── SECTEURS (défilement 0.3s + hold 3s) ────────
                n_sec_pea = min(len(rk.sec_pea), 11)
                n_sec_cto = min(len(rk.sec_cto), 11)
                n_sec_max = max(n_sec_pea, n_sec_cto)
                for v in range(1, n_sec_max + 1):
                    result["frames"].append(shot(
                        html_sectors(rk, snapshot, period_fr,
                                     visible_pea=min(v, n_sec_pea),
                                     visible_cto=min(v, n_sec_cto)),
                        0.3
                    ))
                result["frames"].append(shot(
                    html_sectors(rk, snapshot, period_fr,
                                 visible_pea=n_sec_pea, visible_cto=n_sec_cto),
                    3.0
                ))

                # ── CTA finale (7s) ─────────────────────────────
                result["frames"].append(shot(html_cta(rk, snapshot, period_fr), 7.0))

                browser.close()
        except Exception as e:
            result["error"] = e

    t = threading.Thread(target=_work)
    t.start(); t.join()
    if result["error"]:
        raise result["error"]
    return result["frames"]


def assemble_mp4(frames: list[tuple[Path, float]], output: Path) -> None:
    """Assemble PNG frames into MP4 H.264 — version SIMPLE concat demuxer.
    Garantit que chaque frame est affichée pendant sa durée custom (défilement Top 10)."""
    list_file = output.parent / f"{output.stem}_list.txt"
    total_duration = sum(d for _, d in frames)
    with open(list_file, "w", encoding="utf-8") as f:
        for path, dur in frames:
            f.write(f"file '{path.resolve()}'\n")
            f.write(f"duration {dur:.3f}\n")
        f.write(f"file '{frames[-1][0].resolve()}'\n")

    has_music = MUSIC_FILE.exists() and MUSIC_FILE.is_file()
    if has_music:
        log.info("   🎵 Musique trouvée → %s", MUSIC_FILE)
    else:
        log.info("   🔇 Pas de musique (assets/music.mp3 absent) → piste silencieuse AAC")

    fade_out_start = max(0, total_duration - VIDEO_FADE_OUT)
    v_filter_parts = []
    if VIDEO_FADE_IN > 0:
        v_filter_parts.append(f"fade=t=in:st=0:d={VIDEO_FADE_IN}")
    v_filter_parts.append(f"fade=t=out:st={fade_out_start:.3f}:d={VIDEO_FADE_OUT}")
    if VIGNETTE:
        v_filter_parts.append("vignette=angle=0.5")
    v_filter = ",".join(v_filter_parts)

    cmd = [
        FFMPEG_BIN, "-y", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(list_file),
    ]

    if has_music:
        cmd.extend([
            "-stream_loop", "-1",
            "-i", str(MUSIC_FILE.absolute()),
        ])
        audio_filter = (
            f"volume={AUDIO_VOLUME},"
            f"afade=t=in:st=0:d={AUDIO_FADE_IN},"
            f"afade=t=out:st={fade_out_start:.3f}:d={AUDIO_FADE_OUT}"
        )
    else:
        cmd.extend([
            "-f", "lavfi",
            "-i", "anullsrc=channel_layout=stereo:sample_rate=44100",
        ])
        audio_filter = "anull"

    cmd.extend([
        "-fps_mode", "vfr",
        "-pix_fmt", "yuv420p",
        "-c:v", "libx264",
        "-preset", VIDEO_PRESET,
        "-crf", str(VIDEO_CRF),
        "-vf", v_filter,
        "-c:a", "aac",
        "-b:a", AUDIO_BITRATE,
        "-af", audio_filter,
        "-t", f"{total_duration:.3f}",
        "-shortest",
        "-movflags", "+faststart",
        str(output),
    ])

    subprocess.run(cmd, check=True)
    list_file.unlink(missing_ok=True)


def auto_download_music() -> None:
    """Si AUTO_DOWNLOAD_MUSIC=True et music.mp3 absent → télécharge depuis Internet Archive."""
    if not AUTO_DOWNLOAD_MUSIC:
        return
    if MUSIC_FILE.exists() and MUSIC_FILE.stat().st_size > 100_000:
        log.info("  ✓ Musique déjà présente → %s (%.1f MB)",
                 MUSIC_FILE, MUSIC_FILE.stat().st_size / (1024*1024))
        return

    MUSIC_FILE.parent.mkdir(parents=True, exist_ok=True)
    log.info("\n⚙ Téléchargement musique RF (Internet Archive · CC BY-SA 4.0)…")

    for url in DEFAULT_MUSIC_URLS:
        try:
            name = url.rsplit("/", 1)[-1]
            log.info("   Tentative : %s", name)
            r = requests.get(url, timeout=120, stream=True,
                           headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            with open(MUSIC_FILE, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            size_mb = MUSIC_FILE.stat().st_size / (1024 * 1024)
            log.info("   ✓ Téléchargé : %s (%.1f MB)", MUSIC_FILE.name, size_mb)
            log.info("   ℹ️  Source : Adrian Diaz / Internet Archive · CC BY-SA 4.0")
            return
        except Exception as e:
            log.warning("   ⚠️  Échec %s : %s", name, e)
            if MUSIC_FILE.exists():
                MUSIC_FILE.unlink(missing_ok=True)

    log.warning("⚠️  Auto-download a échoué — la vidéo aura une piste silencieuse AAC")


def make_video(rk: Rankings, snapshot: str, period_fr: str, suffix: str) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    output = OUT_DIR / f"brief_{snapshot}{suffix}.mp4"

    auto_download_music()

    log.info("\n🎬  Génération de la vidéo MP4…")
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        frames = render_frames_to_disk(rk, snapshot, period_fr, tmp_dir)
        log.info("   %d frames rendues", len(frames))
        assemble_mp4(frames, output)

    size_mb = output.stat().st_size / (1024 * 1024)
    duration = sum(d for _, d in frames)
    log.info("✅  vidéo générée → %s  (%.1f MB · %.1fs)", output, size_mb, duration)
    return output


# ═════════════════════════════════════════════════════════════════════
#  15. WEBHOOK  +  UPLOAD LITTERBOX
# ═════════════════════════════════════════════════════════════════════

def upload_to_litterbox(video_path: Path,
                        expiration: str = LITTERBOX_EXPIRATION,
                        max_retries: int = LITTERBOX_MAX_RETRIES) -> str:
    """Upload MP4 sur litterbox.catbox.moe. Retourne l'URL publique."""
    size_mb = video_path.stat().st_size / (1024 * 1024)
    log.info("\n☁️  Upload vers litterbox.catbox.moe  (%.2f MB · expire dans %s)",
             size_mb, expiration)

    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            with open(video_path, "rb") as f:
                r = requests.post(
                    "https://litterbox.catbox.moe/resources/internals/api.php",
                    data={"reqtype": "fileupload", "time": expiration},
                    files={"fileToUpload": (video_path.name, f, "video/mp4")},
                    timeout=120,
                )
            r.raise_for_status()
            url = r.text.strip()
            if not url.startswith("http"):
                raise RuntimeError(f"Réponse inattendue : {url[:200]}")
            log.info("   ✓ URL publique : %s", url)
            return url
        except Exception as e:
            last_err = e
            log.warning("   ⚠️  Tentative %d/%d échouée : %s", attempt, max_retries, e)
            if attempt < max_retries:
                time.sleep(5 * attempt)

    raise RuntimeError(f"Upload litterbox a échoué après {max_retries} tentatives : {last_err}")


def send_webhook(video_path: Path, post_text: str, snapshot: str,
                 period_fr: str, rk: Rankings) -> str | None:
    """Upload la vidéo sur litterbox, puis envoie le webhook avec juste l'URL.
    Retourne l'URL publique de la vidéo (ou None si SEND_TO_WEBHOOK=False)."""
    if not SEND_TO_WEBHOOK:
        log.info("\n⏸  SEND_TO_WEBHOOK = False → POST webhook skippé")
        return None

    # 1. Upload la vidéo sur un host temporaire
    video_url = upload_to_litterbox(video_path)

    # 2. POST sur le webhook Make.com avec juste l'URL
    log.info("\n🚀  Envoi webhook → %s", WEBHOOK_URL)
    n_disp = n_actions_display(rk.n_total)
    payload = {
        "date":        snapshot,
        "period":      period_fr,
        "message":     f"Brief Mensuel Bourse — {period_fr.title()} — {n_disp} actions",
        "filename":    video_path.name,
        "video_url":   video_url,
        "video_mime":  "video/mp4",
        "video_title": f"Brief Mensuel Bourse · {period_fr.title()} · {n_disp} actions analysées",
        "post_text":   post_text,
    }

    try:
        r = requests.post(WEBHOOK_URL, json=payload, timeout=30)
        if r.status_code == 200:
            log.info("✅  webhook OK : %s", r.text[:120])
        else:
            log.error("❌  webhook %d : %s", r.status_code, r.text[:200])
    except Exception as e:
        log.error("❌  webhook erreur : %s", e)

    return video_url


# ═════════════════════════════════════════════════════════════════════
#  16. MAIN
# ═════════════════════════════════════════════════════════════════════

def main() -> None:
    banner("BRIEF MENSUEL EQUITY · GITHUB ACTIONS · PREMIUM v7", "═")

    # ── 0. CHECK : doit-on run aujourd'hui ? ─────────────────────────
    today = date.today()
    force_run = _env_bool("FORCE_RUN", default=False)

    if not force_run and not is_first_business_day_of_month(today):
        d = date(today.year, today.month, 1)
        while d.weekday() >= 5:
            d += timedelta(days=1)
        log.info("\n⏸  Pas le premier jour ouvré du mois.")
        log.info("    Aujourd'hui : %s (%s)", today, today.strftime("%A"))
        log.info("    Attendu     : %s (%s)", d, d.strftime("%A"))
        log.info("    → Skip propre, exit 0\n")
        sys.exit(0)

    if force_run:
        log.info("🔧  FORCE_RUN=true → bypass du check premier jour ouvré")

    log.info("✅  Aujourd'hui = premier jour ouvré (%s %s)",
             today, today.strftime("%A"))

    # ── Validation des variables d'env ───────────────────────────────
    if SEND_TO_WEBHOOK and not WEBHOOK_URL:
        log.error("❌  SEND_TO_WEBHOOK=true mais WEBHOOK_URL est vide.")
        log.error("    Configure le secret GitHub : Settings → Secrets → WEBHOOK_URL")
        sys.exit(1)

    log.info("  TEST_MODE=%s  ·  N_TOP=%d  ·  N_TOP_VIDEO=%d  ·  N_SECTOR=%d",
             TEST_MODE, N_TOP, N_TOP_VIDEO, N_SECTOR)
    log.info("  SEND_TO_WEBHOOK=%s  ·  LITTERBOX_EXPIRATION=%s",
             SEND_TO_WEBHOOK, LITTERBOX_EXPIRATION)
    log.info("  VIDÉO   : %dx%d · CRF=%d · preset=%s · fade=%.1fs/%.1fs · vignette=%s",
             VIDEO_W, VIDEO_H, VIDEO_CRF, VIDEO_PRESET, VIDEO_FADE_IN, VIDEO_FADE_OUT, VIGNETTE)
    log.info("  AUDIO   : %s · %s · fade=%.1fs/%.1fs · volume=%.1f",
             ("music.mp3" if MUSIC_FILE.exists() else "auto-download / silence AAC"),
             AUDIO_BITRATE, AUDIO_FADE_IN, AUDIO_FADE_OUT, AUDIO_VOLUME)
    if TEST_MODE:
        log.info("  ⚠️  Mode TEST : data partielle (~%d tickers/univers)", N_TICKERS_TEST)
    if not SEND_TO_WEBHOOK:
        log.info("  ⏸  SEND_TO_WEBHOOK=false : pas d'upload litterbox ni de webhook")

    # 1. Setup
    ensure_chromium_and_ffmpeg()

    # 2. Data pipeline (df + benchmarks)
    df, benchmarks, snapshot, period, suffix = run_data_pipeline()
    period_fr     = to_fr_period(period)
    prev_month    = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
    prev_month_fr = to_fr_month_year(prev_month)

    # 3. Snapshot + Excel
    save_snapshot(df, suffix)
    xlsx_path = export_xlsx(df, snapshot, suffix)

    # 4. Rankings + Post LinkedIn (avec benchmarks)
    rk = Rankings(df, suffix, benchmarks=benchmarks)
    post = build_linkedin_post(rk, period_fr, prev_month_fr, snapshot)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    post_path = OUT_DIR / f"linkedin_post_{snapshot}{suffix}.txt"
    post_path.write_text(post, encoding="utf-8")
    log.info("\n📝  Post LinkedIn → %s  (%d caractères)", post_path, len(post))

    # 5. Vidéo MP4
    video_path = make_video(rk, snapshot, period_fr, suffix)

    # 6. Webhook
    video_url = send_webhook(video_path, post, snapshot, period_fr, rk)

    # 7. Récap final · LES 3 LIENS DE SORTIE
    banner("✅  BRIEF TERMINÉ · LIENS DE SORTIE", "═")
    log.info("")
    log.info("📊  XLSX           : %s", xlsx_path.absolute())
    log.info("📝  Post LinkedIn  : %s  (%d chars)", post_path.absolute(), len(post))
    log.info("🎬  Vidéo locale   : %s", video_path.absolute())
    if video_url:
        log.info("☁️  URL publique   : %s", video_url)
        log.info("    (litterbox · expire dans %s)", LITTERBOX_EXPIRATION)
    log.info("")


if __name__ == "__main__":
    main()
