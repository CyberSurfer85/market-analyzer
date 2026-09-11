import streamlit as st
import yfinance as yf
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
from pathlib import Path
import json
import requests
import warnings
warnings.filterwarnings('ignore')

# ─── WATCHLIST PERSISTENTE ───────────────────────────────────────────────────
WATCHLIST_FILE = Path(__file__).parent / "watchlist.json"

def load_watchlist() -> dict:
    """Carica watchlist a gruppi. Formato: {gruppo: [ticker, ...]}."""
    try:
        if WATCHLIST_FILE.exists():
            data = json.loads(WATCHLIST_FILE.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return {"Generale": data}
            return data
    except Exception:
        pass
    return {"Generale": []}

def save_watchlist(wl: dict):
    try:
        WATCHLIST_FILE.write_text(json.dumps(wl), encoding="utf-8")
    except Exception:
        pass

# File separato per le note
NOTES_FILE = Path(__file__).parent / "notes.json"

def load_notes() -> dict:
    """Carica note per ticker. Formato: {ticker: nota}."""
    try:
        if NOTES_FILE.exists():
            return json.loads(NOTES_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def save_notes(notes: dict):
    try:
        NOTES_FILE.write_text(json.dumps(notes, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def total_tickers(wl: dict) -> int:
    return sum(len(v) for v in wl.values())

@st.cache_data(ttl=900)
def is_usa_ticker(symbol: str) -> bool:
    """Verifica se il ticker è USA — nessun suffisso di borsa estera."""
    return '.' not in symbol and '=' not in symbol

def fetch_dividend_info(symbol: str) -> dict:
    """Recupera dati dividendo — solo per titoli USA."""
    if not is_usa_ticker(symbol):
        return {"yield": "n/d", "rate": "n/d", "ex_date": "n/d", "freq": "n/d"}
    try:
        info = yf.Ticker(symbol).info
        yield_val  = info.get("dividendYield")
        rate       = info.get("dividendRate")
        ex_date    = info.get("exDividendDate")
        freq       = info.get("dividendFrequency") or info.get("payoutFrequency")

        # Converti timestamp ex-date
        ex_str = "—"
        if ex_date:
            try:
                import datetime
                ex_str = datetime.datetime.fromtimestamp(ex_date).strftime("%d/%m/%Y")
            except Exception:
                ex_str = "—"

        # Frequenza pagamento
        freq_map = {1: "Annuale", 2: "Semestrale", 4: "Trimestrale", 12: "Mensile"}
        freq_str = freq_map.get(freq, "—")

        # Yahoo Finance restituisce yield gia' in percentuale (es. 0.81 = 0.81%)
        # Valori > 25 sono quasi certamente anomali
        yield_str = "—"
        if yield_val and yield_val > 0:
            if yield_val <= 25:  # max 25% yield plausibile
                yield_str = f"{yield_val:.2f}%"
            else:
                yield_str = "⚠️ dato anomalo"

        return {
            "yield":   yield_str,
            "rate":    f"${rate:.2f}" if rate else "—",
            "ex_date": ex_str,
            "freq":    freq_str,
        }
    except Exception:
        return {"yield": "—", "rate": "—", "ex_date": "—", "freq": "—"}

st.set_page_config(
    page_title="Market Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── DISCLAIMER MODALE AD OGNI AVVIO ────────────────────────────────────────
if "disclaimer_accepted" not in st.session_state:
    st.session_state.disclaimer_accepted = False

if not st.session_state.disclaimer_accepted:
    st.markdown("## ⚠️ Disclaimer — Leggere prima di procedere")
    st.markdown("---")
    st.markdown(
        "Questo strumento fornisce esclusivamente **dati e indicatori tecnici** "
        "a scopo informativo e di ricerca personale.\n\n"
        "**Non costituisce** consulenza finanziaria, raccomandazione di investimento "
        "né sollecitazione all\'acquisto o alla vendita di strumenti finanziari "
        "ai sensi della **Direttiva MiFID II (2014/65/UE)** e del **D.Lgs. 58/1998 (TUF)**.\n\n"
        "**Limitazioni dei dati:**\n"
        "- Prezzi con possibile ritardo di 15-20 minuti\n"
        "- Dati fondamentali aggiornati con frequenza trimestrale\n"
        "- I rendimenti passati non costituiscono garanzia di risultati futuri\n\n"
        "**Le decisioni di investimento sono di esclusiva responsabilità dell\'utente.** "
        "L\'utilizzo di questo strumento implica l\'accettazione integrale delle presenti condizioni."
    )
    st.markdown("---")
    st.caption("Market Analyzer — Strumento di analisi dati a uso personale e informativo")
    col1, col2, col3, col4, col5 = st.columns([1, 2, 0.3, 2, 1])
    with col2:
        if st.button("✅ Accetto — Entra nel sistema", use_container_width=True, type="primary"):
            st.session_state.disclaimer_accepted = True
            st.rerun()
    with col4:
        if st.button("❌ Non accetto — Chiudi", use_container_width=True, type="secondary"):
            st.markdown(
                "<meta http-equiv='refresh' content='0; url=about:blank'>",
                unsafe_allow_html=True
            )
            st.info("Puoi chiudere questa scheda del browser.")
            st.stop()
    st.stop()

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Converte DataFrame in bytes Excel scaricabile."""
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dati")
    return buf.getvalue()

# ─── UNIVERSE ────────────────────────────────────────────────────────────────

UNIVERSE = {
    "🇺🇸 Indici USA": {
        "S&P 500":       "^GSPC",
        "Nasdaq 100":    "^NDX",
        "NYSE Comp.":    "^NYA",
        "QQQ ETF":       "QQQ",
        "Dow Jones":     "^DJI",
        "VIX":           "^VIX",
    },
    "🇪🇺 Indici Europa": {
        "DAX":           "^GDAXI",
        "CAC 40":        "^FCHI",
        "Eurostoxx 50":  "^STOXX50E",
        "FTSE MIB":      "FTSEMIB.MI",
        "IBEX 35":       "^IBEX",
        "FTSE 100":      "^FTSE",
    },
    "📦 ETF Globali": {
        "MSCI World (IWDA)":  "IWDA.AS",
        "S&P 500 (SPY)":      "SPY",
        "Emerging Markets":   "EEM",
        "Europe (VGK)":       "VGK",
        "Bond USA 20Y":       "TLT",
        "Bond Euro (IEGA)":   "IEGA.AS",
        "MSCI Italy":         "EWI",
    },
    "🏗️ Materie Prime": {
        "Oro":           "GC=F",
        "Argento":       "SI=F",
        "Petrolio WTI":  "CL=F",
        "Gas Naturale":  "NG=F",
        "Rame":          "HG=F",
        "Grano":         "ZW=F",
    },
    "🇮🇹 Titoli FTSE MIB": {
        "ENI":           "ENI.MI",
        "Enel":          "ENEL.MI",
        "Intesa SP":     "ISP.MI",
        "Unicredit":     "UCG.MI",
        "Stellantis":    "STLAM.MI",
        "Ferrari":       "RACE.MI",
        "Mediobanca":    "MB.MI",
        "Generali":      "G.MI",
    },
    "🇺🇸 Titoli USA Blue Chip": {
        "Apple":         "AAPL",
        "Microsoft":     "MSFT",
        "Nvidia":        "NVDA",
        "Amazon":        "AMZN",
        "Alphabet":      "GOOGL",
        "Berkshire B":   "BRK-B",
        "JPMorgan":      "JPM",
        "Johnson & J.":  "JNJ",
    },
}

PERIODS = {
    "1 mese":   ("1mo", "1d"),
    "3 mesi":   ("3mo", "1d"),
    "6 mesi":   ("6mo", "1d"),
    "1 anno":   ("1y",  "1d"),
    "2 anni":   ("2y",  "1wk"),
    "5 anni":   ("5y",  "1wk"),
}

# ─── HELPERS ─────────────────────────────────────────────────────────────────

@st.cache_data(ttl=900)
def fetch_ticker(symbol: str, period: str, interval: str) -> pd.DataFrame:
    # Primo tentativo — metodo standard
    try:
        t = yf.Ticker(symbol)
        df = t.history(period=period, interval=interval, auto_adjust=True)
        if not df.empty:
            return df
    except Exception:
        pass
    # Secondo tentativo — download diretto (piu robusto su alcuni ticker europei)
    try:
        df = yf.download(symbol, period=period, interval=interval,
                         auto_adjust=True, progress=False, show_errors=False)
        if not df.empty:
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            return df
    except Exception:
        pass
    return pd.DataFrame()

@st.cache_data(ttl=900)
def fetch_info(symbol: str) -> dict:
    try:
        return yf.Ticker(symbol).info
    except Exception:
        return {}

def fetch_fred(series_id: str, limit: int = 60) -> pd.DataFrame:
    """Scarica serie storiche da FRED — senza cache per massima compatibilità."""
    import io
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        date_col = "observation_date" if "observation_date" in df.columns else "DATE"
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.set_index(date_col)
        df.columns = ["value"]
        df = df.replace(".", float("nan"))
        df["value"] = pd.to_numeric(df["value"], errors="coerce")
        df = df.dropna()
        return df.tail(limit)
    except Exception as e:
        st.error(f"Errore FRED {series_id}: {type(e).__name__}: {e}")
        return pd.DataFrame()

def compute_signals(df: pd.DataFrame) -> dict:
    if df.empty or len(df) < 20:
        return {}
    close = df["Close"]
    vol   = df["Volume"]

    ma50  = close.rolling(50).mean().iloc[-1]  if len(df) >= 50  else None
    ma200 = close.rolling(200).mean().iloc[-1] if len(df) >= 200 else None
    price = close.iloc[-1]
    prev  = close.iloc[-2]

    # RSI 14
    delta = close.diff()
    gain  = delta.clip(lower=0).rolling(14).mean()
    loss  = (-delta.clip(upper=0)).rolling(14).mean()
    rs    = gain / loss.replace(0, float('nan'))
    rsi   = (100 - 100 / (1 + rs)).iloc[-1]

    # Trend semplice
    trend = "→ Laterale"
    if ma50 and ma200:
        if price > ma50 > ma200:
            trend = "↑ Rialzista"
        elif price < ma50 < ma200:
            trend = "↓ Ribassista"

    # Volume anomalo (vs media 20gg)
    vol_avg = vol.rolling(20).mean().iloc[-1] if not vol.empty else 0
    vol_ratio = (vol.iloc[-1] / vol_avg) if vol_avg > 0 else 1

    # 52 settimane
    high52 = close.max()
    low52  = close.min()
    pct_from_high = (price - high52) / high52 * 100
    pct_from_low  = (price - low52)  / low52  * 100

    # Variazione giornaliera
    daily_chg = (price - prev) / prev * 100

    return {
        "price":          price,
        "daily_chg":      daily_chg,
        "trend":          trend,
        "ma50":           ma50,
        "ma200":          ma200,
        "rsi":            rsi,
        "vol_ratio":      vol_ratio,
        "high52":         high52,
        "low52":          low52,
        "pct_from_high":  pct_from_high,
        "pct_from_low":   pct_from_low,
    }

def score_signal(s: dict) -> tuple[str, str]:
    """Returns (label, color) for overall signal."""
    if not s:
        return "N/D", "gray"
    points = 0
    if "↑" in s["trend"]:  points += 2
    if "↓" in s["trend"]:  points -= 2
    if s["rsi"] < 30:       points += 1
    if s["rsi"] > 70:       points -= 1
    if s["vol_ratio"] > 1.5: points += 1 if "↑" in s["trend"] else -1
    if points >= 2:   return "🟢 Favorevole",   "#2ecc71"
    if points <= -2:  return "🔴 Sfavorevole",  "#e74c3c"
    return "🟡 Neutro", "#f39c12"

def fmt_price(v):
    if v is None or (isinstance(v, float) and v != v):
        return "—"
    if abs(v) > 1000:
        return f"{v:,.0f}"
    if abs(v) > 10:
        return f"{v:.2f}"
    return f"{v:.4f}"

def fmt_pct(v):
    if v is None or (isinstance(v, float) and v != v):
        return "—"
    arrow = "▲" if v > 0 else "▼" if v < 0 else "—"
    color = "green" if v > 0 else "red" if v < 0 else "gray"
    return f":{color}[{arrow} {abs(v):.2f}%]"

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("⚙️ Impostazioni")
    period_label = st.selectbox("Periodo analisi", list(PERIODS.keys()), index=3)
    period, interval = PERIODS[period_label]

    st.markdown("---")
    st.subheader("⭐ Watchlist personale")
    st.caption("Max 20 titoli per gruppo. Salvata automaticamente.")

    watchlist = load_watchlist()
    gruppi = list(watchlist.keys())

    # Crea / elimina gruppi
    with st.expander("⚙️ Gestisci gruppi"):
        with st.form("form_nuovo_gruppo", clear_on_submit=True):
            nuovo_gruppo = st.text_input("Nuovo gruppo", placeholder="es. Dividendi USA")
            if st.form_submit_button("➕ Crea gruppo") and nuovo_gruppo.strip():
                nome = nuovo_gruppo.strip()
                if nome not in watchlist:
                    watchlist[nome] = []
                    save_watchlist(watchlist)
                    st.rerun()
        if len(gruppi) > 1:
            gruppo_da_eliminare = st.selectbox("Elimina gruppo", gruppi, key="del_gruppo")
            if st.button("🗑️ Elimina gruppo selezionato"):
                del watchlist[gruppo_da_eliminare]
                save_watchlist(watchlist)
                st.rerun()

    # Aggiungi ticker — seleziona gruppo di destinazione + ticker
    # Mantieni gruppo selezionato tra i rerun
    # Selettore gruppo persistente — indipendente dal form
    if "sidebar_gruppo_dest" not in st.session_state:
        st.session_state.sidebar_gruppo_dest = gruppi[0] if gruppi else "Generale"
    if st.session_state.sidebar_gruppo_dest not in gruppi:
        st.session_state.sidebar_gruppo_dest = gruppi[0]

    idx_dest = gruppi.index(st.session_state.sidebar_gruppo_dest)
    gruppo_dest_sel = st.selectbox("Aggiungi a gruppo", gruppi, index=idx_dest, key="gruppo_dest_outer")
    st.session_state.sidebar_gruppo_dest = gruppo_dest_sel

    with st.form("wl_form", clear_on_submit=True):
        new_ticker = st.text_input(f"Ticker per '{gruppo_dest_sel}'", key="wl_input").strip().upper()
        submitted = st.form_submit_button("➕ Aggiungi")
        if submitted and new_ticker:
            if len(watchlist.get(gruppo_dest_sel, [])) >= 20:
                st.warning(f"Limite di 20 titoli per gruppo raggiunto in '{gruppo_dest_sel}'.")
            elif new_ticker in watchlist.get(gruppo_dest_sel, []):
                st.warning(f"{new_ticker} già presente in '{gruppo_dest_sel}'.")
            else:
                watchlist[gruppo_dest_sel].append(new_ticker)
                save_watchlist(watchlist)
                st.rerun()

    # Riepilogo gruppi con rimozione ticker
    for grp in gruppi:
        tickers_grp = watchlist.get(grp, [])
        if tickers_grp:
            st.markdown(f"**{grp}** ({len(tickers_grp)})")
            for i, sym in enumerate(tickers_grp):
                col_sym, col_del = st.columns([4, 1])
                col_sym.markdown(f"`{sym}`")
                if col_del.button("✕", key=f"del_{grp}_{i}_{sym}"):
                    watchlist[grp].pop(i)
                    save_watchlist(watchlist)
                    st.rerun()

    st.caption(f"Totale: {total_tickers(watchlist)} titoli · Max 20 per gruppo")

    st.markdown("---")
    st.subheader("Ticker aggiuntivi (temporanei)")
    custom_raw = st.text_area(
        "Aggiungi ticker (uno per riga, es. TSLA)",
        height=80,
        help="Non salvati. Per analisi rapide. Usa la Watchlist per titoli fissi."
    )
    custom_tickers = [t.strip().upper() for t in custom_raw.splitlines() if t.strip()]

    st.markdown("---")
    show_fundamentals = st.toggle("Mostra fondamentali (solo azioni)", value=True)
    st.markdown("---")
    st.caption("Dati: Yahoo Finance · Ritardo 15-20 min · Aggiornamento ogni 15 min")
    if st.button("🔄 Aggiorna dati"):
        st.cache_data.clear()
        st.rerun()

# ─── MAIN ────────────────────────────────────────────────────────────────────

st.title("📊 Market Analyzer")
st.caption(f"Analisi al {datetime.now().strftime('%d/%m/%Y %H:%M')} · Periodo: {period_label}")

tabs = st.tabs(["⭐ Watchlist", "📋 Scanner", "📈 Grafico & Dettaglio", "🏛️ Fondamentali", "📝 Note", "ℹ️ Guida"])

# ────────────────────────────────────────────────────
# TAB 0 — WATCHLIST
# ────────────────────────────────────────────────────
with tabs[0]:
    wl = load_watchlist()
    tutti_ticker = total_tickers(wl)

    if tutti_ticker == 0:
        st.info("La tua watchlist è vuota. Aggiungi titoli e gruppi dalla barra laterale.")
    else:
        gruppi_disponibili = list(wl.keys())
        # Filtra gruppi non vuoti per il selettore — evita banner "gruppo vuoto" durante refresh
        gruppi_con_ticker = [g for g in gruppi_disponibili if wl.get(g)]
        if not gruppi_con_ticker:
            st.info("Tutti i gruppi sono vuoti. Aggiungi titoli dalla barra laterale.")
        else:
            # Mantieni gruppo selezionato tra i refresh
            if "tab_gruppo_vista" not in st.session_state:
                st.session_state.tab_gruppo_vista = gruppi_con_ticker[0]
            if st.session_state.tab_gruppo_vista not in gruppi_con_ticker:
                st.session_state.tab_gruppo_vista = gruppi_con_ticker[0]

            idx_vista = gruppi_con_ticker.index(st.session_state.tab_gruppo_vista)

            col_sel, col_info = st.columns([2, 3])
            with col_sel:
                gruppo_vista = st.selectbox(
                    "Visualizza gruppo",
                    gruppi_con_ticker,
                    index=idx_vista,
                    key="gruppo_vista"
                )
                st.session_state.tab_gruppo_vista = gruppo_vista
            with col_info:
                n_gruppo = len(wl.get(gruppo_vista, []))
                st.caption(f"**{gruppo_vista}** — {n_gruppo} titoli · Totale watchlist: {tutti_ticker}/20")

            mostra_dividendi = st.toggle("💰 Mostra dati dividendo (solo azioni USA)", value=False)

            ticker_gruppo = wl.get(gruppo_vista, [])

            if not ticker_gruppo:
                st.info(f"Il gruppo '{gruppo_vista}' è vuoto. Aggiungi titoli dalla barra laterale.")
            else:
                import time
                wl_rows = []
                progress_bar = st.progress(0, text=f"Caricamento {gruppo_vista}...")

                for i, sym in enumerate(ticker_gruppo):
                    progress_bar.progress((i + 1) / len(ticker_gruppo), text=f"Caricamento {sym}...")
                    df = fetch_ticker(sym, period, interval)
                    if df.empty:
                        time.sleep(0.3)
                        df = fetch_ticker(sym, period, interval)
                    s = compute_signals(df)
                    label, _ = score_signal(s)
                    div_info = fetch_dividend_info(sym) if mostra_dividendi else {}
                    wl_rows.append({
                        "Simbolo":   sym,
                        "Prezzo":    fmt_price(s.get("price")) if s else "—",
                        "Var. %":    s.get("daily_chg") if s else None,
                        "Trend":     s.get("trend", "—") if s else "—",
                        "RSI(14)":   round(s["rsi"], 1) if s and s.get("rsi") == s.get("rsi") else None,
                        "Vol×":      round(s.get("vol_ratio", 1), 2) if s else None,
                        "% da Max":  round(s.get("pct_from_high", 0), 1) if s else None,
                        "Segnale":   label,
                        "Div.Yield": div_info.get("yield", "—"),
                        "Div.Rate":  div_info.get("rate", "—"),
                        "Frequenza": div_info.get("freq", "—"),
                    })

                progress_bar.empty()

                # Summary gruppo
                fav  = sum(1 for r in wl_rows if "🟢" in r["Segnale"])
                neu  = sum(1 for r in wl_rows if "🟡" in r["Segnale"])
                sfav = sum(1 for r in wl_rows if "🔴" in r["Segnale"])
                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("Titoli nel gruppo", n_gruppo)
                mc2.metric("🟢 Favorevoli", fav)
                mc3.metric("🟡 Neutri", neu)
                mc4.metric("🔴 Sfavorevoli", sfav)

                st.caption(f"Aggiornamento al {datetime.now().strftime('%d/%m/%Y %H:%M')}")

                # Export Watchlist
                export_wl = [{
                    "Gruppo":    gruppo_vista,
                    "Simbolo":   r["Simbolo"],
                    "Prezzo":    r["Prezzo"],
                    "Var. %":    f"{r['Var. %']:.2f}" if r.get("Var. %") is not None else "—",
                    "Trend":     r["Trend"],
                    "RSI(14)":   r["RSI(14)"],
                    "Vol×":      r["Vol×"],
                    "% da Max":  r["% da Max"],
                    "Segnale":   r["Segnale"],
                } for r in wl_rows]
                from datetime import date
                wl_filename = f"watchlist_{gruppo_vista}_{date.today().strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    label="📥 Scarica Watchlist (Excel)",
                    data=df_to_excel_bytes(pd.DataFrame(export_wl)),
                    file_name=wl_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_watchlist"
                )

                st.markdown("---")

                # Intestazione tabella
                if mostra_dividendi:
                    hcols = st.columns([1.2, 0.8, 0.8, 1, 0.7, 0.7, 0.8, 0.9, 0.9, 1.2])
                    headers = ["Simbolo", "Prezzo", "Var.%", "Trend", "RSI", "Vol×", "%Max", "Yield", "Rate", "Segnale"]
                else:
                    hcols = st.columns([1.5, 1, 1, 1.2, 0.8, 0.8, 1.2, 1.5])
                    headers = ["Simbolo", "Prezzo", "Var. %", "Trend", "RSI", "Vol×", "% da Max", "Segnale"]

                for hc, hl in zip(hcols, headers):
                    hc.markdown(f"**{hl}**")

                for row in wl_rows:
                    if mostra_dividendi:
                        cols = st.columns([1.2, 0.8, 0.8, 1, 0.7, 0.7, 0.8, 0.9, 0.9, 1.2])
                    else:
                        cols = st.columns([1.5, 1, 1, 1.2, 0.8, 0.8, 1.2, 1.5])

                    cols[0].markdown(f"`{row['Simbolo']}`")
                    cols[1].markdown(row["Prezzo"])

                    var = row["Var. %"]
                    try:
                        v = float(var) if var is not None else float('nan')
                        if v == v:
                            arrow = "▲" if v > 0 else "▼" if v < 0 else "—"
                            color = "green" if v > 0 else "red" if v < 0 else "gray"
                            cols[2].markdown(f":{color}[{arrow} {abs(v):.2f}%]")
                        else:
                            cols[2].markdown("—")
                    except Exception:
                        cols[2].markdown("—")

                    trend = row["Trend"]
                    t_color = "green" if "↑" in trend else "red" if "↓" in trend else "orange"
                    cols[3].markdown(f":{t_color}[{trend}]")

                    rsi = row["RSI(14)"]
                    if rsi:
                        rsi_color = "red" if rsi > 70 else "green" if rsi < 30 else "gray"
                        cols[4].markdown(f":{rsi_color}[{rsi}]")
                    else:
                        cols[4].markdown("—")

                    vol = row["Vol×"]
                    cols[5].markdown(f"{'🔥' if vol and vol > 2 else ''}{vol if vol else '—'}")

                    pfh = row["% da Max"]
                    try:
                        p = float(pfh) if pfh is not None else float('nan')
                        if p == p:
                            cols[6].markdown(f":{'green' if p > -5 else 'orange' if p > -20 else 'red'}[{p:.1f}%]")
                        else:
                            cols[6].markdown("—")
                    except Exception:
                        cols[6].markdown("—")

                    if mostra_dividendi:
                        cols[7].markdown(row["Div.Yield"])
                        cols[8].markdown(row["Div.Rate"])
                        cols[9].markdown(row["Segnale"])
                    else:
                        cols[7].markdown(row["Segnale"])

# TAB 1 — SCANNER
# ────────────────────────────────────────────────────
with tabs[1]:
    all_rows = []

    for category, assets in UNIVERSE.items():
        for name, symbol in assets.items():
            df = fetch_ticker(symbol, period, interval)
            s  = compute_signals(df)
            label, color = score_signal(s)
            all_rows.append({
                "Categoria":  category,
                "Asset":      name,
                "Simbolo":    symbol,
                "Prezzo":     fmt_price(s.get("price")) if s else "—",
                "Var. %":     s.get("daily_chg") if s else None,
                "Trend":      s.get("trend", "—") if s else "—",
                "RSI(14)":    round(s["rsi"], 1) if s and s.get("rsi") == s.get("rsi") else None,
                "Vol×":       round(s.get("vol_ratio", 1), 2) if s else None,
                "% da Max":   round(s.get("pct_from_high", 0), 1) if s else None,
                "Segnale":    label,
                "_color":     color,
                "_s":         s,
            })

    # Custom tickers
    for sym in custom_tickers:
        df = fetch_ticker(sym, period, interval)
        s  = compute_signals(df)
        label, color = score_signal(s)
        all_rows.append({
            "Categoria": "⭐ Personalizzati",
            "Asset":     sym,
            "Simbolo":   sym,
            "Prezzo":    fmt_price(s.get("price")) if s else "—",
            "Var. %":    s.get("daily_chg") if s else None,
            "Trend":     s.get("trend", "—") if s else "—",
            "RSI(14)":   round(s["rsi"], 1) if s and s.get("rsi") == s.get("rsi") else None,
            "Vol×":      round(s.get("vol_ratio", 1), 2) if s else None,
            "% da Max":  round(s.get("pct_from_high", 0), 1) if s else None,
            "Segnale":   label,
            "_color":    color,
            "_s":        s,
        })

    # Filtri
    col1, col2, col3 = st.columns(3)
    with col1:
        cat_filter = st.multiselect("Filtra categoria", list(UNIVERSE.keys()) + ["⭐ Personalizzati"], default=[])
    with col2:
        sig_filter = st.multiselect("Filtra segnale", ["🟢 Favorevole", "🟡 Neutro", "🔴 Sfavorevole"], default=[])
    with col3:
        trend_filter = st.multiselect("Filtra trend", ["↑ Rialzista", "→ Laterale", "↓ Ribassista"], default=[])

    filtered = all_rows
    if cat_filter:   filtered = [r for r in filtered if r["Categoria"] in cat_filter]
    if sig_filter:   filtered = [r for r in filtered if r["Segnale"] in sig_filter]
    if trend_filter: filtered = [r for r in filtered if any(t in r["Trend"] for t in trend_filter)]

    # Summary cards
    favorevoli = sum(1 for r in filtered if "🟢" in r["Segnale"])
    neutri     = sum(1 for r in filtered if "🟡" in r["Segnale"])
    sfavorevoli= sum(1 for r in filtered if "🔴" in r["Segnale"])

    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Asset analizzati", len(filtered))
    mc2.metric("🟢 Favorevoli",    favorevoli)
    mc3.metric("🟡 Neutri",        neutri)
    mc4.metric("🔴 Sfavorevoli",   sfavorevoli)

    # Export Scanner
    if filtered:
        export_data = [{
            "Categoria": r["Categoria"],
            "Asset":     r["Asset"],
            "Simbolo":   r["Simbolo"],
            "Prezzo":    r["Prezzo"],
            "Var. %":    f"{r['Var. %']:.2f}" if r.get("Var. %") is not None else "—",
            "Trend":     r["Trend"],
            "RSI(14)":   r["RSI(14)"],
            "Vol×":      r["Vol×"],
            "% da Max":  r["% da Max"],
            "Segnale":   r["Segnale"],
        } for r in filtered]
        export_df = pd.DataFrame(export_data)
        from datetime import date
        filename = f"scanner_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="📥 Scarica Scanner (Excel)",
            data=df_to_excel_bytes(export_df),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")

    # Tabella
    for category in (set(r["Categoria"] for r in filtered)):
        cat_rows = [r for r in filtered if r["Categoria"] == category]
        if not cat_rows:
            continue
        st.subheader(category)

        header_cols = st.columns([2.5, 1, 1, 1, 1, 1, 1.5, 1.5])
        headers = ["Asset", "Prezzo", "Var. %", "Trend", "RSI", "Vol×", "% da Max", "Segnale"]
        for hc, hl in zip(header_cols, headers):
            hc.markdown(f"**{hl}**")

        for row in cat_rows:
            cols = st.columns([2.5, 1, 1, 1, 1, 1, 1.5, 1.5])
            cols[0].markdown(f"`{row['Simbolo']}` {row['Asset']}")
            cols[1].markdown(row["Prezzo"])

            var = row["Var. %"]
            if var is not None:
                arrow = "▲" if var > 0 else "▼" if var < 0 else "—"
                color = "green" if var > 0 else "red" if var < 0 else "gray"
                cols[2].markdown(f":{color}[{arrow} {abs(var):.2f}%]")
            else:
                cols[2].markdown("—")

            trend = row["Trend"]
            t_color = "green" if "↑" in trend else "red" if "↓" in trend else "orange"
            cols[3].markdown(f":{t_color}[{trend}]")

            rsi = row["RSI(14)"]
            if rsi:
                rsi_color = "red" if rsi > 70 else "green" if rsi < 30 else "gray"
                cols[4].markdown(f":{rsi_color}[{rsi}]")
            else:
                cols[4].markdown("—")

            vol = row["Vol×"]
            cols[5].markdown(f"{'🔥' if vol and vol > 2 else ''}{vol if vol else '—'}")

            pfh = row["% da Max"]
            if pfh is not None:
                cols[6].markdown(f":{'green' if pfh > -5 else 'orange' if pfh > -20 else 'red'}[{pfh:.1f}%]")
            else:
                cols[6].markdown("—")

            cols[7].markdown(row["Segnale"])

        st.markdown("")

# ────────────────────────────────────────────────────
# TAB 2 — GRAFICO & DETTAGLIO
# ────────────────────────────────────────────────────
with tabs[2]:
    all_symbols = {}
    for cat, assets in UNIVERSE.items():
        for name, sym in assets.items():
            all_symbols[f"{name} ({sym})"] = sym
    for sym in custom_tickers:
        all_symbols[sym] = sym

    selected_label = st.selectbox("Seleziona asset", list(all_symbols.keys()))
    selected_sym   = all_symbols[selected_label]

    df = fetch_ticker(selected_sym, period, interval)

    if df.empty:
        st.error(f"Impossibile caricare dati per {selected_sym}. Verifica il simbolo.")
    else:
        s = compute_signals(df)
        signal_label, _ = score_signal(s)

        # KPI row
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Prezzo",     fmt_price(s.get("price")))
        var = s.get("daily_chg", 0)
        k2.metric("Var. giorno", f"{var:+.2f}%" if var else "—")
        k3.metric("Trend",       s.get("trend", "—"))
        k4.metric("RSI (14)",    f"{s['rsi']:.1f}" if s.get("rsi") == s.get("rsi") else "—")
        k5.metric("Segnale",     signal_label)

        st.markdown("---")

        # Opzioni grafico
        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            usa_candele = st.toggle("🕯️ Candele giapponesi", value=True)
        with col_opt2:
            mostra_bollinger = st.toggle("📊 Bande di Bollinger", value=True)

        # Chart: candele/linea + MA + Bollinger + volume
        close = df["Close"]
        fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                            row_heights=[0.72, 0.28], vertical_spacing=0.04)

        if usa_candele and "Open" in df.columns and "High" in df.columns and "Low" in df.columns:
            fig.add_trace(go.Candlestick(
                x=df.index,
                open=df["Open"], high=df["High"],
                low=df["Low"],   close=df["Close"],
                name="Prezzo",
                increasing_line_color="#2ecc71",
                decreasing_line_color="#e74c3c",
                increasing_fillcolor="#2ecc71",
                decreasing_fillcolor="#e74c3c",
            ), row=1, col=1)
        else:
            fig.add_trace(go.Scatter(
                x=df.index, y=close,
                name="Prezzo", line=dict(color="#4a90d9", width=2)
            ), row=1, col=1)

        if len(df) >= 50:
            ma50_line = close.rolling(50).mean()
            fig.add_trace(go.Scatter(
                x=df.index, y=ma50_line,
                name="MA 50", line=dict(color="#f39c12", width=1.5, dash="dot")
            ), row=1, col=1)

        if len(df) >= 200:
            ma200_line = close.rolling(200).mean()
            fig.add_trace(go.Scatter(
                x=df.index, y=ma200_line,
                name="MA 200", line=dict(color="#e74c3c", width=1.5, dash="dash")
            ), row=1, col=1)

        # Bande di Bollinger (20 periodi, 2 deviazioni standard)
        if mostra_bollinger and len(df) >= 20:
            bb_period = 20
            bb_std    = 2
            bb_ma  = close.rolling(bb_period).mean()
            bb_std_val = close.rolling(bb_period).std()
            bb_upper = bb_ma + bb_std * bb_std_val
            bb_lower = bb_ma - bb_std * bb_std_val

            fig.add_trace(go.Scatter(
                x=df.index, y=bb_upper,
                name="BB Superiore", line=dict(color="#8e44ad", width=1, dash="dot"),
                showlegend=True
            ), row=1, col=1)
            fig.add_trace(go.Scatter(
                x=df.index, y=bb_lower,
                name="BB Inferiore", line=dict(color="#8e44ad", width=1, dash="dot"),
                fill="tonexty", fillcolor="rgba(142,68,173,0.06)",
                showlegend=True
            ), row=1, col=1)
            fig.add_trace(go.Scatter(
                x=df.index, y=bb_ma,
                name="BB Media", line=dict(color="#8e44ad", width=1),
                showlegend=True
            ), row=1, col=1)

        colors_vol = ["#2ecc71" if c >= o else "#e74c3c"
                      for c, o in zip(df["Close"], df["Open"])]
        fig.add_trace(go.Bar(
            x=df.index, y=df["Volume"],
            name="Volume", marker_color=colors_vol, opacity=0.6
        ), row=2, col=1)

        fig.update_layout(
            height=560,
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", y=1.05),
            margin=dict(l=0, r=0, t=40, b=0),
            title=f"{selected_label} · {period_label}",
            xaxis_rangeslider_visible=False,
        )
        fig.update_yaxes(gridcolor="rgba(255,255,255,0.06)")
        fig.update_xaxes(gridcolor="rgba(255,255,255,0.04)")
        st.plotly_chart(fig, use_container_width=True)

        # RSI chart
        delta = close.diff()
        gain  = delta.clip(lower=0).rolling(14).mean()
        loss  = (-delta.clip(upper=0)).rolling(14).mean()
        rs    = gain / loss.replace(0, float('nan'))
        rsi_series = 100 - 100 / (1 + rs)

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=df.index, y=rsi_series,
            name="RSI 14", line=dict(color="#9b59b6", width=1.8)
        ))
        fig2.add_hline(y=70, line_dash="dot", line_color="#e74c3c", annotation_text="Ipercomprato 70")
        fig2.add_hline(y=30, line_dash="dot", line_color="#2ecc71", annotation_text="Ipervenduto 30")
        fig2.update_layout(
            height=180, template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=0, r=0, t=20, b=0), showlegend=False,
            yaxis=dict(range=[0, 100]),
        )
        st.plotly_chart(fig2, use_container_width=True)

        # Dettaglio segnali
        with st.expander("📊 Dettaglio segnali tecnici"):
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("**Medie mobili**")
                if s.get("ma50"):   st.markdown(f"- MA 50:  `{fmt_price(s['ma50'])}`")
                if s.get("ma200"):  st.markdown(f"- MA 200: `{fmt_price(s['ma200'])}`")
                st.markdown(f"- Prezzo attuale: `{fmt_price(s.get('price'))}`")
            with d2:
                st.markdown("**Range periodo**")
                st.markdown(f"- Massimo:     `{fmt_price(s.get('high52'))}`")
                st.markdown(f"- Minimo:      `{fmt_price(s.get('low52'))}`")
                st.markdown(f"- Dist. max:   `{s.get('pct_from_high', 0):.1f}%`")
                st.markdown(f"- Dist. min:   `{s.get('pct_from_low', 0):.1f}%`")

        # ── NOTIZIE TICKER ────────────────────────────────
        st.markdown("---")
        st.markdown("#### 📰 Ultime notizie")
        st.caption("Fonte: Yahoo Finance · Le notizie influenzano i movimenti di breve termine.")
        try:
            ticker_obj = yf.Ticker(selected_sym)
            news = ticker_obj.news
            if news:
                for item in news[:8]:
                    content_item = item.get("content", {}) if isinstance(item, dict) else {}
                    titolo = content_item.get("title", item.get("title", "")) if isinstance(item, dict) else ""
                    summary = content_item.get("summary", "")
                    pub_date = content_item.get("pubDate", "")
                    data_str = ""
                    if pub_date:
                        try:
                            dt = datetime.strptime(pub_date[:10], "%Y-%m-%d")
                            data_str = dt.strftime("%d/%m/%Y")
                        except Exception:
                            data_str = pub_date[:10] if len(pub_date) >= 10 else ""
                    click_url = content_item.get("clickThroughUrl", {})
                    url = click_url.get("url", "") if isinstance(click_url, dict) else ""
                    if titolo:
                        if url:
                            st.markdown(f"**[{titolo}]({url})**")
                        else:
                            st.markdown(f"**{titolo}**")
                        if data_str:
                            st.caption(data_str)
                        if summary:
                            st.markdown(f"{summary[:200]}{'...' if len(summary) > 200 else ''}")
                        st.markdown("")
            else:
                st.caption("Nessuna notizia disponibile per questo ticker.")
        except Exception:
            st.caption("Notizie non disponibili al momento.")

# ────────────────────────────────────────────────────
# TAB 3 — FONDAMENTALI
# ────────────────────────────────────────────────────
with tabs[3]:
    if not show_fundamentals:
        st.info("Abilita 'Mostra fondamentali' nella barra laterale.")
    else:
        st.subheader("Fondamentali — azioni singole")
        st.caption("Solo per azioni. ETF e indici non hanno multipli di valutazione.")

        equity_symbols = {}
        # Ticker personalizzati in cima — priorità all'analisi manuale
        for sym in custom_tickers:
            equity_symbols[f"⭐ {sym} (personalizzato)"] = sym
        for name, sym in UNIVERSE["🇮🇹 Titoli FTSE MIB"].items():
            equity_symbols[f"{name} ({sym})"] = sym
        for name, sym in UNIVERSE["🇺🇸 Titoli USA Blue Chip"].items():
            equity_symbols[f"{name} ({sym})"] = sym

        # Indice default: primo personalizzato se presente, altrimenti primo della lista
        default_idx = 0
        sel_eq = st.selectbox("Seleziona azione", list(equity_symbols.keys()),
                               index=default_idx, key="fund_sel")
        sym_eq = equity_symbols[sel_eq]

        with st.spinner("Caricamento dati fondamentali…"):
            info = fetch_info(sym_eq)

        if not info or info.get("regularMarketPrice") is None and info.get("currentPrice") is None:
            st.warning("Dati fondamentali non disponibili per questo simbolo.")
        else:
            def safe(key, fmt=None):
                v = info.get(key)
                if v is None or v != v:
                    return "—"
                if fmt == "pct":
                    return f"{v*100:.1f}%"
                if fmt == "B":
                    return f"${v/1e9:.1f}B" if v > 1e9 else f"${v/1e6:.0f}M"
                if isinstance(v, float):
                    return f"{v:.2f}"
                return str(v)

            col_a, col_b, col_c = st.columns(3)

            with col_a:
                st.markdown("##### Valutazione")
                st.markdown(f"P/E (ttm): **{safe('trailingPE')}**")
                st.markdown(f"P/E (forw): **{safe('forwardPE')}**")
                st.markdown(f"P/B: **{safe('priceToBook')}**")
                st.markdown(f"EV/EBITDA: **{safe('enterpriseToEbitda')}**")
                st.markdown(f"P/S: **{safe('priceToSalesTrailing12Months')}**")

            with col_b:
                st.markdown("##### Qualità")
                st.markdown(f"Marg. netto: **{safe('profitMargins', 'pct')}**")
                st.markdown(f"ROE: **{safe('returnOnEquity', 'pct')}**")
                st.markdown(f"ROA: **{safe('returnOnAssets', 'pct')}**")
                st.markdown(f"Debt/Equity: **{safe('debtToEquity')}**")
                st.markdown(f"Current ratio: **{safe('currentRatio')}**")

            with col_c:
                st.markdown("##### Crescita & Dividendo")
                st.markdown(f"Rev. growth: **{safe('revenueGrowth', 'pct')}**")
                st.markdown(f"EPS growth: **{safe('earningsGrowth', 'pct')}**")
                st.markdown(f"Div. yield: **{safe('dividendYield', 'pct')}**")
                st.markdown(f"Payout ratio: **{safe('payoutRatio', 'pct')}**")
                st.markdown(f"Market cap: **{safe('marketCap', 'B')}**")

            # Semaforo fondamentale grezzo
            st.markdown("---")
            st.markdown("##### Lettura rapida fondamentali")

            checks = []
            pe = info.get("trailingPE")
            pb = info.get("priceToBook")
            roe = info.get("returnOnEquity")
            debt = info.get("debtToEquity")
            div = info.get("dividendYield")

            if pe and 0 < pe < 25:  checks.append(("🟢", "P/E sotto 25 — valutazione non estrema"))
            elif pe and pe > 40:    checks.append(("🔴", f"P/E a {pe:.0f} — multiplo elevato"))
            else:                   checks.append(("🟡", "P/E nella fascia media o non disponibile"))

            if pb and pb < 3:       checks.append(("🟢", "P/B sotto 3 — non evidente sopravvalutazione"))
            elif pb and pb > 8:     checks.append(("🔴", "P/B elevato"))

            if roe and roe > 0.15:  checks.append(("🟢", f"ROE {roe*100:.1f}% — buona redditività del capitale"))
            elif roe and roe < 0:   checks.append(("🔴", "ROE negativo"))

            if debt and debt < 80:  checks.append(("🟢", "Debito/Equity contenuto"))
            elif debt and debt > 200: checks.append(("🔴", f"Debt/Equity {debt:.0f} — leva elevata"))

            if div and div > 0.02:  checks.append(("🟢", f"Dividendo {div*100:.1f}% — fonte di rendimento"))

            for icon, msg in checks:
                st.markdown(f"{icon} {msg}")

            st.caption("⚠️ Questi dati sono indicativi. Confronta sempre con il settore di riferimento.")

# ────────────────────────────────────────────────────
# TAB 5 — NOTE
# ────────────────────────────────────────────────────
with tabs[4]:
    st.subheader("📝 Note personali sui titoli")
    st.caption("Annota il motivo di interesse, il prezzo di carico, o qualsiasi considerazione personale. Salvate automaticamente.")

    notes = load_notes()

    # Recupera tutti i ticker dalla watchlist
    wl_all = load_watchlist()
    tutti_wl = sorted(set(sym for tickers in wl_all.values() for sym in tickers))

    if not tutti_wl:
        st.info("Aggiungi titoli alla watchlist per poter annotare note su di essi.")
    else:
        # Selettore ticker
        ticker_note = st.selectbox("Seleziona titolo", tutti_wl, key="note_ticker_sel")

        # Nota esistente
        nota_attuale = notes.get(ticker_note, "")

        # Editor nota
        st.markdown(f"**Nota per `{ticker_note}`**")
        nuova_nota = st.text_area(
            "Scrivi la tua nota",
            value=nota_attuale,
            height=150,
            placeholder="Es: Prezzo di carico €45.20 — acquistato 15/03/2026. Dividendo trimestrale atteso a giugno. Tenere fino a target €60.",
            key=f"nota_{ticker_note}"
        )

        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("💾 Salva nota", use_container_width=True):
                if nuova_nota.strip():
                    notes[ticker_note] = nuova_nota.strip()
                else:
                    notes.pop(ticker_note, None)
                save_notes(notes)
                st.success(f"Nota salvata per {ticker_note}")
        with col2:
            if nota_attuale and st.button("🗑️ Elimina nota", use_container_width=True):
                notes.pop(ticker_note, None)
                save_notes(notes)
                st.rerun()

        st.markdown("---")

        # Riepilogo tutte le note
        note_esistenti = {k: v for k, v in notes.items() if k in tutti_wl}
        if note_esistenti:
            st.markdown("#### Riepilogo note")
            for sym, nota in note_esistenti.items():
                with st.expander(f"`{sym}`"):
                    st.markdown(nota)
        else:
            st.caption("Nessuna nota salvata.")

# TAB 4 — GUIDA
# ────────────────────────────────────────────────────
with tabs[5]:
    st.subheader("Guida — Market Analyzer v17")

    st.markdown("""
**Logica del sistema**

Market Analyzer aggrega dati da Yahoo Finance (prezzi, fondamentali) e FRED-Federal Reserve
(dati macro USA) per supportare l'analisi di investimento a medio-lungo termine.

Non genera raccomandazioni di acquisto o vendita. Produce segnali tecnici, fondamentali
e macro per supportare la tua analisi — non per sostituirla.

---

### ⭐ Watchlist
Elenco personale di massimo 20 titoli salvato automaticamente.
- Aggiungi ticker dal campo in sidebar — funziona con Enter o col pulsante ➕
- Il campo non è case sensitive: `aapl`, `AAPL`, `Aapl` sono equivalenti
- Rimuovi titoli con il pulsante ✕ accanto a ciascuno
- I dati si caricano con barra di avanzamento titolo per titolo
- **Frequenza consigliata:** una volta al giorno su mercati aperti

---

### 📋 Scanner
Panoramica completa di tutti i mercati coperti.

| Indicatore | Cosa misura | Come usarlo |
|---|---|---|
| Trend (MA50/MA200) | Direzione primaria del prezzo | Rialzista = prezzo sopra entrambe le medie. Non entrare contro trend primario. |
| RSI 14 | Momentum | <30 = possibile zona di accumulazione. >70 = attenzione a nuovi ingressi. |
| Vol× | Volume vs media 20gg | >1.5 con trend rialzista = possibile accumulo istituzionale. |
| % da Max | Distanza dal massimo del periodo | Utile per stimare margine di recupero o rischio di ritracciamento. |

**Segnale complessivo:**
- 🟢 Favorevole: trend rialzista + RSI non estremo + conferma volume
- 🟡 Neutro: segnali contrastanti o lateralità
- 🔴 Sfavorevole: trend ribassista o RSI ipercomprato

**Filtri disponibili:** categoria, segnale, trend.

---

### 📈 Grafico & Dettaglio
- Grafico prezzo con MA50 (arancio) e MA200 (rosso)
- Volume colorato: verde = chiusura positiva, rosso = chiusura negativa
- RSI 14 in sottografico separato con soglie 30/70
- KPI sintetici: prezzo, variazione, trend, RSI, segnale

---

### 🏛️ Fondamentali
Solo per azioni singole. ETF e indici non hanno multipli di valutazione.

I ticker della watchlist e personalizzati compaiono **in cima** alla lista.

| Metrica | Riferimento indicativo |
|---|---|
| P/E | <15 economico, >40 caro — confronta sempre col settore |
| P/B | <3 ragionevole per la maggior parte dei settori |
| ROE | >15% buona redditività del capitale |
| Debt/Equity | <80 contenuto, >200 leva elevata |
| Div. Yield | Confronta con rendimento obbligazionario corrente |

⚠️ Questi dati sono aggiornati trimestralmente da Yahoo Finance.

---

### 🌐 Macro
Fonte: Federal Reserve di St. Louis (FRED) — dati ufficiali gratuiti.
**Frequenza consigliata: una volta a settimana.**

**Curva dei rendimenti (Treasury 2Y vs 10Y)**
- Spread positivo = curva normale = nessun segnale recessivo
- Spread negativo = curva invertita = segnale storico di rallentamento (anticipo 12-18 mesi)

**Inflazione CPI (YoY%)**
- Target Fed: 2%
- La direzione conta più del numero assoluto
- In calo verso 2% = apertura a tagli dei tassi = favorevole per azioni

**Dollaro DXY**
- Forte = pressione su materie prime e multinazionali USA
- Debole = favorevole per commodity e mercati emergenti

**Lettura integrata:** sintesi automatica dei tre indicatori in un giudizio di contesto.

---

### Aggiungere ticker personalizzati

Dalla sidebar — campo "Ticker aggiuntivi (temporanei)" per analisi rapide non salvate,
oppure direttamente in Watchlist per titoli fissi.

| Tipo | Formato | Esempio |
|---|---|---|
| Azioni italiane | TICKER.MI | ENI.MI, RACE.MI |
| Azioni tedesche | TICKER.DE | SAP.DE, BMW.DE |
| Azioni francesi | TICKER.PA | MC.PA, RMS.PA |
| Criptovalute | TICKER-USD | BTC-USD, ETH-USD |
| Forex | COPPIA=X | EURUSD=X |
| Futures | TICKER=F | GC=F (oro), CL=F (petrolio) |

---

### Limiti del sistema

1. Dati con ritardo 15-20 minuti — irrilevante per medio-lungo termine
2. Fondamentali aggiornati trimestralmente da Yahoo Finance
3. Titoli a bassa capitalizzazione possono non essere coperti da Yahoo Finance
4. Yahoo Finance non concede licenza commerciale — per distribuzione a pagamento valutare Alpha Vantage o Twelve Data
5. I dati macro FRED coprono principalmente l'economia USA

---

### Disclaimer

Questo strumento fornisce esclusivamente dati e indicatori tecnici a scopo informativo
e di ricerca personale. Non costituisce consulenza finanziaria, raccomandazione di
investimento né sollecitazione all'acquisto o alla vendita di strumenti finanziari
ai sensi della Direttiva MiFID II (2014/65/UE) e del D.Lgs. 58/1998 (TUF).
Le decisioni di investimento sono di esclusiva responsabilità dell'utente.
    """)
