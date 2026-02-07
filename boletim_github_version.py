# -*- coding: utf-8 -*-
r"""
Netuno — Boletim de Mercado — GitHub Actions Edition
Versão adaptada para rodar na nuvem via GitHub Actions + SMTP Gmail.
"""

import os
import re
import time
import html
import smtplib
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta, timezone
from typing import List, Dict

import requests
from bs4 import BeautifulSoup
import feedparser
from readability import Document

# ================== CONFIG ==================
DESTINATARIOS = [
    "gustavoportugalhamer@gmail.com",
    "gustavo.hamer@r2fcapital.com.br",
    "gustavo.sernagiotto@r2fcapital.com.br",
    "nadine.dias@r2fcapital.com.br",
    "arthur.hamer@r2fcapital.com.br",
    "arthur@netunoinvestimentos.com.br",
    "paolaperassoli@gmail.com",
    "marketing@netunoinvestimentos.com.br",
    "carlos.ferreira@r2fcapital.com.br",
    "carlos@r2fseguros.com"
]
ASSUNTO_PREFIXO = "[Grupo Netuno] Boletim Diário de Mercado"
TIMEOUT      = 18
SLEEP        = 0.35
JANELA_DIAS  = 2
MAX_BULLETS  = 12
MAX_PER_FONTE= 8
TZ_BR        = timezone(timedelta(hours=-3))
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetunoBot/2.0"}

# ===== Credenciais via variáveis de ambiente =====
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

# ===== Paleta de Cores R2F Capital / Grupo Netuno =====
COLORS = {
    "primary": "#00acad",
    "primary_dark": "#008b8c",
    "secondary": "#1a2744",
    "accent": "#00acad",
    "accent_hover": "#61a229",
    "positive": "#28a745",
    "negative": "#dc3545",
    "neutral": "#6c757d",
    "text_dark": "#000000",
    "text_medium": "#333333",
    "text_light": "#6c757d",
    "bg_light": "#f8f9fa",
    "bg_card": "#ffffff",
    "border": "#e3e1e8",
    "divider": "#dee2e6",
    "header_bg": "#1a2744",
}

# ===== Cotações (Yahoo Finance) =====
B3_TICKERS = {
    "PETR4.SA":"Petrobras PN", "VALE3.SA":"Vale ON", "ITUB4.SA":"Itaú Unibanco PN",
    "BBDC4.SA":"Bradesco PN", "BBAS3.SA":"Banco do Brasil ON", "ABEV3.SA":"Ambev ON",
    "WEGE3.SA":"WEG ON", "SUZB3.SA":"Suzano ON", "B3SA3.SA":"B3 ON", "GGBR4.SA":"Gerdau PN"
}
GLOBAL_TICKERS = {
    "AAPL":"Apple", "MSFT":"Microsoft", "NVDA":"Nvidia", "AMZN":"Amazon",
    "TSLA":"Tesla", "META":"Meta", "GOOGL":"Alphabet", "JPM":"JPMorgan",
    "ES=F":"S&P 500 Fut", "NQ=F":"Nasdaq Fut", "^STOXX50E":"Euro Stoxx 50",
    "^FTSE":"FTSE 100", "^FCHI":"CAC 40", "^N225":"Nikkei 225",
    "CL=F":"WTI", "BZ=F":"Brent", "GC=F":"Ouro", "BTC-USD":"Bitcoin",
    "BRL=X":"USD/BRL", "EURUSD=X":"EUR/USD"
}

# ===== Fontes =====
SOURCES = [
    {"name":"InfoMoney", "cat":"Brasil", "rss":["https://www.infomoney.com.br/feed/"]},
    {"name":"Valor Investe", "cat":"Brasil", "rss":["https://valorinveste.globo.com/rss/"]},
    {"name":"Mais Retorno", "cat":"Brasil", "rss":["https://maisretorno.com/feed"]},
    {"name":"Yubb", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://yubb.com.br/blog", "sel":"a.card", "attr":"href"}]},
    {"name":"Bora Investir / B3", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://borainvestir.b3.com.br/noticias", "sel":"a[href*='/noticias/']", "attr":"href"}]},
    {"name":"Bloomberg Línea BR", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://www.bloomberglinea.com.br/", "sel":"a[href*='/brasil/'], a[href*='/mercados/']", "attr":"href"}]},
    {"name":"Brazil Journal", "cat":"Empresas", "rss":[], "html":[
        {"url":"https://braziljournal.com/", "sel":"article a", "attr":"href"}]},
    {"name":"Bloomberg", "cat":"Internacional", "rss":[], "html":[
        {"url":"https://www.bloomberg.com/markets", "sel":"a[href*='/news/']", "attr":"href"}]},
    {"name":"Reuters Markets", "cat":"Internacional", "rss":["https://www.reuters.com/markets/rss"]},
    {"name":"Financial Times", "cat":"Internacional", "rss":["https://www.ft.com/world/americas/rss"]},
    {"name":"Seeking Alpha", "cat":"Internacional", "rss":["https://seekingalpha.com/market_currents.xml"]},
]

# ================= ESTILOS HTML PREMIUM =================
def get_email_styles() -> str:
    return f"""
    <style type="text/css">
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        body, table, td, p, a, li {{
            -webkit-text-size-adjust: 100%;
            -ms-text-size-adjust: 100%;
        }}
        .email-container {{
            max-width: 680px;
            margin: 0 auto;
            background-color: {COLORS['bg_light']};
        }}
        .header-gradient {{
            background: linear-gradient(135deg, {COLORS['primary']} 0%, {COLORS['secondary']} 100%);
        }}
        @media screen and (max-width: 600px) {{
            .email-container {{ width: 100% !important; }}
            .content-padding {{ padding: 20px !important; }}
            .quote-grid td {{ display: block !important; width: 100% !important; }}
        }}
    </style>
    """

def get_header_html() -> str:
    data_atual = datetime.now(TZ_BR).strftime('%d de %B de %Y').replace(
        'January', 'Janeiro').replace('February', 'Fevereiro').replace('March', 'Marco'
        ).replace('April', 'Abril').replace('May', 'Maio').replace('June', 'Junho'
        ).replace('July', 'Julho').replace('August', 'Agosto').replace('September', 'Setembro'
        ).replace('October', 'Outubro').replace('November', 'Novembro').replace('December', 'Dezembro')
    hora_atual = datetime.now(TZ_BR).strftime('%H:%M')

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['header_bg']};">
        <tr>
            <td style="padding: 35px 30px 30px 30px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td>
                            <table cellpadding="0" cellspacing="0" border="0">
                                <tr>
                                    <td style="padding-right: 15px; vertical-align: middle;">
                                        <table cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['primary']}; border-radius: 8px;">
                                            <tr>
                                                <td style="padding: 12px 14px;">
                                                    <span style="font-family: Arial, sans-serif; font-size: 22px; font-weight: 700; color: #ffffff;">N</span>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="vertical-align: middle;">
                                        <p style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 24px; font-weight: 700; color: #ffffff; letter-spacing: 0.5px;">
                                            GRUPO NETUNO
                                        </p>
                                        <p style="margin: 4px 0 0 0; font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: 600; color: {COLORS['primary']}; text-transform: uppercase; letter-spacing: 2px;">
                                            Inteligência de Mercado
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top: 25px;">
                    <tr>
                        <td>
                            <p style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 28px; font-weight: 700; color: #ffffff; line-height: 1.3;">
                                Boletim Diário de Mercado
                            </p>
                            <table cellpadding="0" cellspacing="0" border="0" style="margin-top: 12px;">
                                <tr>
                                    <td style="background-color: {COLORS['primary']}; border-radius: 4px; padding: 6px 12px;">
                                        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #ffffff;">
                                            {data_atual} &bull; {hora_atual} (Brasilia)
                                        </span>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    """

def get_footer_html() -> str:
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['header_bg']};">
        <tr>
            <td style="padding: 30px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td style="text-align: center;">
                            <p style="margin: 0 0 10px 0; font-family: Arial, Helvetica, sans-serif; font-size: 16px; font-weight: 700; color: {COLORS['primary']};">
                                Grupo Netuno | R2F Capital
                            </p>
                            <p style="margin: 0 0 15px 0; font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: rgba(255,255,255,0.7); line-height: 1.6;">
                                Este boletim é produzido automaticamente com informações de fontes públicas confiáveis.<br>
                                As informações contidas não constituem recomendação de investimento.
                            </p>
                        </td>
                    </tr>
                    <tr>
                        <td style="border-top: 1px solid rgba(255,255,255,0.1); padding-top: 15px; text-align: center;">
                            <p style="margin: 0; font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: rgba(255,255,255,0.5); line-height: 1.6;">
                                Fontes: InfoMoney, Valor Investe, Bloomberg, Reuters, Financial Times, Seeking Alpha, Brazil Journal, Yahoo Finance<br>
                                Gerado em {datetime.now(TZ_BR).strftime('%d/%m/%Y às %H:%M')} (UTC-3) | <a href="https://r2fcapital.com.br" style="color: {COLORS['primary']}; text-decoration: none;">r2fcapital.com.br</a>
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    """

# ================= HTTP helpers =================
def http_get(url: str, **kw) -> requests.Response:
    return requests.get(url, timeout=TIMEOUT, headers=HEADERS, **kw)

def norm_url(base: str, href: str) -> str:
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"):  return requests.compat.urljoin(base, href)
    return href

# =================== Tradução ===================
def translate_to_pt(text: str) -> str:
    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {"client": "gtx", "sl": "auto", "tl": "pt", "dt": "t", "q": text}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        return "".join([seg[0] for seg in r.json()[0]])
    except Exception:
        return text

# ========================= Yahoo Finance =========================
def _yahoo_quote_once(symbols: List[str], host: str):
    url = f"https://{host}/v7/finance/quote"
    r = http_get(url, params={"symbols": ",".join(symbols)})
    r.raise_for_status()
    data = r.json().get("quoteResponse", {}).get("result", [])
    return {d.get("symbol"): d for d in data if d.get("symbol")}

def _yahoo_chart_one(symbol: str, host: str):
    url = f"https://{host}/v8/finance/chart/{symbol}"
    r = http_get(url, params={"range":"1d","interval":"1m"})
    r.raise_for_status()
    j = r.json()
    res = (j.get("chart") or {}).get("result") or []
    if not res: return {}
    meta = res[0].get("meta", {})
    price = meta.get("regularMarketPrice")
    if price is None:
        closes = (res[0].get("indicators", {}).get("quote", [{}])[0].get("close") or [])
        closes = [c for c in closes if isinstance(c,(int,float))]
        if closes: price = closes[-1]
    prev = meta.get("previousClose")
    pct = None
    if isinstance(price,(int,float)) and isinstance(prev,(int,float)) and prev:
        pct = (price/prev - 1.0) * 100.0
    return {"price": price, "pct": pct}

def yahoo_quote(symbols: List[str]) -> Dict[str, dict]:
    out = {}
    if not symbols: return out
    hosts = ["query1.finance.yahoo.com","query2.finance.yahoo.com"]
    chunk = 10
    for i in range(0, len(symbols), chunk):
        syms = symbols[i:i+chunk]
        ok = False
        for h in hosts:
            try:
                out.update(_yahoo_quote_once(syms, h))
                ok = True
                break
            except Exception:
                time.sleep(0.6)
        time.sleep(0.25)
        if not ok:
            for sym in syms:
                try:
                    c = _yahoo_chart_one(sym, hosts[0])
                    if c:
                        out[sym] = {
                            "symbol": sym,
                            "regularMarketPrice": c.get("price"),
                            "regularMarketChangePercent": c.get("pct")
                        }
                except Exception:
                    time.sleep(0.25)
    for sym in symbols:
        d = out.get(sym, {})
        if not d or d.get("regularMarketPrice") is None:
            try:
                c = _yahoo_chart_one(sym, hosts[0])
                if c:
                    d["regularMarketPrice"] = c.get("price")
                    d["regularMarketChangePercent"] = c.get("pct")
                    out[sym] = d
            except Exception:
                pass
    return out

def fmt_price(sym: str, v) -> str:
    if v is None: return "—"
    try:
        prefix = "R$ " if sym.endswith(".SA") else "$"
        return prefix + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"

def fmt_pct(v) -> str:
    if v is None: return "—"
    try:
        return "{:+.2f}%".format(float(v)).replace(".", ",")
    except Exception:
        return "—"

def get_pct_color(v) -> str:
    if v is None: return COLORS['neutral']
    try:
        if float(v) > 0: return COLORS['positive']
        elif float(v) < 0: return COLORS['negative']
    except: pass
    return COLORS['neutral']

def bloco_cotacoes_premium() -> str:
    br = yahoo_quote(list(B3_TICKERS.keys()))
    gl = yahoo_quote(list(GLOBAL_TICKERS.keys()))

    def render_quote_row(sym: str, label: str, payload: dict, show_ticker: bool = True) -> str:
        d = payload.get(sym, {})
        px = d.get("regularMarketPrice")
        pct = d.get("regularMarketChangePercent")
        pct_color = get_pct_color(pct)
        ticker = sym.replace(".SA", "") if show_ticker else ""

        return f"""
        <tr>
            <td style="padding: 8px 0; border-bottom: 1px solid {COLORS['border']};">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: {COLORS['text_dark']}; font-weight: 500;">
                            {html.escape(label)}
                            {f'<span style="color: {COLORS["text_light"]}; font-size: 11px; font-weight: 400;"> ({ticker})</span>' if ticker else ''}
                        </td>
                        <td style="text-align: right; font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: {COLORS['text_dark']}; font-weight: 600; padding-right: 15px;">
                            {fmt_price(sym, px)}
                        </td>
                        <td style="text-align: right; width: 70px;">
                            <span style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #ffffff; font-weight: 600; background-color: {pct_color}; padding: 3px 8px; border-radius: 4px;">
                                {fmt_pct(pct)}
                            </span>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        """

    html_br = "".join([render_quote_row(sym, label, br, True) for sym, label in B3_TICKERS.items()])
    html_gl = "".join([render_quote_row(sym, label, gl, False) for sym, label in GLOBAL_TICKERS.items()])

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['bg_light']};">
        <tr>
            <td style="padding: 20px 30px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0" class="quote-grid">
                    <tr>
                        <td style="width: 48%; vertical-align: top; padding-right: 12px;">
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background: {COLORS['bg_card']}; border-radius: 8px; border: 1px solid {COLORS['border']};">
                                <tr>
                                    <td style="padding: 15px 15px 10px 15px; border-bottom: 2px solid {COLORS['primary']};">
                                        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: 700; color: {COLORS['primary']}; text-transform: uppercase; letter-spacing: 1px;">
                                            B3 Brasil
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 10px 15px 15px 15px;">
                                        <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                            {html_br}
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td style="width: 48%; vertical-align: top; padding-left: 12px;">
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background: {COLORS['bg_card']}; border-radius: 8px; border: 1px solid {COLORS['border']};">
                                <tr>
                                    <td style="padding: 15px 15px 10px 15px; border-bottom: 2px solid {COLORS['header_bg']};">
                                        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: 700; color: {COLORS['header_bg']}; text-transform: uppercase; letter-spacing: 1px;">
                                            Mercados Globais
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 10px 15px 15px 15px;">
                                        <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                            {html_gl}
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    """

# ================= News: RSS + HTML scraping =================
def parse_rss(url: str) -> List[Dict]:
    items = []
    try:
        f = feedparser.parse(url)
        fonte = f.feed.get("title", url)
        for e in f.entries:
            quando = None
            for key in ("published_parsed","updated_parsed"):
                t = getattr(e, key, None) or e.get(key)
                if t:
                    quando = datetime(*t[:6], tzinfo=timezone.utc).astimezone(TZ_BR); break
            if not quando: quando = datetime.now(TZ_BR)
            if quando < datetime.now(TZ_BR) - timedelta(days=JANELA_DIAS):
                continue
            title = (getattr(e,"title","") or e.get("title","")).strip()
            link  = getattr(e,"link","") or e.get("link","")
            summary = getattr(e,"summary","") or e.get("summary","")
            items.append({"fonte":fonte,"titulo":title,"link":link,"quando":quando,"resumo_feed":summary})
            if len(items) >= MAX_PER_FONTE: break
    except Exception:
        pass
    return items

def parse_html_list(url: str, css_sel: str, attr: str) -> List[str]:
    try:
        r = http_get(url); r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        hrefs = []
        for a in soup.select(css_sel)[:20]:
            href = a.get(attr) or ""
            if not href: continue
            hrefs.append(norm_url(url, href))
        seen, uniq = set(), []
        for h in hrefs:
            if h not in seen:
                seen.add(h); uniq.append(h)
        return uniq
    except Exception:
        return []

def _blocked_text(txt: str) -> bool:
    bad = ["Access to this page has been denied", "enable Javascript", "subscribe to read", "Assine para ler"]
    t = txt.lower()
    return any(b.lower() in t for b in bad)

def extract_article_text(url: str) -> str:
    try:
        r = http_get(url); r.raise_for_status()
        doc = Document(r.text)
        html_main = doc.summary(html_partial=True)
        txt = BeautifulSoup(html_main, "lxml").get_text(" ", strip=True)
        txt = re.sub(r"\s+", " ", txt).strip()
        if _blocked_text(txt): return ""
        return txt
    except Exception:
        try:
            soup = BeautifulSoup(r.text, "lxml")
            art = soup.find("article") or soup
            txt = art.get_text(" ", strip=True)
            txt = re.sub(r"\s+", " ", txt).strip()
            if _blocked_text(txt): return ""
            return txt
        except Exception:
            return ""

def summarize(text: str, max_frases=3, max_chars=380) -> str:
    if not text: return ""
    sents = re.split(r"(?<=[\.\!\?])\s+", text)
    sents = [s.strip() for s in sents if len(s.strip()) >= 30][:60]
    if not sents: return ""
    out = " ".join(sents[:max_frases])
    if len(out) > max_chars: out = out[:max_chars].rsplit(" ",1)[0]+"…"
    return out

def collect_news() -> Dict[str, List[Dict]]:
    out: Dict[str, List[Dict]] = {"Brasil":[], "Internacional":[], "Empresas":[]}
    for src in SOURCES:
        for u in src.get("rss", []):
            items = parse_rss(u)
            for it in items:
                it["srcname"] = src["name"]; it["cat"] = src["cat"]
            out[src["cat"]].extend(items); time.sleep(SLEEP)
        for cfg in src.get("html", []):
            links = parse_html_list(cfg["url"], cfg["sel"], cfg["attr"])
            for lk in links[:MAX_PER_FONTE]:
                try:
                    r = http_get(lk); r.raise_for_status()
                    soup = BeautifulSoup(r.text, "lxml")
                    title = (soup.title.get_text(strip=True) if soup.title else lk)
                    quando = datetime.now(TZ_BR)
                    out[src["cat"]].append({"fonte":src["name"], "titulo":title, "link":lk,
                                            "quando":quando, "resumo_feed":"", "srcname":src["name"], "cat":src["cat"]})
                    time.sleep(0.15)
                except Exception:
                    pass
            time.sleep(SLEEP)
    for k in out.keys():
        seen, uniq = set(), []
        for it in sorted(out[k], key=lambda x:x.get("quando", datetime.now(TZ_BR)), reverse=True):
            key = (it.get("titulo",""), it.get("fonte",""))
            if key in seen: continue
            seen.add(key); uniq.append(it)
        out[k] = uniq[:MAX_BULLETS]
    return out

# ================= Montagem Premium =================
def render_news_card(it: Dict, traduzir: bool = False) -> str:
    quando = it.get("quando", datetime.now(TZ_BR)).strftime("%d/%m %H:%M")
    fonte = html.escape(it.get("fonte") or it.get("srcname") or "")
    title_raw = re.sub(r"\s+", " ", (it.get("titulo") or "").strip())
    link = it.get("link", "")

    txt = extract_article_text(link)
    if not txt:
        raw = BeautifulSoup(it.get("resumo_feed", ""), "lxml").get_text(" ", strip=True)
        txt = re.sub(r"\s+", " ", raw)
    if not txt or "Access to this page has been denied" in txt:
        return ""

    resumo = summarize(txt, max_frases=2, max_chars=280)

    if traduzir:
        title = html.escape(translate_to_pt(title_raw))
        resumo = html.escape(translate_to_pt(resumo))
    else:
        title = html.escape(title_raw)
        resumo = html.escape(resumo)

    return f"""
    <tr>
        <td style="padding-bottom: 12px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background: {COLORS['bg_card']}; border-radius: 6px; border: 1px solid {COLORS['border']};">
                <tr>
                    <td style="padding: 15px 18px;">
                        <p style="margin: 0 0 8px 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; font-weight: 700; color: {COLORS['text_dark']}; line-height: 1.4;">
                            <a href="{link}" style="color: {COLORS['text_dark']}; text-decoration: none;">{title}</a>
                        </p>
                        <p style="margin: 0 0 12px 0; font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: {COLORS['text_light']}; line-height: 1.5;">
                            {resumo}
                        </p>
                        <table cellpadding="0" cellspacing="0" border="0">
                            <tr>
                                <td style="padding-right: 12px;">
                                    <span style="font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: {COLORS['text_light']};">
                                        {fonte}
                                    </span>
                                </td>
                                <td style="padding-right: 12px;">
                                    <span style="font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: {COLORS['text_light']};">
                                        {quando}
                                    </span>
                                </td>
                                <td>
                                    <a href="{link}" style="font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: {COLORS['primary']}; text-decoration: none; font-weight: 600;">
                                        Ler mais &rarr;
                                    </a>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    """

def render_section_header(title: str, subtitle: str = "") -> str:
    return f"""
    <tr>
        <td style="padding: 20px 0 12px 0;">
            <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td style="border-left: 4px solid {COLORS['primary']}; padding-left: 12px;">
                        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 16px; font-weight: 700; color: {COLORS['header_bg']}; text-transform: uppercase; letter-spacing: 0.5px;">
                            {title}
                        </span>
                    </td>
                </tr>
                {f'<tr><td style="padding-left: 16px; padding-top: 4px;"><span style="font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: {COLORS["text_light"]};">{subtitle}</span></td></tr>' if subtitle else ''}
            </table>
        </td>
    </tr>
    """

def bloco_destaques_premium(news: Dict[str, List[Dict]], total: int = 4) -> str:
    pool = []
    pool.extend(news.get("Brasil", [])[:2])
    pool.extend(news.get("Empresas", [])[:1])
    pool.extend(news.get("Internacional", [])[:3])

    bullets = []
    for it in pool[:total]:
        link = it.get("link", "")
        title = re.sub(r"\s+", " ", (it.get("titulo") or "").strip())
        if it.get("cat") == "Internacional":
            title = translate_to_pt(title)

        bullets.append(f"""
        <tr>
            <td style="padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.15);">
                <table cellpadding="0" cellspacing="0" border="0" width="100%">
                    <tr>
                        <td style="width: 6px; vertical-align: top; padding-top: 6px; padding-right: 12px;">
                            <span style="display: inline-block; width: 6px; height: 6px; background: {COLORS['primary']}; border-radius: 50%;"></span>
                        </td>
                        <td>
                            <a href="{link}" style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: 500; line-height: 1.5;">
                                {html.escape(title)}
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        """)

    if not bullets:
        return ""

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['bg_light']};">
        <tr>
            <td style="padding: 20px 30px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['header_bg']}; border-radius: 8px; border-left: 4px solid {COLORS['primary']};">
                    <tr>
                        <td style="padding: 20px 25px;">
                            <table cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 12px;">
                                <tr>
                                    <td style="background-color: {COLORS['primary']}; border-radius: 4px; padding: 5px 10px;">
                                        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 11px; font-weight: 700; color: #ffffff; text-transform: uppercase; letter-spacing: 1px;">
                                            Destaques do Dia
                                        </span>
                                    </td>
                                </tr>
                            </table>
                            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                                {''.join(bullets)}
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    """

def bloco_noticias_premium(items: List[Dict], title: str, traduzir: bool = False, limit: int = 5, subtitle: str = "") -> str:
    cards = []
    for it in items[:limit]:
        card = render_news_card(it, traduzir=traduzir)
        if card:
            cards.append(card)

    if not cards:
        return ""

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: {COLORS['bg_light']};">
        <tr>
            <td style="padding: 0 30px 15px 30px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    {render_section_header(title, subtitle)}
                    {''.join(cards)}
                </table>
            </td>
        </tr>
    </table>
    """

# ================= EMAIL (SMTP Gmail) =================
def enviar_email_smtp(dest: List[str], assunto: str, html_corpo: str):
    """Envia email via SMTP do Gmail."""
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        raise ValueError("Credenciais GMAIL_USER e GMAIL_APP_PASSWORD não configuradas!")

    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = f"Grupo Netuno <{GMAIL_USER}>"
    msg['To'] = ", ".join(dest)

    # Anexa o HTML
    part = MIMEText(html_corpo, 'html', 'utf-8')
    msg.attach(part)

    # Conecta ao Gmail SMTP
    print(f"[*] Conectando ao SMTP Gmail...")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, dest, msg.as_string())

    print(f"[OK] E-mail enviado para {len(dest)} destinatários via Gmail SMTP.")

# ================= MAIN =================
def main():
    print("[*] Iniciando coleta de noticias...")
    news = collect_news()
    print(f"[+] Coletadas: Brasil={len(news['Brasil'])}, Internacional={len(news['Internacional'])}, Empresas={len(news['Empresas'])}")

    # Montar email
    html_parts = [
        '<!DOCTYPE html>',
        '<html lang="pt-BR">',
        '<head>',
        '<meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
        '<meta http-equiv="X-UA-Compatible" content="IE=edge">',
        '<title>Boletim Grupo Netuno</title>',
        get_email_styles(),
        '</head>',
        '<body style="margin: 0; padding: 0; background-color: #e5e7eb;">',
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #e5e7eb;">',
        '<tr><td align="center" style="padding: 20px 10px;">',
        '<table class="email-container" width="680" cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 24px rgba(0,0,0,0.08);">',
        '<tr><td>',
        get_header_html(),
        bloco_destaques_premium(news, total=4),
        bloco_cotacoes_premium(),
        bloco_noticias_premium(news.get("Internacional", []), "Internacional", traduzir=True, limit=5),
        bloco_noticias_premium(news.get("Brasil", []), "Brasil", traduzir=False, limit=5),
        bloco_noticias_premium(news.get("Empresas", []), "Empresas", traduzir=False, limit=4),
        get_footer_html(),
        '</td></tr>',
        '</table>',
        '</td></tr>',
        '</table>',
        '</body>',
        '</html>'
    ]

    html_email = '\n'.join(html_parts)
    assunto = f"{ASSUNTO_PREFIXO} — {datetime.now(TZ_BR).strftime('%d/%m/%Y')}"

    print("[*] Enviando email...")
    enviar_email_smtp(DESTINATARIOS, assunto, html_email)

if __name__ == "__main__":
    try:
        main()
        print("[OK] Boletim Premium enviado com sucesso.")
    except Exception as e:
        print("[ERRO] Falha no boletim:", e)
        traceback.print_exc()
        exit(1)
