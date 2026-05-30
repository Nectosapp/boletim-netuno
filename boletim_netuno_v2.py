# -*- coding: utf-8 -*-
r"""
Netuno — Boletim de Mercado — v11 (premium redesign)
- Destaques limpos e elegantes
- Insights do Mercado auto-gerados
- Radar "at a glance" (painel compacto)
- Cotações mobile-friendly
- Unsubscribe link no rodapé
"""

import os
import re
import time
import html
import smtplib
import traceback
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List

import requests
from bs4 import BeautifulSoup
import feedparser
import yfinance as yf

# ================== CONFIG ==================
DESTINATARIOS = [
    "gustavoportugalhamer@gmail.com",
    "arthur.hamer@r2fcapital.com.br",
    "arthur@netunoinvestimentos.com.br",
    "paolaperassoli@gmail.com",
    "marketing@netunoinvestimentos.com.br",
    "carlos@r2fseguros.com",
    "Byanca.tavelli@uranofin.com",
    "Pedro.barros@netunoinvestimentos.com.br",
    "eduardo@netunoinvestimentos.com.br",
    "andre@netunoinvestimentos.com.br",
    "carolina@netunoinvestimentos.com.br",
]
TEST_MODE = os.environ.get("BOLETIM_TEST_MODE", "false").lower() == "true"
ASSUNTO_PREFIXO = "[Grupo Netuno] Boletim Diário de Mercado"
TIMEOUT      = 18
SLEEP        = 0.35
JANELA_DIAS  = 2
MAX_BULLETS  = 12
MAX_PER_FONTE= 8
TZ_BR        = timezone(timedelta(hours=-3))
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetunoBot/2.0"}
UNSUBSCRIBE_BASE = "https://nectosapp.github.io/boletim-netuno/unsubscribe.html"

# ===== Gmail SMTP =====
GMAIL_USER = os.environ.get("GMAIL_USER", "gustavoportugalhamer@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
UNSUBSCRIBED_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unsubscribed.txt")

def load_unsubscribed() -> set:
    if not os.path.exists(UNSUBSCRIBED_FILE): return set()
    with open(UNSUBSCRIBED_FILE, "r", encoding="utf-8") as f:
        return {line.strip().lower() for line in f if line.strip() and not line.startswith("#")}

def filter_destinatarios(dest: list) -> list:
    unsub = load_unsubscribed()
    if not unsub: return dest
    filtered = [e for e in dest if e.strip().lower() not in unsub]
    removed = len(dest) - len(filtered)
    if removed: print(f"[*] {removed} email(s) descadastrado(s) removido(s).")
    return filtered

# ===== Cores =====
C = {
    "teal": "#00acad", "teal_dk": "#008b8c", "navy": "#1a2744",
    "green": "#28a745", "red": "#dc3545", "gray": "#6c757d",
    "white": "#ffffff", "light": "#f8f9fa", "border": "#e3e1e8",
    "txt": "#000000", "txt2": "#333333", "txt3": "#6c757d",
}

# ===== Radar (tickers monitorados) =====
RADAR_BR = {
    "PETR4.SA": "Petrobras",  "VALE3.SA": "Vale",  "ITUB4.SA": "Itaú Unibanco",
    "BBDC4.SA": "Bradesco",   "BBAS3.SA": "Banco do Brasil",  "WEGE3.SA": "WEG",
    "ABEV3.SA": "Ambev",      "SUZB3.SA": "Suzano",
}
RADAR_US = {
    "AAPL": "Apple",  "MSFT": "Microsoft",  "NVDA": "Nvidia",  "AMZN": "Amazon",
    "TSLA": "Tesla",  "META": "Meta",  "GOOGL": "Alphabet",  "JPM": "JPMorgan",
}
RECOMENDACAO_MAP = {"strongBuy": "Compra Forte", "buy": "Compra", "hold": "Neutro",
                    "sell": "Venda", "strongSell": "Venda Forte"}

# ===== Cotações =====
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
        {"url":"https://braziljournal.com/category/negocios/", "sel":"article a", "attr":"href"}]},
    {"name":"Exame", "cat":"Empresas", "rss":["https://exame.com/feed/negocios/"]},
    {"name":"InfoMoney Empresas", "cat":"Empresas", "rss":["https://www.infomoney.com.br/feed/mercados/empresas/"]},
    {"name":"Valor Econômico", "cat":"Empresas", "rss":[], "html":[
        {"url":"https://valor.globo.com/empresas/", "sel":"a.feed-post-link, a[href*='/empresas/']", "attr":"href"}]},
    {"name":"Bloomberg", "cat":"Internacional", "rss":[], "html":[
        {"url":"https://www.bloomberg.com/markets", "sel":"a[href*='/news/']", "attr":"href"}]},
    {"name":"Reuters Markets", "cat":"Internacional", "rss":["https://www.reuters.com/markets/rss"]},
    {"name":"Financial Times", "cat":"Internacional", "rss":["https://www.ft.com/world/americas/rss"]},
    {"name":"Seeking Alpha", "cat":"Internacional", "rss":["https://seekingalpha.com/market_currents.xml"]},
]

EMPRESAS_BLACKLIST = [
    "ata ", "atas ", "comunicado ao mercado", "fato relevante",
    "filme", "filmes", "documentário", "cinema", "teatro",
    "série", "séries", "netflix", "streaming", "hollywood",
    "música", "show ", "livro", "cultura", "entretenimento",
    "receita ", "receitas de ", "culinári", "marketing",
    "fatos relevantes", "empresas | valor",
]

# ================= HTML HELPERS =================
F = "font-family: Arial, Helvetica, sans-serif;"

def _section(bg, content, pad="20px 30px"):
    return f"""<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:{bg};"><tr><td style="padding:{pad};" class="content-padding">{content}</td></tr></table>"""

def _badge(text, bg=None):
    bg = bg or C['teal']
    return f'<table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:{bg};border-radius:4px;padding:4px 10px;"><span style="{F}font-size:10px;font-weight:700;color:#fff;text-transform:uppercase;letter-spacing:1px;">{text}</span></td></tr></table>'

# ================= HEADER / FOOTER =================
def get_header_html() -> str:
    data = datetime.now(TZ_BR).strftime('%d de %B de %Y').replace(
        'January','Janeiro').replace('February','Fevereiro').replace('March','Marco'
        ).replace('April','Abril').replace('May','Maio').replace('June','Junho'
        ).replace('July','Julho').replace('August','Agosto').replace('September','Setembro'
        ).replace('October','Outubro').replace('November','Novembro').replace('December','Dezembro')
    hora = datetime.now(TZ_BR).strftime('%H:%M')
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:{C['navy']};">
        <tr><td style="padding:30px 30px 25px;">
            <table cellpadding="0" cellspacing="0" border="0"><tr>
                <td style="padding-right:12px;vertical-align:middle;">
                    <table cellpadding="0" cellspacing="0" border="0" style="background:{C['teal']};border-radius:8px;"><tr><td style="padding:10px 12px;">
                        <span style="{F}font-size:20px;font-weight:700;color:#fff;">N</span>
                    </td></tr></table>
                </td>
                <td style="vertical-align:middle;">
                    <p style="margin:0;{F}font-size:22px;font-weight:700;color:#fff;letter-spacing:0.5px;">GRUPO NETUNO</p>
                    <p style="margin:2px 0 0;{F}font-size:10px;font-weight:600;color:{C['teal']};text-transform:uppercase;letter-spacing:2px;">Inteligência de Mercado</p>
                </td>
            </tr></table>
            <p style="margin:20px 0 0;{F}font-size:24px;font-weight:700;color:#fff;line-height:1.2;">Boletim Diário de Mercado</p>
            <table cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;"><tr>
                <td style="background:{C['teal']};border-radius:4px;padding:5px 10px;">
                    <span style="{F}font-size:11px;color:#fff;">{data} &bull; {hora} (Brasília)</span>
                </td>
            </tr></table>
        </td></tr>
    </table>"""

def get_footer_html(dest_emails: List[str]) -> str:
    unsub_links = ""
    if len(dest_emails) == 1:
        email = dest_emails[0]
        unsub_links = f'<a href="{UNSUBSCRIBE_BASE}?email={email}" style="color:{C["teal"]};text-decoration:underline;{F}font-size:11px;">Cancelar inscrição</a>'
    else:
        unsub_links = f'<a href="{UNSUBSCRIBE_BASE}" style="color:{C["teal"]};text-decoration:underline;{F}font-size:11px;">Cancelar inscrição</a>'

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:{C['navy']};">
        <tr><td style="padding:25px 30px;text-align:center;">
            <p style="margin:0 0 8px;{F}font-size:14px;font-weight:700;color:{C['teal']};">Grupo Netuno | R2F Capital</p>
            <p style="margin:0 0 12px;{F}font-size:11px;color:rgba(255,255,255,0.6);line-height:1.5;">
                Este boletim é produzido automaticamente com informações de fontes públicas confiáveis.<br>
                As informações não constituem recomendação de investimento.
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="border-top:1px solid rgba(255,255,255,0.1);padding-top:12px;text-align:center;">
                <p style="margin:0 0 8px;{F}font-size:9px;color:rgba(255,255,255,0.4);line-height:1.5;">
                    Fontes: InfoMoney, Valor Investe, Exame, Bloomberg, Reuters, FT, Seeking Alpha, Brazil Journal, Yahoo Finance<br>
                    Gerado em {datetime.now(TZ_BR).strftime('%d/%m/%Y às %H:%M')} (UTC-3) | <a href="https://r2fcapital.com.br" style="color:{C['teal']};text-decoration:none;">r2fcapital.com.br</a>
                </p>
                <p style="margin:0;{F}font-size:11px;">{unsub_links}</p>
            </td></tr></table>
        </td></tr>
    </table>"""

# ================= HTTP / TRANSLATE =================
def http_get(url, **kw): return requests.get(url, timeout=TIMEOUT, headers=HEADERS, **kw)
def norm_url(base, href):
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"): return requests.compat.urljoin(base, href)
    return href

def translate_to_pt(text):
    try:
        r = requests.get("https://translate.googleapis.com/translate_a/single",
                         params={"client":"gtx","sl":"auto","tl":"pt","dt":"t","q":text}, timeout=10)
        r.raise_for_status()
        return "".join([s[0] for s in r.json()[0]])
    except Exception: return text

# ================= YAHOO FINANCE =================
def _yq_once(syms, host):
    r = http_get(f"https://{host}/v7/finance/quote", params={"symbols":",".join(syms)})
    r.raise_for_status()
    return {d["symbol"]: d for d in r.json().get("quoteResponse",{}).get("result",[]) if d.get("symbol")}

def _yc_one(sym, host):
    r = http_get(f"https://{host}/v8/finance/chart/{sym}", params={"range":"1d","interval":"1m"})
    r.raise_for_status()
    res = (r.json().get("chart") or {}).get("result") or []
    if not res: return {}
    m = res[0].get("meta",{})
    p = m.get("regularMarketPrice")
    if p is None:
        cl = [c for c in (res[0].get("indicators",{}).get("quote",[{}])[0].get("close") or []) if isinstance(c,(int,float))]
        if cl: p = cl[-1]
    prev = m.get("previousClose")
    pct = ((p/prev-1)*100) if isinstance(p,(int,float)) and isinstance(prev,(int,float)) and prev else None
    return {"price":p,"pct":pct}

def yahoo_quote(symbols):
    out = {}
    if not symbols: return out
    hosts = ["query1.finance.yahoo.com","query2.finance.yahoo.com"]
    for i in range(0,len(symbols),10):
        syms = symbols[i:i+10]
        ok = False
        for h in hosts:
            try: out.update(_yq_once(syms,h)); ok=True; break
            except: time.sleep(0.6)
        time.sleep(0.25)
        if not ok:
            for s in syms:
                try:
                    c = _yc_one(s,hosts[0])
                    if c: out[s]={"symbol":s,"regularMarketPrice":c.get("price"),"regularMarketChangePercent":c.get("pct")}
                except: time.sleep(0.25)
    for s in symbols:
        d = out.get(s,{})
        if not d or d.get("regularMarketPrice") is None:
            try:
                c = _yc_one(s,hosts[0])
                if c: d["regularMarketPrice"]=c.get("price"); d["regularMarketChangePercent"]=c.get("pct"); out[s]=d
            except: pass
    return out

def fp(sym,v):
    if v is None: return "—"
    try: return ("R$ " if sym.endswith(".SA") else "$")+f"{float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except: return "—"
def fpct(v):
    if v is None: return "—"
    try: return "{:+.2f}%".format(float(v)).replace(".",",")
    except: return "—"
def pcolor(v):
    if v is None: return C['gray']
    try:
        if float(v)>0: return C['green']
        if float(v)<0: return C['red']
    except: pass
    return C['gray']

# ================= RADAR =================
def _get_consensus(t):
    try:
        rec = t.recommendations
        if rec is None or rec.empty: return ""
        last = rec.iloc[-1]
        counts = {k:int(last.get(k,0)) for k in ["strongBuy","buy","hold","sell","strongSell"]}
        return RECOMENDACAO_MAP.get(max(counts,key=counts.get),"")
    except: return ""

def _get_target(t,sym):
    try:
        tg = t.analyst_price_targets
        if tg is None: return ""
        m = tg.get("mean") or tg.get("current")
        if m is None: return ""
        return ("R$ " if sym.endswith(".SA") else "US$ ")+f"{float(m):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except: return ""

def collect_radar(tickers, traduzir=False):
    items = []
    for sym,nome in tickers.items():
        try:
            t = yf.Ticker(sym)
            nl = t.get_news(count=3) if hasattr(t,'get_news') else (t.news or [])
            if not nl: continue
            top = nl[0] if isinstance(nl,list) else None
            if not top or not isinstance(top,dict): continue
            ct = top.get("content",top)
            titulo = ct.get("title","")
            ck = ct.get("clickThroughUrl") or {}; cn = ct.get("canonicalUrl") or {}
            link = ck.get("url") or cn.get("url") or ct.get("link","")
            prov = ct.get("provider") or {}
            pub = prov.get("displayName") if isinstance(prov,dict) else str(prov)
            if not pub: pub = "Yahoo Finance"
            if not titulo: continue
            if traduzir: titulo = translate_to_pt(titulo)
            items.append({"ticker":sym.replace(".SA",""),"nome":nome,"titulo":titulo,"link":link,
                          "pub":pub,"consenso":_get_consensus(t),"alvo":_get_target(t,sym)})
            time.sleep(0.3)
        except Exception as e:
            print(f"[WARN] Radar {sym}: {e}"); continue
    return items

def bloco_radar(items_br, items_us):
    """Radar at-a-glance: linhas compactas, não cards."""
    if not items_br and not items_us: return ""

    def row(it):
        tk = html.escape(it["ticker"]); nm = html.escape(it["nome"])
        tit = html.escape(it["titulo"]); lk = it["link"]
        pub = html.escape(it["pub"])
        meta = []
        if it.get("consenso"): meta.append(f"<b>{html.escape(it['consenso'])}</b>")
        if it.get("alvo"): meta.append(f"Alvo: <b>{html.escape(it['alvo'])}</b>")
        ml = " · ".join(meta)
        return f"""<tr>
            <td style="padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.08);">
                <span style="{F}font-size:12px;font-weight:700;color:#fff;">{nm}</span>
                <span style="{F}font-size:10px;color:rgba(255,255,255,0.5);"> ({tk})</span>
                {f'<span style="{F}font-size:10px;color:{C["teal"]};padding-left:6px;">{ml}</span>' if ml else ''}
                <br>
                <a href="{lk}" style="{F}font-size:11px;color:rgba(255,255,255,0.8);text-decoration:none;line-height:1.3;">{tit}</a>
                <span style="{F}font-size:9px;color:rgba(255,255,255,0.35);"> — {pub}</span>
            </td>
        </tr>"""

    def sub(flag, title, items):
        if not items: return ""
        rows = "".join([row(it) for it in items])
        return f"""<tr><td style="padding:8px 0 2px;">
            <span style="{F}font-size:10px;font-weight:700;color:{C['teal']};text-transform:uppercase;letter-spacing:1px;">{flag} {title}</span>
        </td></tr>{rows}"""

    br = sub("🇧🇷","Brasil",items_br)
    us = sub("🇺🇸","EUA",items_us)
    inner = f"""{_badge('Radar de Abertura de Mercado')}
        <p style="margin:4px 0 0;{F}font-size:10px;color:rgba(255,255,255,0.5);">Principais movimentações das empresas sob monitoramento.</p>
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:6px;">{br}{us}</table>"""
    card = f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:{C["navy"]};border-radius:8px;"><tr><td style="padding:14px 18px;">{inner}</td></tr></table>'
    return _section(C['light'], card, "15px 30px")

# ================= NEWS =================
def parse_rss(url):
    items = []
    try:
        f = feedparser.parse(url); fonte = f.feed.get("title",url)
        for e in f.entries:
            quando = None
            for k in ("published_parsed","updated_parsed"):
                t = getattr(e,k,None) or e.get(k)
                if t: quando = datetime(*t[:6],tzinfo=timezone.utc).astimezone(TZ_BR); break
            if not quando: quando = datetime.now(TZ_BR)
            if quando < datetime.now(TZ_BR)-timedelta(days=JANELA_DIAS): continue
            title = (getattr(e,"title","") or e.get("title","")).strip()
            link = getattr(e,"link","") or e.get("link","")
            items.append({"fonte":fonte,"titulo":title,"link":link,"quando":quando,"resumo_feed":getattr(e,"summary","")})
            if len(items)>=MAX_PER_FONTE: break
    except: pass
    return items

def parse_html_list(url,sel,attr):
    try:
        r = http_get(url); r.raise_for_status()
        soup = BeautifulSoup(r.text,"lxml")
        seen,out = set(),[]
        for a in soup.select(sel)[:20]:
            h = a.get(attr) or ""
            if not h: continue
            h = norm_url(url,h)
            if h not in seen: seen.add(h); out.append(h)
        return out
    except: return []

def collect_news():
    out = {"Brasil":[],"Internacional":[],"Empresas":[]}
    for src in SOURCES:
        for u in src.get("rss",[]):
            items = parse_rss(u)
            for it in items: it["srcname"]=src["name"]; it["cat"]=src["cat"]
            out[src["cat"]].extend(items); time.sleep(SLEEP)
        for cfg in src.get("html",[]):
            links = parse_html_list(cfg["url"],cfg["sel"],cfg["attr"])
            for lk in links[:MAX_PER_FONTE]:
                try:
                    r = http_get(lk); r.raise_for_status()
                    soup = BeautifulSoup(r.text,"lxml")
                    title = (soup.title.get_text(strip=True) if soup.title else lk)
                    out[src["cat"]].append({"fonte":src["name"],"titulo":title,"link":lk,
                        "quando":datetime.now(TZ_BR),"resumo_feed":"","srcname":src["name"],"cat":src["cat"]})
                    time.sleep(0.15)
                except: pass
            time.sleep(SLEEP)
    for k in out:
        seen,uniq = set(),[]
        for it in sorted(out[k], key=lambda x:x.get("quando",datetime.now(TZ_BR)), reverse=True):
            key = (it.get("titulo",""),it.get("fonte",""))
            if key in seen: continue
            if k=="Empresas":
                tl = (it.get("titulo") or "").lower()
                if any(bl in tl for bl in EMPRESAS_BLACKLIST): continue
            seen.add(key); uniq.append(it)
        out[k] = uniq[:MAX_BULLETS]
    return out

# ================= DESTAQUES =================
def bloco_destaques(news):
    """Destaques limpos: numerados, elegantes."""
    def bullet(it, traduzir=False):
        link = it.get("link","")
        title = re.sub(r"\s+"," ",(it.get("titulo") or "").strip())
        if traduzir: title = translate_to_pt(title)
        fonte = html.escape(it.get("fonte") or it.get("srcname") or "")
        return f"""<tr><td style="padding:4px 0;">
            <a href="{link}" style="{F}font-size:12px;color:rgba(255,255,255,0.9);text-decoration:none;line-height:1.4;">
                &bull; {html.escape(title)}
            </a>
            <span style="{F}font-size:9px;color:rgba(255,255,255,0.35);"> — {fonte}</span>
        </td></tr>"""

    def cat(items, emoji, name, traduzir=False, limit=3):
        if not items: return ""
        rows = "".join([bullet(it,traduzir) for it in items[:limit]])
        return f"""<tr><td style="padding:8px 0 2px;">
            <span style="{F}font-size:10px;font-weight:700;color:{C['teal']};text-transform:uppercase;letter-spacing:1px;">{emoji} {name}</span>
        </td></tr>{rows}"""

    intl = cat(news.get("Internacional",[]),"🌎","Internacional",True,3)
    br = cat(news.get("Brasil",[]),"🇧🇷","Brasil",False,3)
    emp = cat(news.get("Empresas",[]),"🏢","Empresas",False,3)
    if not (intl or br or emp): return ""

    inner = f"""{_badge('Destaques do Dia')}
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:8px;">{intl}{br}{emp}</table>"""
    card = f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:{C["navy"]};border-radius:8px;border-left:4px solid {C["teal"]};"><tr><td style="padding:16px 20px;">{inner}</td></tr></table>'
    return _section(C['light'], card, "15px 30px")

# ================= INSIGHTS =================
def _calc_insights_text(quotes_gl):
    """Calcula a lista de insights como texto (sem HTML). Usado por bloco_insights e pelo POST ao Supabase."""
    insights = []

    # Insight 1: S&P direction
    sp = quotes_gl.get("ES=F",{})
    sp_pct = sp.get("regularMarketChangePercent")
    if sp_pct is not None:
        try:
            v = float(sp_pct)
            if v > 0: insights.append(f"Futuros de S&P 500 operam em alta de {fpct(sp_pct)}, sinalizando abertura positiva nos EUA")
            elif v < 0: insights.append(f"Futuros de S&P 500 recuam {fpct(sp_pct)}, indicando cautela nos mercados globais")
        except: pass

    # Insight 2: USD/BRL
    brl = quotes_gl.get("BRL=X",{})
    brl_pct = brl.get("regularMarketChangePercent")
    brl_px = brl.get("regularMarketPrice")
    if brl_pct is not None and brl_px is not None:
        try:
            v = float(brl_pct)
            px = f"{float(brl_px):,.2f}".replace(",","X").replace(".",",").replace("X",".")
            if v > 0: insights.append(f"Dólar avança {fpct(brl_pct)} e é negociado a R$ {px}")
            elif v < 0: insights.append(f"Dólar recua {fpct(brl_pct)} e é negociado a R$ {px}")
            else: insights.append(f"Dólar estável a R$ {px}")
        except: pass

    # Insight 3: Oil / Gold
    oil = quotes_gl.get("CL=F",{})
    oil_pct = oil.get("regularMarketChangePercent")
    if oil_pct is not None:
        try:
            v = float(oil_pct)
            if abs(v) >= 0.5:
                dir = "sobe" if v > 0 else "cai"
                insights.append(f"Petróleo WTI {dir} {fpct(oil_pct)}, influenciando setor de commodities")
        except: pass

    # Insight 4: Crypto
    btc = quotes_gl.get("BTC-USD",{})
    btc_pct = btc.get("regularMarketChangePercent")
    btc_px = btc.get("regularMarketPrice")
    if btc_pct is not None and btc_px is not None:
        try:
            v = float(btc_pct)
            if abs(v) >= 1.0:
                px = f"US$ {float(btc_px):,.0f}".replace(",",".")
                dir = "avança" if v > 0 else "recua"
                insights.append(f"Bitcoin {dir} {fpct(btc_pct)} e é negociado a {px}")
        except: pass

    return insights[:4]


def bloco_insights(quotes_gl):
    """Gera 3 insights automáticos baseados nos dados de mercado."""
    insights = _calc_insights_text(quotes_gl)
    if not insights: return ""
    bullets = "".join([f'<tr><td style="padding:4px 0;"><span style="{F}font-size:12px;color:rgba(255,255,255,0.9);line-height:1.5;">&bull; {i}</span></td></tr>' for i in insights])
    inner = f"""{_badge('💡 Insights do Mercado',C['teal_dk'])}
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:8px;">{bullets}</table>"""
    card = f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:{C["navy"]};border-radius:8px;"><tr><td style="padding:14px 18px;">{inner}</td></tr></table>'
    return _section(C['light'], card, "0 30px 15px")

# ================= COTAÇÕES =================
def bloco_cotacoes(quotes_br, quotes_gl):
    def qrow(sym, label, data, show_tk=True):
        d = data.get(sym,{})
        px = d.get("regularMarketPrice"); pct = d.get("regularMarketChangePercent")
        pc = pcolor(pct); tk = sym.replace(".SA","") if show_tk else ""
        return f"""<tr><td style="padding:5px 0;border-bottom:1px solid {C['border']};">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
                <td style="{F}font-size:12px;color:{C['txt']};font-weight:500;width:45%;">{html.escape(label)}{f'<span style="color:{C["txt3"]};font-size:10px;"> ({tk})</span>' if tk else ''}</td>
                <td style="text-align:right;{F}font-size:12px;color:{C['txt']};font-weight:600;width:30%;">{fp(sym,px)}</td>
                <td style="text-align:right;width:25%;"><span style="{F}font-size:10px;color:#fff;font-weight:600;background:{pc};padding:2px 5px;border-radius:3px;white-space:nowrap;">{fpct(pct)}</span></td>
            </tr></table>
        </td></tr>"""

    br_rows = "".join([qrow(s,l,quotes_br,True) for s,l in B3_TICKERS.items()])
    gl_rows = "".join([qrow(s,l,quotes_gl,False) for s,l in GLOBAL_TICKERS.items()])

    def card(title, emoji, color, rows):
        return f"""<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:{C['white']};border-radius:6px;border:1px solid {C['border']};margin-bottom:12px;">
            <tr><td style="padding:10px 12px 6px;border-bottom:2px solid {color};"><span style="{F}font-size:10px;font-weight:700;color:{color};text-transform:uppercase;letter-spacing:1px;">{emoji} {title}</span></td></tr>
            <tr><td style="padding:6px 12px 10px;"><table width="100%" cellpadding="0" cellspacing="0" border="0">{rows}</table></td></tr>
        </table>"""

    content = card("B3 Brasil","📈",C['teal'],br_rows) + card("Mercados Globais","🌐",C['navy'],gl_rows)
    return _section(C['light'], content, "15px 30px")

# ================= EMAIL =================
CONTA_REMETENTE = "gustavo.hamer@r2fcapital.com.br"

def enviar_email_smtp(dest, assunto, html_corpo):
    """Envia email via SMTP do Gmail."""
    gmail_user = GMAIL_USER.strip()
    gmail_pass = GMAIL_APP_PASSWORD.strip().replace(" ", "")

    if not gmail_user or not gmail_pass:
        print("[WARN] Credenciais Gmail não encontradas. Tentando Outlook...")
        return enviar_email_outlook(dest, assunto, html_corpo)

    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = gmail_user
    msg['To'] = ", ".join(dest)
    msg.attach(MIMEText(html_corpo, 'html', 'utf-8'))

    print(f"[*] Conectando ao SMTP Gmail...")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(gmail_user, gmail_pass)
        server.sendmail(gmail_user, dest, msg.as_string())
    print(f"[OK] Email enviado para {len(dest)} destinatário(s) via Gmail SMTP.")

def enviar_email_outlook(dest, assunto, html_corpo):
    """Fallback: envia via Outlook COM."""
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    try:
        for acc in outlook.Session.Accounts:
            ea = acc.SmtpAddress.lower() if hasattr(acc,'SmtpAddress') else ""
            da = acc.DisplayName.lower() if hasattr(acc,'DisplayName') else ""
            if CONTA_REMETENTE.lower() in ea or CONTA_REMETENTE.lower() in da:
                mail._oleobj_.Invoke(*(64209,0,8,0,acc))
                print(f"[OK] Via: {acc.SmtpAddress}"); break
            elif "r2fcapital.com.br" in ea or "r2fcapital.com.br" in da:
                mail._oleobj_.Invoke(*(64209,0,8,0,acc))
                print(f"[!] Via: {acc.SmtpAddress}"); break
    except Exception as e: print(f"[!] Conta: {e}")
    mail.To = ";".join(dest); mail.Subject = assunto; mail.HTMLBody = html_corpo; mail.Send()
    print("[OK] Email enviado via Outlook.")

# ================= MAIN =================
# ================= SUPABASE INGEST =================
# Posta o boletim do dia no app Netuno Investimentos (edge function
# daily-news-ingest). Requer env SUPABASE_URL e DAILY_NEWS_CRON_SECRET.
# Não é fatal: se falhar, só loga e o email segue normal.

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
DAILY_NEWS_CRON_SECRET = os.environ.get("DAILY_NEWS_CRON_SECRET", "")


def _strip_html_to_text(s: str) -> str:
    if not s: return ""
    try:
        return BeautifulSoup(s, "lxml").get_text(separator=" ", strip=True)[:1400]
    except Exception:
        return re.sub(r"<[^>]+>", " ", s).strip()[:1400]


def _news_items_to_payload(news_dict, radar_br, radar_us, insights_text):
    items = []
    for t in insights_text:
        items.append({"category": "insight", "title": t})

    cat_map = {"Brasil": "brasil", "Internacional": "internacional", "Empresas": "empresas"}
    for src_key, dst_key in cat_map.items():
        for it in news_dict.get(src_key, []):
            title = (it.get("titulo") or "").strip()
            if not title: continue
            if dst_key == "internacional":
                title = translate_to_pt(title)
            items.append({
                "category": dst_key,
                "title": title,
                "summary": _strip_html_to_text(it.get("resumo_feed") or ""),
                "source": it.get("fonte") or it.get("srcname") or "",
                "url": it.get("link") or "",
            })

    for it in radar_br:
        items.append({
            "category": "radar_br",
            "title": (it.get("titulo") or "")[:380],
            "source": it.get("pub") or "",
            "url": it.get("link") or "",
            "ticker": it.get("ticker") or "",
            "consensus": it.get("consenso") or "",
            "price_target": it.get("alvo") or "",
        })
    for it in radar_us:
        items.append({
            "category": "radar_us",
            "title": (it.get("titulo") or "")[:380],
            "source": it.get("pub") or "",
            "url": it.get("link") or "",
            "ticker": it.get("ticker") or "",
            "consensus": it.get("consenso") or "",
            "price_target": it.get("alvo") or "",
        })
    return items


def supabase_ingest_daily_news(news_dict, quotes_gl, radar_br, radar_us):
    if not SUPABASE_URL or not DAILY_NEWS_CRON_SECRET:
        print("[supabase] env SUPABASE_URL/DAILY_NEWS_CRON_SECRET não setadas — skip.")
        return
    edition_date = datetime.now(TZ_BR).strftime("%Y-%m-%d")
    insights_text = _calc_insights_text(quotes_gl)
    items = _news_items_to_payload(news_dict, radar_br, radar_us, insights_text)
    if not items:
        print("[supabase] sem items pra postar — skip.")
        return
    url = f"{SUPABASE_URL}/functions/v1/daily-news-ingest"
    try:
        r = requests.post(
            url,
            headers={
                "Content-Type": "application/json",
                "x-cron-secret": DAILY_NEWS_CRON_SECRET,
            },
            json={"edition_date": edition_date, "items": items},
            timeout=20,
        )
        if r.ok:
            print(f"[supabase] ok ({r.status_code}): {r.text[:200]}")
        else:
            print(f"[supabase] FALHOU ({r.status_code}): {r.text[:400]}")
    except Exception as e:
        print(f"[supabase] erro de rede: {e}")


def main():
    print("[*] Coletando notícias...")
    news = collect_news()
    print(f"[+] BR={len(news['Brasil'])}, INTL={len(news['Internacional'])}, EMP={len(news['Empresas'])}")

    print("[*] Cotações...")
    quotes_br = yahoo_quote(list(B3_TICKERS.keys()))
    quotes_gl = yahoo_quote(list(GLOBAL_TICKERS.keys()))

    print("[*] Radar de Abertura...")
    radar_br, radar_us = [], []
    try:
        radar_br = collect_radar(RADAR_BR, False)
        radar_us = collect_radar(RADAR_US, True)
        print(f"[+] Radar: BR={len(radar_br)}, US={len(radar_us)}")
    except Exception as e: print(f"[WARN] Radar: {e}")

    # Postagem pro app Netuno Investimentos (não-fatal).
    try:
        supabase_ingest_daily_news(news, quotes_gl, radar_br, radar_us)
    except Exception as e:
        print(f"[WARN] Supabase ingest falhou: {e}")

    styles = f"""<style type="text/css">
        body,table,td,p,a{{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}}
        .email-container{{max-width:680px;margin:0 auto;background:{C['light']};}}
        @media screen and (max-width:600px){{.email-container{{width:100%!important;}}.content-padding{{padding:12px!important;}}}}
    </style>"""

    assunto = f"{ASSUNTO_PREFIXO} — {datetime.now(TZ_BR).strftime('%d/%m/%Y')}"

    # Determina destinatários
    if TEST_MODE:
        dest_envio = ["gustavoportugalhamer@gmail.com"]
        print("[TEST MODE] Enviando apenas para teste.")
    else:
        dest_envio = DESTINATARIOS
    dest_envio = filter_destinatarios(dest_envio)

    if not dest_envio:
        print("[!] Nenhum destinatário ativo. Email não enviado.")
        return

    # Envio individual (link de unsubscribe personalizado por destinatário)
    print(f"[*] Enviando para {len(dest_envio)} destinatário(s)...")
    for dest_email in dest_envio:
        parts = [
            '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">',
            '<meta name="viewport" content="width=device-width,initial-scale=1.0">',
            f'<title>Boletim Grupo Netuno</title>{styles}</head>',
            f'<body style="margin:0;padding:0;background:#e5e7eb;">',
            '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#e5e7eb;">',
            '<tr><td align="center" style="padding:15px 8px;">',
            '<table class="email-container" width="680" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border-radius:12px;overflow:hidden;">',
            '<tr><td>',
            get_header_html(),
            bloco_destaques(news),
            bloco_insights(quotes_gl),
            bloco_radar(radar_br, radar_us),
            bloco_cotacoes(quotes_br, quotes_gl),
            get_footer_html([dest_email]),
            '</td></tr></table>',
            '</td></tr></table></body></html>'
        ]
        html_email = "\n".join(parts)

        try:
            enviar_email_smtp([dest_email], assunto, html_email)
        except Exception as e:
            print(f"[ERRO] Falha ao enviar para {dest_email}: {e}")
        time.sleep(0.5)

    # Salva preview para debug
    preview_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "preview_v2.html")
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(html_email)
    print(f"[+] Preview salvo: {preview_path}")

if __name__ == "__main__":
    try:
        main()
        print("[OK] Boletim v2 finalizado.")
    except Exception as e:
        print("[ERRO]", e)
        traceback.print_exc()
