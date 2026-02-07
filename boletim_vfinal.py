# -*- coding: utf-8 -*-
r"""
Netuno — Boletim de Mercado (modelo solicitado) — v7
Versão aprimorada (traduções, links, ações globais, fontes ajustadas)
--------------------------------------------
Envio automático via Outlook mantido.
"""

import re
import time
import html
import math
import argparse
import logging
import traceback
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Optional

import requests
from bs4 import BeautifulSoup
import feedparser
from readability import Document

# ================== CONFIG ==================
DESTINATARIOS = [
    "gustavoportugalhamer@hotmail.com",
    "gustavo.hamer@r2fcapital.com.br"
]

ASSUNTO_PREFIXO = "[Grupo Netuno] Boletim — Internacional • Brasil • Empresas"
TIMEOUT = 18
SLEEP = 0.35
JANELA_DIAS = 2
MAX_BULLETS = 12
MAX_PER_FONTE = 8
TZ_BR = timezone(timedelta(hours=-3))
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetunoBot/1.7"}

# ===== Yahoo Finance (cotações) =====
B3_TICKERS = {
    "PETR4.SA":"Petrobras PN (PETR4)", "VALE3.SA":"Vale ON (VALE3)", "ITUB4.SA":"Itaú Unibanco PN (ITUB4)",
    "BBDC4.SA":"Bradesco PN (BBDC4)", "BBAS3.SA":"Banco do Brasil ON (BBAS3)", "ABEV3.SA":"Ambev ON (ABEV3)",
    "WEGE3.SA":"WEG ON (WEGE3)", "SUZB3.SA":"Suzano ON (SUZB3)", "B3SA3.SA":"B3 ON (B3SA3)", "GGBR4.SA":"Gerdau PN (GGBR4)"
}
GLOBAL_TICKERS = {
    "AAPL":"Apple", "MSFT":"Microsoft", "AMZN":"Amazon", "TSLA":"Tesla",
    "NVDA":"Nvidia", "META":"Meta", "GOOGL":"Alphabet", "JPM":"JPMorgan",
    "ES=F":"S&P 500 Futuro", "NQ=F":"Nasdaq 100 Futuro", "^STOXX50E":"Euro Stoxx 50",
    "^FTSE":"FTSE 100", "^FCHI":"CAC 40", "^N225":"Nikkei 225",
    "CL=F":"Petróleo WTI", "BZ=F":"Petróleo Brent", "GC=F":"Ouro", "BTC-USD":"Bitcoin"
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

# ================= HTTP Helpers =================
def http_get(url, **kw):
    return requests.get(url, timeout=TIMEOUT, headers=HEADERS, **kw)

# ================= Yahoo Finance =================
def yahoo_quote(symbols):
    out = {}
    url = "https://query1.finance.yahoo.com/v7/finance/quote"
    for i in range(0, len(symbols), 10):
        r = http_get(url, params={"symbols": ",".join(symbols[i:i+10])})
        data = r.json().get("quoteResponse", {}).get("result", [])
        out.update({d["symbol"]: d for d in data if d.get("symbol")})
    return out

def fmt_price(sym, v):
    if not v: return "—"
    prefix = "R$" if sym.endswith(".SA") else "US$"
    return f"{prefix} {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_pct(v):
    if v is None: return "—"
    return f"{v:+.2f}%".replace(".", ",")

def bloco_cotacoes():
    br = yahoo_quote(list(B3_TICKERS.keys()))
    gl = yahoo_quote(list(GLOBAL_TICKERS.keys()))

    def linhas(payload, labels):
        out = []
        for sym, name in labels.items():
            d = payload.get(sym, {})
            px = d.get("regularMarketPrice")
            pct = d.get("regularMarketChangePercent")
            out.append(f"• {name} {fmt_pct(pct)} a {fmt_price(sym, px)}")
        return out

    b3 = "<h3>📈 Cotações — Brasil (B3)</h3><p>" + "<br>".join(linhas(br, B3_TICKERS)) + "</p>"
    glb = "<h3>🌐 Cotações — Lá fora / Commodities / Cripto</h3><p>" + "<br>".join(linhas(gl, GLOBAL_TICKERS)) + "</p>"
    return b3 + glb

# ================= News =================
def parse_rss(url):
    items = []
    try:
        f = feedparser.parse(url)
        fonte = f.feed.get("title", url)
        for e in f.entries[:MAX_PER_FONTE]:
            items.append({
                "fonte": fonte,
                "titulo": e.get("title", "").strip(),
                "link": e.get("link", ""),
                "resumo": e.get("summary", ""),
                "quando": datetime.now(TZ_BR)
            })
    except Exception as ex:
        logging.warning(f"RSS fail {url}: {ex}")
    return items

def extract_article_text(url):
    try:
        r = http_get(url)
        r.raise_for_status()
        doc = Document(r.text)
        soup = BeautifulSoup(doc.summary(html_partial=True), "lxml")
        txt = soup.get_text(" ", strip=True)
        return re.sub(r"\s+", " ", txt)
    except Exception:
        return ""

def summarize(text, max_chars=500):
    if not text: return ""
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_chars].rsplit(" ", 1)[0] + "…"

# Tradução via endpoint Google (sem libs externas)
def translate_to_pt(text):
    try:
        r = requests.get(
            "https://translate.googleapis.com/translate_a/single",
            params={"client":"gtx","sl":"auto","tl":"pt","dt":"t","q":text},
            timeout=10
        )
        result = r.json()
        return "".join([t[0] for t in result[0]])
    except Exception:
        return text

def mk_bullets(items, traduzir=False):
    rows = []
    for it in items[:6]:
        resumo = extract_article_text(it["link"]) or it["resumo"]
        resumo = summarize(resumo, 480)
        if traduzir:
            resumo = translate_to_pt(resumo)
        fonte = html.escape(it["fonte"])
        quando = it["quando"].strftime("%d/%m %H:%M")
        link = it["link"]
        rows.append(
            f"• <b><a href='{link}' style='color:#0047ab;'>{html.escape(it['titulo'])}</a></b> — {html.escape(resumo)} "
            f"<a href='{link}' style='color:#0047ab;'>continuar lendo</a> "
            f"<span style='color:#777'>({fonte}, {quando})</span>"
        )
    return rows

def collect_news():
    out = {"Brasil":[], "Internacional":[], "Empresas":[]}
    for s in SOURCES:
        for rss in s.get("rss", []):
            for n in parse_rss(rss):
                n["cat"] = s["cat"]
                n["fonte"] = s["name"]
                out[s["cat"]].append(n)
        time.sleep(SLEEP)
    return out

def bloco_internacional(items):
    bullets = mk_bullets(items, traduzir=True)
    if not bullets: return ""
    return "<h3>🌎 INTERNACIONAL</h3><p>" + "<br>".join(bullets) + "</p><p style='color:#777;font-size:12px;'>(Tradução automática pelo Google)</p>"

def bloco_brasil(items):
    bullets = mk_bullets(items)
    return "<h3>🇧🇷 BRASIL</h3><p>" + "<br>".join(bullets) + "</p>"

def bloco_empresas(items):
    bullets = mk_bullets(items)
    return "<h3>🏢 EMPRESAS</h3><p>" + "<br>".join(bullets) + "</p>"

# ================= Outlook =================
def enviar_email_outlook(to_addr, cc_addrs, bcc_addrs, assunto, html_corpo, MODE="SEND", from_hint=None, delay_seconds=20):
    import win32com.client as win32, pythoncom
    try:
        pythoncom.CoInitialize()
        outlook = win32.DispatchEx("Outlook.Application")
        session = outlook.GetNamespace("MAPI")
        try: session.Logon("", "", False, False)
        except Exception: pass
        mail = outlook.CreateItem(0)
        mail.Subject = assunto
        mail.HTMLBody = html_corpo
        if to_addr: mail.To = to_addr
        if cc_addrs: mail.CC = "; ".join(cc_addrs)
        if bcc_addrs: mail.BCC = "; ".join(bcc_addrs)
        if delay_seconds > 0:
            logging.info(f"Aguardando {delay_seconds}s antes do envio...")
            time.sleep(delay_seconds)
        if MODE.upper() == "DISPLAY_ONLY":
            mail.Display()
        elif MODE.upper() == "SAVE_DRAFT":
            mail.Save()
        else:
            mail.Send()
        logging.info("E-mail enviado com sucesso via Outlook.")
    except Exception as e:
        logging.error(f"Erro no envio Outlook: {e}")
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()

# ================= MAIN =================
def run(mode="SEND", from_hint=None, use_bcc=True, delay_seconds=20):
    news = collect_news()
    cot = bloco_cotacoes()
    intl = bloco_internacional(news["Internacional"])
    br = bloco_brasil(news["Brasil"])
    emp = bloco_empresas(news["Empresas"])

    html_email = (
        f"<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;'>"
        f"<h2>Boletim — Internacional • Brasil • Empresas</h2>"
        f"<p style='color:#444;'>Gerado em {datetime.now(TZ_BR).strftime('%d/%m/%Y %H:%M')}</p>"
        f"<hr>{cot}{intl}{br}{emp}"
        f"<hr><p style='color:#777;font-size:12px;'>"
        f"Fontes: InfoMoney, Valor Investe, Mais Retorno, Yubb, Bora Investir/B3, "
        f"Bloomberg Línea Brasil, Brazil Journal, Bloomberg, Reuters, Financial Times, Seeking Alpha."
        f"</p></div>"
    )

    assunto = f"{ASSUNTO_PREFIXO} — {datetime.now(TZ_BR).strftime('%Y-%m-%d')}"
    if use_bcc:
        to_addr = DESTINATARIOS[1]
        bcc_list = [DESTINATARIOS[0]]
        cc_list = []
    else:
        to_addr = DESTINATARIOS[0]
        cc_list = DESTINATARIOS[1:]
        bcc_list = []
    enviar_email_outlook(to_addr, cc_list, bcc_list, assunto, html_email, MODE=mode, from_hint=from_hint, delay_seconds=delay_seconds)

if __name__ == "__main__":
    try:
        ap = argparse.ArgumentParser(description="Boletim Netuno — Internacional • Brasil • Empresas")
        g = ap.add_mutually_exclusive_group()
        g.add_argument("--preview", action="store_true")
        g.add_argument("--save-draft", action="store_true")
        g.add_argument("--send", action="store_true")
        ap.add_argument("--from", dest="from_hint", default=None)
        ap.add_argument("--no-bcc", action="store_true")
        ap.add_argument("--delay", type=int, default=20)
        ap.add_argument("--log-path", default="netuno_boletim.log")
        args = ap.parse_args()

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[logging.FileHandler(args.log_path, encoding="utf-8"), logging.StreamHandler()]
        )

        mode = "SEND"
        if args.preview: mode = "DISPLAY_ONLY"
        elif args.save_draft: mode = "SAVE_DRAFT"

        run(mode=mode, from_hint=args.from_hint, use_bcc=(not args.no_bcc), delay_seconds=args.delay)
        logging.info("Boletim processado com sucesso.")
    except Exception as e:
        logging.error("Falha no boletim: %s", e)
        traceback.print_exc()
