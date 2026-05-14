# -*- coding: utf-8 -*-
r"""
Netuno — Boletim de Mercado — v9.5b
- Assunto: "Boletim Grupo Netuno — dd-mm-yyyy"
- Anexo PDF (Convite - Grupo Netuno.pdf)
- Envio via Microsoft Graph (Device Flow, sem client_secret), com cache
- Log Excel sem timezone (append)
- Destaques no topo; Cotações (uma única vez); Internacional traduzido com links
- Yahoo Finance v7 + fallback v8/chart
"""

import os
import re
import io
import time
import html
import json
import math
import traceback
from datetime import datetime, timedelta, timezone
from typing import List, Dict

import requests
from bs4 import BeautifulSoup
import feedparser
from readability import Document

import pandas as pd

# ================== CONFIG GERAL ==================
TZ_BR = timezone(timedelta(hours=-3))  # America/Sao_Paulo

DESTINATARIOS = [
    "gustavoportugalhamer@gmail.com",
    "arthur.hamer@r2fcapital.com.br",
    "arthur@netunoinvestimentos.com.br",
    "carloshferreira75@hotmail.com"
]

ASSUNTO_PREFIXO = "Boletim Grupo Netuno"
ANEXO_PDF = r"C:\Users\gusta\OneDrive - R2F Capital\Automacao\Convite - Grupo Netuno.pdf"
ARQ_LOG = r"C:\Users\gusta\OneDrive - R2F Capital\Automacao\BOLETIM(NEWS)\log_boletim_netuno.xlsx"

TIMEOUT = 18
SLEEP = 0.35
JANELA_DIAS = 2
MAX_BULLETS = 12
MAX_PER_FONTE = 8

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetunoBot/1.9"}

# ================== MICROSOFT GRAPH (Device Flow) ==================
# App (Public Client) — habilitar "Allow public client flows" no Azure
MS_CLIENT_ID = "44062437-451e-47fa-b742-3a91fd9faf94"     # App (client) ID
MS_TENANT_ID = "9d049248-9621-4bf0-baa4-e8f9a1c4d019"     # Directory (tenant) ID
MS_SCOPES = ["https://graph.microsoft.com/Mail.Send"]
TOKEN_CACHE_ARQ = os.path.join(os.path.expanduser("~"), ".netuno_graph_token_cache.json")

# ================== COTAÇÕES (Yahoo Finance) ==================
B3_TICKERS = {
    "PETR4.SA": "Petrobras PN (PETR4)",
    "VALE3.SA": "Vale ON (VALE3)",
    "ITUB4.SA": "Itaú Unibanco PN (ITUB4)",
    "BBDC4.SA": "Bradesco PN (BBDC4)",
    "BBAS3.SA": "Banco do Brasil ON (BBAS3)",
    "ABEV3.SA": "Ambev ON (ABEV3)",
    "WEGE3.SA": "WEG ON (WEGE3)",
    "SUZB3.SA": "Suzano ON (SUZB3)",
    "B3SA3.SA": "B3 ON (B3SA3)",
    "GGBR4.SA": "Gerdau PN (GGBR4)"
}
GLOBAL_TICKERS = {
    "AAPL": "Apple", "MSFT": "Microsoft", "NVDA": "Nvidia", "AMZN": "Amazon",
    "TSLA": "Tesla", "META": "Meta", "GOOGL": "Alphabet", "JPM": "JPMorgan",
    "ES=F": "S&P 500 Futuro", "NQ=F": "Nasdaq 100 Futuro", "^STOXX50E": "Euro Stoxx 50",
    "^FTSE": "FTSE 100", "^FCHI": "CAC 40", "^N225": "Nikkei 225",
    "CL=F": "Petróleo WTI", "BZ=F": "Petróleo Brent", "GC=F": "Ouro", "BTC-USD": "Bitcoin",
    "BRL=X": "USD/BRL", "EURUSD=X": "EUR/USD"
}

# ================== FONTES DE NOTÍCIAS ==================
SOURCES = [
    # Brasil
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

    # Internacional
    {"name":"Bloomberg", "cat":"Internacional", "rss":[], "html":[
        {"url":"https://www.bloomberg.com/markets", "sel":"a[href*='/news/']", "attr":"href"}]},
    {"name":"Reuters Markets", "cat":"Internacional", "rss":["https://www.reuters.com/markets/rss"]},
    {"name":"Financial Times", "cat":"Internacional", "rss":["https://www.ft.com/world/americas/rss"]},
    {"name":"Seeking Alpha", "cat":"Internacional", "rss":["https://seekingalpha.com/market_currents.xml"]},
]

# ================== HELPERS HTTP ==================
def http_get(url: str, **kw) -> requests.Response:
    return requests.get(url, timeout=TIMEOUT, headers=HEADERS, **kw)

def norm_url(base: str, href: str) -> str:
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"):  return requests.compat.urljoin(base, href)
    return href

# ================== TRADUÇÃO LEVE ==================
def translate_to_pt(text: str) -> str:
    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {"client": "gtx", "sl": "auto", "tl": "pt", "dt": "t", "q": text}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        return "".join(seg[0] for seg in r.json()[0])
    except Exception:
        return text

# ================== YAHOO FINANCE ==================
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
    hosts = ["query1.finance.yahoo.com", "query2.finance.yahoo.com"]
    # v7 em chunks + fallback v8
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
                time.sleep(0.5)
        if not ok:
            for sym in syms:
                for h in hosts:
                    try:
                        c = _yahoo_chart_one(sym, h)
                        if c:
                            out[sym] = {
                                "symbol": sym,
                                "regularMarketPrice": c.get("price"),
                                "regularMarketChangePercent": c.get("pct")
                            }
                            break
                    except Exception:
                        time.sleep(0.2)
        time.sleep(0.25)
    # completar faltantes via v8
    for sym in symbols:
        d = out.get(sym, {})
        if not d or d.get("regularMarketPrice") is None:
            for h in hosts:
                try:
                    c = _yahoo_chart_one(sym, h)
                    if c:
                        d = d or {"symbol": sym}
                        d["regularMarketPrice"] = d.get("regularMarketPrice") or c.get("price")
                        if d.get("regularMarketChangePercent") is None:
                            d["regularMarketChangePercent"] = c.get("pct")
                        out[sym] = d
                        break
                except Exception:
                    pass
    return out

def fmt_price(sym: str, v) -> str:
    if v is None: return "—"
    try:
        return ("R$ " if sym.endswith(".SA") else "US$ ") + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"

def fmt_pct(v) -> str:
    if v is None: return "—"
    try:
        return "{:+.2f}%".format(float(v)).replace(".", ",")
    except Exception:
        return "—"

def bloco_cotacoes() -> str:
    br = yahoo_quote(list(B3_TICKERS.keys()))
    gl = yahoo_quote(list(GLOBAL_TICKERS.keys()))

    def linhas(payload, labels):
        out = []
        for sym, label in labels.items():
            d = payload.get(sym, {})
            px = d.get("regularMarketPrice")
            pct = d.get("regularMarketChangePercent")
            out.append(f"• {label}: {fmt_pct(pct)} a {fmt_price(sym, px)}")
        return out

    b3 = "<h3 style='margin:16px 0 6px;'>📈 Brasil — Principais Ações</h3>" \
         f"<p style='margin:4px 0 10px'>{'<br>'.join(linhas(br, B3_TICKERS))}</p>"
    glob = "<h3 style='margin:8px 0 6px;'>🌎 Global — Índices e Ações</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(linhas(gl, GLOBAL_TICKERS))}</p>"
    return b3 + glob

# ================== NOTÍCIAS ==================
def is_recent(dt: datetime) -> bool:
    return dt >= datetime.now(TZ_BR) - timedelta(days=JANELA_DIAS)

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
            if not is_recent(quando): continue
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

def summarize(text: str, max_frases=3, max_chars=420) -> str:
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
        # RSS
        for u in src.get("rss", []):
            items = parse_rss(u)
            for it in items:
                it["srcname"] = src["name"]; it["cat"] = src["cat"]
            out[src["cat"]].extend(items); time.sleep(SLEEP)
        # HTML
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
    # dedup + corte
    for k in out.keys():
        seen, uniq = set(), []
        for it in sorted(out[k], key=lambda x:x.get("quando", datetime.now(TZ_BR)), reverse=True):
            key = (it.get("titulo",""), it.get("fonte",""))
            if key in seen: continue
            seen.add(key); uniq.append(it)
        out[k] = uniq[:MAX_BULLETS]
    return out

# ================== MONTAGEM ==================
def mk_bullets(items: List[Dict], traduzir=False, limit=6) -> List[str]:
    rows = []
    for it in items[:limit]:
        quando = it.get("quando", datetime.now(TZ_BR)).strftime("%d/%m %H:%M")
        fonte  = html.escape(it.get("fonte") or it.get("srcname") or "")
        title_raw = re.sub(r"\s+"," ", (it.get("titulo") or "").strip())
        link  = it.get("link","")
        txt = extract_article_text(link)
        if not txt:
            raw = BeautifulSoup(it.get("resumo_feed",""), "lxml").get_text(" ", strip=True)
            txt = re.sub(r"\s+"," ", raw)
        if not txt or "Access to this page has been denied" in txt:
            continue
        resumo = summarize(txt, max_frases=3, max_chars=420)
        if traduzir:
            title = html.escape(translate_to_pt(title_raw))
            resumo = html.escape(translate_to_pt(resumo))
        else:
            title = html.escape(title_raw)
            resumo = html.escape(resumo)
        rows.append(
            f"• <b><a href='{link}' style='color:#0047ab;'>{title}</a></b> — {resumo} "
            f"<a href='{link}' style='color:#0047ab;'>continuar lendo</a> "
            f"<span style='color:#777'>({fonte}, {quando})</span>"
        )
    return rows

def mk_destaques(news: Dict[str, List[Dict]], total:int=4) -> str:
    pool = []
    pool.extend(news.get("Brasil", [])[:2])
    pool.extend(news.get("Empresas", [])[:1])
    pool.extend(news.get("Internacional", [])[:3])
    bullets = []
    for it in pool[:total]:
        link = it.get("link","")
        title = re.sub(r"\s+"," ", (it.get("titulo") or "").strip())
        if it.get("cat") == "Internacional":
            title = translate_to_pt(title)
        bullets.append(f"• <a href='{link}' style='color:#0047ab;'>{html.escape(title)}</a>")
    if not bullets:
        return ""
    return "<h3 style='margin:0 0 6px;'>📌 Principais destaques</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def bloco_internacional(items: List[Dict]) -> str:
    bullets = mk_bullets(items, traduzir=True, limit=6)
    if not bullets: return ""
    nota = "<p style='color:#777;font-size:12px;margin:6px 0 0;'>(Tradução automática pelo Google)</p>"
    return (
        "<h3 style='margin:16px 0 6px;'>🌎 INTERNACIONAL</h3>"
        f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>" + nota
    )

def bloco_brasil(items: List[Dict]) -> str:
    bullets = mk_bullets(items, traduzir=False, limit=6)
    if not bullets: return ""
    return "<h3 style='margin:16px 0 6px;'>🇧🇷 BRASIL</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def bloco_empresas(items: List[Dict]) -> str:
    bullets = mk_bullets(items, traduzir=False, limit=6)
    if not bullets: return ""
    return "<h3 style='margin:16px 0 6px;'>🏢 EMPRESAS</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def corpo_email(news: Dict[str, List[Dict]]) -> str:
    topo = (
        "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;'>"
        "<h2 style='margin:0;'>Boletim — Internacional • Brasil • Empresas</h2>"
        f"<p style='margin:6px 0 8px;color:#444;'>Gerado em {datetime.now(TZ_BR).strftime('%d/%m/%Y %H:%M')} (UTC-03)</p>"
    )
    destaques = mk_destaques(news, total=4)
    sep = "<hr style='border:none;border-top:1px solid #e1e1e1;margin:8px 0;'>"
    cot = bloco_cotacoes()                      # aparece UMA VEZ no topo
    intl = bloco_internacional(news.get("Internacional", []))
    br   = bloco_brasil(news.get("Brasil", []))
    emp  = bloco_empresas(news.get("Empresas", []))
    rodape = (
        "<hr style='border:none;border-top:1px solid #e1e1e1;margin:12px 0;'>"
        "<p style='color:#777;font-size:12px;'>"
        "Fontes: InfoMoney, Valor Investe, Mais Retorno, Yubb, Bora Investir/B3, "
        "Bloomberg Línea Brasil, Brazil Journal, Bloomberg, Reuters Markets, Financial Times e Seeking Alpha. "
        "Todas as notícias incluem link para leitura no site original."
        "</p></div>"
    )
    return topo + (destaques or "") + sep + cot + intl + br + emp + rodape

# ================== GRAPH: AUTH & SEND ==================
def _load_cache():
    if os.path.exists(TOKEN_CACHE_ARQ):
        try:
            with open(TOKEN_CACHE_ARQ, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _save_cache(data: dict):
    try:
        with open(TOKEN_CACHE_ARQ, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass

def obter_token_device_flow() -> str:
    """
    Device Flow sem msal (para evitar dependência): usa endpoints OAuth2 v2.0.
    Passo 1: solicitar user_code; Passo 2: polling até authorized; salva refresh_token.
    """
    base = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0"
    # Tenta refresh_token do cache
    cache = _load_cache()
    if "refresh_token" in cache:
        data = {
            "client_id": MS_CLIENT_ID,
            "scope": " ".join(MS_SCOPES),
            "grant_type": "refresh_token",
            "refresh_token": cache["refresh_token"]
        }
        r = requests.post(base + "/token", data=data, timeout=30)
        if r.ok:
            tok = r.json()
            cache["access_token"] = tok.get("access_token")
            if tok.get("refresh_token"):
                cache["refresh_token"] = tok.get("refresh_token")
            _save_cache(cache)
            return cache["access_token"]
        else:
            # limpa refresh inválido
            cache.pop("refresh_token", None)
            _save_cache(cache)

    # Device code
    data = {
        "client_id": MS_CLIENT_ID,
        "scope": " ".join(MS_SCOPES)
    }
    r = requests.post(base + "/devicecode", data=data, timeout=30)
    r.raise_for_status()
    dc = r.json()
    user_code = dc["user_code"]
    verification_uri = dc["verification_uri"]
    device_code = dc["device_code"]
    interval = int(dc.get("interval", 5))
    print(f"\n⚠️ Autorização necessária uma única vez:")
    print(f"1) Acesse: {verification_uri}")
    print(f"2) Digite o código: {user_code}")
    print(f"Aguarde, validando autorização...\n")

    # Polling
    while True:
        time.sleep(interval)
        token_req = {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": MS_CLIENT_ID,
            "device_code": device_code
        }
        tr = requests.post(base + "/token", data=token_req, timeout=30)
        if tr.ok:
            tok = tr.json()
            access_token = tok.get("access_token")
            refresh_token = tok.get("refresh_token")
            cache = {"access_token": access_token}
            if refresh_token:
                cache["refresh_token"] = refresh_token
            _save_cache(cache)
            return access_token
        else:
            err = tr.json().get("error")
            if err in ("authorization_pending", "slow_down"):
                continue
            raise Exception(f"Falha no Device Flow: {tr.text}")

def enviar_email_graph(destinatarios: List[str], assunto: str, html_corpo: str, anexo_pdf_path: str):
    token = obter_token_device_flow()
    # Monta a mensagem
    msg = {
        "message": {
            "subject": assunto,
            "body": {"contentType": "HTML", "content": html_corpo},
            "toRecipients": [{"emailAddress": {"address": d}} for d in destinatarios],
        },
        "saveToSentItems": "true"
    }

    # Anexo (se existir)
    if anexo_pdf_path and os.path.exists(anexo_pdf_path):
        with open(anexo_pdf_path, "rb") as f:
            arq_bytes = f.read()
        import base64
        anex = {
            "@odata.type": "#microsoft.outlookservices.FileAttachment",
            "name": os.path.basename(anexo_pdf_path),
            "contentType": "application/pdf",
            "contentBytes": base64.b64encode(arq_bytes).decode("utf-8")
        }
        msg["message"]["attachments"] = [anex]

    # Envia
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, data=json.dumps(msg), timeout=30)
    if not r.ok:
        raise Exception(f"Graph sendMail falhou: {r.status_code} - {r.text}")

# ================== LOG ==================
def registrar_log(status: str):
    agora_local = datetime.now(TZ_BR)
    # sem timezone para Excel
    agora_naive = datetime(agora_local.year, agora_local.month, agora_local.day,
                           agora_local.hour, agora_local.minute, agora_local.second)
    linha = {
        "data_envio": agora_naive,
        "status": status
    }
    df = pd.DataFrame([linha])
    if os.path.exists(ARQ_LOG):
        try:
            antigo = pd.read_excel(ARQ_LOG)
            df = pd.concat([antigo, df], ignore_index=True)
        except Exception:
            # Se planilha estiver corrompida, recria
            pass
    os.makedirs(os.path.dirname(ARQ_LOG), exist_ok=True)
    df.to_excel(ARQ_LOG, index=False)

# ================== MAIN ==================
def main():
    news = collect_news()
    corpo = corpo_email(news)

    # Assunto com data dd-mm-yyyy
    assunto = f"{ASSUNTO_PREFIXO} — {datetime.now(TZ_BR).strftime('%d-%m-%Y')}"
    enviar_email_graph(DESTINATARIOS, assunto, corpo, ANEXO_PDF)
    registrar_log("✅ Enviado")

if __name__ == "__main__":
    try:
        main()
        print("✅ Boletim enviado com sucesso.")
    except Exception as e:
        try:
            registrar_log(f"❌ Erro: {e}")
        except Exception:
            pass
        print("❌ Falha no boletim:", e)
        traceback.print_exc()
