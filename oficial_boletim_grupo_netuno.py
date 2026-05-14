# -*- coding: utf-8 -*-
r"""
Netuno — Boletim de Mercado — v9.6b
- Envio automático via Microsoft Graph (Device Code Flow) — sem client_secret
- Assunto: "Boletim Grupo Netuno — dd-mm-yyyy"
- Destaques no topo
- Cotações (B3 + globais) com preço e variação
- Internacional (Reuters + Bloomberg /markets) 4–6 notícias traduzidas e link "continuar lendo"
- Links clicáveis em todas as notícias
- Log Excel sem timezone-aware
- Anexo opcional (se existir)

Requisitos:
  pip install requests feedparser beautifulsoup4 readability-lxml msal pandas openpyxl
"""

import os
import re
import io
import time
import base64
import html
import math
import json
import traceback
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Optional

import requests
from bs4 import BeautifulSoup
import feedparser
from readability import Document
import msal
import pandas as pd

# ================== CONFIG ==================
CLIENT_ID  = "44062437-451e-47fa-b742-3a91fd9faf94"  # App (client) ID
TENANT_ID  = "9d049248-9621-4bf0-baa4-e8f9a1c4d019"  # Directory (tenant) ID
SCOPES     = ["https://graph.microsoft.com/Mail.Send"]

DESTINATARIOS = [
    "gustavoportugalhamer@gmail.com",
    "arthur.hamer@r2fcapital.com.br",
    "arthur@netunoinvestimentos.com.br",
    "carloshferreira75@hotmail.com"
]

ASSUNTO_PREFIXO = "Boletim Grupo Netuno"
TIMEOUT      = 18
SLEEP        = 0.35
JANELA_DIAS  = 2
MAX_BULLETS  = 12
MAX_PER_FONTE= 8
TZ_BR        = timezone(timedelta(hours=-3))  # America/Sao_Paulo
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetunoBot/1.9"}

# Caminhos locais
BASE_DIR  = r"C:\Users\gusta\OneDrive - R2F Capital\Automacao\BOLETIM(NEWS)"
ARQ_LOG   = os.path.join(BASE_DIR, "log_boletim_netuno.xlsx")
ANEXO_PDF = os.path.join(BASE_DIR, "Convite - Grupo Netuno.pdf")  # opcional

# ===== Cotações (Yahoo Finance) =====
B3_TICKERS = {
    "PETR4.SA":"Petrobras PN (PETR4)", "VALE3.SA":"Vale ON (VALE3)", "ITUB4.SA":"Itaú Unibanco PN (ITUB4)",
    "BBDC4.SA":"Bradesco PN (BBDC4)", "BBAS3.SA":"Banco do Brasil ON (BBAS3)", "ABEV3.SA":"Ambev ON (ABEV3)",
    "WEGE3.SA":"WEG ON (WEGE3)", "SUZB3.SA":"Suzano ON (SUZB3)", "B3SA3.SA":"B3 ON (B3SA3)", "GGBR4.SA":"Gerdau PN (GGBR4)"
}
GLOBAL_TICKERS = {
    # Ações globais (preço + %)
    "AAPL":"Apple", "MSFT":"Microsoft", "NVDA":"Nvidia", "AMZN":"Amazon",
    "META":"Meta", "GOOGL":"Alphabet",
    # Índices
    "ES=F":"S&P 500 Futuro", "NQ=F":"Nasdaq 100 Futuro",
    # Commodities/cripto/câmbio
    "CL=F":"Petróleo WTI", "BZ=F":"Petróleo Brent", "GC=F":"Ouro", "BTC-USD":"Bitcoin",
    "BRL=X":"USD/BRL"
}

# ===== Fontes (Internacional foca em Reuters + Bloomberg) =====
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

    # Internacional — foco
    {"name":"Reuters Markets", "cat":"Internacional", "rss":["https://www.reuters.com/markets/rss"]},
    {"name":"Bloomberg", "cat":"Internacional", "rss":[], "html":[
        {"url":"https://www.bloomberg.com/markets", "sel":"a[href*='/news/']", "attr":"href"}]},
]

# ================= HTTP helpers =================
def http_get(url: str, **kw) -> requests.Response:
    return requests.get(url, timeout=TIMEOUT, headers=HEADERS, **kw)

def norm_url(base: str, href: str) -> str:
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"):  return requests.compat.urljoin(base, href)
    return href

# =================== Tradução (leve) ===================
def translate_to_pt(text: str) -> str:
    """Google translate endpoint não-oficial (fallback: texto original)."""
    try:
        if not text: return ""
        url = "https://translate.googleapis.com/translate_a/single"
        params = {"client": "gtx", "sl": "auto", "tl": "pt", "dt": "t", "q": text}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        data = r.json()
        return "".join(seg[0] for seg in data[0])
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
    # v7 em chunks
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
        time.sleep(0.2)
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
                    time.sleep(0.2)
    # fallback por símbolo para campos faltantes
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
    b3 = "<h3 style='margin:12px 0 6px;'>📈 Brasil — Principais Ações</h3>" \
         f"<p style='margin:4px 0 10px'>{'<br>'.join(linhas(br, B3_TICKERS))}</p>"
    glob = "<h3 style='margin:8px 0 6px;'>🌎 Global — Índices e Ações</h3>" \
         f"<p style='margin:4px 0 10px'>{'<br>'.join(linhas(gl, GLOBAL_TICKERS))}</p>"
    return b3 + glob

# ================= News: Reuters + Bloomberg (Internacional) =================
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

def parse_bloomberg_markets() -> List[Dict]:
    """Captura links de /markets; dá timestamp de agora para manter recente."""
    out = []
    try:
        url = "https://www.bloomberg.com/markets"
        r = http_get(url); r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        links = []
        for a in soup.select("a[href*='/news/']")[:25]:
            href = a.get("href") or ""
            if not href: continue
            links.append(norm_url(url, href))
        # dedup
        seen, uniq = set(), []
        for h in links:
            if h not in seen:
                seen.add(h); uniq.append(h)
        for lk in uniq[:10]:
            try:
                rr = http_get(lk); rr.raise_for_status()
                ss = BeautifulSoup(rr.text, "lxml")
                title = (ss.title.get_text(strip=True) if ss.title else lk)
                out.append({
                    "fonte":"Bloomberg",
                    "titulo":title,
                    "link":lk,
                    "quando":datetime.now(TZ_BR),
                    "resumo_feed":""
                })
                time.sleep(0.15)
            except Exception:
                continue
    except Exception:
        pass
    return out

def extract_article_text(url: str) -> str:
    try:
        r = http_get(url); r.raise_for_status()
        doc = Document(r.text)
        html_main = doc.summary(html_partial=True)
        txt = BeautifulSoup(html_main, "lxml").get_text(" ", strip=True)
        txt = re.sub(r"\s+", " ", txt).strip()
        # filtro básico de paywall/bloqueio
        bad = ["access to this page has been denied", "enable javascript", "subscribe to read", "assine para ler"]
        tl = txt.lower()
        if any(b in tl for b in bad): return ""
        return txt
    except Exception:
        try:
            soup = BeautifulSoup(r.text, "lxml")
            art = soup.find("article") or soup
            txt = art.get_text(" ", strip=True)
            txt = re.sub(r"\s+", " ", txt).strip()
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

def collect_news_internacional() -> List[Dict]:
    """Gera pool Internacional mesclando Reuters (RSS) + Bloomberg (HTML)."""
    pool = []
    # Reuters RSS
    for u in ["https://www.reuters.com/markets/rss"]:
        pool.extend(parse_rss(u))
        time.sleep(SLEEP)
    # Bloomberg /markets
    pool.extend(parse_bloomberg_markets())

    # dedup por (titulo, fonte) + ordena por quando desc
    seen, uniq = set(), []
    for it in sorted(pool, key=lambda x: x.get("quando", datetime.now(TZ_BR)), reverse=True):
        key = (it.get("titulo","").strip(), it.get("fonte","").strip())
        if key in seen: continue
        seen.add(key); uniq.append(it)
    return uniq[:12]

# ================= Montagem =================
def mk_bullets(items: List[Dict], traduzir=False, limit=6) -> List[str]:
    rows = []
    for it in items[:limit]:
        quando = it.get("quando", datetime.now(TZ_BR)).strftime("%d/%m %H:%M")
        fonte  = html.escape(it.get("fonte") or "")
        title_raw = re.sub(r"\s+"," ", (it.get("titulo") or "").strip())
        link  = it.get("link","")
        txt = extract_article_text(link)
        if not txt:
            raw = BeautifulSoup(it.get("resumo_feed",""), "lxml").get_text(" ", strip=True)
            txt = re.sub(r"\s+"," ", raw)
        if not txt:
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

def mk_destaques(news_all: Dict[str, List[Dict]], intl_pool: List[Dict], total:int=4) -> str:
    """Destaques misturando BR, Empresas e Internacional traduzido."""
    pool = []
    pool.extend(news_all.get("Brasil", [])[:2])
    pool.extend(news_all.get("Empresas", [])[:1])
    pool.extend(intl_pool[:3])
    bullets = []
    for it in pool[:total]:
        link = it.get("link","")
        title = re.sub(r"\s+"," ", (it.get("titulo") or "").strip())
        if (it.get("fonte") or it.get("cat")) in ("Bloomberg","Reuters Markets","Internacional"):
            title = translate_to_pt(title)
        bullets.append(f"• <a href='{link}' style='color:#0047ab;'>{html.escape(title)}</a>")
    if not bullets:
        return ""
    return "<h3 style='margin:0 0 6px;'>📌 Principais destaques</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def bloco_brasil(items: List[Dict]) -> str:
    bullets = mk_bullets(items, traduzir=False, limit=6)
    if not bullets: return ""
    return "<h3 style='margin:16px 0 6px;'>🇧🇷 Brasil</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def bloco_empresas(items: List[Dict]) -> str:
    bullets = mk_bullets(items, traduzir=False, limit=6)
    if not bullets: return ""
    return "<h3 style='margin:16px 0 6px;'>🏢 Empresas</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def bloco_internacional_traduzido(intl_pool: List[Dict]) -> str:
    bullets = mk_bullets(intl_pool, traduzir=True, limit=6)
    if not bullets: return ""
    nota = "<p style='color:#777;font-size:12px;margin:6px 0 0;'>(Tradução automática via Google)</p>"
    return (
        "<h3 style='margin:16px 0 6px;'>🌍 Internacional (Traduzido)</h3>"
        f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>" + nota
    )

# ======= Coleta geral BR/Empresas (aproveita seu pipeline antigo) =======
def collect_news_geral() -> Dict[str, List[Dict]]:
    out: Dict[str, List[Dict]] = {"Brasil":[], "Internacional":[], "Empresas":[]}
    for src in SOURCES:
        if src["cat"] == "Internacional":
            # Internacional é tratado à parte
            continue
        # RSS
        for u in src.get("rss", []):
            try:
                items = parse_rss(u)
                for it in items:
                    it["srcname"] = src["name"]; it["cat"] = src["cat"]
                out[src["cat"]].extend(items); time.sleep(SLEEP)
            except Exception:
                pass
        # HTML
        for cfg in src.get("html", []):
            try:
                links = []
                r = http_get(cfg["url"]); r.raise_for_status()
                soup = BeautifulSoup(r.text, "lxml")
                for a in soup.select(cfg["sel"])[:20]:
                    href = a.get(cfg["attr"]) or ""
                    if not href: continue
                    links.append(norm_url(cfg["url"], href))
                # dedup
                seen, uniq = set(), []
                for h in links:
                    if h not in seen:
                        seen.add(h); uniq.append(h)
                for lk in uniq[:MAX_PER_FONTE]:
                    try:
                        rr = http_get(lk); rr.raise_for_status()
                        ss = BeautifulSoup(rr.text, "lxml")
                        title = (ss.title.get_text(strip=True) if ss.title else lk)
                        out[src["cat"]].append({
                            "fonte":src["name"], "titulo":title, "link":lk,
                            "quando":datetime.now(TZ_BR), "resumo_feed":"", "srcname":src["name"], "cat":src["cat"]
                        })
                        time.sleep(0.15)
                    except Exception:
                        continue
                time.sleep(SLEEP)
            except Exception:
                pass
    # dedup + corte
    for k in ("Brasil","Empresas"):
        seen, uniq = set(), []
        for it in sorted(out[k], key=lambda x: x.get("quando", datetime.now(TZ_BR)), reverse=True):
            key = (it.get("titulo",""), it.get("fonte",""))
            if key in seen: continue
            seen.add(key); uniq.append(it)
        out[k] = uniq[:MAX_BULLETS]
    return out

# ======================= Microsoft Graph =======================
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def obter_token_device_code() -> dict:
    """Device Code Flow (delegated). Não requer client_secret."""
    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"❌ Erro MSAL: {json.dumps(flow, ensure_ascii=False)}")
    print(f"🔐 Para autorizar, acesse {flow['verification_uri']} e entre com o código: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise Exception(f"❌ Erro MSAL: {result.get('error_description','Falha ao obter token')}")
    return result

def enviar_email_graph(dest: List[str], assunto: str, html_corpo: str, anexos: Optional[List[dict]]=None):
    token = obter_token_device_code()
    headers = {
        "Authorization": f"Bearer {token['access_token']}",
        "Content-Type": "application/json"
    }
    msg = {
        "message": {
            "subject": assunto,
            "body": {"contentType": "HTML", "content": html_corpo},
            "toRecipients": [{"emailAddress": {"address": d}} for d in dest],
        },
        "saveToSentItems": True
    }
    if anexos:
        msg["message"]["attachments"] = anexos
    url = f"{GRAPH_BASE}/me/sendMail"
    r = requests.post(url, headers=headers, json=msg, timeout=30)
    if r.status_code not in (202, 200):
        raise Exception(f"❌ Falha Graph sendMail: {r.status_code} — {r.text}")

def montar_anexo_pdf(path_pdf: str) -> Optional[dict]:
    if not os.path.isfile(path_pdf):
        return None
    with open(path_pdf, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("utf-8")
    nome = os.path.basename(path_pdf)
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": nome, "contentBytes": content_b64, "contentType": "application/pdf"
    }

# ======================= Log Excel =======================
def registrar_log(status_msg: str):
    agora = datetime.now().replace(tzinfo=None)  # timezone-unaware
    linha = {
        "timestamp": agora.strftime("%Y-%m-%d %H:%M:%S"),
        "status": status_msg
    }
    df = pd.DataFrame([linha])
    if os.path.isfile(ARQ_LOG):
        try:
            antigo = pd.read_excel(ARQ_LOG, engine="openpyxl")
        except Exception:
            antigo = pd.DataFrame(columns=df.columns)
        # evita FutureWarning juntando apenas colunas existentes
        colunas = list(set(antigo.columns) | set(df.columns))
        antigo = antigo.reindex(columns=colunas)
        df = df.reindex(columns=colunas)
        df = pd.concat([antigo, df], ignore_index=True)
    # salva
    df.to_excel(ARQ_LOG, index=False)

# ======================= MAIN =======================
def main():
    # Coleta BR/Empresas
    news_geral = collect_news_geral()
    # Coleta Internacional (Reuters + Bloomberg)
    intl_pool = collect_news_internacional()

    # Título e cabeçalho
    hoje_ddmmyyyy = datetime.now(TZ_BR).strftime("%d-%m-%Y")
    topo = (
        "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;'>"
        f"<h2 style='margin:0;'>Boletim — Internacional • Brasil • Empresas</h2>"
        f"<p style='margin:6px 0 8px;color:#444;'>Gerado em {datetime.now(TZ_BR).strftime('%d/%m/%Y %H:%M')} (UTC-03)</p>"
    )

    destaques = mk_destaques(news_geral, intl_pool, total=4)
    sep = "<hr style='border:none;border-top:1px solid #e1e1e1;margin:8px 0;'>"

    # Cotações uma única vez
    cot = bloco_cotacoes()
    # Blocos
    intl = bloco_internacional_traduzido(intl_pool)
    br   = bloco_brasil(news_geral.get("Brasil", []))
    emp  = bloco_empresas(news_geral.get("Empresas", []))

    rodape = (
        "<hr style='border:none;border-top:1px solid #e1e1e1;margin:12px 0;'>"
        "<p style='color:#777;font-size:12px;'>"
        "Fontes: InfoMoney, Valor Investe, Mais Retorno, Yubb, B3 (Bora Investir), Bloomberg Línea BR, Brazil Journal, "
        "Reuters e Bloomberg (Internacional). Todas as notícias incluem link para leitura no site original."
        "</p></div>"
    )

    html_email = topo + (destaques or "") + sep + cot + intl + br + emp + rodape
    assunto = f"{ASSUNTO_PREFIXO} — {hoje_ddmmyyyy}"

    # Anexo opcional
    anexos = []
    anex = montar_anexo_pdf(ANEXO_PDF)
    if anex: anexos.append(anex)

    # Envio
    enviar_email_graph(DESTINATARIOS, assunto, html_email, anexos if anexos else None)
    registrar_log("✅ Enviado com sucesso")

if __name__ == "__main__":
    try:
        main()
        print("✅ Boletim enviado.")
    except Exception as e:
        print("❌ Falha no boletim:", e)
        registrar_log(f"❌ Erro: {e}")
        traceback.print_exc()
