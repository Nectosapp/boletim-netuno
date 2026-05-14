# -*- coding: utf-8 -*-
r"""
Netuno — Boletim de Mercado (modelo solicitado) — v2 com Fallback de Cotações e modos de envio

Fontes:
- BR: InfoMoney, Valor Investe, XP (conteúdos), Mais Retorno, Yubb, Bora Investir/B3,
      Bloomberg Línea Brasil, Investing.com (BR), Brazil Journal
- Internacional: Bloomberg, Reuters Markets, Financial Times, Seeking Alpha

Layout do e-mail (sem links):
  🌎 INTERNACIONAL -> parágrafo + bullets de índices/commodities/cripto
  🇧🇷 BRASIL -> bullets
  🏢 EMPRESAS -> bullets

Uso:
  py netuno_boletim_modelo.py            # envia (padrão)
  py netuno_boletim_modelo.py --preview  # abre o e-mail pra revisar
  py netuno_boletim_modelo.py --save-draft
  py netuno_boletim_modelo.py --send     # explicita envio
  py netuno_boletim_modelo.py --from "seu_email@dominio.com"
"""

import re
import time
import html
import math
import argparse
import traceback
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Optional

import requests
from bs4 import BeautifulSoup
import feedparser
from readability import Document

# ================== CONFIG ==================
DESTINATARIOS = [
    "gustavoportugalhamer@gmail.com",
]
ASSUNTO_PREFIXO = "[Grupo Netuno] Boletim — Internacional • Brasil • Empresas"
TIMEOUT      = 18
SLEEP        = 0.35
JANELA_DIAS  = 2       # últimas N dias
MAX_BULLETS  = 12      # por seção
MAX_PER_FONTE= 8
TZ_BR        = timezone(timedelta(hours=-3))  # America/Sao_Paulo

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetunoBot/1.3"}

# ===== Yahoo Finance (cotações) =====
B3_TICKERS = {
    "PETR4.SA":"Petrobras PN (PETR4)", "VALE3.SA":"Vale ON (VALE3)", "ITUB4.SA":"Itaú Unibanco PN (ITUB4)",
    "BBDC4.SA":"Bradesco PN (BBDC4)", "BBAS3.SA":"Banco do Brasil ON (BBAS3)", "ABEV3.SA":"Ambev ON (ABEV3)",
    "WEGE3.SA":"WEG ON (WEGE3)", "SUZB3.SA":"Suzano ON (SUZB3)", "B3SA3.SA":"B3 ON (B3SA3)", "GGBR4.SA":"Gerdau PN (GGBR4)"
}
GLOBAL_TICKERS = {
    "ES=F":"S&P 500 Futuro", "NQ=F":"Nasdaq 100 Futuro", "^STOXX50E":"Euro Stoxx 50", "^FTSE":"FTSE 100",
    "^FCHI":"CAC 40", "^N225":"Nikkei 225", "^HSI":"Hang Seng", "000001.SS":"Shanghai SE Comp.",
    "CL=F":"Petróleo WTI", "BZ=F":"Petróleo Brent", "GC=F":"Ouro", "BTC-USD":"Bitcoin"
}

# ===== Fontes: RSS (quando possível) + scraping leve =====
SOURCES = [
    # --- Brasil ---
    {"name":"InfoMoney", "cat":"Brasil", "rss":["https://www.infomoney.com.br/feed/"]},
    {"name":"Valor Investe", "cat":"Brasil", "rss":["https://valorinveste.globo.com/rss/"]},
    {"name":"XP Conteúdos", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://www.xpi.com.br/conteudos", "sel":"a[href*='/conteudos/']", "attr":"href"}]},
    {"name":"Mais Retorno", "cat":"Brasil", "rss":["https://maisretorno.com/feed"]},
    {"name":"Yubb", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://yubb.com.br/blog", "sel":"a.card", "attr":"href"}]},
    {"name":"Bora Investir / B3", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://borainvestir.b3.com.br/noticias", "sel":"a[href*='/noticias/']", "attr":"href"}]},
    {"name":"Bloomberg Línea BR", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://www.bloomberglinea.com.br/", "sel":"a[href*='/brasil/'], a[href*='/mercados/']", "attr":"href"}]},
    {"name":"Investing.com BR", "cat":"Brasil", "rss":[], "html":[
        {"url":"https://br.investing.com/news/stock-market-news", "sel":"a.title", "attr":"href"}]},
    {"name":"Brazil Journal", "cat":"Empresas", "rss":[], "html":[
        {"url":"https://braziljournal.com/", "sel":"article a", "attr":"href"}]},

    # --- Internacional ---
    {"name":"Bloomberg", "cat":"Internacional", "rss":[], "html":[
        {"url":"https://www.bloomberg.com/markets", "sel":"a[href*='/news/']", "attr":"href"}]},
    {"name":"Reuters Markets", "cat":"Internacional", "rss":["https://www.reuters.com/markets/rss"]},
    {"name":"Financial Times", "cat":"Internacional", "rss":["https://www.ft.com/world/americas/rss"]},
    {"name":"Seeking Alpha", "cat":"Internacional", "rss":["https://seekingalpha.com/market_currents.xml"]},
]

# ================= HTTP helpers =================
def http_get(url: str, **kw) -> requests.Response:
    return requests.get(url, timeout=TIMEOUT, headers=HEADERS, **kw)

def norm_url(base: str, href: str) -> str:
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"):  return requests.compat.urljoin(base, href)
    return href

# ========================= QUOTES =========================
def _yahoo_quote_once(symbols: List[str], host: str):
    url = f"https://{host}/v7/finance/quote"
    r = http_get(url, params={"symbols": ",".join(symbols)})
    r.raise_for_status()
    data = r.json().get("quoteResponse", {}).get("result", [])
    return {d.get("symbol"): d for d in data if d.get("symbol")}

def _yahoo_chart_one(symbol: str, host: str):
    """Fallback por símbolo usando /v8/finance/chart: calcula % a partir de previousClose."""
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
    """Primário: v7/quote (chunks + retries + fallback host). Fallback: v8/chart por símbolo."""
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
            except Exception as ex:
                print(f"[WARN] Yahoo v7 {h} falhou {syms}: {ex}")
                time.sleep(0.6)
        if not ok:
            print(f"[WARN] v7 falhou de vez {syms}")
        time.sleep(0.25)

    # completa faltantes com v8/chart
    for sym in symbols:
        d = out.get(sym, {})
        has_price = d.get("regularMarketPrice") or d.get("postMarketPrice") or d.get("preMarketPrice")
        has_pct   = d.get("regularMarketChangePercent") or d.get("postMarketChangePercent") or d.get("preMarketChangePercent")
        if has_price and (has_pct is not None):
            continue
        for h in hosts:
            try:
                c = _yahoo_chart_one(sym, h)
                if c:
                    d = out.get(sym, {}) or {"symbol": sym}
                    if not has_price and c.get("price") is not None:
                        d["regularMarketPrice"] = c["price"]
                    if has_pct is None and c.get("pct") is not None:
                        d["regularMarketChangePercent"] = c["pct"]
                    out[sym] = d
                    break
            except Exception as ex:
                print(f"[WARN] Yahoo v8 {h} falhou {sym}: {ex}")
                time.sleep(0.4)
    return out

def fmt_price(sym: str, v) -> str:
    try:
        if v is None or (isinstance(v,float) and (math.isnan(v) or math.isinf(v))):
            return "—"
        if sym.endswith(".SA"):
            return "R$ {:,.2f}".format(float(v)).replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return "US$ {:,.2f}".format(float(v)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"

def fmt_pct(v) -> str:
    try:
        if v is None or (isinstance(v,float) and (math.isnan(v) or math.isinf(v))):
            return "—"
        return "{:+.2f}%".format(float(v)).replace(".", ",")
    except Exception:
        return "—"

def bloco_cotacoes() -> str:
    br = yahoo_quote(list(B3_TICKERS.keys()))
    gl = yahoo_quote(list(GLOBAL_TICKERS.keys()))

    def linhas(payload: Dict[str, dict], labels: Dict[str,str]) -> List[str]:
        out = []
        for sym, label in labels.items():
            d = payload.get(sym, {})
            px = d.get("regularMarketPrice") or d.get("postMarketPrice") or d.get("preMarketPrice")
            pct= d.get("regularMarketChangePercent") or d.get("postMarketChangePercent") or d.get("preMarketChangePercent")
            out.append(f"• {label} {fmt_pct(pct)} a {fmt_price(sym, px)}")
        return out

    b3 = "<h3 style='margin:16px 0 6px;'>📈 Cotações — Brasil (B3)</h3>" \
         f"<p style='margin:4px 0 10px'>{'<br>'.join(linhas(br, B3_TICKERS))}</p>"
    glob = "<h3 style='margin:8px 0 6px;'>🌐 Cotações — Lá fora / Commodities / Cripto</h3>" \
         f"<p style='margin:4px 0 10px'>{'<br>'.join(linhas(gl, GLOBAL_TICKERS))}</p>"
    return b3 + glob

# ================= News: RSS + HTML scraping =================
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
    except Exception as ex:
        print(f"[WARN] RSS fail {url}: {ex}")
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
        # dedup
        seen, uniq = set(), []
        for h in hrefs:
            if h not in seen:
                seen.add(h); uniq.append(h)
        return uniq
    except Exception as ex:
        print(f"[WARN] HTML list fail {url}: {ex}")
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
        if _blocked_text(txt):
            return ""
        return txt
    except Exception:
        try:
            soup = BeautifulSoup(r.text, "lxml")
            art = soup.find("article") or soup
            txt = art.get_text(" ", strip=True)
            txt = re.sub(r"\s+", " ", txt).strip()
            if _blocked_text(txt):
                return ""
            return txt
        except Exception as ex:
            print(f"[WARN] extract fail {url}: {ex}")
            return ""

def summarize(text: str, max_frases=2, max_chars=260) -> str:
    if not text: return ""
    sents = re.split(r"(?<=[\.\!\?])\s+", text)
    sents = [s.strip() for s in sents if len(s.strip()) >= 30][:50]
    if not sents: return ""
    def toks(s): return [t.lower() for t in re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9%]+", s)]
    STOP_PT = {"a","o","as","os","de","da","do","das","dos","e","é","em","no","na","nos","nas","para","por","com","um","uma",
               "ao","aos","à","às","como","que","se","sua","seu","suas","seus","mais","menos","entre","sobre","já","também",
               "até","há","ser","ter","foi","são","pela","pelo","pelas","pelos","ou","onde","quando","porque"}
    fre = {}
    for s in sents:
        for t in toks(s):
            if t in STOP_PT: continue
            fre[t] = fre.get(t,0)+1
    scored = []
    for i,s in enumerate(sents):
        sc = sum(fre.get(t,0) for t in toks(s))
        if i == 0: sc *= 1.15
        scored.append((sc,i,s))
    top = [t[2] for t in sorted(sorted(scored, key=lambda x:x[0], reverse=True)[:max_frases], key=lambda x:x[1])]
    out = " ".join(top)
    if len(out) > max_chars:
        out = out[:max_chars].rsplit(" ",1)[0]+"…"
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
                except Exception as ex:
                    print(f"[WARN] fetch link fail {lk}: {ex}")
            time.sleep(SLEEP)
    # dedup + corta por seção
    for k in out.keys():
        seen, uniq = set(), []
        for it in sorted(out[k], key=lambda x:x["quando"], reverse=True):
            key = (it.get("titulo",""), it.get("fonte",""))
            if key in seen: continue
            seen.add(key); uniq.append(it)
        out[k] = uniq[:MAX_BULLETS]
    return out

# ================= Montagem de texto por seção =================
def mk_paragrafo_internacional(items: List[Dict], max_chars=700) -> str:
    parts = []
    for it in items[:6]:
        txt = extract_article_text(it.get("link",""))
        if not txt:
            raw = BeautifulSoup(it.get("resumo_feed",""), "lxml").get_text(" ", strip=True)
            txt = re.sub(r"\s+"," ", raw)
        if not txt or "Access to this page has been denied" in txt:
            continue
        parts.append(summarize(txt, max_frases=1, max_chars=220))
    par = " ".join([p for p in parts if p])
    if len(par) > max_chars:
        par = par[:max_chars].rsplit(" ",1)[0]+"…"
    return html.escape(par)

def mk_bullets(items: List[Dict]) -> List[str]:
    rows = []
    for it in items:
        quando = it.get("quando", datetime.now(TZ_BR)).strftime("%d/%m %H:%M")
        fonte  = html.escape(it.get("fonte") or it.get("srcname") or "")
        title  = html.escape(re.sub(r"\s+"," ", (it.get("titulo") or "").strip()))
        txt = extract_article_text(it.get("link",""))
        if not txt:
            raw = BeautifulSoup(it.get("resumo_feed",""), "lxml").get_text(" ", strip=True)
            txt = re.sub(r"\s+"," ", raw)
        if not txt or "Access to this page has been denied" in txt:
            continue
        resumo = html.escape(summarize(txt, max_frases=2, max_chars=260))
        rows.append(f"• <b>{title}</b> — {resumo} <span style='color:#777'>({fonte}, {quando})</span>")
    return rows

def bloco_internacional(items: List[Dict]) -> str:
    if not items: return ""
    par = mk_paragrafo_internacional(items)
    cot = bloco_cotacoes()
    parts = cot.split("<h3")
    glob = "<h3" + parts[2] if len(parts) >= 3 else ""
    return (
        "<h3 style='margin:16px 0 6px;'>🌎 INTERNACIONAL</h3>"
        f"<p style='margin:6px 0 10px; color:#222; line-height:1.45'>{par}</p>"
        + glob
    )

def bloco_brasil(items: List[Dict]) -> str:
    bullets = mk_bullets(items)
    if not bullets: return ""
    return "<h3 style='margin:16px 0 6px;'>🇧🇷 BRASIL</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

def bloco_empresas(items: List[Dict]) -> str:
    bullets = mk_bullets(items)
    if not bullets: return ""
    return "<h3 style='margin:16px 0 6px;'>🏢 EMPRESAS</h3>" \
           f"<p style='margin:4px 0 10px'>{'<br>'.join(bullets)}</p>"

# ================= EMAIL (Outlook) =================
def _try_set_sender_account(mail, outlook_session, from_hint: Optional[str]) -> None:
    """Seleciona a conta remetente quando houver múltiplas contas.
       from_hint pode ser parte do display name ou o e-mail completo."""
    if not from_hint:
        return
    try:
        # percorre contas e tenta match
        for i in range(1, outlook_session.Accounts.Count + 1):
            acc = outlook_session.Accounts.Item(i)
            name = str(getattr(acc, "SmtpAddress", "") or getattr(acc, "DisplayName", "") or "").lower()
            if from_hint.lower() in name:
                # set SentOnBehalfOfName (simples e suficiente na maioria dos casos)
                mail.SentOnBehalfOfName = acc.SmtpAddress or acc.DisplayName
                return
    except Exception as ex:
        print("[WARN] Falha ao selecionar conta remetente:", ex)

def enviar_email_outlook(dest, assunto, html_corpo, MODE="SEND", from_account_hint: Optional[str]=None):
    """
    MODE:
      - "DISPLAY_ONLY": abre o e-mail na tela para revisão
      - "SAVE_DRAFT": salva em Rascunhos
      - "SEND": envia direto (padrão)
    """
    import win32com.client as win32
    import pythoncom
    try:
        pythoncom.CoInitialize()
        outlook = win32.DispatchEx("Outlook.Application")
        session = outlook.GetNamespace("MAPI")
        try:
            session.Logon("", "", False, False)
        except Exception:
            pass

        mail = outlook.CreateItem(0)  # olMailItem
        mail.Subject = assunto
        mail.HTMLBody = html_corpo

        # define remetente (se pedido)
        _try_set_sender_account(mail, session, from_account_hint)

        # adiciona destinatários e resolve
        recips = mail.Recipients
        for addr in dest:
            if not addr:
                continue
            r = recips.Add(addr)
            r.Type = 1  # Para
        if not recips.ResolveAll():
            raise RuntimeError("Falha ao resolver um ou mais destinatários. Verifique os e-mails.")

        if MODE.upper() == "DISPLAY_ONLY":
            mail.Display()
            print("E-mail aberto para revisão (DISPLAY_ONLY).")
        elif MODE.upper() == "SAVE_DRAFT":
            mail.Save()
            print("E-mail salvo em Rascunhos (SAVE_DRAFT).")
        else:
            mail.Send()
            print("E-mail enviado (SEND).")

    except Exception as e:
        print("Falha no envio via Outlook:", e)
        print("Dicas:")
        print("- Abra o Outlook e deixe logado no perfil padrão.")
        print("- Rode o Python com a MESMA elevação do Outlook (ambos admin ou ambos normal).")
        print("- Outlook > Arquivo > Opções > Central de Confiabilidade > Acesso Programático.")
        raise
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

# ================= MAIN =================
def run(mode: str = "SEND", from_hint: Optional[str] = None):
    news = collect_news()  # {"Brasil":[...], "Internacional":[...], "Empresas":[...]}

    topo = (
        "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;'>"
        "<h2 style='margin:0;'>Boletim — Internacional • Brasil • Empresas</h2>"
        f"<p style='margin:6px 0 8px;color:#444;'>Gerado em {datetime.now(TZ_BR).strftime('%d/%m/%Y %H:%M')} (UTC-03). "
        f"Janela: últimos {JANELA_DIAS} dia(s); até {MAX_BULLETS} tópicos por seção.</p>"
        "<hr style='border:none;border-top:1px solid #e1e1e1;margin:8px 0;'>"
    )

    cot = bloco_cotacoes()
    intl = bloco_internacional(news.get("Internacional", []))
    br   = bloco_brasil(news.get("Brasil", []))
    emp  = bloco_empresas(news.get("Empresas", []))

    rodape = (
        "<hr style='border:none;border-top:1px solid #e1e1e1;margin:12px 0;'>"
        "<p style='color:#777;font-size:12px;'>"
        "Fontes: InfoMoney, Valor Investe, XP Conteúdos, Mais Retorno, Yubb, Bora Investir/B3, "
        "Bloomberg Línea Brasil, Investing.com (BR), Brazil Journal, Bloomberg, Reuters Markets, "
        "Financial Times e Seeking Alpha. "
        "Este e-mail apresenta resumos sem hiperlinks; consulte os portais oficiais para os conteúdos integrais."
        "</p></div>"
    )

    html_email = topo + cot + intl + br + emp + rodape
    assunto = f"{ASSUNTO_PREFIXO} — {datetime.now(TZ_BR).strftime('%Y-%m-%d')}"
    enviar_email_outlook(DESTINATARIOS, assunto, html_email, MODE=mode, from_account_hint=from_hint)

if __name__ == "__main__":
    try:
        ap = argparse.ArgumentParser(description="Boletim Netuno — Internacional • Brasil • Empresas")
        g = ap.add_mutually_exclusive_group()
        g.add_argument("--preview", action="store_true", help="abre o e-mail para revisão (não envia)")
        g.add_argument("--save-draft", action="store_true", help="salva o e-mail em Rascunhos (não envia)")
        g.add_argument("--send", action="store_true", help="envia direto (padrão)")
        ap.add_argument("--from", dest="from_hint", default=None,
                        help="e-mail ou parte do display name da conta remetente a usar")
        args = ap.parse_args()

        mode = "SEND"
        if args.preview:
            mode = "DISPLAY_ONLY"
        elif args.save_draft:
            mode = "SAVE_DRAFT"
        elif args.send:
            mode = "SEND"

        run(mode=mode, from_hint=args.from_hint)
        print("Boletim processado com sucesso.")
    except Exception as e:
        print("Falha no boletim:", e)
        traceback.print_exc()
