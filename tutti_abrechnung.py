#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tutti Tschutti / R1 - Event-Abrechnung (Variante 3: voll automatisierte Pipeline)

Holt die Verkaufsdaten automatisch aus der SumUp-API, liest die TWINT/RaiseNow-CSV
ein und erstellt pro Event eine fertige Abrechnung als Excel (.xlsx) und PDF.

Berechnet pro Event:
  1) Gesamtumsatz (SumUp)        -> 13% Umsatzbeteiligung
  2) Hausgetraenke-Verbrauch     -> Menge x Einkaufspreis (Gratis/Barteam als CHF-0-Artikel
                                    sind automatisch dabei, weil sie als Artikel getippt werden)
  3) Zahlungs-Split (Abgleich)   -> Karte / TWINT / Bargeld

WICHTIG - KEIN DOPPELZAEHLEN:
  In eurem Workflow wird JEDE Bestellung in SumUp getippt, auch TWINT-Zahlungen
  (gebucht als 'Bargeld'). Der SumUp-Gesamtumsatz ist daher bereits vollstaendig.
  Die RaiseNow/TWINT-CSV wird NICHT zum Umsatz addiert, sondern nur fuer den
  Zahlungs-Abgleich verwendet:  Bargeld (echt) = SumUp-Cash - TWINT.

VORAUSSETZUNG:
  In SumUp muss jedes Getraenk ein Artikel im Katalog sein (nicht als freier Betrag
  tippen). Nur dann liefert die API die Einzelpositionen, die fuer die EK-Verrechnung
  noetig sind.

Aufruf (Beispiele):
  # Erst testen, ohne API/Keys - mit eingebauten Demodaten:
  python tutti_abrechnung.py --demo

  # Echtlauf fuer ein Event:
  python tutti_abrechnung.py \
      --event "Eröffnungsabend" \
      --von "2026-06-11T17:00:00+02:00" \
      --bis "2026-06-12T02:00:00+02:00" \
      --twint raisenow_export.csv \
      --config config.json \
      --ek ek_preise.csv

  # Welche payment_type / card_type Werte hat mein Konto? (zum Konfigurieren):
  python tutti_abrechnung.py --inspect --von ... --bis ... --config config.json
"""

import argparse
import csv
import datetime as dt
import json
import os
import sys
import unicodedata

# ----------------------------------------------------------------------------
# SumUp API-Endpunkte (zentral, falls SumUp die Versionen aendert -> hier anpassen).
# Pruefe im Zweifel: https://developer.sumup.com/api
# ----------------------------------------------------------------------------
API_BASE = "https://api.sumup.com"
EP_ME = API_BASE + "/v0.1/me"
EP_TX_HISTORY = API_BASE + "/v0.1/me/transactions/history"
EP_RECEIPT = API_BASE + "/v1.1/receipts/{tx_id}?mid={mid}"

# Wie die SumUp-Zahlungsarten in "Karte" vs "Bargeld(+TWINT)" einsortiert werden.
# Defaults sind ein erster Versuch. Mit --inspect siehst du die echten Werte deines
# Kontos und kannst sie hier bei Bedarf anpassen.
CARD_PAYMENT_TYPES = {"ECOM", "POS", "BOLETO", "RECURRING"}  # alles mit Karte/Reader
CASH_PAYMENT_TYPES = {"CASH"}                                # bar getippt (= bar + TWINT)


# ----------------------------------------------------------------------------
# Hilfsfunktionen
# ----------------------------------------------------------------------------
def norm(name):
    """Produktnamen vereinheitlichen (klein, ohne Akzente/Doppel-Leerzeichen).

    Mitarbeiter-Gratisartikel (0 CHF, Kategorie 'Mitarbeiter') matchen auf
    denselben EK wie der normale Artikel: Suffixe/Praefixe wie 'Mitarbeiter',
    '(MA)', 'Personal', 'Gratis' werden entfernt."""
    if name is None:
        return ""
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    s = " ".join(s.lower().split())
    for tag in ("mitarbeiter", "personal", "gratis", "(ma)", "ma:"):
        s = s.replace(tag, " ")
    s = " ".join(s.replace("(", " ").replace(")", " ").split())
    # Reihenfolge-unabhaengig: "Alkoholfrei Lager 3dl" == "Lager Alkoholfrei 3dl"
    return " ".join(sorted(s.split()))


def parse_iso(s):
    """ISO-8601 String -> aware datetime."""
    s = s.strip().replace("Z", "+00:00")
    return dt.datetime.fromisoformat(s)


def load_config(path):
    if not path or not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_ek_prices(path):
    """CSV 'produkt,ek' -> dict {normierter_name: (anzeigename, ek)}."""
    ek = {}
    if not path or not os.path.exists(path):
        return ek
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or len(row) < 2:
                continue
            name, raw = row[0].strip(), row[1].strip()
            if norm(name) in ("produkt", "name", "getraenk", "getränk"):
                continue  # Header
            try:
                ek[norm(name)] = (name, float(str(raw).replace(",", ".")))
            except ValueError:
                continue
    return ek


def load_twint_csv(path, von, bis, cfg):
    """RaiseNow/TWINT-Export einlesen, Betraege im Zeitfenster summieren.

    Spaltennamen sind je nach Export unterschiedlich. Standardmaessig wird nach
    einer Datums- und einer Betrags-Spalte gesucht; per config['twint'] ueberschreibbar.
    """
    if not path or not os.path.exists(path):
        return {"total": 0.0, "count": 0, "note": "keine TWINT-CSV angegeben"}

    tw = (cfg.get("twint") or {})
    amount_keys = [tw.get("amount_col")] if tw.get("amount_col") else \
        ["amount", "betrag", "gross", "brutto", "value", "total"]
    date_keys = [tw.get("date_col")] if tw.get("date_col") else \
        ["date", "datum", "created", "created_at", "timestamp", "zeit", "time"]

    total, count = 0.0, 0
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers_norm = {norm(h): h for h in (reader.fieldnames or [])}
        amount_col = next((headers_norm[norm(k)] for k in amount_keys
                           if k and norm(k) in headers_norm), None)
        date_col = next((headers_norm[norm(k)] for k in date_keys
                         if k and norm(k) in headers_norm), None)
        if amount_col is None:
            return {"total": 0.0, "count": 0,
                    "note": "Betrags-Spalte nicht gefunden -> in config['twint']['amount_col'] setzen. "
                            "Gefundene Spalten: " + ", ".join(reader.fieldnames or [])}
        for r in reader:
            # Datumsfilter (best effort - wenn kein Datum erkennbar, alles zaehlen)
            if date_col and r.get(date_col):
                try:
                    d = parse_iso(str(r[date_col])[:25])
                    if d.tzinfo is None:
                        d = d.replace(tzinfo=von.tzinfo or dt.timezone.utc)
                    if not (von <= d <= bis):
                        continue
                except Exception:
                    pass
            raw = str(r.get(amount_col, "")).replace("'", "").replace(",", ".").strip()
            try:
                total += float(raw)
                count += 1
            except ValueError:
                continue
    return {"total": round(total, 2), "count": count,
            "note": "Datumsfilter aktiv" if date_col else "ohne Datumsfilter (ganze Datei)"}


# ----------------------------------------------------------------------------
# SumUp API
# ----------------------------------------------------------------------------
def _session(api_key):
    import requests
    s = requests.Session()
    s.headers.update({"Authorization": "Bearer " + api_key,
                      "Accept": "application/json"})
    return s


def get_merchant_code(s, cfg):
    if cfg.get("merchant_code"):
        return cfg["merchant_code"]
    r = s.get(EP_ME, timeout=30)
    r.raise_for_status()
    data = r.json()
    return (data.get("merchant_profile") or {}).get("merchant_code") \
        or data.get("merchant_code")


def fetch_transactions(s, von, bis):
    """Alle Transaktionen im Zeitfenster holen (mit Pagination)."""
    items, url = [], EP_TX_HISTORY
    params = {"limit": 100, "order": "ascending",
              "oldest_time": von.isoformat(), "newest_time": bis.isoformat()}
    seen = 0
    while url:
        r = s.get(url, params=params if url == EP_TX_HISTORY else None, timeout=30)
        r.raise_for_status()
        data = r.json()
        batch = data.get("items", data if isinstance(data, list) else [])
        items.extend(batch)
        seen += len(batch)
        # naechste Seite via links[rel=next]. SumUp liefert href meist nur als
        # Query-String ("limit=100&oldest_ref=..."), nicht als volle URL.
        nxt = None
        for link in data.get("links", []) if isinstance(data, dict) else []:
            if link.get("rel") == "next":
                href = link.get("href", "")
                if href.startswith("http"):
                    nxt = href
                elif href.startswith("/"):
                    nxt = API_BASE + href
                elif href:
                    nxt = EP_TX_HISTORY + "?" + href.lstrip("?")
        url, params = nxt, None
        if seen > 50000:  # Sicherheitsbremse
            break
    return items


def fetch_receipt_items(s, mid, tx):
    """Einzelpositionen (Artikel) einer Transaktion holen."""
    tx_id = tx.get("transaction_id") or tx.get("id")
    if not tx_id:
        return []
    r = s.get(EP_RECEIPT.format(tx_id=tx_id, mid=mid), timeout=30)
    if r.status_code != 200:
        return []
    td = (r.json().get("transaction_data") or {})
    out = []
    for p in td.get("products", []) or []:
        out.append({
            "name": p.get("name") or p.get("description") or "Unbenannt",
            "quantity": float(p.get("quantity") or 0),
            "price": float(p.get("price") or 0),
            "total_price": float(p.get("total_price")
                                 or (float(p.get("price") or 0) * float(p.get("quantity") or 0))),
        })
    return out


# ----------------------------------------------------------------------------
# Demodaten (fuer --demo, damit du den Output ohne Keys siehst)
# ----------------------------------------------------------------------------
def demo_data():
    tx = [
        {"transaction_id": "t1", "amount": 35.0, "status": "SUCCESSFUL", "type": "PAYMENT", "payment_type": "POS"},
        {"transaction_id": "t2", "amount": 12.0, "status": "SUCCESSFUL", "type": "PAYMENT", "payment_type": "CASH"},
        {"transaction_id": "t3", "amount": 24.0, "status": "SUCCESSFUL", "type": "PAYMENT", "payment_type": "CASH"},
        {"transaction_id": "t4", "amount": 0.0,  "status": "SUCCESSFUL", "type": "PAYMENT", "payment_type": "CASH"},
        {"transaction_id": "t5", "amount": 18.0, "status": "SUCCESSFUL", "type": "PAYMENT", "payment_type": "POS"},
    ]
    receipts = {
        "t1": [{"name": "Bier 0.5", "quantity": 5, "price": 5.0, "total_price": 25.0},
               {"name": "Vodka Shot 4cl", "quantity": 2, "price": 5.0, "total_price": 10.0}],
        "t2": [{"name": "Bier 0.5", "quantity": 2, "price": 5.0, "total_price": 10.0},
               {"name": "Mineral", "quantity": 1, "price": 2.0, "total_price": 2.0}],
        "t3": [{"name": "Aperol Spritz", "quantity": 3, "price": 8.0, "total_price": 24.0}],
        "t4": [{"name": "Bier 0.5 (Personal)", "quantity": 2, "price": 0.0, "total_price": 0.0}],
        "t5": [{"name": "Bier 0.5", "quantity": 2, "price": 5.0, "total_price": 10.0},
               {"name": "Vodka Shot 4cl", "quantity": 1, "price": 5.0, "total_price": 5.0},
               {"name": "Chips", "quantity": 1, "price": 3.0, "total_price": 3.0}],
    }
    # Demo-EK-Liste
    ek = {norm("Bier 0.5"): ("Bier 0.5", 0.90),
          norm("Bier 0.5 (Personal)"): ("Bier 0.5 (Personal)", 0.90),
          norm("Vodka Shot 4cl"): ("Vodka Shot 4cl", 0.80),
          norm("Aperol Spritz"): ("Aperol Spritz", 2.20),
          norm("Mineral"): ("Mineral", 0.40)}
          # "Chips" fehlt absichtlich -> wird als unbekannt markiert
    twint = {"total": 24.0, "count": 1, "note": "Demo (entspricht t3, bar getippt)"}
    return tx, receipts, ek, twint


# ----------------------------------------------------------------------------
# Kernberechnung
# ----------------------------------------------------------------------------
def compute(transactions, receipts_fn, ek, twint, rate=0.13):
    umsatz = 0.0
    card_total = 0.0
    cash_total = 0.0
    products = {}   # norm -> {"name":..., "qty":..., }
    no_items = 0

    for tx in transactions:
        if tx.get("status") not in (None, "SUCCESSFUL"):
            continue
        ttype = tx.get("type", "PAYMENT")
        amt = float(tx.get("amount") or 0)
        sign = -1.0 if ttype == "REFUND" else 1.0
        umsatz += sign * amt

        ptype = (tx.get("payment_type") or "").upper()
        has_card = bool(tx.get("card_type")) or ptype in CARD_PAYMENT_TYPES
        if ptype in CASH_PAYMENT_TYPES and not tx.get("card_type"):
            cash_total += sign * amt
        elif has_card:
            card_total += sign * amt
        else:
            cash_total += sign * amt  # unbekannt -> zu Bargeld(+TWINT)

        items = receipts_fn(tx)
        for it in items:
            is_free = (float(it.get("price") or 0) == 0)
            key = (norm(it["name"]), is_free)
            if key not in products:
                products[key] = {"name": it["name"], "qty": 0.0}
            products[key]["qty"] += sign * it["quantity"]
        if not items:
            no_items += 1

    # EK je Produkt zuordnen (Gratis-/Mitarbeiter-Items als eigene Zeile)
    rows, unknown = [], []
    getraenkeschuld = 0.0
    for (nkey, is_free), p in sorted(products.items(),
                                     key=lambda kv: (kv[0][1], kv[1]["name"].lower())):
        disp, ekval = ek.get(nkey, (p["name"], None))
        label = f"{disp} (gratis/MA)" if is_free else disp
        qty = round(p["qty"], 3)
        if ekval is None:
            unknown.append(label)
            line = 0.0
            rows.append({"name": label, "qty": qty, "ek": None, "line": line, "unknown": True})
        else:
            line = round(qty * ekval, 2)
            getraenkeschuld += line
            rows.append({"name": label, "qty": qty, "ek": ekval, "line": line, "unknown": False})

    umsatz = round(umsatz, 2)
    card_total = round(card_total, 2)
    cash_total = round(cash_total, 2)
    twint_total = round(float(twint.get("total", 0.0)), 2)
    bargeld_real = round(cash_total - twint_total, 2)
    beteiligung = round(umsatz * rate, 2)
    getraenkeschuld = round(getraenkeschuld, 2)
    total_owed = round(beteiligung + getraenkeschuld, 2)

    return {
        "umsatz": umsatz, "rate": rate, "beteiligung": beteiligung,
        "rows": rows, "unknown": unknown, "getraenkeschuld": getraenkeschuld,
        "total_owed": total_owed,
        "card_total": card_total, "cash_total": cash_total,
        "twint_total": twint_total, "bargeld_real": bargeld_real,
        "twint_note": twint.get("note", ""), "no_items": no_items,
        "n_tx": len(transactions),
    }


# ----------------------------------------------------------------------------
# Ausgabe: Excel (mit Formeln) + PDF
# ----------------------------------------------------------------------------
def write_excel(res, meta, path, zahlen=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    FONT = "Arial"
    bold = Font(name=FONT, bold=True)
    base = Font(name=FONT)
    title = Font(name=FONT, bold=True, size=14)
    head_fill = PatternFill("solid", start_color="1F2937")
    head_font = Font(name=FONT, bold=True, color="FFFFFF")
    warn_fill = PatternFill("solid", start_color="FFF3CD")
    money = '#,##0.00;(#,##0.00);"-"'
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Abrechnung"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16

    def cell(ref, val, f=base, num=None, al=None, fill=None):
        c = ws[ref]; c.value = val; c.font = f
        if num: c.number_format = num
        if al: c.alignment = Alignment(horizontal=al)
        if fill: c.fill = fill
        return c

    cell("A1", "Event-Abrechnung", title)
    cell("A2", meta["event"], bold)
    cell("A3", f"Zeitraum: {meta['von']}  bis  {meta['bis']}", base)
    cell("A4", f"Erstellt: {meta['created']}   Transaktionen: {res['n_tx']}", base)

    r = 6
    cell(f"A{r}", "1) Einnahmen — eingegangen auf dem R1-Konto", bold)
    r += 1
    cell(f"A{r}", "Karte via SumUp (brutto)"); cell(f"D{r}", res["card_total"], base, money, "right")
    card_row = r
    r += 1
    z = zahlen or {}
    cell(f"A{r}", "TWINT (in SumUp als Bargeld getippt, brutto)")
    cell(f"D{r}", z.get("twint_brutto", res["cash_total"]), base, money, "right")
    twintb_row = r
    r += 1
    cell(f"A{r}", "− SumUp-Gebühren (2.5%)")
    cell(f"D{r}", -z.get("sumup_geb", 0), base, money, "right"); sgeb_row = r
    r += 1
    cell(f"A{r}", "− TWINT-Gebühren (1.3%)")
    cell(f"D{r}", -z.get("twint_geb", 0), base, money, "right"); tgeb_row = r
    r += 1
    cell(f"A{r}", "Einnahmen netto", bold)
    cell(f"D{r}", f"=D{card_row}+D{twintb_row}+D{sgeb_row}+D{tgeb_row}", bold, money, "right")
    netto_row = r

    r += 2
    cell(f"A{r}", "2) Abzüge R1 (bleiben beim Trägerverein)", bold)
    r += 1
    cell(f"A{r}", f"Gesamtumsatz (Basis, alle Kategorien)"); cell(f"D{r}", res["umsatz"], base, money, "right")
    umsatz_row = r
    r += 1
    cell(f"A{r}", "Umsatzbeteiligungssatz"); cell(f"D{r}", res["rate"], base, "0.0%", "right")
    rate_row = r
    r += 1
    cell(f"A{r}", "13% Umsatzbeteiligung", bold)
    cell(f"D{r}", f"=D{umsatz_row}*D{rate_row}", bold, money, "right")
    beteiligung_row = r
    r += 1
    cell(f"A{r}", "Einkauf Getränke (Detail unten, inkl. Mitarbeiter/Gratis)", bold)
    ek_row_ref = r  # Formel wird nach der Detailtabelle gesetzt
    r += 1
    cell(f"A{r}", "Miete (nach Wochentag)")
    cell(f"D{r}", z.get("miete", 0), base, money, "right"); miete_row = r
    r += 1
    cell(f"A{r}", "Total Abzüge", bold)
    cell(f"D{r}", f"=D{beteiligung_row}+D{ek_row_ref}+D{miete_row}", bold, money, "right")
    abzuege_row = r

    r += 2
    cell(f"A{r}", "3) ÜBERWEISUNG an die Nutzung (Reingewinn)", title)
    cell(f"D{r}", f"=D{netto_row}-D{abzuege_row}", title, money, "right")
    cell(f"A{r}", "3) ÜBERWEISUNG an die Nutzung (Reingewinn)", title).fill = PatternFill(
        "solid", start_color="DCFCE7")
    ws[f"D{r}"].fill = PatternFill("solid", start_color="DCFCE7")

    r += 2
    cell(f"A{r}", "Detail: Getränke-Einkauf (Menge × Einkaufspreis)", bold)
    r += 1
    for col, txt in zip("ABCD", ["Produkt", "Menge", "EK/Stk", "Total"]):
        c = cell(f"{col}{r}", txt, head_font, fill=head_fill)
        c.alignment = Alignment(horizontal="left" if col == "A" else "right")
        c.border = border
    r += 1
    first = r
    for row in res["rows"]:
        cell(f"A{r}", row["name"], base).border = border
        cell(f"B{r}", row["qty"], base, "#,##0.###", "right").border = border
        if row["unknown"]:
            cell(f"C{r}", "kein EK", base, None, "right").border = border
            cell(f"D{r}", 0, base, money, "right").border = border
        else:
            cell(f"C{r}", row["ek"], base, money, "right").border = border
            cell(f"D{r}", f"=B{r}*C{r}", base, money, "right").border = border
        r += 1
    last = r - 1
    cell(f"A{r}", "Einkauf Getränke (Summe)", bold)
    cell(f"D{r}", f"=SUM(D{first}:D{last})" if last >= first else 0, bold, money, "right")
    # Abzugszeile oben mit der Detailsumme verknuepfen
    cell(f"D{ek_row_ref}", f"=D{r}", bold, money, "right")



    r += 2
    if res["unknown"]:
        cell(f"A{r}", "Ohne EK-Abzug (Fremdkategorie oder nicht in EK-Liste):", bold)
        r += 1
        cell(f"A{r}", ", ".join(res["unknown"]), base)
        r += 1
    if res["no_items"]:
        cell(f"A{r}", f"⚠ {res['no_items']} Transaktion(en) ohne Artikel "
                      f"(als freier Betrag getippt?) — nicht in EK enthalten.", base).fill = warn_fill
        r += 1

    wb.save(path)
    return path


def recalc_excel(path):
    """Formeln in echte Werte rechnen (LibreOffice), falls vorhanden — sonst egal."""
    try:
        import subprocess
        skill_recalc = "/mnt/skills/public/xlsx/scripts/recalc.py"
        if os.path.exists(skill_recalc):
            subprocess.run([sys.executable, skill_recalc, path], timeout=120,
                           capture_output=True)
    except Exception:
        pass


R1_LOGO_B64 = "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"


def write_pdf(res, meta, path, zahlen=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT, TA_LEFT

    # Farbpalette (angelehnt an das gewuenschte Invoice-Design)
    BLUE = colors.HexColor("#3642BA")
    INK = BLUE                          # komplette Schrift blau
    GREY = colors.HexColor("#7C84C4")   # etwas helleres Blau für Nebentext
    HAIR = colors.HexColor("#3642BA")   # feine blaue Trennlinien

    def chf(x): return f"{x:,.2f}".replace(",", "'")

    styles = getSampleStyleSheet()
    def S(name, **kw):
        kw.setdefault("parent", styles["Normal"])
        return ParagraphStyle(name, **kw)
    title = S("t", fontName="Helvetica-Bold", fontSize=34, textColor=BLUE, leading=36)
    label = S("l", fontName="Helvetica-Bold", fontSize=8, textColor=BLUE, leading=11)
    lblR = S("lr", parent=label, alignment=TA_RIGHT)
    big = S("b", fontName="Helvetica-Bold", fontSize=11, textColor=INK, leading=14)
    bigR = S("br", parent=big, alignment=TA_RIGHT)
    body = S("bd", fontName="Helvetica", fontSize=9.5, textColor=INK, leading=13)
    bodyR = S("bdr", parent=body, alignment=TA_RIGHT)
    small = S("sm", fontName="Helvetica", fontSize=8, textColor=GREY, leading=11)
    smallR = S("smr", parent=small, alignment=TA_RIGHT)

    doc = SimpleDocTemplate(path, pagesize=A4, topMargin=18*mm, bottomMargin=18*mm,
                            leftMargin=18*mm, rightMargin=18*mm)
    W = doc.width
    el = []

    def rule(space_before=6, space_after=8, thick=1):
        t = Table([[""]], colWidths=[W])
        t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), thick, HAIR),
                               ("TOPPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        return [Spacer(1, space_before), t, Spacer(1, space_after)]

    # ---- Logo (R1) aus eingebettetem Base64 ------------------------------
    import base64, io
    from reportlab.platypus import Image as RLImage
    logo_h = 46                          # Höhe in pt (~16mm)
    logo = RLImage(io.BytesIO(base64.b64decode(R1_LOGO_B64)),
                   width=logo_h * 0.721, height=logo_h)
    logo.hAlign = "RIGHT"

    # ---- Kopf: Titel links, Logo + Adresse rechts ------------------------
    head = Table([[
        Paragraph("ABRECHNUNG", title), logo,
    ], [
        Paragraph(f"{meta['created']}", body),
        Paragraph(f"R1 Trägerverein<br/>Uferstrasse 70<br/>4057 Basel", smallR),
    ]], colWidths=[W*0.6, W*0.4])
    head.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, 0), 0), ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 1), (-1, 1), 0)]))
    el += [head]
    el += rule(6, 10)

    # ---- Event / Zeitraum -------------------------------------------------
    meta_block = Table([[
        Paragraph(meta["event"], big), Paragraph("ZEITRAUM", lblR),
    ], [
        Paragraph("", body),
        Paragraph(f"{meta['von']}<br/>bis {meta['bis']}", bodyR),
    ], [
        Paragraph("", body), Paragraph(f"Transaktionen: {res['n_tx']}", smallR),
    ]], colWidths=[W*0.5, W*0.5])
    meta_block.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    el += [meta_block]
    el += rule(8, 10)

    # ---- Zusammenfassung (Einnahmen / Abzuege / Ueberweisung) -------------
    z = zahlen or {}
    def r_(desc, amount, bold=False, sub=False):
        ds = big if (bold or sub) else body
        vs = bigR if (bold or sub) else bodyR
        return [Paragraph(desc, ds), Paragraph(amount, vs)]

    rows_sum = [
        [Paragraph("BESCHREIBUNG", label), Paragraph("BETRAG (CHF)", lblR)],
        r_("Einnahmen R1-Konto — Karte (brutto)", chf(res["card_total"])),
        r_("Einnahmen R1-Konto — TWINT (brutto)", chf(z.get("twint_brutto", res["cash_total"]))),
        r_("SumUp-Gebühren (2.5%)", "− " + chf(z.get("sumup_geb", 0))),
        r_("TWINT-Gebühren (1.3%)", "− " + chf(z.get("twint_geb", 0))),
        r_("Einnahmen netto", chf(z.get("einnahmen", res["umsatz"])), sub=True),
        r_(f"13% Umsatzbeteiligung (auf Umsatz {chf(res['umsatz'])})", "− " + chf(res["beteiligung"])),
        r_("Einkauf Getränke (inkl. Mitarbeiter/Gratis)", "− " + chf(res["getraenkeschuld"])),
        r_("Miete", "− " + chf(z.get("miete", 0))),
    ]
    st = Table(rows_sum, colWidths=[W*0.72, W*0.28])
    n = len(rows_sum)
    style = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
             ("LINEBELOW", (0, 0), (-1, 0), 1, HAIR),          # unter Kopfzeile
             ("LINEBELOW", (0, 5), (-1, 5), 1, HAIR),          # unter "Einnahmen netto"
             ("LINEABOVE", (0, 1), (-1, 1), 0, colors.white)]
    for i in range(1, n):
        if i == 5:            # unter "Einnahmen netto" liegt bereits die blaue Linie
            continue
        style.append(("LINEBELOW", (0, i), (-1, i), 0.25, colors.HexColor("#D8D8E8")))
    st.setStyle(TableStyle(style))
    el += [st]

    # TOTAL-Zeile (Ueberweisung) gross und blau, rechtsbuendig
    total = Table([[Paragraph("ÜBERWEISUNG AN DIE NUTZUNG",
                              S("tt", fontName="Helvetica-Bold", fontSize=11,
                                textColor=colors.white)),
                    Paragraph("CHF " + chf(z.get("auszahlung", 0)),
                              S("tv", fontName="Helvetica-Bold", fontSize=13,
                                textColor=colors.white, alignment=TA_RIGHT))]],
                  colWidths=[W*0.72, W*0.28])
    total.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), BLUE),
                               ("TOPPADDING", (0, 0), (-1, -1), 8),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                               ("LEFTPADDING", (0, 0), (0, 0), 8),
                               ("RIGHTPADDING", (-1, -1), (-1, -1), 8),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    el += [Spacer(1, 2), total]

    # ---- Detail Getraenke-Einkauf ----------------------------------------
    el += [Spacer(1, 16), Paragraph("DETAIL — GETRÄNKE-EINKAUF", label), Spacer(1, 4)]
    data = [[Paragraph("PRODUKT", label), Paragraph("MENGE", lblR),
             Paragraph("EK / STK", lblR), Paragraph("BETRAG", lblR)]]
    for row in res["rows"]:
        data.append([Paragraph(row["name"], body),
                     Paragraph(f"{row['qty']:g}", bodyR),
                     Paragraph("—" if row["unknown"] else chf(row["ek"]), bodyR),
                     Paragraph(chf(row["line"]), bodyR)])
    data.append([Paragraph("Einkauf Getränke — Summe", big), Paragraph("", bodyR),
                 Paragraph("", bodyR), Paragraph(chf(res["getraenkeschuld"]), bigR)])
    dt = Table(data, colWidths=[W*0.58, W*0.14, W*0.14, W*0.14])
    dstyle = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
              ("LINEBELOW", (0, 0), (-1, 0), 1, HAIR),
              ("LINEABOVE", (0, -1), (-1, -1), 1, HAIR)]
    for i in range(1, len(data) - 1):
        if i == len(data) - 2:   # direkt über der Summe liegt bereits die blaue Linie
            continue
        dstyle.append(("LINEBELOW", (0, i), (-1, i), 0.25, colors.HexColor("#E4E4F0")))
    dt.setStyle(TableStyle(dstyle))
    el += [dt]

    doc.build(el)
    return path


# ----------------------------------------------------------------------------
# Optional: KI-Schritt (Abrechnungs-E-Mail formulieren). Standardmaessig AUS.
# Benoetigt ANTHROPIC_API_KEY in der Umgebung.
# ----------------------------------------------------------------------------
def ki_email_text(res, meta):
    import requests
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    prompt = (
        "Schreibe eine kurze, freundliche Abrechnungs-E-Mail auf Deutsch (Du-Form, "
        "schweizerischer Ton) an eine Nutzung unseres Kulturraums. Zahlen:\n"
        f"- Event: {meta['event']}\n- Gesamtumsatz: CHF {res['umsatz']:.2f}\n"
        f"- Umsatzbeteiligung {res['rate']*100:.0f}%: CHF {res['beteiligung']:.2f}\n"
        f"- Getränkeschuld (Einkaufspreis): CHF {res['getraenkeschuld']:.2f}\n"
        f"- TOTAL offen: CHF {res['total_owed']:.2f}\n"
        "Kontext: Das Geld ist bereits beim Trägerverein eingegangen; wir überweisen "
        "der Nutzung ihren Reingewinn. Maximal 8 Sätze, keine Floskeln.")
    r = requests.post("https://api.anthropic.com/v1/messages",
                      headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                               "content-type": "application/json"},
                      json={"model": "claude-sonnet-4-6", "max_tokens": 500,
                            "messages": [{"role": "user", "content": prompt}]},
                      timeout=60)
    r.raise_for_status()
    return "".join(b.get("text", "") for b in r.json().get("content", []))


# ----------------------------------------------------------------------------
# EVENTBLATT R1: Vorlage mit Tokens bauen + Tokens befuellen
#
# Es werden NUR finanzielle Felder befuellt. Nicht-finanzielle Felder (Eventangaben,
# Tickets, Unterschriften, sonstige Kosten) bleiben leer fuer die manuelle Eingabe.
# ----------------------------------------------------------------------------
def chf(x):
    if x is None or x == "":
        return ""
    return f"{float(x):,.2f}".replace(",", "'")


def build_eventblatt_vorlage(path):
    """Erstellt eine EVENTBLATT-Vorlage mit {{TOKENS}} in den Finanzfeldern."""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    d = Document()
    base = d.styles["Normal"]
    base.font.name = "Arial"
    base.font.size = Pt(10)

    def h(text, size=14):
        p = d.add_paragraph()
        r = p.add_run(text); r.bold = True; r.font.size = Pt(size)
        return p

    def it(text):
        p = d.add_paragraph()
        r = p.add_run(text); r.italic = True; r.font.size = Pt(8)
        return p

    def kv(rows):
        t = d.add_table(rows=0, cols=2)
        t.style = "Table Grid"
        t.columns[0].width = Pt(260); t.columns[1].width = Pt(120)
        for label, token in rows:
            c = t.add_row().cells
            rp = c[0].paragraphs[0]; rr = rp.add_run(label); rr.bold = True
            if token:
                c[1].paragraphs[0].add_run(token)
        return t

    h("EVENTBLATT", 16)
    h("Abrechnung Getränke und Ticketeinnahmen", 11)
    it("Beilage zum Kooperationsvertrag – von verantwortlicher Person zu unterzeichnen")

    h("Eventangaben", 12)
    it("Erster Teil direkt nach dem Event und noch vor Ort ausfüllen und unterzeichnen!")
    kv([("Event:", ""), ("Datum:", ""),
        ("Verantwortliche Person Kooperationspartner:", ""),
        ("Tel. Verantwortliche Person Kooperationspartner:", ""),
        ("Abendverantwortung R1:", "")])
    d.add_paragraph()
    kv([("Bruttoeinnahmen SumUp (CHF):", ""),
        ("Anzahl Abendkasse Eintritte (CHF):", ""),
        ("Abendkasse Kosten pro Eintritt (CHF):", "")])
    it("Die unterzeichnende verantwortliche Person bestätigt die Richtigkeit der obenstehenden Angaben.")
    d.add_paragraph("Ort, Datum: ______________________   Unterschrift Kooperationspartner: ______________________")
    d.add_paragraph("Ort, Datum: ______________________   Unterschrift Abendverantwortung: ______________________")

    d.add_paragraph("―" * 40)
    h("ABRECHNUNG", 14)
    d.add_paragraph("Abrechnung durchgeführt von: ______________________________")

    h("Einnahmen SumUp", 12)
    kv([("Total Einnahmen Brutto (CHF):", "{{SUMUP_BRUTTO}}"),
        ("Gebühren (CHF):", "{{SUMUP_GEBUEHR}}"),
        ("Total Einnahmen Netto (CHF):", "{{SUMUP_NETTO}}")])

    h("Einnahmen Twint", 12)
    kv([("Total Einnahmen Brutto (CHF):", "{{TWINT_BRUTTO}}"),
        ("Gebühren (CHF):", "{{TWINT_GEBUEHR}}"),
        ("Total Einnahmen Netto (CHF):", "{{TWINT_NETTO}}")])

    h("Einnahmen Tickets", 12)
    it("Manuell – externe Ticketing-Plattform.")
    kv([("Ticketing-Plattform:", ""), ("Anzahl verkaufte Tickets:", ""),
        ("Kosten pro Ticket (CHF):", ""), ("Total Einnahmen Brutto (CHF):", ""),
        ("Gebühren (CHF):", ""), ("Total Einnahmen Netto (CHF):", ""),
        ("Nachweis beigefügt (Rapport/Screenshot):", "")])

    h("Abzüge", 12)
    kv([("13% Umsatzbeteiligung (CHF):", "{{UMSATZBETEILIGUNG}}"),
        ("Tagespauschale (CHF):", "{{TAGESPAUSCHALE}}"),
        ("Reinigungskosten (CHF):", "{{REINIGUNG}}"),
        ("Einkaufskosten Getränke (CHF):", "{{EK_GETRAENKE}}"),
        ("Abendverantwortung (CHF):", "{{ABENDVERANTWORTUNG}}"),
        ("Abzug sonstige Kosten (CHF):", ""),
        ("Total Abzüge (CHF):", "{{TOTAL_ABZUEGE}}")])

    h("Gesamtabrechnung", 12)
    kv([("Total Einnahmen (CHF):", "{{TOTAL_EINNAHMEN}}"),
        ("Total Abzüge (CHF):", "{{TOTAL_ABZUEGE}}"),
        ("Auszahlung an Kooperationspartner (CHF):", "{{AUSZAHLUNG}}")])
    it("Die unterzeichnende verantwortliche Person bestätigt die Richtigkeit der obenstehenden Angaben.")
    d.add_paragraph("Ort, Datum: ______________________   Unterschrift Kooperationspartner: ______________________")
    d.add_paragraph("Ort, Datum: ______________________   Unterschrift Abrechnungsverantwortung Trägerverein: ______________________")

    d.save(path)
    return path


def _replace_tokens_in_paragraph(p, mapping):
    """Token in einem Absatz ersetzen (auch wenn ueber mehrere Runs verteilt)."""
    full = "".join(run.text for run in p.runs)
    if "{{" not in full:
        return
    new = full
    for k, v in mapping.items():
        new = new.replace(k, v)
    if new != full and p.runs:
        p.runs[0].text = new
        for run in p.runs[1:]:
            run.text = ""


def fill_eventblatt(template_path, output_path, mapping):
    """Befuellt alle {{TOKENS}} in einem docx (Absaetze + Tabellenzellen)."""
    from docx import Document
    d = Document(template_path)

    def walk_tables(tables):
        for t in tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_tokens_in_paragraph(p, mapping)
                    walk_tables(cell.tables)

    for p in d.paragraphs:
        _replace_tokens_in_paragraph(p, mapping)
    walk_tables(d.tables)
    d.save(output_path)
    return output_path


def miete_fuer_datum(d):
    """Raummiete nach Wochentag des Eventstarts: So-Mi 50, Do 100, Fr+Sa 200."""
    wd = d.weekday()  # Mo=0 ... So=6
    if wd in (6, 0, 1, 2):   # So, Mo, Di, Mi
        return 50.0
    if wd == 3:              # Do
        return 100.0
    return 200.0             # Fr, Sa


def eventblatt_values(res, cfg, ticket_netto=0.0, event_date=None):
    """Berechnet die Finanzwerte fuers Eventblatt aus dem Abrechnungs-Ergebnis.

    Gebuehren-Logik (Schaetzung, konfigurierbar):
      - SumUp-Gebuehr auf alles, was mit Karte getippt wurde (Default 2.5%).
      - TWINT-Gebuehr auf alles, was als 'Bargeld' getippt wurde (Default 1.3%),
        denn im R1-Workflow steht 'bar getippt' fuer TWINT-QR-Zahlungen.
        Liegt eine RaiseNow-CSV vor, wird stattdessen deren Total verwendet.
      - Miete automatisch nach Wochentag (So-Mi 50 / Do 100 / Fr+Sa 200),
        via config['eventblatt']['tagespauschale'] uebersteuerbar.
    """
    eb = (cfg.get("eventblatt") or {})
    sumup_rate = float(eb.get("sumup_fee_rate", 0.025))
    twint_rate = float(eb.get("twint_fee_rate", 0.013))

    if cfg.get("_card_only"):
        # Nachhol-Modus fuer Alt-Events: vor dem Bargeld-Knopf lief nichts als
        # TWINT/bar ueber SumUp -> alles als Karte werten, keine TWINT-Gebuehr.
        sumup_brutto = res["card_total"] + res["cash_total"]
        twint_brutto = 0.0
    else:
        sumup_brutto = res["card_total"]                 # Karte via SumUp
        twint_brutto = res["cash_total"]                 # bar getippt = TWINT
    sumup_geb = round(sumup_brutto * sumup_rate, 2)
    twint_geb = round(twint_brutto * twint_rate, 2)
    sumup_netto = round(sumup_brutto - sumup_geb, 2)
    twint_netto = round(twint_brutto - twint_geb, 2)

    umsatzbeteiligung = res["beteiligung"]
    ek = res["getraenkeschuld"]
    tagespauschale = eb.get("tagespauschale")
    if tagespauschale in (None, "") and event_date is not None:
        tagespauschale = miete_fuer_datum(event_date)
    reinigung = eb.get("reinigung")
    abendverantwortung = eb.get("abendverantwortung")

    abzuege = umsatzbeteiligung + ek
    for v in (tagespauschale, reinigung, abendverantwortung):
        if v not in (None, ""):
            abzuege += float(v)
    abzuege = round(abzuege, 2)
    ticket_netto = float(ticket_netto or 0.0)
    einnahmen = round(sumup_netto + twint_netto + ticket_netto, 2)
    auszahlung = round(einnahmen - abzuege, 2)

    def opt(v):  # optionaler Pauschalwert -> Zahl oder leer
        return chf(v) if v not in (None, "") else ""

    mapping = {
        "{{SUMUP_BRUTTO}}": chf(sumup_brutto), "{{SUMUP_GEBUEHR}}": chf(sumup_geb),
        "{{SUMUP_NETTO}}": chf(sumup_netto),
        "{{TWINT_BRUTTO}}": chf(twint_brutto), "{{TWINT_GEBUEHR}}": chf(twint_geb),
        "{{TWINT_NETTO}}": chf(twint_netto),
        "{{UMSATZBETEILIGUNG}}": chf(umsatzbeteiligung),
        "{{EK_GETRAENKE}}": chf(ek),
        "{{TAGESPAUSCHALE}}": opt(tagespauschale), "{{REINIGUNG}}": opt(reinigung),
        "{{ABENDVERANTWORTUNG}}": opt(abendverantwortung),
        "{{TOTAL_ABZUEGE}}": chf(abzuege), "{{TOTAL_EINNAHMEN}}": chf(einnahmen),
        "{{AUSZAHLUNG}}": chf(auszahlung),
    }
    zahlen = {"sumup_brutto": sumup_brutto, "sumup_geb": sumup_geb,
              "sumup_netto": sumup_netto, "twint_brutto": twint_brutto,
              "twint_geb": twint_geb, "twint_netto": twint_netto,
              "umsatzbeteiligung": umsatzbeteiligung, "ek": ek,
              "miete": float(tagespauschale or 0), "abzuege": abzuege,
              "einnahmen": einnahmen, "auszahlung": auszahlung,
              "ticket_netto": ticket_netto}
    return mapping, zahlen


# ----------------------------------------------------------------------------
# AUTOMATIK: Event-Erkennung (4h Stille = Event zu Ende), Abrechnung, Mailversand
# ----------------------------------------------------------------------------
def tx_time(t):
    for k in ("timestamp", "time", "created_at", "local_time"):
        if t.get(k):
            try:
                return parse_iso(str(t[k]))
            except Exception:
                continue
    return None


def cluster_events(transactions, gap_hours=4.0):
    """Sortiert Transaktionen zeitlich und trennt Events an Luecken > gap_hours."""
    ts = [(tx_time(t), t) for t in transactions]
    ts = sorted([x for x in ts if x[0] is not None], key=lambda x: x[0])
    events, cur = [], []
    for when, t in ts:
        if cur and (when - cur[-1][0]).total_seconds() > gap_hours * 3600:
            events.append(cur)
            cur = []
        cur.append((when, t))
    if cur:
        events.append(cur)
    return events  # Liste von Listen [(zeit, tx), ...]


def load_state(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"abgerechnet": []}


def save_state(path, state):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def send_mail(cfg, subject, body, attachments):
    """Versendet die Abrechnung per SMTP (z.B. Gmail mit App-Passwort)."""
    import smtplib
    from email.message import EmailMessage
    import mimetypes

    em = (cfg.get("email") or {})
    host = em.get("smtp_host", "smtp.gmail.com")
    port = int(em.get("smtp_port", 587))
    user = em.get("user") or os.environ.get("MAIL_USER")
    pw = em.get("app_password") or os.environ.get("MAIL_APP_PASSWORD")
    to = em.get("to") or os.environ.get("MAIL_TO")
    if not (user and pw and to):
        raise RuntimeError("E-Mail nicht konfiguriert: config.json -> email "
                           "{smtp_host, smtp_port, user, app_password, to}")

    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    for path in attachments:
        if not path or not os.path.exists(path):
            continue
        ctype, _ = mimetypes.guess_type(path)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype,
                               filename=os.path.basename(path))
    with smtplib.SMTP(host, port, timeout=60) as srv:
        srv.starttls()
        srv.login(user, pw)
        srv.send_message(msg)


def run_event(tx, receipts_fn, ek, twint, cfg, event_name, von_s, bis_s,
              out_dir, rate, event_date=None):
    """Kompletter Abrechnungslauf fuer ein Event -> (res, zahlen, pfade)."""
    res = compute(tx, receipts_fn, ek, twint, rate=rate)
    mapping, zahlen = eventblatt_values(res, cfg, event_date=event_date)
    created = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = {"event": event_name, "von": von_s, "bis": bis_s, "created": created}
    safe = "".join(c if c.isalnum() or c in "-_ " else "_"
                   for c in event_name).strip().replace(" ", "_")
    os.makedirs(out_dir, exist_ok=True)
    xlsx_path = os.path.join(out_dir, f"Abrechnung_{safe}.xlsx")
    pdf_path = os.path.join(out_dir, f"Abrechnung_{safe}.pdf")
    write_excel(res, meta, xlsx_path, zahlen)
    recalc_excel(xlsx_path)
    write_pdf(res, meta, pdf_path, zahlen)

    eventblatt_path = None
    try:
        eb_cfg = (cfg.get("eventblatt") or {})
        template = eb_cfg.get("template")
        if not template or not os.path.exists(template):
            template = os.path.join(out_dir, "EVENTBLATT_Vorlage.docx")
            if not os.path.exists(template):
                build_eventblatt_vorlage(template)
        eventblatt_path = os.path.join(out_dir, f"Eventblatt_{safe}.docx")
        fill_eventblatt(template, eventblatt_path, mapping)
    except Exception as e:
        print(f"  ⚠ Eventblatt konnte nicht erstellt werden: {e}")

    return res, zahlen, {"xlsx": xlsx_path, "pdf": pdf_path,
                         "eventblatt": eventblatt_path}


def mail_body(event_name, res, zahlen):
    z = zahlen or {}
    lines = [
        f"Automatische Abrechnung: {event_name}",
        "",
        f"Gesamtumsatz (alle Kategorien):  CHF {res['umsatz']:.2f}",
        f"  davon Karte (SumUp):           CHF {res['card_total']:.2f}",
        f"  davon TWINT (bar getippt):     CHF {res['cash_total']:.2f}",
        "",
        "Abzüge:",
        f"  13% Umsatzbeteiligung:         CHF {res['beteiligung']:.2f}",
        f"  Einkauf Getränke (inkl. Mitarbeiter/Gratis): CHF {res['getraenkeschuld']:.2f}",
    ]
    if z:
        lines += [
            f"  Miete (nach Wochentag):        CHF {z['miete']:.2f}",
            f"  SumUp-Gebühren (2.5% auf Karte): CHF {z['sumup_geb']:.2f}",
            f"  TWINT-Gebühren (1.3% auf TWINT): CHF {z['twint_geb']:.2f}",
            "",
            f"Einnahmen netto:                 CHF {z['einnahmen']:.2f}",
            f"Total Abzüge:                    CHF {z['abzuege']:.2f}",
            f"ÜBERWEISUNG an euch (Reingewinn): CHF {z['auszahlung']:.2f}",
        ]
    if res.get("unknown"):
        lines += ["", "Ohne EK-Abzug (Fremdkategorie): " + ", ".join(res["unknown"])]
    if res.get("no_items"):
        lines += [f"Hinweis: {res['no_items']} Transaktion(en) ohne Artikel "
                  f"(freier Betrag getippt) – ohne EK-Abzug."]
    lines += ["", "Details im angehängten Eventblatt (Word), Excel und PDF.",
              "Gebührensätze: SumUp 2.5% auf Kartenzahlungen, TWINT 1.3%."]
    return "\n".join(lines)


def catchup_run(args, cfg):
    """Einmalige Nachhol-Abrechnung fuer einen festen Zeitraum (--von/--bis).

    Erkennt alle Events darin (4h-Stille-Clustering) und mailt jedes einzeln.
    Mit --card-only wird alles als Karte gewertet (fuer Alt-Events ohne TWINT)."""
    api_key = cfg.get("sumup_api_key") or os.environ.get("SUMUP_API_KEY")
    if not api_key:
        raise SystemExit("SumUp API-Key fehlt (config.json -> sumup_api_key "
                         "oder Umgebungsvariable SUMUP_API_KEY).")
    if not (args.von and args.bis):
        raise SystemExit("--catchup braucht --von und --bis (z.B. "
                         "--von 2026-06-30T00:00:00+02:00 --bis 2026-07-20T00:00:00+02:00).")
    s = _session(api_key)
    mid = get_merchant_code(s, cfg)
    ek = load_ek_prices(args.ek)
    gap = float((cfg.get("auto") or {}).get("gap_hours", 4))
    if args.card_only:
        cfg = dict(cfg); cfg["_card_only"] = True

    von, bis = parse_iso(args.von), parse_iso(args.bis)
    print(f"Nachhol-Abrechnung {von:%d.%m.%Y} – {bis:%d.%m.%Y}"
          f"{'  (alles als Karte)' if args.card_only else ''}")
    tx = fetch_transactions(s, von, bis)
    events = cluster_events(tx, gap_hours=gap)
    print(f"{len(tx)} Transaktionen, {len(events)} Event(s) erkannt.\n")

    for ev in events:
        first_t, last_t = ev[0][0], ev[-1][0]
        lokal = first_t.astimezone()
        name = f"Event_{lokal:%Y-%m-%d}"
        print(f"→ {name}: {first_t.astimezone():%d.%m %H:%M} – "
              f"{last_t.astimezone():%H:%M}  ({len(ev)} Transaktionen)")
        ev_tx = [t for _, t in ev]
        twint = {"total": 0.0, "count": 0, "note": "Nachhol-Abrechnung"}
        receipts_fn = lambda t: fetch_receipt_items(s, mid, t)
        res, zahlen, paths = run_event(
            ev_tx, receipts_fn, ek, twint, cfg, name,
            first_t.astimezone().isoformat(), last_t.astimezone().isoformat(),
            args.out, args.rate, event_date=lokal)
        print(f"   Umsatz {res['umsatz']:.2f} | Auszahlung "
              f"{(zahlen or {}).get('auszahlung', 0):.2f}")
        if args.no_mail:
            continue
        try:
            send_mail(cfg, f"R1 Nachhol-Abrechnung {name} – Auszahlung CHF "
                           f"{(zahlen or {}).get('auszahlung', 0):.2f}",
                      "NACHTRÄGLICHE ABRECHNUNG (Alt-Event)\n\n" + mail_body(name, res, zahlen),
                      [paths["eventblatt"], paths["xlsx"], paths["pdf"]])
            print(f"   ✉ Mail versendet an {(cfg.get('email') or {}).get('to')}")
        except Exception as e:
            print(f"   ⚠ Mailversand fehlgeschlagen: {e} (Dateien in {args.out})")
    print("\nFertig.")


def auto_run(args, cfg, once=True):
    """Prueft auf beendete Events (>=4h Stille) und rechnet neue automatisch ab."""
    api_key = cfg.get("sumup_api_key") or os.environ.get("SUMUP_API_KEY")
    if not api_key:
        raise SystemExit("SumUp API-Key fehlt (config.json -> sumup_api_key "
                         "oder Umgebungsvariable SUMUP_API_KEY).")
    s = _session(api_key)
    mid = get_merchant_code(s, cfg)
    ek = load_ek_prices(args.ek)
    gap = float((cfg.get("auto") or {}).get("gap_hours", 4))
    lookback = float((cfg.get("auto") or {}).get("lookback_hours", args.lookback_hours))
    state_path = os.path.join(args.out, "auto_state.json")

    while True:
        now = dt.datetime.now(dt.timezone.utc)
        von = now - dt.timedelta(hours=lookback)
        try:
            tx = fetch_transactions(s, von, now)
        except Exception as e:
            print(f"[{now:%H:%M}] API-Fehler: {e}")
            tx = []
        state = load_state(state_path)
        done_ids = set(state.get("abgerechnet", []))
        events = cluster_events(tx, gap_hours=gap)
        neu = 0
        for ev in events:
            first_t, last_t = ev[0][0], ev[-1][0]
            ev_id = last_t.isoformat()
            fertig = (now - last_t).total_seconds() >= gap * 3600
            if not fertig or ev_id in done_ids:
                continue
            lokal = first_t.astimezone()
            name = f"Event_{lokal:%Y-%m-%d}"
            print(f"→ Abgeschlossenes Event erkannt: {name} "
                  f"({first_t:%d.%m %H:%M} – {last_t:%d.%m %H:%M} UTC, "
                  f"{len(ev)} Transaktionen)")
            ev_tx = [t for _, t in ev]
            twint = {"total": 0.0, "count": 0,
                     "note": "Auto-Modus: bar getippt = TWINT (Gebührenbasis)"}
            receipts_fn = lambda t: fetch_receipt_items(s, mid, t)
            res, zahlen, paths = run_event(
                ev_tx, receipts_fn, ek, twint, cfg, name,
                first_t.astimezone().isoformat(), last_t.astimezone().isoformat(),
                args.out, args.rate, event_date=lokal)
            try:
                send_mail(cfg, f"R1 Abrechnung {name} – Auszahlung CHF "
                               f"{(zahlen or {}).get('auszahlung', 0):.2f}",
                          mail_body(name, res, zahlen),
                          [paths["eventblatt"], paths["xlsx"], paths["pdf"]])
                print(f"  ✉ Mail versendet an {(cfg.get('email') or {}).get('to')}")
            except Exception as e:
                print(f"  ⚠ Mailversand fehlgeschlagen: {e} "
                      f"(Dateien liegen in {args.out})")
            done_ids.add(ev_id)
            state["abgerechnet"] = sorted(done_ids)[-200:]
            save_state(state_path, state)
            neu += 1
        if neu == 0:
            print(f"[{dt.datetime.now():%d.%m %H:%M}] Kein neues beendetes Event.")
        if once:
            return
        import time
        time.sleep(15 * 60)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Tutti Tschutti Event-Abrechnung (Variante 3)")
    ap.add_argument("--event", default="Event")
    ap.add_argument("--von", help="Start ISO-8601, z.B. 2026-06-11T17:00:00+02:00")
    ap.add_argument("--bis", help="Ende ISO-8601")
    ap.add_argument("--twint", help="Pfad zur RaiseNow/TWINT-CSV")
    ap.add_argument("--ek", default="ek_preise.csv", help="Pfad zur EK-Preisliste (CSV)")
    ap.add_argument("--config", default="config.json", help="Pfad zur config.json")
    ap.add_argument("--rate", type=float, default=0.13, help="Umsatzbeteiligung (Default 0.13)")
    ap.add_argument("--out", default=".", help="Ausgabeordner")
    ap.add_argument("--demo", action="store_true", help="Mit eingebauten Demodaten testen")
    ap.add_argument("--inspect", action="store_true",
                    help="Nur die vorkommenden payment_type/card_type Werte ausgeben")
    ap.add_argument("--ki-email", action="store_true", help="Abrechnungs-E-Mail via KI erzeugen")
    ap.add_argument("--test-connection", action="store_true",
                    help="Prüft nur, ob der SumUp-API-Key funktioniert (ruft /me auf).")
    ap.add_argument("--auto", action="store_true",
                    help="Automatik: beendete Events (>=4h Stille) abrechnen + mailen, dann Ende. Für cron.")
    ap.add_argument("--watch", action="store_true",
                    help="Wie --auto, aber als Dauerschleife (prüft alle 15 Min).")
    ap.add_argument("--lookback-hours", type=float, default=72,
                    help="Automatik: wie weit zurück nach Events gesucht wird (Default 72h).")
    ap.add_argument("--catchup", action="store_true",
                    help="Einmalige Nachhol-Abrechnung für --von/--bis (alle Events darin).")
    ap.add_argument("--card-only", action="store_true",
                    help="Catchup: alles als Karte werten (Alt-Events ohne TWINT/bar).")
    ap.add_argument("--no-mail", action="store_true",
                    help="Catchup: keine Mails senden, nur Dateien erzeugen.")
    args = ap.parse_args()

    cfg = load_config(args.config)

    if args.test_connection:
        api_key = cfg.get("sumup_api_key") or os.environ.get("SUMUP_API_KEY")
        if not api_key:
            ap.error("SumUp API-Key fehlt (config.json -> sumup_api_key oder Umgebungsvariable SUMUP_API_KEY).")
        s = _session(api_key)
        r = s.get(EP_ME, timeout=30)
        if r.status_code != 200:
            print(f"❌ Verbindung fehlgeschlagen (HTTP {r.status_code}): {r.text[:200]}")
            print("   -> Key falsch/abgelaufen oder ohne Leserechte?")
            return
        data = r.json()
        prof = data.get("merchant_profile") or {}
        print("✅ Verbindung OK")
        print("   Händler:", prof.get("company_name") or prof.get("legal_name") or "(unbekannt)")
        print("   merchant_code:", prof.get("merchant_code") or data.get("merchant_code") or "(nicht gefunden)")
        return

    if args.catchup:
        catchup_run(args, cfg)
        return

    if args.auto or args.watch:
        auto_run(args, cfg, once=args.auto and not args.watch)
        return

    if args.demo:
        tx, receipts, ek, twint = demo_data()
        receipts_fn = lambda t: receipts.get(t.get("transaction_id"), [])
        von_s, bis_s = "Demo", "Demo"
        event_date = dt.datetime.now()
    else:
        if not (args.von and args.bis):
            ap.error("--von und --bis sind im Echtlauf nötig (oder nutze --demo).")
        von, bis = parse_iso(args.von), parse_iso(args.bis)
        von_s, bis_s = args.von, args.bis
        api_key = cfg.get("sumup_api_key") or os.environ.get("SUMUP_API_KEY")
        if not api_key:
            ap.error("SumUp API-Key fehlt (config.json -> sumup_api_key oder Umgebungsvariable SUMUP_API_KEY).")
        s = _session(api_key)
        mid = get_merchant_code(s, cfg)
        tx = fetch_transactions(s, von, bis)

        if args.inspect:
            seen = {}
            for t in tx:
                k = (t.get("payment_type"), t.get("card_type"), t.get("type"))
                seen[k] = seen.get(k, 0) + 1
            print("Vorkommende (payment_type, card_type, type) Werte:")
            for k, n in sorted(seen.items(), key=lambda x: -x[1]):
                print(f"  {k}: {n}x")
            print("\n-> CARD_PAYMENT_TYPES / CASH_PAYMENT_TYPES im Script ggf. anpassen.")
            return

        ek = load_ek_prices(args.ek)
        twint = load_twint_csv(args.twint, von, bis, cfg)
        receipts_fn = lambda t: fetch_receipt_items(s, mid, t)
        event_date = von

    res, zahlen, paths = run_event(tx, receipts_fn, ek, twint, cfg, args.event,
                                   von_s, bis_s, args.out, args.rate,
                                   event_date=event_date)
    xlsx_path, pdf_path = paths["xlsx"], paths["pdf"]
    eventblatt_path = paths["eventblatt"]
    safe = "".join(c if c.isalnum() or c in "-_ " else "_"
                   for c in args.event).strip().replace(" ", "_")
    created = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = {"event": args.event, "von": von_s, "bis": bis_s, "created": created}

    print(f"\n=== {args.event} ===")
    print(f"Umsatz:            CHF {res['umsatz']:.2f}")
    print(f"Beteiligung {args.rate*100:.0f}%:   CHF {res['beteiligung']:.2f}")
    print(f"Getränkeschuld:    CHF {res['getraenkeschuld']:.2f}")
    print(f"Abzüge R1 (13%+EK): CHF {res['total_owed']:.2f}")
    print(f"  Karte: {res['card_total']:.2f} | TWINT (bar getippt): {res['cash_total']:.2f}")
    if res["unknown"]:
        print("  ohne EK-Abzug (Fremdkategorie):", ", ".join(res["unknown"]))
    if zahlen:
        print(f"  Miete: {zahlen['miete']:.2f} | SumUp-Geb. {zahlen['sumup_geb']:.2f} | "
              f"TWINT-Geb. {zahlen['twint_geb']:.2f}")
        print(f"  ÜBERWEISUNG an Nutzung: CHF {zahlen['auszahlung']:.2f}")
    print(f"\nGeschrieben:\n  {xlsx_path}\n  {pdf_path}")
    if eventblatt_path:
        print(f"  {eventblatt_path}")

    if args.ki_email:
        txt = ki_email_text(res, meta)
        if txt:
            with open(os.path.join(args.out, f"Email_{safe}.txt"), "w", encoding="utf-8") as f:
                f.write(txt)
            print(f"  {os.path.join(args.out, f'Email_{safe}.txt')}")
        else:
            print("  (KI-E-Mail übersprungen: ANTHROPIC_API_KEY nicht gesetzt)")


if __name__ == "__main__":
    main()
