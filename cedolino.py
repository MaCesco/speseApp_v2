"""Lettura del cedolino PDF (Electrolux/Zucchetti) con pdfplumber.
Un cedolino puo' avere piu' pagine (di solito 2 o 3): la prima ha le voci e il calendario, l'ultima i totali,
la tabella ferie/PAR, la legenda e i dati fiscali. Ogni pagina viene letta e i dati vengono uniti."""
import io, re

MESI = "gennaio febbraio marzo aprile maggio giugno luglio agosto settembre ottobre novembre dicembre".split()
NUM = re.compile(r"^-?\d[\d.]*,\d+$")
CAMPI_PAG = ("minimo", "scatti", "sup_ass", "tot_retr", "trattenute", "competenze", "netto", "giorni", "cometa", "tfr_fondo", "tfe",
             "fe_prec", "fe_spett", "fe_god", "_fe_saldo", "par_prec", "par_spett", "par_god", "_par_saldo")


def _n(t):
    try:
        return float(t.replace(".", "").replace(",", "."))
    except (ValueError, AttributeError):
        return None


def _righe(words, tol=3):
    righe = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if righe and abs(righe[-1][0] - w["top"]) <= tol:
            righe[-1][1].append(w)
        else:
            righe.append([w["top"], [w]])
    return [(t, sorted(ws, key=lambda w: w["x0"])) for t, ws in righe]


def _pagina(pg):
    """Legge una pagina: restituisce {valori, voci, eventi (ore per codice), legenda}."""
    testo = pg.extract_text(x_tolerance=2, y_tolerance=3) or ""
    words = pg.extract_words(x_tolerance=2)
    righe = _righe(words)
    v = {}
    # testata: minimo, scatti, sup. ass. (numeri sotto le intestazioni)
    for t, ws in righe:
        if any(w["text"] == "MINIMO" for w in ws) and any(w["text"] == "SCATTI" for w in ws):
            intest = {w["text"]: (w["x0"] + w["x1"]) / 2 for w in ws if w["text"] in ("MINIMO", "SCATTI", "SUP.ASS.")}
            for w in words:
                if t + 4 < w["top"] < t + 22 and NUM.match(w["text"]):
                    c = (w["x0"] + w["x1"]) / 2
                    k = min(intest, key=lambda k: abs(intest[k] - c))
                    if abs(intest[k] - c) < 30:
                        v[{"MINIMO": "minimo", "SCATTI": "scatti", "SUP.ASS.": "sup_ass"}[k]] = _n(w["text"])
            break
    for w in words:  # totale retributivo: numero in alto a destra
        if 255 < w["top"] < 300 and w["x0"] > 400 and NUM.match(w["text"]):
            v["tot_retr"] = _n(w["text"])
    # colonne della tabella voci
    inte = {}
    for w in words:
        if w["text"] in ("Giorni", "unitario", "Trattenute", "Competenze") and 295 < w["top"] < 310 and w["x0"] < 400:
            inte[w["text"]] = w["x1"]
    voci = []
    if len(inte) == 4:
        lim = [(inte["Giorni"] + inte["unitario"]) / 2, (inte["unitario"] + inte["Trattenute"]) / 2, (inte["Trattenute"] + inte["Competenze"]) / 2]
        col = lambda x1: "ore" if x1 < lim[0] else "unit" if x1 < lim[1] else "tratt" if x1 < lim[2] else "comp"
        y_tot = min([w["top"] for w in words if w["text"] == "Totale" and w["top"] > 500] or [640])
        for t, ws in righe:
            if not (315 < t < y_tot - 8):
                continue
            ws = [w for w in ws if w["x0"] < 405]
            cod = next((w for w in ws if re.fullmatch(r"\d{4}", w["text"]) and w["x0"] < 80), None)
            if not cod:
                continue
            r = {"cod": cod["text"], "desc": " ".join(w["text"] for w in ws if w["x0"] > cod["x0"] and not NUM.match(w["text"]) and w["x0"] < 200).strip()}
            for w in ws:
                if NUM.match(w["text"]) and w["x0"] > cod["x0"]:
                    r[col(w["x1"])] = _n(w["text"])
            voci.append(r)
        # totali (riga "Totale": trattenute e competenze)
        for w in words:
            if NUM.match(w["text"]) and y_tot - 22 < w["top"] < y_tot + 3 and w["x0"] < 405 and w["x1"] > lim[1]:
                c = col(w["x1"])
                if c in ("tratt", "comp"):
                    v["trattenute" if c == "tratt" else "competenze"] = _n(w["text"])
    mn = re.search(r"([\d.]+,\d\d)\s+Netto da pagare", testo)
    if mn:
        v["netto"] = _n(mn.group(1))
    def cerca(p):
        mm = re.search(p, testo)
        return _n(mm.group(1)) if mm else None
    v["giorni"] = cerca(r"G\.LAV\.TI\s+([\d.]+,\d+)")
    v["cometa"] = cerca(r"COMETA \(METAL\s+([\d.]+,\d+)")
    v["tfr_fondo"] = cerca(r"TFR\s+F\.DI\s+([\d.]+,\d+)")
    v["tfe"] = cerca(r"TFR 12 AP\s+([\d.]+,\d+)")
    # tabella ferie / PAR (sulla pagina dove e' vuota non ci sono le righe FERIE / PAR)
    hdr = {}
    for w in words:
        if w["top"] > 600 and w["text"] in ("prec.", "Spett.", "Godute", "Saldo") and w["x0"] > 290:
            hdr[w["text"]] = w["x1"]
    if len(hdr) == 4:
        for nome_r, pre in (("FERIE", "fe"), ("PAR", "par")):
            rw = next((w for w in words if w["text"] == nome_r and w["x0"] > 290 and w["top"] > 600), None)
            nums = [w for w in words if rw and abs(w["top"] - rw["top"]) < 3 and w["x0"] > rw["x1"] and NUM.match(w["text"])]
            if not nums:
                continue  # tabella vuota in questa pagina
            for k in ("prec", "spett", "god", "saldo"):
                v[f"{pre}_{k}"] = None
            for w in nums:
                k = min(hdr, key=lambda k: abs(hdr[k] - w["x1"]))
                v[f"{pre}_{ {'prec.': 'prec', 'Spett.': 'spett', 'Godute': 'god', 'Saldo': 'saldo'}[k] }"] = _n(w["text"])
            v[f"_{pre}_saldo"] = v.pop(f"{pre}_saldo", None)
    # calendario presenze: ore per codice evento (le ore sono scritte subito prima del codice: "8 00 WS")
    ev = {}
    for t, ws in righe:
        r = [w for w in ws if w["x0"] > 350]
        if r and re.fullmatch(r"[LMGVSD]\d\d\*?", r[0]["text"]):
            for i, w in enumerate(r):
                if re.fullmatch(r"[A-Z]{2}", w["text"]) and i >= 2 and r[i - 1]["text"].isdigit() and r[i - 2]["text"].isdigit():
                    ev[w["text"]] = ev.get(w["text"], 0.0) + float(r[i - 2]["text"]) + float(r[i - 1]["text"]) / 100
    leg = {}
    for cod, val in re.findall(r"\b([A-Z]{2})=[A-Z' ]+?\s+([\d.]+,\d+)", testo):
        leg[cod] = _n(val)
    m = re.search(r"Periodo:\s*([A-Za-zÀ-ú]+)\s+(\d{4})", testo)
    return {"v": v, "voci": voci, "ev": ev, "leg": leg, "periodo": (m.group(1).lower(), int(m.group(2))) if m else None}


def leggi_cedolino(dati):
    """Restituisce (valori, voci, controlli) oppure solleva ValueError se non sembra un cedolino."""
    import pdfplumber
    try:
        pdf = pdfplumber.open(io.BytesIO(dati))
        pagine = [_pagina(pg) for pg in pdf.pages]
    except Exception as e:
        raise ValueError("PDF non leggibile") from e
    per = next((p["periodo"] for p in pagine if p["periodo"]), None)
    if not per:
        raise ValueError("Non trovo il periodo: non sembra un cedolino")
    nome, anno = per
    tipo = "mensile"
    if nome.startswith(("tredic", "13")):
        mese, tipo = 12, "tredicesima"
    elif nome in MESI:
        mese = MESI.index(nome) + 1
    else:
        raise ValueError(f"Periodo non riconosciuto: {nome} {anno}")
    v = {"anno": anno, "mese": mese, "tipo": tipo}
    voci, ev, leg = [], {}, {}
    for p in pagine:  # unione: per ogni dato vale la prima pagina dove e' presente
        for k, val in p["v"].items():
            if val is not None and v.get(k) is None:
                v[k] = val
        voci += p["voci"]
        for k, val in p["ev"].items():
            ev[k] = ev.get(k, 0.0) + val
        for k, val in p["leg"].items():
            leg.setdefault(k, val)
    for k in ("fe_god", "par_god"):  # nelle tabelle la colonna "Godute" puo' essere vuota
        v.setdefault(k, None)
    if v.get("competenze") is None or v.get("trattenute") is None:
        raise ValueError("Non trovo i totali del cedolino (competenze e trattenute)")
    def voce(cod, campo):
        return sum(r.get(campo) or 0 for r in voci if r["cod"] == cod) if any(r["cod"] == cod for r in voci) else None
    v["edp"], v["anticipo"] = voce("2978", "comp"), voce("3693", "comp")
    v["gg_mensa"], v["mensa_tratt"] = voce("4990", "ore"), voce("4990", "tratt")
    v["unitario"] = next((r.get("unit") for r in voci if r["cod"] == "2000"), None)
    v["_sw_ore"], v["_fe_desc"], v["_par_desc"] = voce("1273", "ore"), voce("4802", "ore"), voce("4806", "ore")
    # controlli interni al PDF
    c = []
    def ctrl(nome, a, b, tol=0.02, um="€"):
        if a is None or b is None:
            return
        c.append({"nome": nome, "a": a, "b": b, "ok": abs(a - b) <= tol, "um": um})
    ctrl("Somma delle voci a competenza = totale competenze", round(sum(r.get("comp") or 0 for r in voci), 2), v.get("competenze"))
    ctrl("Somma delle voci a trattenuta = totale trattenute", round(sum(r.get("tratt") or 0 for r in voci), 2), v.get("trattenute"))
    if v.get("netto") is not None:
        ctrl("Competenze − trattenute = netto da pagare", round(v["competenze"] - v["trattenute"], 2), v["netto"])
    for cod, nome_e, desc in (("FR", "Ferie", v.get("_fe_desc")), ("WS", "Smart working", v.get("_sw_ore")), ("PA", "Permessi (PAR)", v.get("_par_desc"))):
        if cod in leg and cod in ev:
            ctrl(f"{nome_e}: ore nel calendario = ore in legenda", round(ev[cod], 2), leg[cod], um="h")
        if desc is not None and cod in leg:
            ctrl(f"{nome_e}: ore in descrizione = ore in legenda", desc, leg[cod], um="h")
        elif desc is not None and cod in ev:
            ctrl(f"{nome_e}: ore in descrizione = ore nel calendario", desc, round(ev[cod], 2), um="h")
    for pre, nome_r in (("fe", "Ferie"), ("par", "PAR")):
        if all(v.get(f"{pre}_{k}") is not None for k in ("prec", "spett")) and v.get(f"_{pre}_saldo") is not None:
            ctrl(f"{nome_r}: riporto + spettanti − godute = saldo", round(v[f"{pre}_prec"] + v[f"{pre}_spett"] - (v.get(f"{pre}_god") or 0), 2), v[f"_{pre}_saldo"], um="h")
    v["_ev"], v["_leg"], v["_pagine"] = {k: round(x, 2) for k, x in ev.items()}, leg, len(pagine)
    return v, voci, c
