import io, os, re, secrets
from datetime import date, datetime, time, timedelta, timezone
from functools import wraps
import click
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, abort, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

app = Flask(__name__)
uri = os.environ.get("DATABASE_URL", "sqlite:///spese.db")
if uri.startswith("postgres"):  # accetta qualsiasi forma della stringa di Neon
    uri = "postgresql+psycopg://" + uri.split("://", 1)[1]
app.config["SQLALCHEMY_DATABASE_URI"] = uri
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = "DATABASE_URL" in os.environ  # solo online (HTTPS)
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True, "pool_recycle": 280}  # database che va in pausa
app.secret_key = os.environ.get("SECRET_KEY", "cambiami-in-produzione")
db = SQLAlchemy(app)

ENTRATE = ["Stipendio", "Ripetizioni", "Concerti/Matrimoni", "Entrate extra"]
USCITE = ["Benzina auto", "Benzina moto", "Ricariche", "Biglietti treno/Autobus", "Parcheggio auto / moto",
          "Autostrada", "Libri", "Lavaggio Auto", "Abbigliamento", "Parrucchiera/Barbiere", "Supermercato",
          "Bar/Aperitivi", "Pranzi/Cene", "Cinema", "Discoteca", "Mostre/Fiere/Concerti/Giostre", "Viaggi",
          "Accessori/Computer", "Sport", "Assicurazione auto / moto", "Tagliandi auto / moto", "Accessori moto",
          "Farmaci / medicinali / Visite mediche", "Regali", "Corsi", "Uscite extra"]


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)


class Movimento(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    data = db.Column(db.Date, nullable=False)
    tipo = db.Column(db.String(10), nullable=False)  # entrata | uscita
    categoria = db.Column(db.String(80), nullable=False)
    importo = db.Column(db.Float, nullable=False)
    descrizione = db.Column(db.String(200), default="")
    fonte = db.Column(db.String(10), default="manuale")  # manuale | import | ricorrente


class Budget(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    categoria = db.Column(db.String(80), nullable=False)
    importo = db.Column(db.Float, nullable=False)


class Ricorrente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    tipo = db.Column(db.String(10), nullable=False)
    categoria = db.Column(db.String(80), nullable=False)
    descrizione = db.Column(db.String(200), default="")
    importo = db.Column(db.Float, nullable=False)
    giorno = db.Column(db.Integer, nullable=False)
    ultimo = db.Column(db.String(7), nullable=False)  # ultimo mese generato, YYYY-MM


class Studente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    nome = db.Column(db.String(120), nullable=False)
    tariffa = db.Column(db.Float, default=20)
    note = db.Column(db.String(500), default="")
    credito = db.Column(db.Float, default=0)  # euro pagati in anticipo, non ancora usati


class Lezione(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    studente_id = db.Column(db.Integer, db.ForeignKey("studente.id"), nullable=False, index=True)
    data = db.Column(db.Date, nullable=False)
    ora_inizio = db.Column(db.String(5), default="")
    ora_fine = db.Column(db.String(5), default="")
    ore = db.Column(db.Float, nullable=False)
    materia = db.Column(db.String(80), default="")
    modalita = db.Column(db.String(10), default="")  # Presenza | Online
    pagata = db.Column(db.Boolean, default=False, nullable=False)
    data_pagamento = db.Column(db.Date)
    fonte = db.Column(db.String(10), default="manuale")
    tariffa = db.Column(db.Float)  # tariffa applicata a questa lezione
    studente = db.relationship("Studente", backref="lezioni")


class FileCaricato(db.Model):
    """Ultimo file Excel caricato (uno per tipo): resta salvato nel database online."""
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    tipo = db.Column(db.String(20), nullable=False)
    nome = db.Column(db.String(200), nullable=False)
    caricato_il = db.Column(db.DateTime, nullable=False)  # UTC
    contenuto = db.Column(db.LargeBinary, nullable=False)


def migra():
    """Aggiunge le colonne nuove alle tabelle gia' esistenti (database online)."""
    from sqlalchemy import inspect, text
    insp = inspect(db.engine)
    for tab, col, tipo in (("studente", "note", "VARCHAR(500)"), ("lezione", "tariffa", "FLOAT"), ("studente", "credito", "FLOAT")):
        if col not in [c["name"] for c in insp.get_columns(tab)]:
            db.session.execute(text(f"ALTER TABLE {tab} ADD COLUMN {col} {tipo}"))
    db.session.execute(text("UPDATE lezione SET tariffa = (SELECT studente.tariffa FROM studente WHERE studente.id = lezione.studente_id) WHERE tariffa IS NULL"))
    db.session.execute(text("UPDATE studente SET credito = 0 WHERE credito IS NULL"))
    db.session.commit()


with app.app_context():
    db.create_all()
    migra()


# ---------- sicurezza ----------
def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if "uid" not in session:
            return redirect(url_for("login"))
        return f(*a, **k)
    return w


@app.before_request
def csrf():
    if request.method == "POST" and request.form.get("_t") != session.get("_t"):
        abort(400)


@app.context_processor
def ctx():
    session.setdefault("_t", secrets.token_hex(16))
    return {"csrf": session["_t"], "ENTRATE": ENTRATE, "USCITE": USCITE}


@app.cli.command("crea-utente")
@click.argument("username")
@click.password_option()
def crea_utente(username, password):
    if User.query.first():
        print("Esiste gia' un utente. Per cambiare password usa: flask --app app reimposta-password")
        return
    if len(password) < 10:
        print("Password troppo corta (minimo 10 caratteri).")
        return
    db.session.add(User(username=username, password=generate_password_hash(password)))
    db.session.commit()
    print("Utente creato.")


@app.cli.command("reimposta-password")
@click.password_option()
def reimposta_password(password):
    u = User.query.first()
    if not u:
        print("Nessun utente: usa prima crea-utente.")
    elif len(password) < 10:
        print("Password troppo corta (minimo 10 caratteri).")
    else:
        u.password = generate_password_hash(password)
        db.session.commit()
        print(f"Password aggiornata per l'utente: {u.username}")


# ---------- utilita ----------
def num(s):
    return float(str(s).replace(",", "."))


def applica_ricorrenti(uid):
    oggi = date.today()
    for r in Ricorrente.query.filter_by(user_id=uid):
        y, m = map(int, r.ultimo.split("-"))
        while True:
            m += 1
            if m > 12:
                y, m = y + 1, 1
            d = date(y, m, min(r.giorno, 28))
            if d > oggi:
                break
            db.session.add(Movimento(user_id=uid, data=d, tipo=r.tipo, categoria=r.categoria,
                                     importo=r.importo, descrizione=r.descrizione, fonte="ricorrente"))
            r.ultimo = f"{y}-{m:02d}"
    db.session.commit()


# ---------- pagine ----------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = User.query.filter_by(username=request.form["username"].strip()).first()
        if u and check_password_hash(u.password, request.form["password"]):
            session.clear()
            session["uid"] = u.id
            return redirect(url_for("home"))
        flash("Username o password errati.")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def home():
    applica_ricorrenti(session["uid"])
    return render_template("dashboard.html")


@app.route("/api/dati")
@login_required
def dati():
    uid = session["uid"]
    mesi = {}
    q = db.session.query(Movimento.data, Movimento.tipo, Movimento.categoria, Movimento.importo)
    for d, t, c, i in q.filter(Movimento.user_id == uid):
        k = mesi.setdefault(f"{d:%Y-%m}", {"entrata": {}, "uscita": {}})[t]
        k[c] = round(k.get(c, 0) + i, 2)
    budget = {b.categoria: b.importo for b in Budget.query.filter_by(user_id=uid)}
    return jsonify(mesi=mesi, budget=budget)


@app.route("/nuovo", methods=["GET", "POST"])
@login_required
def nuovo():
    if request.method == "POST":
        f = request.form
        try:
            imp = abs(num(f["importo"]))
            d = datetime.strptime(f["data"], "%Y-%m-%d").date()
            cat = f["categoria"].strip()
            if not cat:
                raise ValueError
        except (ValueError, KeyError):
            flash("Controlla importo, data e categoria.")
            return redirect(url_for("nuovo"))
        db.session.add(Movimento(user_id=session["uid"], data=d, tipo="entrata" if f.get("tipo") == "entrata" else "uscita",
                                 categoria=cat, importo=imp, descrizione=f.get("descrizione", "").strip()))
        db.session.commit()
        flash("Movimento salvato.")
        return redirect(url_for("nuovo"))
    return render_template("nuovo.html", oggi=date.today().isoformat())


@app.route("/movimenti")
@login_required
def movimenti():
    mese = request.args.get("mese", "")
    try:
        y, m = map(int, mese.split("-"))
        a = date(y, m, 1)
    except ValueError:
        a = date.today().replace(day=1)
    b = date(a.year + (a.month == 12), a.month % 12 + 1, 1)
    righe = (Movimento.query.filter(Movimento.user_id == session["uid"], Movimento.data >= a, Movimento.data < b)
             .order_by(Movimento.data.desc(), Movimento.id.desc()).all())
    return render_template("movimenti.html", righe=righe, mese=f"{a:%Y-%m}")


@app.route("/elimina/<int:id>", methods=["POST"])
@login_required
def elimina(id):
    r = db.get_or_404(Movimento, id)
    if r.user_id != session["uid"]:
        abort(404)
    db.session.delete(r)
    db.session.commit()
    return redirect(request.referrer or url_for("movimenti"))


@app.route("/impostazioni", methods=["GET", "POST"])
@login_required
def impostazioni():
    uid = session["uid"]
    if request.method == "POST":
        f = request.form
        az = f.get("azione")
        try:
            if az == "budget":
                b = Budget.query.filter_by(user_id=uid, categoria=f["categoria"]).first() or Budget(user_id=uid, categoria=f["categoria"])
                b.importo = num(f["importo"])
                db.session.add(b)
            elif az == "ricorrente":
                t = date.today()
                y, m = (t.year, t.month - 1) if t.month > 1 else (t.year - 1, 12)
                db.session.add(Ricorrente(user_id=uid, tipo="entrata" if f["tipo"] == "entrata" else "uscita",
                                          categoria=f["categoria"].strip(), descrizione=f.get("descrizione", "").strip(),
                                          importo=abs(num(f["importo"])), giorno=min(max(int(f["giorno"]), 1), 28),
                                          ultimo=f"{y}-{m:02d}"))
            elif az == "del_budget":
                Budget.query.filter_by(id=int(f["id"]), user_id=uid).delete()
            elif az == "del_ric":
                Ricorrente.query.filter_by(id=int(f["id"]), user_id=uid).delete()
            db.session.commit()
        except (ValueError, KeyError):
            db.session.rollback()
            flash("Dati non validi.")
        return redirect(url_for("impostazioni"))
    return render_template("impostazioni.html", budget=Budget.query.filter_by(user_id=uid).all(),
                           ric=Ricorrente.query.filter_by(user_id=uid).all())


@app.route("/password", methods=["POST"])
@login_required
def cambia_password():
    u = db.session.get(User, session["uid"])
    f = request.form
    if not check_password_hash(u.password, f.get("vecchia", "")):
        flash("La password attuale non e' corretta.")
    elif len(f.get("nuova", "")) < 10 or f["nuova"] != f.get("conferma"):
        flash("La nuova password deve avere almeno 10 caratteri ed essere uguale nei due campi.")
    else:
        u.password = generate_password_hash(f["nuova"])
        db.session.commit()
        flash("Password aggiornata.")
    return redirect(url_for("impostazioni"))


@app.route("/importa", methods=["POST"])
@login_required
def importa():
    uid = session["uid"]
    n, anni, errori = 0, set(), []
    for file in request.files.getlist("file"):
        try:
            wb = load_workbook(file.stream, data_only=True)
        except Exception:
            errori.append(file.filename)
            continue
        nuovi = []
        for ws in wb:
            if ws.title.strip().upper() == "ANNUALE":
                continue
            for c in range(3, ws.max_column + 1):
                d = ws.cell(2, c).value
                if not isinstance(d, datetime):
                    continue
                for r in range(3, 33):
                    v, cat = ws.cell(r, c).value, ws.cell(r, 1).value
                    if cat and isinstance(v, (int, float)) and v:
                        cat = str(cat).strip()
                        nuovi.append(Movimento(user_id=uid, data=d.date(), tipo="entrata" if cat in ENTRATE else "uscita",
                                               categoria=cat, importo=abs(v), fonte="import"))
        if not nuovi:
            errori.append(file.filename)
            continue
        ys = {m.data.year for m in nuovi}
        for y in ys:  # sostituisce solo gli anni presenti nel file
            Movimento.query.filter(Movimento.user_id == uid, Movimento.fonte == "import",
                                   Movimento.data >= date(y, 1, 1), Movimento.data < date(y + 1, 1, 1)).delete()
        db.session.add_all(nuovi)
        n += len(nuovi)
        anni |= ys
    db.session.commit()
    if anni:
        flash(f"Importati {n} movimenti per gli anni: {', '.join(map(str, sorted(anni)))}.")
    if errori:
        flash("File non letti: " + ", ".join(errori))
    return redirect(url_for("impostazioni"))


# ---------- ripetizioni ----------
MESI_IT = {m: i for i, m in enumerate("gennaio febbraio marzo aprile maggio giugno luglio agosto settembre ottobre novembre dicembre".split(), 1)}
PAG = re.compile(r"pagat|bonifico|contanti|paypal|ricevut|incass|pagamento|regalat", re.I)


def mio(model, id):
    o = db.get_or_404(model, id)
    if o.user_id != session["uid"]:
        abort(404)
    return o


def parse_data(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, str):
        m = re.search(r"(\d{1,2})\s+([a-z]+)\s+(\d{4})", v.lower())
        if m and m.group(2) in MESI_IT:
            return date(int(m.group(3)), MESI_IT[m.group(2)], int(m.group(1)))
    return None


def rosso(c):
    """Il testo della cella e' rosso? (rosso = lezione da pagare, verde = pagata)"""
    col = c.font.color
    if col is None or col.type != "rgb" or not isinstance(col.rgb, str):
        return False
    h = col.rgb[-6:]
    return int(h[0:2], 16) > 180 and int(h[2:4], 16) < 90 and int(h[4:6], 16) < 90


def leggi_credito(ws):
    """Credito prepagato scritto dal sito nel foglio (colonna L). None se il foglio non lo contiene."""
    v = ws.cell(2, 12).value
    if str(ws.cell(1, 12).value or "").strip().lower() == "credito eur" and isinstance(v, (int, float)) and v >= 0:
        return float(v)
    return None


def leggi_foglio(ws):
    tariffa = 20.0
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if isinstance(a, str) and a.strip().lower().startswith("importi"):
            m = re.search(r"\d+(?:[.,]\d+)?", str(b or ""))
            if m:
                tariffa = float(m.group().replace(",", "."))
            break
    out = []
    usa_j = str(ws.cell(1, 10).value or "").strip().lower().startswith("tariffa")

    def riga_tar(r):
        v = ws.cell(r, 10).value if usa_j else None
        return float(v) if isinstance(v, (int, float)) and v >= 0 else tariffa
    for r in range(2, ws.max_row + 1):
        d, ore = parse_data(ws.cell(r, 1).value), ws.cell(r, 4).value
        if not d or not isinstance(ore, (int, float)) or ore <= 0:
            continue
        note = [str(ws.cell(r, c).value) for c in range(5, 10) if isinstance(ws.cell(r, c).value, str)]
        testo = " ".join(note).lower()
        e = ws.cell(r, 5).value
        modalita = "Presenza" if "presenza" in testo else "Online" if re.search(r"dad|skype|meet", testo) else ""
        hh = lambda v: v.strftime("%H:%M") if isinstance(v, (time, datetime)) else ""
        out.append(dict(data=d, ora_inizio=hh(ws.cell(r, 2).value), ora_fine=hh(ws.cell(r, 3).value), ore=float(ore),
                        materia=(str(e).strip()[:80] if isinstance(e, str) else ""), modalita=modalita,
                        pagata=not any(rosso(ws.cell(r, c)) for c in range(1, 6)), tariffa=riga_tar(r)))
    nota = ""
    for r in (2, 3):
        for c in range(6, 10):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() and not PAG.search(v) and not re.fullmatch(r"(in )?presenza|dad|skype|google meet|meet", v.strip(), re.I):
                nota = nota or v.strip()[:300]
    k2 = ws.cell(2, 11).value
    if str(ws.cell(1, 11).value or "").strip().lower() == "note" and isinstance(k2, str) and k2.strip():
        nota = k2.strip()[:500]
    return tariffa, nota, out


def txt_pre(k):
    return "nessuna lezione prepagata" if k == 0 else "1 lezione già pagata da svolgere" if k == 1 else f"{k} lezioni già pagate da svolgere"


def n_prepagate(s):
    """Quante lezioni intere copre il credito dell'allievo."""
    t = s.tariffa or 0
    return int((s.credito or 0) / t + 1e-9) if t > 0 else 0


@app.route("/ripetizioni")
@login_required
def ripetizioni():
    uid = session["uid"]
    al = {s.id: dict(s=s, n=0, dovuto=0.0, incassato=0.0, prima=None, ultima=None) for s in Studente.query.filter_by(user_id=uid)}
    inc_mesi = {}
    q = db.session.query(Lezione.studente_id, Lezione.ore, Lezione.tariffa, Lezione.pagata, Lezione.data).filter(Lezione.user_id == uid)
    for sid, ore, tar, pag, d in q:
        a = al[sid]
        euro = ore * (tar or a["s"].tariffa or 0)
        if pag:
            a["incassato"] += euro
            k = d.strftime("%Y-%m")
            inc_mesi[k] = round(inc_mesi.get(k, 0) + euro, 2)
        else:
            a["n"] += 1
            a["dovuto"] += euro
        if not a["ultima"] or d > a["ultima"]:
            a["ultima"] = d
        if not a["prima"] or d < a["prima"]:
            a["prima"] = d
    # prima chi ha lezioni da pagare; poi i piu' recenti (prima lezione piu' vicina a oggi) in alto, i piu' storici in basso
    for a in al.values():
        a["pre"] = n_prepagate(a["s"])
    elenco = sorted(al.values(), key=lambda a: (a["n"] == 0, -(a["prima"] or date.max).toordinal(), -a["s"].id))
    anno = date.today() - timedelta(days=365)
    tutti = request.args.get("tutti") == "1"
    visibili = [a for a in elenco if tutti or a["n"] or (a["ultima"] and a["ultima"] >= anno)]
    return render_template("ripetizioni.html", elenco=elenco, visibili=visibili, tutti=tutti, oggi=date.today().isoformat(),
                           inc_mesi=inc_mesi, tot=sum(a["dovuto"] for a in elenco), tot_n=sum(a["n"] for a in elenco),
                           n_deb=sum(1 for a in elenco if a["n"]), p=ore_periodi(uid), ultimo_file=ultimo_file(uid))


@app.route("/ripetizioni/<int:id>")
@login_required
def allievo(id):
    s = mio(Studente, id)
    lez = Lezione.query.filter_by(studente_id=id).order_by(Lezione.data.desc(), Lezione.ora_inizio.desc()).all()
    n = sum(1 for l in lez if not l.pagata)
    return render_template("allievo.html", s=s, lez=lez, n=n, n_tot=len(lez), ore_tot=sum(l.ore for l in lez),
                           dovuto=sum(l.ore * (l.tariffa or s.tariffa or 0) for l in lez if not l.pagata),
                           incassato=sum(l.ore * (l.tariffa or s.tariffa or 0) for l in lez if l.pagata),
                           serie=[[l.data.isoformat(), l.ore] for l in lez],
                           credito=s.credito or 0, n_pre=n_prepagate(s))


@app.route("/ripetizioni/nuova", methods=["POST"])
@login_required
def rip_nuova():
    uid, f = session["uid"], request.form
    try:
        d = datetime.strptime(f["data"], "%Y-%m-%d").date()
        ore = round((datetime.strptime(f["fine"], "%H:%M") - datetime.strptime(f["inizio"], "%H:%M")).total_seconds() / 3600, 2)
        nome = f["studente"].strip()
        if ore <= 0 or not nome:
            raise ValueError
    except (ValueError, KeyError):
        flash("Controlla allievo, data e orari.")
        return redirect(url_for("ripetizioni"))
    s = Studente.query.filter_by(user_id=uid, nome=nome).first()
    if not s:
        try:
            t = num(f.get("tariffa") or 20)
        except ValueError:
            t = 20
        s = Studente(user_id=uid, nome=nome, tariffa=t)
        db.session.add(s)
        db.session.flush()
    l = Lezione(user_id=uid, studente_id=s.id, data=d, ora_inizio=f["inizio"], ora_fine=f["fine"], ore=ore,
                materia=f.get("materia", "").strip(), modalita=f.get("modalita", ""), tariffa=s.tariffa)
    costo = ore * (s.tariffa or 0)
    arretrati = Lezione.query.filter_by(studente_id=s.id, pagata=False).count()
    prepagata = arretrati == 0 and costo > 0 and (s.credito or 0) >= costo - 0.001
    if prepagata:  # l'allievo aveva gia' pagato in anticipo: la lezione nasce pagata e scala il credito
        l.pagata, l.data_pagamento = True, date.today()
        s.credito = round(max((s.credito or 0) - costo, 0), 2)
    db.session.add(l)
    db.session.commit()
    flash(f"Lezione aggiunta e già pagata con il credito. Restano: {txt_pre(n_prepagate(s))}." if prepagata else "Lezione aggiunta.")
    return redirect(url_for("allievo", id=s.id))


@app.route("/ripetizioni/<int:id>/dati", methods=["POST"])
@login_required
def rip_dati(id):
    s, f = mio(Studente, id), request.form
    try:
        nome, tar = f["nome"].strip(), abs(num(f["tariffa"]))
        if not nome:
            raise ValueError
    except (ValueError, KeyError):
        flash("Nome o tariffa non validi.")
        return redirect(url_for("allievo", id=id))
    s.nome, s.tariffa, s.note = nome, tar, f.get("note", "").strip()[:500]
    if "credito" in f:
        try:
            s.credito = abs(num(f["credito"] or 0))
        except ValueError:
            flash("Credito non valido: gli altri dati sono stati salvati.")
    if f.get("applica"):
        Lezione.query.filter_by(studente_id=id, pagata=False).update({"tariffa": tar})
    db.session.commit()
    flash("Dati salvati.")
    return redirect(url_for("allievo", id=id))


@app.route("/ripetizioni/<int:id>/elimina", methods=["POST"])
@login_required
def rip_elimina(id):
    s = mio(Studente, id)
    Lezione.query.filter_by(studente_id=id).delete()
    db.session.delete(s)
    db.session.commit()
    flash("Allievo eliminato.")
    return redirect(url_for("ripetizioni"))


@app.route("/ripetizioni/<int:id>/incassa", methods=["POST"])
@login_required
def rip_incassa(id):
    s, f = mio(Studente, id), request.form
    q = Lezione.query.filter_by(studente_id=id, pagata=False).order_by(Lezione.data, Lezione.id)
    n = 0
    try:
        residuo = None if f.get("tutte") else abs(num(f["importo"])) + (s.credito or 0)  # importo + credito gia' presente
    except (ValueError, KeyError):
        flash("Importo non valido.")
        return redirect(url_for("allievo", id=id))
    for l in q:
        costo = l.ore * (l.tariffa or s.tariffa or 0)
        if residuo is not None:
            if costo > residuo + 0.001:
                break
            residuo -= costo
        l.pagata, l.data_pagamento = True, date.today()
        n += 1
    msg = f"{n} lezioni segnate come pagate."
    if residuo is not None:
        s.credito = round(max(residuo, 0), 2)  # quello che avanza e' un anticipo per le prossime lezioni
        if s.credito:
            k = n_prepagate(s)
            msg += f" Avanzano {s.credito:.0f} €: {txt_pre(k)}." if k else f" Avanzano {s.credito:.0f} € di credito."
    db.session.commit()
    flash(msg)
    return redirect(url_for("allievo", id=id))


@app.route("/lezione/<int:id>/<azione>", methods=["POST"])
@login_required
def rip_lezione(id, azione):
    l = mio(Lezione, id)
    sid = l.studente_id
    if azione == "paga":
        l.pagata = not l.pagata
        l.data_pagamento = date.today() if l.pagata else None
    elif azione == "elimina":
        db.session.delete(l)
    db.session.commit()
    return redirect(url_for("allievo", id=sid))


@app.route("/ripetizioni/importa", methods=["POST"])
@login_required
def rip_importa():
    uid = session["uid"]
    f = request.files.get("file")
    dati = f.read() if f else b""
    try:
        wb = load_workbook(io.BytesIO(dati), data_only=True)
        fogli = [(ws.title.strip(),) + leggi_foglio(ws) for ws in wb]  # (nome, tariffa, nota, lezioni)
        crediti = {ws.title.strip(): leggi_credito(ws) for ws in wb}
    except Exception:
        flash("Il file non e' un Excel (.xlsx) valido: non ho cambiato nulla.")
        return redirect(url_for("ripetizioni"))
    fogli = [x for x in fogli if x[3]]
    if not fogli:
        flash("Nel file non ho trovato lezioni: non ho cambiato nulla.")
        return redirect(url_for("ripetizioni"))
    try:
        completo = (wb.properties.title or "") == "ripetizioni-export"  # file scaricato dal sito: contiene tutto
        q = Lezione.query.filter_by(user_id=uid)
        (q if completo else q.filter_by(fonte="import")).delete()  # il nuovo file sostituisce il precedente
        n = nuovi = 0
        for nome, tariffa, nota, lez in fogli:
            s = Studente.query.filter_by(user_id=uid, nome=nome).first()
            if not s:
                s = Studente(user_id=uid, nome=nome, tariffa=tariffa, note=nota)
                db.session.add(s)
                db.session.flush()
                nuovi += 1
            else:
                s.tariffa = tariffa
                s.note = s.note or nota
            if completo and crediti.get(nome) is not None:  # file scaricato dal sito: ripristina anche il credito
                s.credito = crediti[nome]
            db.session.add_all(Lezione(user_id=uid, studente_id=s.id, fonte="import", **l) for l in lez)
            n += len(lez)
        db.session.flush()
        orfani = [i for (i,) in db.session.query(Studente.id).outerjoin(Lezione, Lezione.studente_id == Studente.id)
                  .filter(Studente.user_id == uid, Lezione.id.is_(None), db.func.coalesce(Studente.credito, 0) <= 0)]  # allievi non piu' presenti nel file
        if orfani:
            Studente.query.filter(Studente.id.in_(orfani)).delete(synchronize_session=False)
        rec = FileCaricato.query.filter_by(user_id=uid, tipo="ripetizioni").first() or FileCaricato(user_id=uid, tipo="ripetizioni")
        rec.nome, rec.contenuto = (f.filename or "Ripetizioni.xlsx")[:200], dati
        rec.caricato_il = datetime.now(timezone.utc).replace(tzinfo=None)
        db.session.add(rec)
        db.session.commit()
    except Exception:
        db.session.rollback()
        flash("Errore durante l'importazione: i dati precedenti non sono stati modificati.")
        return redirect(url_for("ripetizioni"))
    flash(f"File salvato e dati aggiornati: {n} lezioni, {len(fogli)} allievi ({nuovi} nuovi, {len(orfani)} rimossi)." + (" Era un file scaricato dal sito: sostituiti anche i dati inseriti a mano." if completo else ""))
    return redirect(url_for("ripetizioni"))


@app.route("/modifica-lezione/<int:id>", methods=["GET", "POST"])
@login_required
def rip_modifica(id):
    l = mio(Lezione, id)
    if request.method == "POST":
        f = request.form
        try:
            l.data = datetime.strptime(f["data"], "%Y-%m-%d").date()
            l.ore, l.tariffa = num(f["ore"]), abs(num(f["tariffa"]))
            if l.ore <= 0:
                raise ValueError
            l.ora_inizio, l.ora_fine = f.get("inizio", ""), f.get("fine", "")
            l.materia, l.modalita = f.get("materia", "").strip(), f.get("modalita", "")
            nuova = f.get("pagata") == "1"
            if nuova != l.pagata:
                l.pagata, l.data_pagamento = nuova, (date.today() if nuova else None)
            db.session.commit()
            return redirect(url_for("allievo", id=l.studente_id))
        except (ValueError, KeyError):
            db.session.rollback()
            flash("Controlla i dati inseriti.")
            return redirect(url_for("rip_modifica", id=id))
    return render_template("lezione.html", l=l)


@app.route("/ripetizioni/confronti")
@login_required
def rip_confronti():
    dati = {}  # "AAAA-MM" -> [lezioni, ore, incassato]
    q = (db.session.query(Lezione.data, Lezione.ore, Lezione.tariffa, Lezione.pagata, Studente.tariffa)
         .join(Studente, Lezione.studente_id == Studente.id).filter(Lezione.user_id == session["uid"]))
    for d, ore, tar, pag, st in q:
        r = dati.setdefault(d.strftime("%Y-%m"), [0, 0.0, 0.0])
        r[0] += 1
        r[1] = round(r[1] + ore, 2)
        if pag:
            r[2] = round(r[2] + ore * (tar or st or 0), 2)
    return render_template("confronti.html", dati=dati)


MESI_LUNGHI = list(MESI_IT)


@app.template_filter("ore")
def filtro_ore(x):
    return f"{x:.2f}".rstrip("0").rstrip(".")


def ore_tra(uid, a, b):
    """Ore di lezione con data da a (incluso) a b (escluso)."""
    return db.session.query(db.func.coalesce(db.func.sum(Lezione.ore), 0)).filter(
        Lezione.user_id == uid, Lezione.data >= a, Lezione.data < b).scalar() or 0


def ore_periodi(uid):
    oggi, g = date.today(), timedelta(days=1)
    m0 = oggi.replace(day=1)
    pm0 = (m0 - g).replace(day=1)
    m1 = (m0 + timedelta(days=32)).replace(day=1)
    s0 = oggi - timedelta(days=oggi.weekday())
    ps0 = s0 - timedelta(days=7)
    fine_ps = s0 - g
    return dict(
        mese=ore_tra(uid, m0, oggi + g), mese_prog=ore_tra(uid, oggi + g, m1),
        mese_prec=ore_tra(uid, pm0, m0), mese_prec_pari=ore_tra(uid, pm0, min(m0, pm0 + timedelta(days=oggi.day))),
        sett=ore_tra(uid, s0, oggi + g), sett_prog=ore_tra(uid, oggi + g, s0 + timedelta(days=7)),
        sett_prec=ore_tra(uid, ps0, s0), sett_prec_pari=ore_tra(uid, ps0, ps0 + timedelta(days=oggi.weekday() + 1)),
        nome_mese_prec=f"{MESI_LUNGHI[pm0.month - 1]} {pm0.year}",
        periodo_sett_prec=f"dal {ps0.day}/{ps0.month} al {fine_ps.day}/{fine_ps.month}",
        periodo_sett=f"dal {s0.day}/{s0.month} a oggi")


@app.route("/ripetizioni/calendario")
@login_required
def rip_calendario():
    try:
        rif = datetime.strptime(request.args.get("data", ""), "%Y-%m-%d").date()
    except ValueError:
        rif = date.today()
    lun = rif - timedelta(days=rif.weekday())
    dom = lun + timedelta(days=6)
    q = (db.session.query(Lezione, Studente).join(Studente, Lezione.studente_id == Studente.id)
         .filter(Lezione.user_id == session["uid"], Lezione.data >= lun, Lezione.data <= dom)
         .order_by(Lezione.data, Lezione.ora_inizio))
    giorni = [dict(d=lun + timedelta(days=i), ev=[], senza=[]) for i in range(7)]
    n, ore_tot = 0, 0.0
    for l, s in q:
        e = dict(sid=s.id, nome=s.nome, ini=l.ora_inizio, fine=l.ora_fine, ore=l.ore, materia=l.materia, mod=l.modalita, pagata=l.pagata)
        g = giorni[(l.data - lun).days]
        try:
            h, m = map(int, l.ora_inizio.split(":"))
            e["start"] = h * 60 + m
            e["end"] = e["start"] + int(l.ore * 60)
            if l.ora_fine:
                h2, m2 = map(int, l.ora_fine.split(":"))
                if h2 * 60 + m2 > e["start"]:
                    e["end"] = h2 * 60 + m2
            g["ev"].append(e)
        except ValueError:
            g["senza"].append(e)
        n += 1
        ore_tot += l.ore
    inizi = [e["start"] // 60 for g in giorni for e in g["ev"]]
    fini = [-(-e["end"] // 60) for g in giorni for e in g["ev"]]
    h0, h1 = min([8] + inizi), min(max([21] + fini), 24)
    titolo = (f"{lun.day}–{dom.day} {MESI_LUNGHI[dom.month - 1]} {dom.year}" if lun.month == dom.month
              else f"{lun.day} {MESI_LUNGHI[lun.month - 1]} – {dom.day} {MESI_LUNGHI[dom.month - 1]} {dom.year}")
    return render_template("calendario.html", giorni=giorni, h0=h0, h1=h1, titolo=titolo, n=n, ore=ore_tot, oggi=date.today(),
                           prec=(lun - timedelta(days=7)).isoformat(), succ=(lun + timedelta(days=7)).isoformat(),
                           nomi=["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"], mesi_brevi=[m[:3] for m in MESI_LUNGHI])


def ultimo_file(uid):
    r = (db.session.query(FileCaricato.nome, FileCaricato.caricato_il)
         .filter(FileCaricato.user_id == uid, FileCaricato.tipo == "ripetizioni").first())
    return dict(nome=r.nome, iso=r.caricato_il.isoformat() + "Z") if r else None


@app.route("/ripetizioni/file")
@login_required
def rip_file():
    r = FileCaricato.query.filter_by(user_id=session["uid"], tipo="ripetizioni").first_or_404()
    return send_file(io.BytesIO(r.contenuto), as_attachment=True, download_name=r.nome)


@app.route("/ripetizioni/esporta-fogli")
@login_required
def rip_esporta_fogli():
    """Excel nel formato originale: un foglio per allievo, testo rosso/verde, tariffa in fondo. Si puo' ricaricare."""
    def ora(v):
        try:
            h, m = map(int, v.split(":"))
            return time(h, m)
        except (ValueError, AttributeError):
            return None
    wb = Workbook()
    wb.remove(wb.active)
    rosso, verde, grassetto = Font(color="FF0000"), Font(color="00B050"), Font(bold=True)
    usati = set()
    for s in Studente.query.filter_by(user_id=session["uid"]).order_by(Studente.id):
        lez = Lezione.query.filter_by(studente_id=s.id).order_by(Lezione.data, Lezione.ora_inizio, Lezione.id).all()
        if not lez:
            continue
        base = re.sub(r"[\[\]:*?/\\]", "", s.nome).strip()[:31] or "Allievo"
        nome, i = base, 2
        while nome.lower() in usati:
            suff = f" {i}"
            nome, i = base[:31 - len(suff)] + suff, i + 1
        usati.add(nome.lower())
        ws = wb.create_sheet(nome)
        ws.append(["Data", "Ora inizio", "Ora fine", "Ore totali", "Materia", "Presenza/DAD", None, None, None, "Tariffa EUR/h", "Note", "Credito EUR"])
        for c in ws[1]:
            c.font = grassetto
        for l in lez:
            ws.append([l.data, ora(l.ora_inizio), ora(l.ora_fine), l.ore, l.materia or None,
                       {"Presenza": "Presenza", "Online": "DAD"}.get(l.modalita), None, None, None, l.tariffa or s.tariffa])
            r = ws.max_row
            for c in range(1, 7):
                ws.cell(r, c).font = verde if l.pagata else rosso
            ws.cell(r, 1).number_format = "DD/MM/YYYY"
            ws.cell(r, 2).number_format = ws.cell(r, 3).number_format = "HH:MM"
        ws["K2"] = s.note or None
        ws["L2"] = round(s.credito or 0, 2)
        ws.append([])
        ws.append(["Importi/ora", s.tariffa])
        for col, w in zip("ABCDEFJKL", (14, 11, 10, 11, 22, 14, 13, 40, 13)):
            ws.column_dimensions[col].width = w
    wb.properties.title = "ripetizioni-export"  # riconosce i file scaricati da qui
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"Ripetizioni_aggiornato_{date.today()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
