import io, json, os, re, secrets
from datetime import date, datetime, time, timedelta, timezone
from functools import wraps
import click
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, abort, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
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
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=14)  # la sessione scade dopo 14 giorni
_sk = os.environ.get("SECRET_KEY")
if not _sk:
    if "DATABASE_URL" in os.environ:  # online: senza chiave segreta le sessioni sarebbero falsificabili
        raise RuntimeError("Imposta la variabile d'ambiente SECRET_KEY (una stringa casuale lunga).")
    _sk = "solo-sviluppo-locale"
app.secret_key = _sk
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
    file_id = db.Column(db.Integer)  # file Excel (FileBilancio) da cui arriva, se importato


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


class FileBilancio(db.Model):
    """Elenco dei file Excel del bilancio importati (solo nome e data: i movimenti stanno in Movimento)."""
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    nome = db.Column(db.String(200), nullable=False)
    caricato_il = db.Column(db.DateTime, nullable=False)  # UTC
    contenuto = db.Column(db.LargeBinary)  # il file originale, cosi' si puo' riscaricare (vuoto per i file caricati prima)


class FileCaricato(db.Model):
    """Ultimo file Excel caricato (uno per tipo): resta salvato nel database online."""
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    tipo = db.Column(db.String(20), nullable=False)
    nome = db.Column(db.String(200), nullable=False)
    caricato_il = db.Column(db.DateTime, nullable=False)  # UTC
    contenuto = db.Column(db.LargeBinary, nullable=False)


class Stipendio(db.Model):
    """Una riga di Check_stipendi: i dati di un cedolino (mensile o tredicesima)."""
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    anno = db.Column(db.Integer, nullable=False)
    mese = db.Column(db.Integer, nullable=False)
    tipo = db.Column(db.String(10), nullable=False, default="mensile")  # mensile | tredicesima
    fonte = db.Column(db.String(10), default="sito")  # import | sito
    note = db.Column(db.String(1000))
    pdf = db.Column(db.String(200))


STIP_CAMPI = [  # (colonna Excel, attributo, titolo)
    ("B", "giorni", "Giorni lavorati"), ("C", "gg_mensa", "Giorni (trattenuta mensa)"), ("D", "scatti", "Scatti"),
    ("E", "unitario", "Compenso unitario"), ("F", "minimo", "Minimo"), ("G", "sup_ass", "Sup. ass."),
    ("H", "tot_retr", "Tot. retributivo"), ("I", "competenze", "Tot. competenze"), ("J", "trattenute", "Trattenute"),
    ("L", "fe_prec", "Ferie: anno prec."), ("M", "fe_spett", "Ferie: spettanti"), ("N", "fe_god", "Ferie: godute"),
    ("P", "par_prec", "PAR: anno prec."), ("Q", "par_spett", "PAR: spettanti"), ("R", "par_god", "PAR: godute"),
    ("U", "edp", "EDP"), ("V", "anticipo", "Anticipo for the better"), ("W", "mensa_tratt", "Trattenuta mensa"),
    ("X", "mensa_sup", "Supplemento mensa"), ("Z", "cometa", "Cometa"), ("AA", "tfr_fondo", "TFR fondo"), ("AB", "tfe", "TFE 12 AP")]
for _c, _a, _t in STIP_CAMPI:
    setattr(Stipendio, _a, db.Column(db.Float))


class CedolinoPdf(db.Model):
    """Il PDF di un cedolino, con i valori letti dal PDF (JSON)."""
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    anno = db.Column(db.Integer, nullable=False)
    mese = db.Column(db.Integer, nullable=False)
    tipo = db.Column(db.String(10), nullable=False, default="mensile")
    nome = db.Column(db.String(200), nullable=False)
    caricato_il = db.Column(db.DateTime, nullable=False)  # UTC
    contenuto = db.Column(db.LargeBinary, nullable=False)
    dati = db.Column(db.Text)  # {"v": valori, "voci": [...], "c": controlli}


def migra():
    """Aggiunge le colonne nuove alle tabelle gia' esistenti (database online)."""
    from sqlalchemy import inspect, text
    insp = inspect(db.engine)
    for tab, col, tipo in (("studente", "note", "VARCHAR(500)"), ("lezione", "tariffa", "FLOAT"), ("studente", "credito", "FLOAT"), ("movimento", "file_id", "INTEGER")):
        if col not in [c["name"] for c in insp.get_columns(tab)]:
            db.session.execute(text(f"ALTER TABLE {tab} ADD COLUMN {col} {tipo}"))
    if "contenuto" not in [c["name"] for c in insp.get_columns("file_bilancio")]:
        db.session.execute(text("ALTER TABLE file_bilancio ADD COLUMN contenuto " + ("BYTEA" if db.engine.dialect.name == "postgresql" else "BLOB")))
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
    if request.method == "POST":
        atteso = session.get("_t")
        if not atteso or not secrets.compare_digest(request.form.get("_t", ""), atteso):
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


def nota_cella(c):
    """Testo della nota (commento) attaccata a una cella Excel, ripulito; stringa vuota se non c'e'."""
    if not c.comment:
        return ""
    t = (c.comment.text or "").strip()
    m = re.search(r"Comment:\s*\n(.*)", t, re.S)  # commenti "a catena" dell'Excel moderno
    if m:
        t = m.group(1).strip()
    else:  # nota classica: la prima riga e' spesso solo il nome dell'autore ("Marco:")
        righe = t.split("\n")
        if len(righe) > 1 and righe[0].rstrip().endswith(":") and len(righe[0]) <= 40:
            t = "\n".join(righe[1:]).strip()
    return " ".join(t.split())[:200]  # il campo descrizione accetta al massimo 200 caratteri


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
TENTATIVI = {}  # ip -> [istanti dei tentativi falliti]; in memoria, si azzera al riavvio
MAX_TENTATIVI, FINESTRA = 5, 15 * 60  # 5 errori ogni 15 minuti


def tentativi_recenti(ip):
    ora = datetime.now().timestamp()
    lista = [t for t in TENTATIVI.get(ip, []) if ora - t < FINESTRA]
    if lista:
        TENTATIVI[ip] = lista
    else:
        TENTATIVI.pop(ip, None)
    return lista


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        ip = request.remote_addr or "?"
        if len(tentativi_recenti(ip)) >= MAX_TENTATIVI:
            flash("Troppi tentativi falliti. Riprova tra 15 minuti.")
            return render_template("login.html"), 429
        u = User.query.filter_by(username=request.form.get("username", "").strip()).first()
        if u and check_password_hash(u.password, request.form.get("password", "")):
            TENTATIVI.pop(ip, None)
            session.clear()
            session.permanent = True
            session["uid"] = u.id
            return redirect(url_for("home"))
        TENTATIVI.setdefault(ip, []).append(datetime.now().timestamp())
        flash("Username o password errati.")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def home():
    return render_template("home.html")


@app.route("/dashboard")
@login_required
def dashboard():
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
    mese = f"{r.data:%Y-%m}"
    db.session.delete(r)
    db.session.commit()
    return redirect(url_for("movimenti", mese=mese))


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
    file_bilanci, anni_senza_file = elenco_bilanci(uid)
    return render_template("impostazioni.html", budget=Budget.query.filter_by(user_id=uid).all(),
                           ric=Ricorrente.query.filter_by(user_id=uid).all(),
                           file_bilanci=file_bilanci, anni_senza_file=anni_senza_file,
                           anni_export=sorted({int(y) for (y,) in db.session.query(db.extract("year", Movimento.data)).filter(Movimento.user_id == uid).distinct()}))


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
        raw = file.read()
        try:
            wb = load_workbook(io.BytesIO(raw), data_only=True)
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
                                               categoria=cat, importo=abs(v), fonte="import",
                                               descrizione=nota_cella(ws.cell(r, c))))
        if not nuovi:
            errori.append(file.filename)
            continue
        ys = {m.data.year for m in nuovi}
        da_sito = (wb.properties.title or "") == "spese-export"  # scaricato da qui: contiene gia' anche i movimenti inseriti dal sito
        for y in ys:  # sostituisce solo gli anni presenti nel file
            q = Movimento.query.filter(Movimento.user_id == uid, Movimento.data >= date(y, 1, 1), Movimento.data < date(y + 1, 1, 1))
            (q if da_sito else q.filter(Movimento.fonte == "import")).delete()
        rec = FileBilancio(user_id=uid, nome=(file.filename or "Bilancio.xlsx")[:200],
                           caricato_il=datetime.now(timezone.utc).replace(tzinfo=None), contenuto=raw)
        db.session.add(rec)
        db.session.flush()
        for m in nuovi:
            m.file_id = rec.id
        db.session.add_all(nuovi)
        n += len(nuovi)
        anni |= ys
    db.session.flush()
    vuoti = [i for (i,) in db.session.query(FileBilancio.id).outerjoin(Movimento, Movimento.file_id == FileBilancio.id)
             .filter(FileBilancio.user_id == uid, Movimento.id.is_(None))]  # file i cui anni sono stati sostituiti da uno nuovo
    if vuoti:
        FileBilancio.query.filter(FileBilancio.id.in_(vuoti)).delete(synchronize_session=False)
    db.session.commit()
    if anni:
        flash(f"Importati {n} movimenti per gli anni: {', '.join(map(str, sorted(anni)))}.")
    if errori:
        flash("File non letti: " + ", ".join(errori))
    return redirect(url_for("impostazioni"))


@app.route("/importa/scarica/<int:id>")
@login_required
def importa_scarica(id):
    """Il file Excel esattamente come era stato caricato."""
    f = mio(FileBilancio, id)
    if not f.contenuto:
        abort(404)
    return send_file(io.BytesIO(f.contenuto), as_attachment=True, download_name=f.nome)


@app.route("/esporta/<int:anno>")
@login_required
def esporta_anno(anno):
    """Excel dell'anno con tutti i movimenti attuali (importati e inseriti dal sito), nello stesso formato dell'importazione:
    un foglio per mese, i giorni in colonna, le categorie in riga, le note come commenti della cella. Si puo' ricaricare."""
    if not 1900 <= anno <= 2200:
        abort(404)
    mov = (Movimento.query.filter(Movimento.user_id == session["uid"], Movimento.data >= date(anno, 1, 1), Movimento.data < date(anno + 1, 1, 1))
           .order_by(Movimento.data, Movimento.id).all())
    if not mov:
        flash(f"Non ci sono movimenti nel {anno} da esportare.")
        return redirect(url_for("impostazioni"))
    mesi = "Gennaio Febbraio Marzo Aprile Maggio Giugno Luglio Agosto Settembre Ottobre Novembre Dicembre".split()
    grassetto = Font(bold=True)
    ordine = {c: i for i, c in enumerate(ENTRATE + USCITE)}
    wb = Workbook()
    ann = wb.active
    ann.title = "ANNUALE"
    cats = sorted({m.categoria for m in mov}, key=lambda c: (c not in ENTRATE, ordine.get(c, 999), c))
    ann.append(["Categoria", "Totale"] + mesi)
    for c in ann[1]:
        c.font = grassetto
    for cat in cats:
        riga = [round(sum(m.importo for m in mov if m.categoria == cat and m.data.month == i), 2) or None for i in range(1, 13)]
        ann.append([cat, round(sum(m.importo for m in mov if m.categoria == cat), 2)] + riga)
    ann.column_dimensions["A"].width = 24
    for i in range(1, 13):
        ws = wb.create_sheet(mesi[i - 1])
        dm = [m for m in mov if m.data.month == i]
        ws["A1"] = f"{mesi[i - 1]} {anno}"
        ws["A1"].font = grassetto
        ws["A2"], ws["B2"] = "Categoria", "Totale"
        ws["A2"].font = ws["B2"].font = grassetto
        giorni = (date(anno + (i == 12), i % 12 + 1, 1) - date(anno, i, 1)).days
        for g in range(1, giorni + 1):
            c = ws.cell(2, 2 + g, datetime(anno, i, g))
            c.number_format = "D"
            c.font = grassetto
        cm = sorted({m.categoria for m in dm}, key=lambda c: (c not in ENTRATE, ordine.get(c, 999), c))
        for r, cat in enumerate(cm, 3):
            ws.cell(r, 1, cat)
            tot = 0
            for g in range(1, giorni + 1):
                qui = [m for m in dm if m.categoria == cat and m.data.day == g]
                if not qui:
                    continue
                v = round(sum(m.importo for m in qui), 2)
                tot += v
                cella = ws.cell(r, 2 + g, v)
                note = " · ".join(dict.fromkeys(m.descrizione.strip() for m in qui if m.descrizione and m.descrizione.strip()))
                if note:
                    cella.comment = Comment(note, "Le mie spese")
            ws.cell(r, 2, round(tot, 2))
        r = max(35, 4 + len(cm))  # i totali stanno sotto le righe delle categorie, cosi' non vengono letti come movimenti
        ws.cell(r, 1, "Totale entrate").font = grassetto
        ws.cell(r, 2, round(sum(m.importo for m in dm if m.tipo == "entrata"), 2))
        ws.cell(r + 1, 1, "Totale uscite").font = grassetto
        ws.cell(r + 1, 2, round(sum(m.importo for m in dm if m.tipo == "uscita"), 2))
        ws.column_dimensions["A"].width = 24
        ws.freeze_panes = "C3"
    wb.properties.title = "spese-export"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"Bilancio_{anno}_aggiornato_{date.today()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/importa/elimina/<int:id>", methods=["POST"])
@login_required
def importa_elimina(id):
    f = mio(FileBilancio, id)
    nome = f.nome
    n = Movimento.query.filter_by(user_id=session["uid"], file_id=id, fonte="import").delete()
    db.session.delete(f)
    db.session.commit()
    flash(f"File «{nome}» eliminato, con i suoi {n} movimenti. I movimenti inseriti a mano non sono stati toccati.")
    return redirect(url_for("impostazioni"))


@app.route("/importa/elimina-anno/<int:anno>", methods=["POST"])
@login_required
def importa_elimina_anno(anno):
    """Per i movimenti importati prima che i file venissero registrati (nessun file associato)."""
    if not 1900 <= anno <= 2200:
        abort(404)
    n = Movimento.query.filter(Movimento.user_id == session["uid"], Movimento.fonte == "import", Movimento.file_id.is_(None),
                               Movimento.data >= date(anno, 1, 1), Movimento.data < date(anno + 1, 1, 1)).delete()
    db.session.commit()
    flash(f"Eliminati {n} movimenti importati del {anno}. I movimenti inseriti a mano non sono stati toccati.")
    return redirect(url_for("impostazioni"))


def elenco_bilanci(uid):
    """File importati (con anni e numero di movimenti) + anni importati senza file associato."""
    anno = db.extract("year", Movimento.data)
    per_file = {}
    q = (db.session.query(Movimento.file_id, anno, db.func.count())
         .filter(Movimento.user_id == uid, Movimento.fonte == "import").group_by(Movimento.file_id, anno))
    for fid, y, c in q:
        per_file.setdefault(fid, {})[int(y)] = c
    con_originale = {i for (i,) in db.session.query(FileBilancio.id).filter(FileBilancio.user_id == uid, FileBilancio.contenuto.isnot(None))}
    file = [dict(f=f, anni=sorted(per_file.get(f.id, {})), n=sum(per_file.get(f.id, {}).values()),
                 iso=f.caricato_il.isoformat() + "Z", originale=f.id in con_originale)
            for f in FileBilancio.query.options(db.defer(FileBilancio.contenuto)).filter_by(user_id=uid)
            .order_by(FileBilancio.caricato_il.desc(), FileBilancio.id.desc())]
    vecchi = sorted(per_file.get(None, {}).items())
    return file, vecchi


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


NOTA_PAG = re.compile(r"(\d+)\s*lezion[ei]\s*(?:gi[aà]\s*)?pagat", re.I)
NOTA_EURO = re.compile(r"(\d+(?:[.,]\d+)?)\s*(?:€|eur)", re.I)


def credito_da_note(ws, tariffa):
    """Nel tuo Excel un pagamento anticipato e' scritto come nota sulla prima lezione che copre, ad es. «4 lezioni pagate bonifico 16/10 - 100€».
    Se da quella riga in poi ci sono meno lezioni di quelle pagate, le altre sono gia' pagate e non ancora svolte: restano come credito."""
    righe = []  # (riga, numero di lezioni pagate scritte nella nota, importo nella nota)
    lez = []
    for r in range(2, ws.max_row + 1):
        d, ore = parse_data(ws.cell(r, 1).value), ws.cell(r, 4).value
        if not d or not isinstance(ore, (int, float)) or ore <= 0:
            continue
        lez.append(r)
        for c in range(5, 10):
            v = ws.cell(r, c).value
            m = NOTA_PAG.search(v) if isinstance(v, str) else None
            if m:
                e = NOTA_EURO.search(v)
                righe.append((len(lez) - 1, int(m.group(1)), float(e.group(1).replace(",", ".")) if e else None))
    if not righe:
        return 0.0
    i, n, euro = righe[-1]
    resto = n - (len(lez) - i)
    if resto <= 0 or n <= 0 or any(rosso(ws.cell(r, c)) for r in lez[i:] for c in range(1, 6)):
        return 0.0  # tutte svolte, oppure ci sono lezioni segnate da pagare: non indovino nulla
    return round((euro / n if euro else tariffa) * resto, 2)


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
    inc_mesi, ult, oggi_d = {}, {}, date.today()
    q = db.session.query(Lezione.studente_id, Lezione.ore, Lezione.tariffa, Lezione.pagata, Lezione.data).filter(Lezione.user_id == uid)
    for sid, ore, tar, pag, d in q:
        a = al[sid]
        euro = ore * (tar or a["s"].tariffa or 0)
        kk = (d <= oggi_d, d)  # l'ultima lezione gia' svolta; se non ce ne sono, la piu' recente
        if sid not in ult or kk > ult[sid][0]:
            ult[sid] = (kk, tar if tar is not None else a["s"].tariffa)
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
    tariffe = {a["s"].nome: (ult[a["s"].id][1] if a["s"].id in ult else a["s"].tariffa) or 0 for a in elenco}
    return render_template("ripetizioni.html", tariffe=tariffe, elenco=elenco, visibili=visibili, tutti=tutti, oggi=date.today().isoformat(),
                           inc_mesi=inc_mesi, tot=sum(a["dovuto"] for a in elenco), tot_n=sum(a["n"] for a in elenco),
                           n_deb=sum(1 for a in elenco if a["n"]), p=ore_periodi(uid), ultimo_file=ultimo_file(uid))


@app.route("/ripetizioni/<int:id>")
@login_required
def allievo(id):
    s = mio(Studente, id)
    lez = Lezione.query.filter_by(studente_id=id).order_by(Lezione.data.desc(), Lezione.ora_inizio.desc()).all()
    n = sum(1 for l in lez if not l.pagata)
    oggi = date.today()
    credito = s.credito or 0
    return render_template("allievo.html", s=s, lez=lez, n=n, n_tot=len(lez),
                           ore_svolte=sum(l.ore for l in lez if l.data <= oggi), ore_prog=sum(l.ore for l in lez if l.data > oggi),
                           dovuto=sum(l.ore * (l.tariffa or s.tariffa or 0) for l in lez if not l.pagata),
                           # totale incassato: lezioni pagate (anche future) + credito versato in anticipo
                           incassato=sum(l.ore * (l.tariffa or s.tariffa or 0) for l in lez if l.pagata) + credito,
                           debito_ore=sum(l.ore for l in lez if not l.pagata),
                           credito_ore=round(credito / s.tariffa, 2) if (s.tariffa or 0) > 0 else 0,
                           serie=[[l.data.isoformat(), l.ore] for l in lez],
                           credito=credito, n_pre=n_prepagate(s))


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
    try:  # l'importo scritto nel modulo vale sempre per questa lezione, anche per un allievo gia' presente
        t_in = abs(num(f["tariffa"])) if (f.get("tariffa") or "").strip() else None
    except ValueError:
        t_in = None
    if not s:
        s = Studente(user_id=uid, nome=nome, tariffa=t_in or 0)
        db.session.add(s)
        db.session.flush()
        t_lez = s.tariffa
    elif t_in is not None:
        t_lez = t_in
    else:  # campo vuoto: l'importo dell'ultima lezione svolta
        base = Lezione.query.filter_by(studente_id=s.id).filter(Lezione.tariffa.isnot(None))
        ordine = (Lezione.data.desc(), Lezione.ora_inizio.desc())
        u = base.filter(Lezione.data <= date.today()).order_by(*ordine).first() or base.order_by(*ordine).first()
        t_lez = u.tariffa if u else (s.tariffa or 0)
    l = Lezione(user_id=uid, studente_id=s.id, data=d, ora_inizio=f["inizio"], ora_fine=f["fine"], ore=ore,
                materia=f.get("materia", "").strip(), modalita=f.get("modalita", ""), tariffa=t_lez)
    costo = ore * (t_lez or 0)
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
    msg = "1 lezione segnata come pagata." if n == 1 else f"{n} lezioni segnate come pagate."
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
        da_note = {nome: credito_da_note(ws, tar) for ws, (nome, tar, _n, _l) in zip(wb, fogli)}  # solo se il foglio non ha la colonna Credito
    except Exception:
        flash("Il file non e' un Excel (.xlsx) valido: non ho cambiato nulla.")
        return redirect(url_for("ripetizioni"))
    fogli = [x for x in fogli if x[3]]
    if not fogli:
        flash("Nel file non ho trovato lezioni: non ho cambiato nulla.")
        return redirect(url_for("ripetizioni"))
    try:
        # il nuovo file SOSTITUISCE tutto: lezioni (anche quelle inserite a mano), pagamenti, crediti, note e allievi
        Lezione.query.filter_by(user_id=uid).delete()
        esistenti = {s.nome: s for s in Studente.query.filter_by(user_id=uid)}
        n = nuovi = 0
        nel_file = set()
        for nome, tariffa, nota, lez in fogli:
            s = esistenti.get(nome)
            if not s:
                s = Studente(user_id=uid, nome=nome)
                db.session.add(s)
                nuovi += 1
            # dati dell'allievo presi solo dal file (se ne' tariffa ne' credito sono nel file, ripartono da zero)
            s.tariffa, s.note = tariffa, nota
            s.credito = crediti[nome] if crediti.get(nome) is not None else da_note.get(nome, 0)
            db.session.flush()
            nel_file.add(s.id)
            db.session.add_all(Lezione(user_id=uid, studente_id=s.id, fonte="import", **l) for l in lez)
            n += len(lez)
        rimossi = Studente.query.filter(Studente.user_id == uid, ~Studente.id.in_(nel_file)).delete(synchronize_session=False)  # allievi non piu' nel file
        rec = FileCaricato.query.filter_by(user_id=uid, tipo="ripetizioni").first() or FileCaricato(user_id=uid, tipo="ripetizioni")
        rec.nome, rec.contenuto = (f.filename or "Ripetizioni.xlsx")[:200], dati
        rec.caricato_il = datetime.now(timezone.utc).replace(tzinfo=None)
        db.session.add(rec)
        db.session.commit()
    except Exception:
        db.session.rollback()
        flash("Errore durante l'importazione: i dati precedenti non sono stati modificati.")
        return redirect(url_for("ripetizioni"))
    flash(f"File caricato: i dati precedenti sono stati sostituiti. Ora ci sono {n} lezioni e {len(fogli)} allievi ({nuovi} nuovi, {rimossi} rimossi).")
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
    for g in giorni:  # lezioni alla stessa ora (o che si sovrappongono): ognuna ha la sua colonna, affiancate
        fine_col, gruppo, fine_gruppo = [], [], 0
        for e in sorted(g["ev"], key=lambda x: (x["start"], x["end"])):
            vis = max(e["end"], e["start"] + 30)  # le lezioni brevi sono disegnate alte almeno 30 minuti
            if gruppo and e["start"] >= fine_gruppo:
                for x in gruppo:
                    x["ncols"] = len(fine_col)
                fine_col, gruppo, fine_gruppo = [], [], 0
            for i, fc in enumerate(fine_col):
                if fc <= e["start"]:
                    e["col"], fine_col[i] = i, vis
                    break
            else:
                e["col"] = len(fine_col)
                fine_col.append(vis)
            gruppo.append(e)
            fine_gruppo = max(fine_gruppo, vis)
        for x in gruppo:
            x["ncols"] = len(fine_col)
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


@app.route("/ripetizioni/file/elimina", methods=["POST"])
@login_required
def rip_file_elimina():
    """Toglie il file caricato e i dati che arrivano da lui (lezioni importate e allievi rimasti senza lezioni).
    Le lezioni inserite a mano dal sito non vengono toccate."""
    uid = session["uid"]
    rec = FileCaricato.query.filter_by(user_id=uid, tipo="ripetizioni").first_or_404()
    nome = rec.nome
    ids = {i for (i,) in db.session.query(Lezione.studente_id).filter_by(user_id=uid, fonte="import").distinct()}
    n = Lezione.query.filter_by(user_id=uid, fonte="import").delete()
    vuoti = [i for i in ids if not Lezione.query.filter_by(studente_id=i).count()]
    if vuoti:
        Studente.query.filter(Studente.id.in_(vuoti)).delete(synchronize_session=False)
    db.session.delete(rec)
    db.session.commit()
    flash(f"File «{nome}» eliminato, con le sue {n} lezioni e {len(vuoti)} allievi. Le lezioni inserite a mano non sono state toccate.")
    return redirect(url_for("ripetizioni"))


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
    if not wb.sheetnames:  # nessuna lezione: un Excel senza fogli non si puo' salvare
        flash("Non ci sono ancora lezioni da esportare.")
        return redirect(url_for("ripetizioni"))
    wb.properties.title = "ripetizioni-export"  # riconosce i file scaricati da qui
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"Ripetizioni_aggiornato_{date.today()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------- stipendi (check cedolini) ----------
def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _nota_cella(c):
    if not c.comment:
        return ""
    t = re.sub(r"^\s*Marco Cescon:?\s*", "", c.comment.text.strip())
    return " · ".join(p.strip() for p in t.splitlines() if p.strip())


def leggi_stipendi(wb_val, wb_form):
    """Legge il foglio 'Stipendi' (riga 2 = titoli, dati dalla riga 3). Restituisce l'elenco di dizionari."""
    ws = wb_val["Stipendi"] if "Stipendi" in wb_val.sheetnames else wb_val.worksheets[0]
    wf = wb_form[ws.title]
    out = []
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, 1).value
        d = {"tipo": "mensile"}
        if isinstance(a, (datetime, date)):
            d["anno"], d["mese"] = a.year, a.month
        elif isinstance(a, str) and a.strip():
            m = re.match(r"\s*tredicesima\W*(\d{2,4})", a, re.I)
            if not m:
                continue
            y = int(m.group(1))
            d.update(tipo="tredicesima", anno=y + 2000 if y < 100 else y, mese=12)
        else:
            continue
        for col, att, _t in STIP_CAMPI:
            d[att] = _num(ws[f"{col}{r}"].value)
        if d["competenze"] is None and d["trattenute"] is None:
            continue
        link = ws[f"T{r}"].value
        d["pdf"] = str(link).strip() if link else None
        note = [_nota_cella(wf[f"{c}{r}"]) for c in "IJNRA"]
        d["note"] = " · ".join(n for n in note if n)[:1000] or None
        out.append(d)
    return out


def stip_dict(s):
    d = {"id": s.id, "anno": s.anno, "mese": s.mese, "tipo": s.tipo, "note": s.note or "", "pdf": s.pdf or ""}
    for _c, a, _t in STIP_CAMPI:
        d[a] = getattr(s, a)
    d["netto"] = round((s.competenze or 0) - (s.trattenute or 0), 2)
    d["fe_saldo"] = None if s.fe_prec is None and s.fe_spett is None else round((s.fe_prec or 0) + (s.fe_spett or 0) - (s.fe_god or 0), 2)
    d["par_saldo"] = None if s.par_prec is None and s.par_spett is None else round((s.par_prec or 0) + (s.par_spett or 0) - (s.par_god or 0), 2)
    d["mensa_gg"] = round(s.mensa_tratt / s.gg_mensa, 2) if s.mensa_tratt and s.gg_mensa else None
    return d


def controlli_stipendi(righe):
    """Controlli di coerenza tra mesi consecutivi di ferie e PAR (A. prec. / Spett. / Godute):
    - a gennaio il riporto dell'anno precedente deve essere il saldo di dicembre;
    - le ore spettanti crescono ogni mese del solito accredito (quello piu' frequente);
    - le ore godute non possono diminuire nell'anno.
    Restituisce {id: [avvisi]}."""
    from collections import Counter
    avv = {}
    mens = [r for r in righe if r["tipo"] == "mensile"]
    gruppi = (("ferie", "fe_prec", "fe_spett", "fe_god", "fe_saldo"), ("PAR", "par_prec", "par_spett", "par_god", "par_saldo"))
    acc = {}
    for nome, _p, sp, _g, _s in gruppi:
        c = Counter()
        for prec, r in zip(mens, mens[1:]):
            if r["anno"] == prec["anno"] and r["mese"] == prec["mese"] + 1 and r[sp] is not None and prec[sp] is not None:
                c[round(r[sp] - prec[sp], 2)] += 1
        acc[nome] = c.most_common(1)[0][0] if c else None
    def add(r, t):
        avv.setdefault(r["id"], []).append(t)
    for prec, r in zip(mens, mens[1:]):
        if (r["anno"], r["mese"]) != (prec["anno"] + (prec["mese"] == 12), prec["mese"] % 12 + 1):
            continue  # mese mancante: nessun confronto
        for nome, pr, sp, gd, sl in gruppi:
            if r["mese"] == 1:
                if prec[sl] is not None and r[pr] is not None and abs(prec[sl] - r[pr]) > 0.05:
                    add(r, f"{nome[0].upper() + nome[1:]}: il riporto ({r[pr]:g} h) non coincide con il saldo di dicembre ({prec[sl]:g} h)")
                continue
            if acc[nome] is not None and prec[sp] is not None and r[sp] is not None and abs((r[sp] - prec[sp]) - acc[nome]) > 0.3:
                add(r, f"{nome[0].upper() + nome[1:]}: maturate {r[sp] - prec[sp]:g} h invece delle solite {acc[nome]:g} h")
            if prec[gd] is not None and r[gd] is not None and r[gd] < prec[gd] - 0.05:
                add(r, f"{nome[0].upper() + nome[1:]}: le ore godute sono scese ({prec[gd]:g} → {r[gd]:g} h)")
    costi = [r["mensa_gg"] for r in mens if r["mensa_gg"]]
    if len(costi) >= 6:
        med = sorted(costi)[len(costi) // 2]
        for r in mens:
            if r["mensa_gg"] and abs(r["mensa_gg"] - med) > med * 0.15:
                add(r, f"Mensa: {r['mensa_gg']:.2f} € al giorno (di solito {med:.2f} €)")
    return avv


MESI_LUN = [m.capitalize() for m in MESI_IT]


@app.route("/stipendi")
@login_required
def stipendi():
    uid = session["uid"]
    righe = [stip_dict(s) for s in Stipendio.query.filter_by(user_id=uid).order_by(Stipendio.anno, Stipendio.mese, Stipendio.tipo)]
    avv = controlli_stipendi(righe)
    for r in righe:
        r["avvisi"] = avv.get(r["id"], [])
    f = FileCaricato.query.filter_by(user_id=uid, tipo="stipendi").first()
    ultimo = dict(nome=f.nome, iso=f.caricato_il.replace(tzinfo=timezone.utc).isoformat()) if f else None
    pdfs = {}
    for p in CedolinoPdf.query.filter_by(user_id=uid).with_entities(CedolinoPdf.id, CedolinoPdf.anno, CedolinoPdf.mese, CedolinoPdf.tipo, CedolinoPdf.nome):
        pdfs[f"{p.anno}-{p.mese}-{p.tipo}"] = {"id": p.id, "nome": p.nome, "n": None}
    for r in righe:
        p = pdfs.get(f"{r['anno']}-{r['mese']}-{r['tipo']}")
        if p:
            rec = db.session.get(CedolinoPdf, p["id"])
            p["n"] = sum(1 for x in confronta_pdf(rec, righe) if not x["ok"])
            r["pdf_id"], r["pdf_n"] = p["id"], p["n"]
    elenco = sorted(({"id": p["id"], "nome": p["nome"], "k": k, "n": p["n"]} for k, p in pdfs.items()), key=lambda x: [int(i) if i.isdigit() else i for i in x["k"].split("-")], reverse=True)
    peso = db.session.query(db.func.coalesce(db.func.sum(db.func.length(CedolinoPdf.contenuto)), 0)).filter_by(user_id=uid).scalar() or 0
    return render_template("stipendi.html", righe=righe, ultimo_file=ultimo, mesi=MESI_LUN, pdfs=elenco, peso_pdf=peso,
                           campi=[(a, t) for _c, a, t in STIP_CAMPI])


@app.route("/stipendi/importa", methods=["POST"])
@login_required
def stipendi_importa():
    uid = session["uid"]
    f = request.files.get("file")
    dati = f.read() if f else b""
    try:
        righe = leggi_stipendi(load_workbook(io.BytesIO(dati), data_only=True), load_workbook(io.BytesIO(dati)))
    except Exception:
        flash("Il file non e' un Excel (.xlsx) valido: non ho cambiato nulla.")
        return redirect(url_for("stipendi"))
    if not righe:
        flash("Non ho trovato righe di stipendi nel file (serve il foglio «Stipendi» con le date dei mesi nella colonna A): non ho cambiato nulla.")
        return redirect(url_for("stipendi"))
    Stipendio.query.filter_by(user_id=uid).delete()
    for d in righe:
        db.session.add(Stipendio(user_id=uid, fonte="import", **d))
    FileCaricato.query.filter_by(user_id=uid, tipo="stipendi").delete()
    db.session.add(FileCaricato(user_id=uid, tipo="stipendi", nome=(f.filename or "Stipendi.xlsx")[:200],
                                caricato_il=datetime.now(timezone.utc).replace(tzinfo=None), contenuto=dati))
    db.session.commit()
    flash(f"Importati {len(righe)} cedolini.")
    return redirect(url_for("stipendi"))


@app.route("/stipendi/file")
@login_required
def stipendi_file():
    r = FileCaricato.query.filter_by(user_id=session["uid"], tipo="stipendi").first_or_404()
    return send_file(io.BytesIO(r.contenuto), as_attachment=True, download_name=r.nome)


@app.route("/stipendi/file/elimina", methods=["POST"])
@login_required
def stipendi_file_elimina():
    uid = session["uid"]
    rec = FileCaricato.query.filter_by(user_id=uid, tipo="stipendi").first_or_404()
    n = Stipendio.query.filter_by(user_id=uid, fonte="import").delete()
    nome = rec.nome
    db.session.delete(rec)
    db.session.commit()
    flash(f"File «{nome}» eliminato, con i suoi {n} cedolini. I mesi inseriti a mano dal sito non sono stati toccati.")
    return redirect(url_for("stipendi"))


@app.route("/stipendi/salva", methods=["POST"])
@login_required
def stipendi_salva():
    uid = session["uid"]
    try:
        anno, mese = int(request.form["anno"]), int(request.form["mese"])
        assert 2000 <= anno <= 2100 and 1 <= mese <= 12
    except (KeyError, ValueError, AssertionError):
        flash("Anno o mese non validi.")
        return redirect(url_for("stipendi"))
    tipo = "tredicesima" if request.form.get("tipo") == "tredicesima" else "mensile"
    s = Stipendio.query.filter_by(user_id=uid, anno=anno, mese=mese, tipo=tipo).first()
    nuovo = s is None
    if nuovo:
        s = Stipendio(user_id=uid, anno=anno, mese=mese, tipo=tipo, fonte="sito")
        db.session.add(s)
    for _c, a, _t in STIP_CAMPI:
        v = request.form.get(a, "").strip().replace(",", ".")
        setattr(s, a, _num(v))
    s.note = request.form.get("note", "").strip()[:1000] or None
    s.pdf = request.form.get("pdf", "").strip()[:200] or None
    db.session.commit()
    flash(f"Cedolino {'aggiunto' if nuovo else 'aggiornato'}: {MESI_LUN[mese - 1]} {anno}{' (tredicesima)' if tipo == 'tredicesima' else ''}.")
    return redirect(url_for("stipendi"))


@app.route("/stipendi/<int:id>/elimina", methods=["POST"])
@login_required
def stipendi_elimina(id):
    s = Stipendio.query.filter_by(id=id, user_id=session["uid"]).first_or_404()
    db.session.delete(s)
    db.session.commit()
    flash("Cedolino eliminato.")
    return redirect(url_for("stipendi"))


@app.route("/stipendi/esporta")
@login_required
def stipendi_esporta():
    """Excel nello stesso formato di Check_stipendi (si puo' ricaricare)."""
    righe = Stipendio.query.filter_by(user_id=session["uid"]).order_by(Stipendio.anno, Stipendio.mese, Stipendio.tipo).all()
    if not righe:
        flash("Non ci sono ancora cedolini da esportare.")
        return redirect(url_for("stipendi"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Stipendi"
    ws["L1"], ws["P1"] = "FERIE", "PAR"
    tit = {"K": "NETTO", "O": "Saldo", "S": "Saldo", "T": "LINK PDF", "Y": "Costo mensa giornaliero medio"}
    tit.update({c: t for c, _a, t in STIP_CAMPI})
    for c, t in tit.items():
        ws[f"{c}2"] = t
    for c in ws[2]:
        c.font = Font(bold=True)
    for i, s in enumerate(righe):
        r = i + 3
        if s.tipo == "tredicesima":
            ws[f"A{r}"] = f"tredicesima-{str(s.anno)[2:]}"
        else:
            ws[f"A{r}"] = date(s.anno, s.mese, 1)
            ws[f"A{r}"].number_format = "mmmm yyyy"
        for c, a, _t in STIP_CAMPI:
            ws[f"{c}{r}"] = getattr(s, a)
        ws[f"K{r}"] = f"=I{r}-J{r}"
        ws[f"O{r}"] = f"=L{r}+M{r}-N{r}"
        ws[f"S{r}"] = f"=P{r}+Q{r}-R{r}"
        ws[f"Y{r}"] = f'=IF(AND(C{r}>0,W{r}<>""),W{r}/C{r},"")'
        ws[f"T{r}"] = s.pdf
        if s.note:
            ws[f"I{r}"].comment = Comment(s.note, "Le mie spese")
    ws.column_dimensions["A"].width = 16
    ws.freeze_panes = "B3"
    wb.properties.title = "stipendi-export"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"Check_stipendi_aggiornato_{date.today()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------- cedolini PDF ----------
PDF_CAMPI = [(a, t) for _c, a, t in STIP_CAMPI if a not in ("mensa_sup",)]


def confronta_pdf(rec, righe=None):
    """Confronta i valori letti dal PDF con la riga di Check_stipendi (stesso mese). Restituisce un elenco di righe."""
    v = json.loads(rec.dati)["v"]
    if righe is None:
        righe = [stip_dict(s) for s in Stipendio.query.filter_by(user_id=rec.user_id).order_by(Stipendio.anno, Stipendio.mese)]
    riga = next((r for r in righe if (r["anno"], r["mese"], r["tipo"]) == (rec.anno, rec.mese, rec.tipo)), None)
    out = []
    if not riga:
        return out
    campi = [(a, t) for a, t in PDF_CAMPI]
    campi.append(("netto", "Netto da pagare"))
    for a, t in campi:
        p, x = v.get(a), riga.get(a)
        if a in ("fe_god", "par_god"):
            p, x = p or 0, x or 0
        if p is None and x is None:
            continue
        if p is None:
            continue  # il PDF non riporta quel dato: non si puo' confrontare
        tol = 0.001 if a == "unitario" else 0.011
        out.append({"campo": a, "titolo": t, "pdf": p, "riga": x, "ok": x is not None and abs(p - x) <= tol})
    # ore scalate nel mese, calcolate dai saldi: saldo del mese prima + ore maturate nel mese - saldo di questo mese.
    # (non basta guardare la colonna "Godute": a volte le ore vengono tolte dal riporto dell'anno precedente)
    prec = next((r for r in righe if r["tipo"] == "mensile" and (r["anno"], r["mese"]) == ((rec.anno, rec.mese - 1) if rec.mese > 1 else (rec.anno - 1, 12))), None)
    if rec.tipo == "mensile" and prec:
        for pre, nome, chiave in (("fe", "Ferie", "_fe_desc"), ("par", "PAR", "_par_desc")):
            ps, sp, sl_prec = v.get(f"{pre}_prec"), v.get(f"{pre}_spett"), prec.get(f"{pre}_saldo")
            if None in (ps, sp, sl_prec):
                continue
            sl = ps + sp - (v.get(f"{pre}_god") or 0)
            maturate = sp - ((prec.get(f"{pre}_spett") or 0) if rec.mese > 1 else 0)
            scalate = round(sl_prec + maturate - sl, 2)
            desc = v.get(chiave) or 0.0
            out.append({"campo": f"_{pre}_mese", "titolo": f"{nome} del mese: ore in descrizione / ore scalate dal saldo", "pdf": desc, "riga": scalate, "ok": abs(desc - scalate) <= 0.011,
                        "nota": f"Saldo {MESI_LUN[prec['mese'] - 1].lower()} {prec.get(f'{pre}_saldo'):g} h + maturate nel mese {maturate:g} h − saldo di questo mese {sl:g} h = {scalate:g} h scalate"})
    return out


def _applica_valori(uid, v):
    s = Stipendio.query.filter_by(user_id=uid, anno=v["anno"], mese=v["mese"], tipo=v["tipo"]).first()
    nuovo = s is None
    if nuovo:
        s = Stipendio(user_id=uid, anno=v["anno"], mese=v["mese"], tipo=v["tipo"], fonte="pdf")
        db.session.add(s)
    for a, _t in PDF_CAMPI:
        if v.get(a) is not None or a in ("fe_god", "par_god"):
            setattr(s, a, v.get(a))
    return nuovo


@app.route("/stipendi/pdf", methods=["POST"])
@login_required
def stipendi_pdf():
    from cedolino import leggi_cedolino
    uid = session["uid"]
    msg = []
    for f in request.files.getlist("file"):
        dati = f.read()
        if not dati:
            continue
        try:
            v, voci, ctrl = leggi_cedolino(dati)
        except ValueError as e:
            msg.append(f"{f.filename}: {e}.")
            continue
        rec = CedolinoPdf.query.filter_by(user_id=uid, anno=v["anno"], mese=v["mese"], tipo=v["tipo"]).first()
        if not rec:
            rec = CedolinoPdf(user_id=uid, anno=v["anno"], mese=v["mese"], tipo=v["tipo"])
            db.session.add(rec)
        rec.nome = (f.filename or "cedolino.pdf")[:200]
        rec.caricato_il = datetime.now(timezone.utc).replace(tzinfo=None)
        rec.contenuto = dati
        rec.dati = json.dumps({"v": v, "voci": voci, "c": ctrl})
        db.session.flush()
        creato = False
        if not Stipendio.query.filter_by(user_id=uid, anno=v["anno"], mese=v["mese"], tipo=v["tipo"]).first():
            creato = _applica_valori(uid, v)
        db.session.flush()
        n = sum(1 for x in confronta_pdf(rec) if not x["ok"]) + sum(1 for x in ctrl if not x["ok"])
        msg.append(f"{MESI_LUN[v['mese'] - 1]} {v['anno']}: {'riga creata dal PDF, ' if creato else ''}" + ("tutto coincide." if not n else f"{n} da controllare."))
    db.session.commit()
    flash(" ".join(msg) or "Nessun file ricevuto.")
    return redirect(url_for("stipendi"))


@app.route("/stipendi/pdf/rileggi", methods=["POST"])
@login_required
def stipendi_pdf_rileggi():
    """Rilegge tutti i PDF salvati con il lettore attuale (utile dopo un miglioramento del lettore)."""
    from cedolino import leggi_cedolino
    n = err = 0
    for rec in CedolinoPdf.query.filter_by(user_id=session["uid"]):
        try:
            v, voci, ctrl = leggi_cedolino(rec.contenuto)
        except ValueError:
            err += 1
            continue
        rec.dati = json.dumps({"v": v, "voci": voci, "c": ctrl})
        n += 1
    db.session.commit()
    flash(f"Riletti {n} PDF" + (f"; {err} non leggibili." if err else "."))
    return redirect(url_for("stipendi"))


@app.route("/stipendi/pdf/<int:id>")
@login_required
def stipendi_pdf_dettaglio(id):
    rec = CedolinoPdf.query.filter_by(id=id, user_id=session["uid"]).first_or_404()
    d = json.loads(rec.dati)
    righe = [stip_dict(s) for s in Stipendio.query.filter_by(user_id=rec.user_id)]
    return render_template("cedolino.html", rec=rec, conf=confronta_pdf(rec, righe), voci=d["voci"], ctrl=d["c"],
                           ha_riga=any((r["anno"], r["mese"], r["tipo"]) == (rec.anno, rec.mese, rec.tipo) for r in righe),
                           mese=MESI_LUN[rec.mese - 1], iso=rec.caricato_il.replace(tzinfo=timezone.utc).isoformat())


@app.route("/stipendi/pdf/<int:id>/file")
@login_required
def stipendi_pdf_file(id):
    rec = CedolinoPdf.query.filter_by(id=id, user_id=session["uid"]).first_or_404()
    return send_file(io.BytesIO(rec.contenuto), as_attachment=True, download_name=rec.nome, mimetype="application/pdf")


@app.route("/stipendi/pdf/<int:id>/elimina", methods=["POST"])
@login_required
def stipendi_pdf_elimina(id):
    rec = CedolinoPdf.query.filter_by(id=id, user_id=session["uid"]).first_or_404()
    db.session.delete(rec)
    db.session.commit()
    flash("PDF eliminato (la riga del cedolino resta).")
    return redirect(url_for("stipendi"))


@app.route("/stipendi/pdf/<int:id>/applica", methods=["POST"])
@login_required
def stipendi_pdf_applica(id):
    rec = CedolinoPdf.query.filter_by(id=id, user_id=session["uid"]).first_or_404()
    _applica_valori(session["uid"], json.loads(rec.dati)["v"])
    db.session.commit()
    flash("Riga aggiornata con i valori del PDF.")
    return redirect(url_for("stipendi_pdf_dettaglio", id=id))
