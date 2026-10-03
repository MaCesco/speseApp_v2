import os, secrets
from datetime import date, datetime
from functools import wraps
import click
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, abort
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook

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


with app.app_context():
    db.create_all()


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
    db.session.add(User(username=username, password=generate_password_hash(password)))
    db.session.commit()
    print("Utente creato.")


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
