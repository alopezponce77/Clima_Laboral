import csv
import io
import os
import sqlite3
from datetime import datetime
from functools import wraps

from flask import (
    Flask,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
except ImportError:
    canvas = None
    letter = None


BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATABASE = os.path.join(BASE_DIR, "database", "clima_laboral.db")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("CLIMA_SECRET_KEY", "cambiar-esta-clave-en-produccion")


CATALOGOS = {
    "sexo": ["Masculino", "Femenino"],
    "escolaridad": ["Primaria", "Secundaria", "Comercial", "Tecnica", "Preparatoria", "Profesional", "Otra"],
    "estado_civil": ["Soltero", "Casado", "Union libre", "Divorciado", "Viudo"],
    "antiguedad": ["0 a 6 meses", "6 meses a 1 ano", "1 a 3 anos", "3 a 5 anos", "Mas de 5 anos"],
    "edad_rango": ["18-25", "26-30", "31-35", "36-40", "41-45", "46-50", "51-55", "56 o mas"],
}

PREGUNTAS_ABIERTAS = [
    (72, "Cuando tienes problemas en tu trabajo, a que persona le pides apoyo para su solucion?"),
    (73, "Quien dentro del hotel crees que es un ejemplo a seguir en aspectos de trabajo y companerismo?"),
    (74, "A quien consideras mas interesado en organizar deportes o actividades entre el personal?"),
    (75, "Comentarios adicionales que desee hacer"),
]

PREGUNTAS = [
    ("Trabajo", "Me siento contento de trabajar en esta empresa", 0),
    ("Trabajo", "Me gusta el trabajo que desempeno", 0),
    ("Trabajo", "Si me ofrecieran el mismo sueldo en otra empresa me cambiaria", 1),
    ("Trabajo", "Tengo el equipo y material necesario para llevar a cabo mi trabajo", 0),
    ("Trabajo", "El equipo y material necesario para mi trabajo esta en buen estado", 0),
    ("Trabajo", "Trabajar en esta empresa me hace sentir importante", 0),
    ("Trabajo", "Mi trabajo es importante para el funcionamiento de mi departamento", 0),
    ("Trabajo", "En mi departamento se promueve que todos realicemos el trabajo excelente", 0),
    ("Trabajo", "Si realizo bien mi trabajo me sentire seguro de permanecer en esta empresa", 0),
    ("Trabajo", "Mi horario de trabajo es suficiente para realizar mis tareas diarias", 0),
    ("Trabajo", "Tengo el uniforme completo y adecuado para realizar mi trabajo", 0),
    ("Lugar de Trabajo", "El lugar donde trabajo esta bien ventilado e iluminado", 0),
    ("Lugar de Trabajo", "Me siento orgulloso de trabajar en este hotel", 0),
    ("Lugar de Trabajo", "Considero que este hotel es el mejor de la ciudad", 0),
    ("Lugar de Trabajo", "Me encuentro a gusto en mi departamento", 0),
    ("Lugar de Trabajo", "En esta empresa se preocupan por el bienestar de sus empleados", 0),
    ("Lugar de Trabajo", "En esta empresa me tratan con dignidad y respeto", 0),
    ("Comunicacion", "Las relaciones de trabajo en esta empresa se hacen en un ambiente de respeto", 0),
    ("Comunicacion", "Mi jefe me informa sobre mi desempeno en el trabajo", 0),
    ("Comunicacion", "Conozco companeros que han sido promovidos en el hotel o el grupo", 0),
    ("Comunicacion", "Recibo informacion para mi trabajo a traves de cursos o conferencias", 0),
    ("Comunicacion", "Recibo informacion para mi trabajo en tableros de comunicacion", 0),
    ("Comunicacion", "Recibo informacion para mi trabajo en juntas o reuniones", 0),
    ("Comunicacion", "Conozco cual es la mision del hotel", 0),
    ("Comunicacion", "Conozco cual es la mision de Optima Hoteles de Mexico", 0),
    ("Comunicacion", "Conozco cual es la vision de Optima Hoteles de Mexico", 0),
    ("Comunicacion", "Conozco cuales son los valores de Optima Hoteles de Mexico", 0),
    ("Comunicacion", "Conozco las instalaciones de los hoteles", 0),
    ("Comunicacion", "He platicado con el Director General de Optima", 0),
    ("Supervision", "Mi jefe directo tiene una buena relacion de trabajo conmigo", 0),
    ("Supervision", "Yo puedo decirle a mi jefe libremente cuando estoy en desacuerdo con el", 0),
    ("Supervision", "Mi jefe directo me da instrucciones claras sobre mi trabajo", 0),
    ("Supervision", "Mi jefe me permite tomar decisiones de acuerdo a mi puesto", 0),
    ("Supervision", "Mi jefe me apoya y me ayuda a tomar decisiones", 0),
    ("Supervision", "Cuando hago bien un trabajo, mi jefe me felicita", 0),
    ("Supervision", "Cuando mi jefe me llama la atencion lo hace con justicia", 0),
    ("Supervision", "Cuando mi jefe me llama la atencion lo hace con respeto", 0),
    ("Supervision", "En este hotel existe favoritismo", 1),
    ("Promocion", "Tengo posibilidades de progresar en este hotel", 0),
    ("Promocion", "En el hotel la forma de subir de puesto es justa", 0),
    ("Promocion", "En el hotel existen posibilidades para lograr una promocion", 0),
    ("Promocion", "He sido tomado en cuenta para una promocion", 0),
    ("Promocion", "Conozco el proceso de promocion y desarrollo en la empresa", 0),
    ("Integracion", "En este hotel existen buenas relaciones de trabajo entre companeros", 0),
    ("Integracion", "En este hotel existen buenas relaciones de trabajo entre departamentos", 0),
    ("Integracion", "En el hotel existen buenas relaciones de trabajo entre gerentes", 0),
    ("Integracion", "Recibo ayuda de mis companeros de trabajo cuando lo solicito", 0),
    ("Integracion", "Se me brinda la oportunidad de hacer sugerencias para mejorar el trabajo de mi departamento", 0),
    ("Capacitacion", "Cuando entre a trabajar a este hotel, asisti al curso de induccion", 0),
    ("Capacitacion", "El programa de induccion de esta empresa es bueno", 0),
    ("Capacitacion", "El hotel se ocupa por preparar bien a su personal", 0),
    ("Capacitacion", "Conozco bien como se hace mi trabajo", 0),
    ("Capacitacion", "En el hotel me dan capacitacion para hacer mejor mi trabajo", 0),
    ("Capacitacion", "Recibo la ayuda necesaria para poder realizar mi trabajo cuando lo solicito", 0),
    ("Seguridad", "He recibido entrenamiento en aspectos de seguridad", 0),
    ("Seguridad", "Se lo que significa Clave 100", 0),
    ("Seguridad", "Conozco las rutas de evacuacion", 0),
    ("Seguridad", "Conozco al jefe de brigadas de emergencias", 0),
    ("Seguridad", "Se que hacer en caso de incendio", 0),
    ("Higienicos", "El comedor esta en buen estado", 0),
    ("Higienicos", "La comida que se proporciona es buena", 0),
    ("Higienicos", "Los roles de horarios son justos", 0),
    ("Higienicos", "Las instalaciones son buenas: banos, vestidores y casilleros", 0),
    ("Higienicos", "El horario de trabajo es adecuado", 0),
    ("Sueldo", "La empresa se preocupa por mejorar los ingresos de sus colaboradores", 0),
    ("Sueldo", "Las prestaciones que otorga el hotel son buenas comparadas con otros hoteles", 0),
    ("Sueldo", "El sueldo de los empleados que trabajamos en el hotel es justo", 0),
    ("Sueldo", "Mi sueldo es justo comparado con el que reciben otros companeros", 0),
    ("Sueldo", "El sueldo que recibo es adecuado por el trabajo que hago", 0),
    ("Sueldo", "En el hotel toman en cuenta los resultados de mi trabajo para aumentarme el sueldo", 0),
    ("Sueldo", "Estoy de acuerdo con el sueldo que recibo", 0),
]


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_error):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    os.makedirs(os.path.dirname(DATABASE), exist_ok=True)
    os.makedirs(EXPORT_DIR, exist_ok=True)
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    cur = db.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            usuario TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            rol TEXT NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1,
            fecha_creacion TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS hoteles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            codigo TEXT,
            activo INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS departamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS campanas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            fecha_inicio TEXT,
            fecha_fin TEXT,
            activa INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS preguntas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero INTEGER NOT NULL UNIQUE,
            factor TEXT NOT NULL,
            texto TEXT NOT NULL,
            valor_invertido INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS encuestas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            campana_id INTEGER NOT NULL,
            hotel_id INTEGER NOT NULL,
            departamento_id INTEGER NOT NULL,
            fecha_captura TEXT NOT NULL,
            sexo TEXT,
            escolaridad TEXT,
            estado_civil TEXT,
            antiguedad TEXT,
            edad_rango TEXT
        );
        CREATE TABLE IF NOT EXISTS respuestas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            encuesta_id INTEGER NOT NULL,
            pregunta_id INTEGER NOT NULL,
            respuesta TEXT NOT NULL,
            valor INTEGER NOT NULL
        );
        CREATE TABLE IF NOT EXISTS respuestas_abiertas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            encuesta_id INTEGER NOT NULL,
            numero INTEGER NOT NULL,
            pregunta TEXT NOT NULL,
            respuesta TEXT NOT NULL
        );
        """
    )
    if cur.execute("SELECT COUNT(*) FROM preguntas").fetchone()[0] == 0:
        cur.executemany(
            "INSERT INTO preguntas (numero, factor, texto, valor_invertido) VALUES (?, ?, ?, ?)",
            [(i + 1, factor, texto, invertida) for i, (factor, texto, invertida) in enumerate(PREGUNTAS)],
        )
    if cur.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO usuarios (nombre, usuario, password_hash, rol, activo, fecha_creacion) VALUES (?, ?, ?, ?, 1, ?)",
            ("Administrador inicial", "admin", generate_password_hash("admin123"), "Administrador", datetime.now().isoformat(timespec="seconds")),
        )
    if cur.execute("SELECT COUNT(*) FROM campanas").fetchone()[0] == 0:
        cur.execute("INSERT INTO campanas (nombre, fecha_inicio, fecha_fin, activa) VALUES (?, ?, ?, 1)", ("Campana inicial", "", ""))
    if cur.execute("SELECT COUNT(*) FROM hoteles").fetchone()[0] == 0:
        cur.execute("INSERT INTO hoteles (nombre, codigo, activo) VALUES (?, ?, 1)", ("Hotel principal", "H001"))
    if cur.execute("SELECT COUNT(*) FROM departamentos").fetchone()[0] == 0:
        cur.execute("INSERT INTO departamentos (nombre, activo) VALUES (?, 1)", ("General",))
    db.commit()
    db.close()


def current_user():
    if not session.get("usuario_id"):
        return None
    return get_db().execute("SELECT * FROM usuarios WHERE id = ?", (session["usuario_id"],)).fetchone()


@app.before_request
def load_user():
    g.user = current_user()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if g.user is None:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def roles_required(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if g.user is None:
                return redirect(url_for("login", next=request.path))
            if roles and g.user["rol"] not in roles:
                flash("No tiene permiso para acceder a esta seccion", "warning")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapped
    return decorator


def catalog_data(active_only=True):
    db = get_db()
    active_sql = " WHERE activo = 1" if active_only else ""
    campana_active_sql = " WHERE activa = 1" if active_only else ""
    return {
        "campanas": db.execute("SELECT * FROM campanas" + campana_active_sql + " ORDER BY activa DESC, nombre").fetchall(),
        "hoteles": db.execute("SELECT * FROM hoteles" + active_sql + " ORDER BY nombre").fetchall(),
        "departamentos": db.execute("SELECT * FROM departamentos" + active_sql + " ORDER BY nombre").fetchall(),
        "catalogos": CATALOGOS,
    }


def build_filter_clause(prefix="e"):
    clauses = []
    params = []
    mapping = {
        "campana_id": "campana_id",
        "hotel_id": "hotel_id",
        "departamento_id": "departamento_id",
        "sexo": "sexo",
        "edad_rango": "edad_rango",
        "antiguedad": "antiguedad",
        "escolaridad": "escolaridad",
    }
    for arg, column in mapping.items():
        value = request.args.get(arg)
        if value:
            clauses.append(f"{prefix}.{column} = ?")
            params.append(value)
    if request.args.get("fecha_inicio"):
        clauses.append(f"date({prefix}.fecha_captura) >= date(?)")
        params.append(request.args["fecha_inicio"])
    if request.args.get("fecha_fin"):
        clauses.append(f"date({prefix}.fecha_captura) <= date(?)")
        params.append(request.args["fecha_fin"])
    return (" AND ".join(clauses), params)


def pct(value, total):
    return round((value / total * 100), 1) if total else 0


def color_for(value):
    if value >= 85:
        return "success"
    if value >= 70:
        return "warning"
    return "danger"


def query_metric(group_field=None, join_extra="", select_label=None):
    where, params = build_filter_clause("e")
    sql_where = f"WHERE {where}" if where else ""
    select = f"{select_label or group_field} AS label, " if group_field else ""
    group = f"GROUP BY {group_field} ORDER BY resultado DESC" if group_field else ""
    sql = f"""
        SELECT {select} ROUND(AVG(r.valor) * 100, 1) AS resultado, COUNT(r.id) AS total
        FROM respuestas r
        JOIN encuestas e ON e.id = r.encuesta_id
        JOIN preguntas p ON p.id = r.pregunta_id
        {join_extra}
        {sql_where}
        {group}
    """
    return get_db().execute(sql, params).fetchall()


def dashboard_data():
    db = get_db()
    where, params = build_filter_clause("e")
    sql_where = f"WHERE {where}" if where else ""
    total_encuestas = db.execute(f"SELECT COUNT(*) FROM encuestas e {sql_where}", params).fetchone()[0]
    general_row = db.execute(
        f"SELECT ROUND(AVG(r.valor) * 100, 1) FROM respuestas r JOIN encuestas e ON e.id = r.encuesta_id {sql_where}",
        params,
    ).fetchone()
    general = general_row[0] or 0
    factores = query_metric("p.factor", select_label="p.factor")
    hoteles = query_metric("h.nombre", "JOIN hoteles h ON h.id = e.hotel_id", "h.nombre")
    departamentos = query_metric("d.nombre", "JOIN departamentos d ON d.id = e.departamento_id", "d.nombre")
    preguntas = query_metric("p.numero || '. ' || p.texto", select_label="p.numero || '. ' || p.texto")
    sexo = count_distribution("sexo")
    escolaridad = count_distribution("escolaridad")
    edad = count_distribution("edad_rango")
    antiguedad = count_distribution("antiguedad")
    campañas = campaign_trend()
    return {
        "total_encuestas": total_encuestas,
        "indice_general": general,
        "color_general": color_for(general),
        "factores": rows_to_chart(factores),
        "hoteles": rows_to_chart(hoteles),
        "departamentos": rows_to_chart(departamentos),
        "preguntas": rows_to_chart(preguntas),
        "sexo": rows_to_chart(sexo, value_key="total"),
        "escolaridad": rows_to_chart(escolaridad, value_key="total"),
        "edad": rows_to_chart(edad, value_key="total"),
        "antiguedad": rows_to_chart(antiguedad, value_key="total"),
        "campanas": rows_to_chart(campañas),
        "mejor_factor": label_at(factores, 0),
        "factor_bajo": label_at(list(reversed(factores)), 0),
        "mejor_hotel": label_at(hoteles, 0),
        "hotel_bajo": label_at(list(reversed(hoteles)), 0),
    }


def rows_to_chart(rows, value_key="resultado"):
    return {"labels": [row["label"] for row in rows], "values": [row[value_key] or 0 for row in rows]}


def label_at(rows, index):
    if not rows:
        return "Sin datos"
    row = rows[index]
    return f"{row['label']} ({row['resultado']}%)"


def count_distribution(column):
    where, params = build_filter_clause("e")
    extra = f" AND {where}" if where else ""
    return get_db().execute(
        f"SELECT COALESCE({column}, 'Sin dato') AS label, COUNT(*) AS total FROM encuestas e WHERE 1=1 {extra} GROUP BY {column} ORDER BY total DESC",
        params,
    ).fetchall()


def campaign_trend():
    return get_db().execute(
        """
        SELECT c.nombre AS label, ROUND(AVG(r.valor) * 100, 1) AS resultado
        FROM respuestas r
        JOIN encuestas e ON e.id = r.encuesta_id
        JOIN campanas c ON c.id = e.campana_id
        GROUP BY c.id
        ORDER BY c.fecha_inicio, c.id
        """
    ).fetchall()


@app.route("/")
def home():
    return render_template("home.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = get_db().execute("SELECT * FROM usuarios WHERE usuario = ? AND activo = 1", (request.form["usuario"],)).fetchone()
        if user and check_password_hash(user["password_hash"], request.form["password"]):
            session.clear()
            session["usuario_id"] = user["id"]
            return redirect(request.args.get("next") or url_for("dashboard"))
        flash("Usuario o contrasena incorrectos", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sesion cerrada correctamente", "success")
    return redirect(url_for("home"))


@app.route("/encuesta", methods=["GET", "POST"])
def encuesta():
    db = get_db()
    preguntas = db.execute("SELECT * FROM preguntas ORDER BY numero").fetchall()
    paso = request.args.get("paso", "campana")
    data = catalog_data()
    encuesta_datos = session.get("encuesta_datos", {})

    if request.method == "GET":
        if paso == "datos" and not encuesta_datos.get("campana_id"):
            return redirect(url_for("encuesta", paso="campana"))
        if paso == "preguntas" and any(not encuesta_datos.get(field) for field in ["campana_id", "hotel_id", "departamento_id", "sexo", "escolaridad", "estado_civil", "antiguedad", "edad_rango"]):
            return redirect(url_for("encuesta", paso="campana"))

    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "campana":
            if not request.form.get("campana_id"):
                flash("Seleccione una campana para continuar", "warning")
                return redirect(url_for("encuesta", paso="campana"))
            session["encuesta_datos"] = {"campana_id": request.form["campana_id"]}
            return redirect(url_for("encuesta", paso="datos"))

        if accion == "datos":
            required = ["hotel_id", "departamento_id", "sexo", "escolaridad", "estado_civil", "antiguedad", "edad_rango"]
            if any(not request.form.get(field) for field in required):
                flash("Complete todos los datos generales antes de continuar", "warning")
                return redirect(url_for("encuesta", paso="datos"))
            encuesta_datos = session.get("encuesta_datos", {})
            if not encuesta_datos.get("campana_id"):
                flash("Seleccione una campana para iniciar la encuesta", "warning")
                return redirect(url_for("encuesta", paso="campana"))
            for field in required:
                encuesta_datos[field] = request.form[field]
            session["encuesta_datos"] = encuesta_datos
            return redirect(url_for("encuesta", paso="preguntas"))

        encuesta_datos = session.get("encuesta_datos", {})
        required = ["campana_id", "hotel_id", "departamento_id", "sexo", "escolaridad", "estado_civil", "antiguedad", "edad_rango"]
        if any(not encuesta_datos.get(field) for field in required):
            flash("Complete la campana y los datos generales antes de responder", "warning")
            return redirect(url_for("encuesta", paso="campana"))
        missing = [str(p["numero"]) for p in preguntas if request.form.get(f"pregunta_{p['id']}") not in ("Si", "No")]
        missing_open = [str(numero) for numero, _texto in PREGUNTAS_ABIERTAS if not request.form.get(f"abierta_{numero}", "").strip()]
        if missing or missing_open:
            flash("No se pueden dejar preguntas sin responder", "warning")
            paso = "preguntas"
        else:
            cur = db.cursor()
            cur.execute(
                """
                INSERT INTO encuestas
                (campana_id, hotel_id, departamento_id, fecha_captura, sexo, escolaridad, estado_civil, antiguedad, edad_rango)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    encuesta_datos["campana_id"],
                    encuesta_datos["hotel_id"],
                    encuesta_datos["departamento_id"],
                    datetime.now().isoformat(timespec="seconds"),
                    encuesta_datos["sexo"],
                    encuesta_datos["escolaridad"],
                    encuesta_datos["estado_civil"],
                    encuesta_datos["antiguedad"],
                    encuesta_datos["edad_rango"],
                ),
            )
            encuesta_id = cur.lastrowid
            rows = []
            for p in preguntas:
                respuesta = request.form[f"pregunta_{p['id']}"]
                positive = 1 if respuesta == "Si" else 0
                valor = 1 - positive if p["valor_invertido"] else positive
                rows.append((encuesta_id, p["id"], respuesta, valor))
            cur.executemany("INSERT INTO respuestas (encuesta_id, pregunta_id, respuesta, valor) VALUES (?, ?, ?, ?)", rows)
            cur.executemany(
                "INSERT INTO respuestas_abiertas (encuesta_id, numero, pregunta, respuesta) VALUES (?, ?, ?, ?)",
                [(encuesta_id, numero, texto, request.form[f"abierta_{numero}"].strip()) for numero, texto in PREGUNTAS_ABIERTAS],
            )
            db.commit()
            session.pop("encuesta_datos", None)
            return redirect(url_for("encuesta_gracias"))

    grouped = {}
    for p in preguntas:
        grouped.setdefault(p["factor"], []).append(p)
    data.update({
        "paso": paso,
        "encuesta_datos": session.get("encuesta_datos", {}),
        "preguntas_por_factor": grouped,
        "preguntas_abiertas": PREGUNTAS_ABIERTAS,
    })
    return render_template("encuesta.html", **data)


@app.route("/encuesta/gracias")
def encuesta_gracias():
    session.pop("encuesta_datos", None)
    return render_template("encuesta_gracias.html")


@app.route("/dashboard")
@login_required
def dashboard():
    data = dashboard_data()
    return render_template("dashboard.html", data=data, **catalog_data(active_only=False))


@app.route("/api/dashboard")
@login_required
def api_dashboard():
    return jsonify(dashboard_data())


@app.route("/resultados/<vista>")
@login_required
def resultados(vista):
    titulos = {
        "general": "Resultados generales",
        "hotel": "Resultados por hotel",
        "departamento": "Resultados por departamento",
        "factor": "Resultados por factor",
    }
    return render_template("resultados.html", vista=vista, titulo=titulos.get(vista, "Resultados"), data=dashboard_data(), **catalog_data(active_only=False))


@app.route("/admin/usuarios", methods=["GET", "POST"])
@login_required
def usuarios():
    db = get_db()
    if request.method == "POST":
        user_id = request.form.get("id")
        activo = 1 if "activo" in request.form else 0
        if user_id:
            if request.form.get("password"):
                db.execute(
                    "UPDATE usuarios SET nombre=?, usuario=?, password_hash=?, rol=?, activo=? WHERE id=?",
                    (request.form["nombre"], request.form["usuario"], generate_password_hash(request.form["password"]), "Administrador", activo, user_id),
                )
            else:
                db.execute(
                    "UPDATE usuarios SET nombre=?, usuario=?, rol=?, activo=? WHERE id=?",
                    (request.form["nombre"], request.form["usuario"], "Administrador", activo, user_id),
                )
        else:
            db.execute(
                "INSERT INTO usuarios (nombre, usuario, password_hash, rol, activo, fecha_creacion) VALUES (?, ?, ?, ?, ?, ?)",
                (request.form["nombre"], request.form["usuario"], generate_password_hash(request.form["password"]), "Administrador", activo, datetime.now().isoformat(timespec="seconds")),
            )
        db.commit()
        flash("Usuario guardado correctamente", "success")
        return redirect(url_for("usuarios"))
    rows = db.execute("SELECT * FROM usuarios ORDER BY nombre").fetchall()
    return render_template("usuarios.html", rows=rows)


@app.route("/admin/<catalogo>", methods=["GET", "POST"])
@login_required
def admin_catalogo(catalogo):
    config = {
        "hoteles": {"table": "hoteles", "title": "Hoteles", "fields": ["nombre", "codigo"]},
        "departamentos": {"table": "departamentos", "title": "Departamentos", "fields": ["nombre"]},
        "campanas": {"table": "campanas", "title": "Campanas", "fields": ["nombre", "fecha_inicio", "fecha_fin"]},
    }.get(catalogo)
    if not config:
        return redirect(url_for("dashboard"))
    db = get_db()
    if request.method == "POST":
        fields = config["fields"]
        values = [request.form.get(field, "") for field in fields]
        activo_field = "activa" if catalogo == "campanas" else "activo"
        active = 1 if "activo" in request.form else 0
        row_id = request.form.get("id")
        if row_id:
            assignments = ", ".join([f"{field}=?" for field in fields] + [f"{activo_field}=?"])
            db.execute(f"UPDATE {config['table']} SET {assignments} WHERE id=?", values + [active, row_id])
        else:
            cols = ", ".join(fields + [activo_field])
            marks = ", ".join(["?"] * (len(fields) + 1))
            db.execute(f"INSERT INTO {config['table']} ({cols}) VALUES ({marks})", values + [active])
        db.commit()
        flash("Registro guardado correctamente", "success")
        return redirect(url_for("admin_catalogo", catalogo=catalogo))
    active_col = "activa" if catalogo == "campanas" else "activo"
    rows = db.execute(f"SELECT *, {active_col} AS activo_visible FROM {config['table']} ORDER BY nombre").fetchall()
    return render_template("catalogo.html", catalogo=catalogo, config=config, rows=rows)


@app.route("/export/<tipo>")
@login_required
def exportar(tipo):
    if tipo == "dashboard-pdf":
        return export_pdf()
    return export_excel(tipo)


def export_rows():
    where, params = build_filter_clause("e")
    sql_where = f"WHERE {where}" if where else ""
    return get_db().execute(
        f"""
        SELECT e.id encuesta_id, e.fecha_captura, c.nombre campana, h.nombre hotel, d.nombre departamento,
               e.sexo, e.escolaridad, e.estado_civil, e.antiguedad, e.edad_rango,
               p.numero, p.factor, p.texto pregunta, r.respuesta, r.valor,
               (SELECT respuesta FROM respuestas_abiertas WHERE encuesta_id = e.id AND numero = 72) apoyo_solucion,
               (SELECT respuesta FROM respuestas_abiertas WHERE encuesta_id = e.id AND numero = 73) ejemplo_trabajo,
               (SELECT respuesta FROM respuestas_abiertas WHERE encuesta_id = e.id AND numero = 74) actividades_personal,
               (SELECT respuesta FROM respuestas_abiertas WHERE encuesta_id = e.id AND numero = 75) comentarios_adicionales
        FROM encuestas e
        JOIN campanas c ON c.id = e.campana_id
        JOIN hoteles h ON h.id = e.hotel_id
        JOIN departamentos d ON d.id = e.departamento_id
        JOIN respuestas r ON r.encuesta_id = e.id
        JOIN preguntas p ON p.id = r.pregunta_id
        {sql_where}
        ORDER BY e.id, p.numero
        """,
        params,
    ).fetchall()


def export_excel(tipo):
    rows = export_rows()
    filename = f"{tipo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    if Workbook is None:
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(rows[0].keys() if rows else ["sin_datos"])
        for row in rows:
            writer.writerow([row[key] for key in row.keys()])
        return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")), as_attachment=True, download_name=filename.replace(".xlsx", ".csv"), mimetype="text/csv")
    wb = Workbook()
    ws = wb.active
    ws.title = "Detalle"
    if rows:
        ws.append(list(rows[0].keys()))
        for row in rows:
            ws.append([row[key] for key in row.keys()])
    else:
        ws.append(["Sin datos"])
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def export_pdf():
    data = dashboard_data()
    filename = f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    stream = io.BytesIO()
    if canvas is None:
        text = f"Dashboard Clima Laboral\nEncuestas: {data['total_encuestas']}\nIndice: {data['indice_general']}%"
        return send_file(io.BytesIO(text.encode("utf-8")), as_attachment=True, download_name=filename.replace(".pdf", ".txt"), mimetype="text/plain")
    pdf = canvas.Canvas(stream, pagesize=letter)
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(72, 740, "Dashboard Clima Laboral Optima")
    pdf.setFont("Helvetica", 11)
    pdf.drawString(72, 710, f"Total de encuestas: {data['total_encuestas']}")
    pdf.drawString(72, 690, f"Indice general: {data['indice_general']}%")
    pdf.drawString(72, 670, f"Mejor factor: {data['mejor_factor']}")
    pdf.drawString(72, 650, f"Factor mas bajo: {data['factor_bajo']}")
    y = 615
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(72, y, "Resultado por factor")
    pdf.setFont("Helvetica", 10)
    for label, value in zip(data["factores"]["labels"], data["factores"]["values"]):
        y -= 18
        pdf.drawString(90, y, f"{label}: {value}%")
        if y < 80:
            pdf.showPage()
            y = 740
    pdf.save()
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name=filename, mimetype="application/pdf")

with app.app_context():
    init_db()

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5004, debug=False)
