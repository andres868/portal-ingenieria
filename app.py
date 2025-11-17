import os
import sqlite3
import csv
from io import BytesIO, StringIO
from datetime import datetime, date
from pathlib import Path
from functools import wraps
from flask import (
    Flask, request, redirect, url_for, render_template, flash, send_from_directory,
    make_response, session
)
from werkzeug.utils import secure_filename
import smtplib
import mimetypes
from email.message import EmailMessage
try:
    import pythoncom  # Inicializa COM por hilo cuando usamos Outlook
except Exception:
    pythoncom = None
from jinja2 import DictLoader
import logging
from logging.handlers import RotatingFileHandler

# ------------------------------
# Configuración básica
# ------------------------------
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "tickets.db"
UPLOAD_FOLDER = BASE_DIR / "uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf"}

# Correos / Envío
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.telecom.local")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")
MAIL_FROM = os.getenv("MAIL_FROM", "noreply@telecom.com.ar")
MAIL_CC_ON_CLOSE = os.getenv("MAIL_CC_ON_CLOSE", "iga-notify@telecom.com.ar")  # múltiples separados por coma
USE_OUTLOOK = os.getenv("USE_OUTLOOK", "1").lower() in ("1", "true", "yes", "y", "on")

# Sistema externo (IGA/JIRA/Remedy/etc.)
EXTERNAL_SYSTEM_NAME = os.getenv("EXTERNAL_SYSTEM_NAME", "IGA")

# Adjuntos por correo
ENABLE_CREATE_ATTACH_PDF = os.getenv("ENABLE_CREATE_ATTACH_PDF", "1").lower() in ("1","true","yes","y","on")
ENABLE_CLOSE_ATTACH_PDF  = os.getenv("ENABLE_CLOSE_ATTACH_PDF",  "1").lower() in ("1","true","yes","y","on")

# Protección admin / portal
ADMIN_PASSWORD = "admin123"
PORTAL_PASSWORD = os.getenv("PORTAL_PASSWORD", "portal123")

# Constantes
PRIORITIES = ["Urgente", "Normal", "Baja"]

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "secret-key-dev")
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)

# ------------------------------
# Logging a archivo portal.log
# ------------------------------
LOG_PATH = BASE_DIR / "portal.log"
logger = logging.getLogger("portal")
logger.setLevel(logging.INFO)
_handler = RotatingFileHandler(LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
_handler.setFormatter(_formatter)
logger.addHandler(_handler)

# ------------------------------
# TEMPLATES en memoria (DictLoader)
# ------------------------------
TEMPLATES = {
    "layout.html": r"""
    <!doctype html>
    <html lang=es>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>{{ title or 'Portal Tickets Ingeniería' }}</title>
      <link href="https://cdn.jsdelivr.net/npm/bootswatch@5.3.3/dist/lux/bootstrap.min.css" rel="stylesheet">
      <style>
        .container { max-width: 1160px; }
        .card { border-radius: 1rem; }
        .badge-status { font-size: .9rem; }
        .pointer { cursor: pointer; }
        code { background: #f6f8fa; padding: .2rem .4rem; border-radius: .25rem; }
      </style>
    </head>
    <body>
      <nav class="navbar navbar-expand-lg navbar-dark bg-primary mb-4">
        <div class="container">
          <a class="navbar-brand" href="{{ url_for('home') }}">Portal Tickets Ingeniería</a>
          <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarsExample" aria-controls="navbarsExample" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
          </button>
          <div class="collapse navbar-collapse" id="navbarsExample">
            <ul class="navbar-nav me-auto">
              <li class="nav-item"><a class="nav-link" href="{{ url_for('new_ticket') }}">Nuevo Ticket</a></li>
              <li class="nav-item"><a class="nav-link" href="{{ url_for('search') }}">Buscar</a></li>
              <li class="nav-item dropdown">
                <a class="nav-link dropdown-toggle" data-bs-toggle="dropdown" href="#">Admin</a>
                <ul class="dropdown-menu">
                  <li><a class="dropdown-item" href="{{ url_for('admin_types') }}">Tipos de Modernización</a></li>
                  <li><a class="dropdown-item" href="{{ url_for('admin_assignees') }}">Responsables</a></li>
                </ul>
              </li>
            </ul>
            <ul class="navbar-nav ms-auto">
              {% if session.get('logged_in') %}
                <li class="nav-item"><a class="nav-link" href="{{ url_for('logout') }}">Salir</a></li>
              {% else %}
                <li class="nav-item"><a class="nav-link" href="{{ url_for('login') }}">Ingresar</a></li>
              {% endif %}
            </ul>
          </div>
        </div>
      </nav>
      <div class="container">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            {% for category, message in messages %}
              <div class="alert alert-{{ category }} alert-dismissible fade show" role="alert">
                {{ message|safe }}
                <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
              </div>
            {% endfor %}
          {% endif %}
        {% endwith %}
        {% block content %}{% endblock %}
      </div>
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    </body>
    </html>
    """,
    "login.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <div class="row justify-content-center">
        <div class="col-md-5">
          <div class="card shadow-sm p-4">
            <h4 class="mb-3">Ingresar</h4>
            <form method="post">
              <div class="mb-3">
                <label class="form-label">Password</label>
                <input type="password" class="form-control" name="password" placeholder="••••••" required>
              </div>
              <div class="d-grid gap-2">
                <button class="btn btn-primary" type="submit">Entrar</button>
              </div>
              <p class="text-muted mt-3 mb-0">* Placeholder de SSO/LDAP. Para productivo, integrar con directorio corporativo.</p>
            </form>
          </div>
        </div>
      </div>
    {% endblock %}
    """,
    "home.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <div class="d-flex align-items-center mb-3">
        <h3 class="me-3">Resumen</h3>
        <a class="btn btn-sm btn-primary" href="{{ url_for('new_ticket') }}">Crear ticket</a>
        <div class="ms-auto d-flex gap-2">
          <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('export_csv') }}">Exportar CSV</a>
          <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('export_xlsx') }}">Exportar Excel</a>
        </div>
      </div>

      <div class="row g-3">
        {% for card in summary_cards %}
        <div class="col-md-4">
          <div class="card shadow-sm">
            <div class="card-body">
              <div class="d-flex justify-content-between align-items-center">
                <h5 class="card-title mb-0">{{ card.title }}</h5>
                <span class="badge bg-secondary">{{ card.count }}</span>
              </div>
              <p class="text-muted mt-2 mb-0">{{ card.desc }}</p>
            </div>
          </div>
        </div>
        {% endfor %}
      </div>

      <hr class="my-4">
      <h4 class="mb-3">Últimos tickets</h4>
      <div class="list-group">
        {% for t in last_tickets %}
          <a class="list-group-item list-group-item-action" href="{{ url_for('ticket_detail', ticket_id=t['id']) }}">
            <div class="d-flex w-100 justify-content-between">
              <h5 class="mb-1">#{{ t['id'] }} · {{ t['site_name'] }}</h5>
              <small class="text-muted">{{ t['created_at'] }}</small>
            </div>
            <small>{{ t['modernization_type_name'] or '—' }} · {{ t['priority'] }} · {{ t['assignee_name'] }} · Estado: {{ t['status'] }}</small>
          </a>
        {% else %}
          <div class="text-muted">No hay tickets aún.</div>
        {% endfor %}
      </div>
    {% endblock %}
    """,
    "new_ticket.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <h3 class="mb-3">Nuevo Ticket</h3>
      <form class="card p-4 shadow-sm" method="post" enctype="multipart/form-data">
        <div class="row g-3">
          <div class="col-md-6">
            <label class="form-label">Nombre del sitio</label>
            <input required type="text" name="site_name" class="form-control" placeholder="Ej: AMBA_UTN_MEDRANO" />
          </div>
          <div class="col-md-6">
            <label class="form-label">Tipo de Modernización</label>
            <div class="input-group">
              <select required name="modernization_type_id" class="form-select">
                <option value="" disabled selected>Seleccionar…</option>
                {% for mt in modernization_types %}
                  <option value="{{ mt['id'] }}">{{ mt['name'] }}</option>
                {% endfor %}
              </select>
              <a href="{{ url_for('admin_types') }}" class="btn btn-outline-secondary">+ tipos</a>
            </div>
          </div>
          <div class="col-md-4">
            <label class="form-label">Fecha de solicitud</label>
            <input required type="date" name="request_date" class="form-control" />
          </div>
          <div class="col-md-4">
            <label class="form-label">Prioridad</label>
            <select required name="priority" class="form-select">
              {% for p in priorities %}
                <option value="{{ p }}">{{ p }}</option>
              {% endfor %}
            </select>
          </div>
          <div class="col-md-4">
            <label class="form-label">Asignado a</label>
            <div class="input-group">
              <select required name="assignee_id" class="form-select">
                <option value="" disabled selected>Seleccionar…</option>
                {% for a in assignees %}
                  <option value="{{ a['id'] }}">{{ a['name'] }} ({{ a['email'] or 'sin email' }})</option>
                {% endfor %}
              </select>
              <a href="{{ url_for('admin_assignees') }}" class="btn btn-outline-secondary">+ responsables</a>
            </div>
          </div>
          <div class="col-md-8">
            <label class="form-label">Email del requirente (recibe confirmación)</label>
            <input required type="email" name="creator_email" class="form-control" placeholder="usuario@telecom.com.ar" />
          </div>
          <div class="col-md-4">
            <label class="form-label">Adjuntar Ingeniería (PDF)</label>
            <input required class="form-control" type="file" name="pdf_file" accept="application/pdf" />
          </div>
        </div>
        <div class="mt-4 d-flex gap-2">
          <button class="btn btn-primary" type="submit">Crear ticket</button>
          <a class="btn btn-light" href="{{ url_for('home') }}">Cancelar</a>
        </div>
      </form>
    {% endblock %}
    """,
    "ticket_detail.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h3>Ticket #{{ t['id'] }} · {{ t['site_name'] }}</h3>
        <span class="badge bg-{% if t['status']=='Cerrado' %}success{% else %}warning{% endif %} badge-status">{{ t['status'] }}</span>
      </div>
      <div class="card shadow-sm p-3 mb-3">
        <div class="row g-3">
          <div class="col-md-6">
            <div><strong>Tipo de Modernización:</strong> {{ t['modernization_type_name'] or '—' }}</div>
            <div><strong>Prioridad:</strong> {{ t['priority'] }}</div>
            <div><strong>Asignado a:</strong> {{ t['assignee_name'] }} ({{ t['assignee_email'] or 'sin email' }})</div>
          </div>
          <div class="col-md-6">
            <div><strong>Fecha solicitud:</strong> {{ t['request_date'] }}</div>
            <div><strong>Días transcurridos:</strong> {{ t['days_passed'] }}</div>
            <div><strong>Creado por:</strong> {{ t['creator_email'] }}</div>
          </div>
        </div>
      </div>
      <div class="card shadow-sm p-3 mb-3">
        <div class="row g-3">
          <div class="col-md-4">
            <label class="form-label">N° de caso {{ EXTSYS }}</label>
            <input form="closeForm" class="form-control" type="text" name="iga_case_number" value="{{ t['iga_case_number'] or '' }}" {% if t['status']=='Cerrado' %}disabled{% endif %}>
          </div>
          <div class="col-md-8">
            <label class="form-label">Link {{ EXTSYS }}</label>
            <input form="closeForm" class="form-control" type="url" name="iga_link" value="{{ t['iga_link'] or '' }}" {% if t['status']=='Cerrado' %}disabled{% endif %}>
            {% if t['iga_link'] %}<small><a href="{{ t['iga_link'] }}" target="_blank">Abrir {{ EXTSYS }}</a></small>{% endif %}
          </div>
        </div>
      </div>

      <!-- Password admin solo para eliminar ticket -->
      <div class="mb-2">
        <label class="form-label">Password admin (solo para eliminar ticket)</label>
        <input type="password" class="form-control form-control-sm" id="adminPwdInput" placeholder="********">
        <small class="text-muted">Solo quien conoce la contraseña de administrador puede eliminar tickets.</small>
      </div>

      <div class="d-flex gap-2 mb-3">
        {% if t['pdf_filename'] %}
          <a class="btn btn-outline-secondary" href="{{ url_for('download_pdf', filename=t['pdf_filename']) }}">Descargar PDF</a>
        {% endif %}
        {% if t['status'] != 'Cerrado' %}
          <form id="closeForm" method="post" action="{{ url_for('close_ticket', ticket_id=t['id']) }}" onsubmit="return confirm('¿Cerrar el caso como COMPLETADO?');">
            <button class="btn btn-success" type="submit">Cerrar caso (Completado)</button>
          </form>
        {% endif %}

        <form id="deleteForm"
              method="post"
              action="{{ url_for('delete_ticket', ticket_id=t['id']) }}"
              onsubmit="return confirm('⚠ Esta acción eliminará el ticket y su PDF. ¿Confirmás?');">
          <input type="hidden" name="admin_password" id="adminPwdHidden">
          <button class="btn btn-danger" type="submit">Eliminar ticket</button>
        </form>

        <a class="btn btn-light" href="{{ url_for('search') }}">Volver</a>
      </div>

      <script>
        (function() {
          const input = document.getElementById('adminPwdInput');
          const hidden = document.getElementById('adminPwdHidden');
          const deleteForm = document.getElementById('deleteForm');

          if (deleteForm && input && hidden) {
            deleteForm.addEventListener('submit', function() {
              hidden.value = input.value;
            });
          }
        })();
      </script>
    {% endblock %}
    
    """,
    "search.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <h3 class="mb-3">Buscar Tickets</h3>
      <form class="row g-2 mb-3" method="get">
        <div class="col-md-3">
          <input type="text" class="form-control" name="q" placeholder="#ticket o Sitio" value="{{ request.args.get('q','') }}">
        </div>
        <div class="col-md-2">
          <select class="form-select" name="status">
            <option value="">Estado (todos)</option>
            <option value="Abierto" {{ 'selected' if request.args.get('status')=='Abierto' }}>Abierto</option>
            <option value="Cerrado" {{ 'selected' if request.args.get('status')=='Cerrado' }}>Cerrado</option>
          </select>
        </div>
        <div class="col-md-2">
          <select class="form-select" name="priority">
            <option value="">Prioridad (todas)</option>
            {% for p in priorities %}
              <option value="{{ p }}" {{ 'selected' if request.args.get('priority')==p }}>{{ p }}</option>
            {% endfor %}
          </select>
        </div>
        <div class="col-md-2">
          <select class="form-select" name="assignee_id">
            <option value="">Responsable (todos)</option>
            {% for a in assignees %}
              <option value="{{ a['id'] }}" {{ 'selected' if (request.args.get('assignee_id')|int) == a['id'] }}>{{ a['name'] }}</option>
            {% endfor %}
          </select>
        </div>
        <div class="col-md-3 d-grid gap-2 d-md-flex justify-content-md-end">
          <button class="btn btn-primary" type="submit">Buscar</button>
          <a class="btn btn-outline-secondary" href="{{ url_for('export_csv', **request.args) }}">CSV</a>
          <a class="btn btn-outline-secondary" href="{{ url_for('export_xlsx', **request.args) }}">Excel</a>
        </div>
      </form>

      <div class="list-group">
        {% for t in results %}
          <a class="list-group-item list-group-item-action" href="{{ url_for('ticket_detail', ticket_id=t['id']) }}">
            <div class="d-flex w-100 justify-content-between">
              <h5 class="mb-1">#{{ t['id'] }} · {{ t['site_name'] }}</h5>
              <small class="text-muted">{{ t['created_at'] }}</small>
            </div>
            <small>{{ t['modernization_type_name'] or '—' }} · {{ t['priority'] }} · {{ t['assignee_name'] }} · Estado: {{ t['status'] }}</small>
          </a>
        {% else %}
          <div class="text-muted">Sin resultados.</div>
        {% endfor %}
      </div>
    {% endblock %}
    """,
    "admin_types.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <h3 class="mb-3">Administrar Tipos de Modernización</h3>
      <form class="card p-3 shadow-sm mb-3" method="post">
        <div class="row g-2 align-items-end">
          <div class="col-md-4">
            <label class="form-label">Password admin</label>
            <input required class="form-control" type="password" name="password" placeholder="********">
          </div>
          <div class="col-md-6">
            <label class="form-label">Nuevo tipo</label>
            <input class="form-control" type="text" name="new_type" placeholder="Ej: SWAP 4G→5G, Sectorización, Cambio AAU…">
          </div>
          <div class="col-md-2 d-grid">
            <button class="btn btn-primary" type="submit">Agregar</button>
          </div>
        </div>
      </form>

      <div class="card p-3 shadow-sm">
        <h5>Tipos existentes</h5>
        <ul class="list-group">
          {% for t in modernization_types %}
            <li class="list-group-item d-flex justify-content-between align-items-center">
              {{ t['name'] }}
              <form method="post" action="{{ url_for('delete_type', type_id=t['id']) }}" onsubmit="return confirm('¿Eliminar tipo?');">
                <input type="hidden" name="password" value="">
                <button class="btn btn-sm btn-outline-danger" type="submit">Eliminar</button>
              </form>
            </li>
          {% else %}
            <li class="list-group-item text-muted">No hay tipos aún.</li>
          {% endfor %}
        </ul>
        <p class="text-muted mt-2 mb-0">* Para eliminar, ingresá primero la contraseña arriba y luego clic en "Eliminar" (el campo se completa por JS).</p>
      </div>

      <script>
        const pwdInput = document.querySelector('input[name="password"]');
        document.querySelectorAll('form[action*="/admin/types/delete/"]').forEach(f => {
          f.addEventListener('submit', (e) => {
            const hidden = f.querySelector('input[type="hidden"][name="password"]');
            hidden.value = pwdInput.value;
          })
        })
      </script>
    {% endblock %}
    """,
    "admin_assignees.html": r"""
    {% extends 'layout.html' %}
    {% block content %}
      <h3 class="mb-3">Administrar Responsables</h3>
      <form class="card p-3 shadow-sm mb-3" method="post">
        <div class="row g-2 align-items-end">
          <div class="col-md-3">
            <label class="form-label">Password admin</label>
            <input required class="form-control" type="password" name="password" placeholder="********">
          </div>
          <div class="col-md-4">
            <label class="form-label">Nombre</label>
            <input class="form-control" type="text" name="name" placeholder="Ej: Andres Martinez">
          </div>
          <div class="col-md-5">
            <label class="form-label">Email</label>
            <input class="form-control" type="email" name="email" placeholder="usuario@telecom.com.ar">
          </div>
          <div class="col-12 d-grid mt-2">
            <button class="btn btn-primary" type="submit">Agregar / Actualizar</button>
          </div>
        </div>
      </form>

      <div class="card p-3 shadow-sm">
        <h5>Responsables</h5>
        <ul class="list-group">
          {% for a in assignees %}
            <li class="list-group-item d-flex justify-content-between align-items-center">
              <div>
                <strong>{{ a['name'] }}</strong>
                <div class="text-muted">{{ a['email'] or 'sin email' }}</div>
              </div>
              <form method="post" action="{{ url_for('delete_assignee', assignee_id=a['id']) }}" onsubmit="return confirm('¿Eliminar responsable?');">
                <input type="hidden" name="password" value="">
                <button class="btn btn-sm btn-outline-danger" type="submit">Eliminar</button>
              </form>
            </li>
          {% else %}
            <li class="list-group-item text-muted">No hay responsables aún.</li>
          {% endfor %}
        </ul>
        <p class="text-muted mt-2 mb-0">* Para eliminar, ingresá primero la contraseña arriba y luego clic en "Eliminar" (el campo se completa por JS).</p>
      </div>

      <script>
        const pwdInput = document.querySelector('input[name="password"]');
        document.querySelectorAll('form[action*="/admin/assignees/delete/"]').forEach(f => {
          f.addEventListener('submit', (e) => {
            const hidden = f.querySelector('input[type="hidden"][name="password"]');
            hidden.value = pwdInput.value;
          })
        })
      </script>
    {% endblock %}
    """,
}

app.jinja_loader = DictLoader(TEMPLATES)
# Exponer nombre de sistema externo a templates
app.jinja_env.globals['EXTSYS'] = EXTERNAL_SYSTEM_NAME

# ------------------------------
# Utilidades
# ------------------------------

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login', next=request.path))
        return fn(*args, **kwargs)
    return wrapper


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS modernization_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS assignees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            email TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS tickets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_name TEXT NOT NULL,
            modernization_type_id INTEGER,
            request_date TEXT NOT NULL,
            priority TEXT NOT NULL,
            assignee_id INTEGER NOT NULL,
            creator_email TEXT NOT NULL,
            pdf_filename TEXT,
            iga_case_number TEXT,
            iga_link TEXT,
            status TEXT NOT NULL DEFAULT 'Abierto',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (modernization_type_id) REFERENCES modernization_types(id),
            FOREIGN KEY (assignee_id) REFERENCES assignees(id)
        )
        """
    )
    conn.commit()
    conn.close()


def get_type_name(conn, type_id):
    if not type_id:
        return None
    cur = conn.cursor()
    cur.execute("SELECT name FROM modernization_types WHERE id=?", (type_id,))
    row = cur.fetchone()
    return row["name"] if row else None


def get_assignee(conn, assignee_id):
    cur = conn.cursor()
    cur.execute("SELECT id, name, email FROM assignees WHERE id=?", (assignee_id,))
    row = cur.fetchone()
    return dict(row) if row else None


def send_mail(subject: str, to: str | list[str], body_html: str, cc: str | list[str] | None = None, attachments: list[str] | None = None):
    logger.info(f"[MAIL] preparing subject={subject} to={to} cc={cc} attachments={attachments}")
    recipients = [to] if isinstance(to, str) else list(to)
    cc_list = []
    if cc:
        cc_list = [cc] if isinstance(cc, str) else list(cc)
    attachments = attachments or []

    # Intento Outlook
    if USE_OUTLOOK:
        try:
            # Inicializa COM en el hilo actual (cada request puede usar un hilo distinto)
            if pythoncom:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            import win32com.client as win32
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # olMailItem
            mail.To = "; ".join([r for r in recipients if r])
            if cc_list:
                mail.CC = "; ".join([c for c in cc_list if c])
            mail.Subject = subject
            mail.HTMLBody = body_html
            for path in attachments:
                try:
                    if path and os.path.exists(path):
                        mail.Attachments.Add(str(path))
                except Exception as e:
                    logger.warning(f"[MAIL] No se pudo adjuntar {path} a Outlook: {e}")
            # Seleccionar cuenta si coincide con MAIL_FROM
            try:
                if MAIL_FROM:
                    session_out = outlook.Session
                    for account in session_out.Accounts:
                        try:
                            smtp = account.SmtpAddress
                        except Exception:
                            smtp = None
                        if smtp and smtp.lower() == MAIL_FROM.lower():
                            mail._oleobj_.Invoke(64209, 0, 8, 0, account)  # PR_SEND_USING_ACCOUNT
                            break
            except Exception as e2:
                logger.warning(f"[MAIL] No se pudo seleccionar la cuenta {MAIL_FROM}: {e2}")
            mail.Send()
            logger.info("[MAIL] Sent via Outlook")
            return
        except ImportError:
            logger.warning("[MAIL] pywin32 no instalado; usando SMTP fallback.")
        except Exception as e:
            logger.warning(f"[MAIL] Error enviando por Outlook: {e}; usando SMTP fallback.")

    # SMTP fallback
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join([r for r in recipients if r])
    if cc_list:
        msg["Cc"] = ", ".join([c for c in cc_list if c])
    msg.set_content("Este mensaje requiere un cliente compatible con HTML.")
    msg.add_alternative(body_html, subtype="html")

    for path in attachments:
        try:
            if path and os.path.exists(path):
                ctype, encoding = mimetypes.guess_type(str(path))
                if ctype is None:
                    ctype = "application/octet-stream"
                maintype, subtype = ctype.split("/", 1)
                with open(path, "rb") as f:
                    msg.add_attachment(
                        f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(path)
                    )
        except Exception as e:
            logger.warning(f"[MAIL] No se pudo adjuntar {path}: {e}")

    # Si no hay SMTP configurado, omitimos fallback
    if not SMTP_HOST:
        logger.warning("[MAIL] SMTP no configurado (SMTP_HOST vacío); se omite fallback. Email NO enviado.")
        return
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls()
            if SMTP_USER:
                s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
            logger.info("[MAIL] Sent via SMTP")
    except Exception as e:
        logger.warning(f"[MAIL] Error enviando por SMTP: {e}")


def human_date(d: str) -> str:
    try:
        return datetime.strptime(d, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return d


def query_tickets(conn, q=None, status=None, priority=None, assignee_id=None):
    sql = (
        "SELECT t.*, mt.name AS modernization_type_name, "
        "a.name AS assignee_name, a.email AS assignee_email "
        "FROM tickets t "
        "LEFT JOIN modernization_types mt ON mt.id = t.modernization_type_id "
        "LEFT JOIN assignees a ON a.id = t.assignee_id "
        "WHERE 1=1 "
    )
    params = []
    if q:
        if q.isdigit():
            sql += "AND t.id = ? "
            params.append(int(q))
        else:
            sql += "AND t.site_name LIKE ? "
            params.append(f"%{q}%")
    if status:
        sql += "AND t.status = ? "
        params.append(status)
    if priority:
        sql += "AND t.priority = ? "
        params.append(priority)
    if assignee_id:
        sql += "AND t.assignee_id = ? "
        params.append(int(assignee_id))
    sql += "ORDER BY t.id DESC"
    cur = conn.cursor()
    cur.execute(sql, params)
    return [dict(r) for r in cur.fetchall()]

# ------------------------------
# Autenticación básica (placeholder LDAP)
# ------------------------------
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        pwd = request.form.get('password','')
        if pwd == PORTAL_PASSWORD:
            session['logged_in'] = True
            flash('Bienvenido.', 'success')
            next_url = request.args.get('next') or url_for('home')
            return redirect(next_url)
        flash('Password incorrecto.', 'warning')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    session.clear()
    flash('Sesión cerrada.', 'info')
    return redirect(url_for('login'))

# ------------------------------
# Rutas principales
# ------------------------------
@app.route("/")
@login_required
def home():
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM tickets WHERE status='Abierto'")
    open_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM tickets WHERE status='Cerrado'")
    closed_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM tickets")
    total_count = cur.fetchone()[0]

    cur.execute(
        """
        SELECT t.*, mt.name as modernization_type_name, a.name as assignee_name
        FROM tickets t
        LEFT JOIN modernization_types mt ON mt.id = t.modernization_type_id
        LEFT JOIN assignees a ON a.id = t.assignee_id
        ORDER BY t.id DESC LIMIT 10
        """
    )
    rows = cur.fetchall()
    last_tickets = []
    for r in rows:
        last_tickets.append({
            "id": r["id"],
            "site_name": r["site_name"],
            "created_at": datetime.fromisoformat(r["created_at"]).strftime("%d/%m/%Y %H:%M"),
            "modernization_type_name": r["modernization_type_name"],
            "priority": r["priority"],
            "assignee_name": r["assignee_name"],
            "status": r["status"],
        })

    conn.close()

    summary_cards = [
        {"title": "Abiertos", "count": open_count, "desc": "Tickets en curso"},
        {"title": "Cerrados", "count": closed_count, "desc": "Tickets completados"},
        {"title": "Total", "count": total_count, "desc": "Acumulado histórico"},
    ]

    return render_template("home.html", title="Inicio", summary_cards=summary_cards, last_tickets=last_tickets)


@app.route("/tickets/new", methods=["GET", "POST"])
@login_required
def new_ticket():
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM modernization_types ORDER BY name ASC")
    modernization_types = [dict(row) for row in cur.fetchall()]
    cur.execute("SELECT id, name, email FROM assignees ORDER BY name ASC")
    assignees = [dict(row) for row in cur.fetchall()]

    if request.method == "POST":
        site_name = request.form.get("site_name", "").strip()
        modernization_type_id = request.form.get("modernization_type_id")
        request_date = request.form.get("request_date")
        priority = request.form.get("priority")
        assignee_id = request.form.get("assignee_id")
        creator_email = request.form.get("creator_email", "").strip()
        file = request.files.get("pdf_file")

        if not site_name or not request_date or not priority or not assignee_id or not creator_email or not file:
            flash("Completá todos los campos.", "warning")
            return render_template("new_ticket.html", modernization_types=modernization_types, priorities=PRIORITIES, assignees=assignees)

        if not allowed_file(file.filename):
            flash("El archivo debe ser PDF.", "warning")
            return render_template("new_ticket.html", modernization_types=modernization_types, priorities=PRIORITIES, assignees=assignees)

        filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{secure_filename(file.filename)}"
        save_path = UPLOAD_FOLDER / filename
        file.save(save_path)

        now_iso = datetime.now().isoformat(timespec='seconds')
        cur.execute(
            """
            INSERT INTO tickets (site_name, modernization_type_id, request_date, priority, assignee_id, creator_email, pdf_filename, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'Abierto', ?, ?)
            """,
            (
                site_name,
                int(modernization_type_id) if modernization_type_id else None,
                request_date,
                priority,
                int(assignee_id),
                creator_email,
                filename,
                now_iso,
                now_iso,
            ),
        )
        new_ticket_id = cur.lastrowid
        conn.commit()

        # ------- Notificaciones por email (creación) -------
        try:
            assignee = get_assignee(conn, int(assignee_id))
            type_name = get_type_name(conn, int(modernization_type_id)) if modernization_type_id else None
            subject = f"[Portal Ingeniería] Nuevo Ticket #{new_ticket_id} — {site_name}"
            recipients = [creator_email]
            if assignee and assignee.get('email'):
                recipients.append(assignee['email'])
            ticket_url = url_for('ticket_detail', ticket_id=new_ticket_id, _external=True)
            body_html = f"""
            <h3>Se creó un ticket de ingeniería</h3>
            <p><b>Ticket:</b> #{new_ticket_id}<br>
            <b>Sitio:</b> {site_name}<br>
            <b>Tipo:</b> {type_name or '—'}<br>
            <b>Prioridad:</b> {priority}<br>
            <b>Asignado a:</b> {(assignee['name'] if assignee else '—')}<br>
            <b>Fecha solicitud:</b> {human_date(request_date)}<br>
            <b>Creado por:</b> {creator_email}</p>
            <p><a href="{ticket_url}">Ver detalles del ticket</a></p>
            """
            attachments = [str(save_path)] if ENABLE_CREATE_ATTACH_PDF else []
            send_mail(subject, to=recipients, body_html=body_html, attachments=attachments)
        except Exception as e:
            logger.warning(f"[MAIL] Error envío creación #{new_ticket_id}: {e}")
            flash(f"Ticket #{new_ticket_id} creado, pero hubo un error al enviar la notificación por email.", "warning")

        flash(f'Ticket <a href="{url_for("ticket_detail", ticket_id=new_ticket_id)}">#{new_ticket_id}</a> creado con éxito. Se envió una notificación.', "success")
        conn.close()
        return redirect(url_for("home"))

    conn.close()
    return render_template("new_ticket.html", title="Nuevo Ticket", modernization_types=modernization_types, priorities=PRIORITIES, assignees=assignees)


@app.route("/tickets/<int:ticket_id>")
@login_required
def ticket_detail(ticket_id: int):
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT t.*, mt.name as modernization_type_name, a.name as assignee_name, a.email as assignee_email
        FROM tickets t
        LEFT JOIN modernization_types mt ON mt.id = t.modernization_type_id
        LEFT JOIN assignees a ON a.id = t.assignee_id
        WHERE t.id=?
        """,
        (ticket_id,),
    )
    r = cur.fetchone()
    conn.close()
    if not r:
        flash("Ticket no encontrado.", "warning")
        return redirect(url_for("search"))

    try:
        req_d = datetime.strptime(r["request_date"], "%Y-%m-%d").date()
        days_passed = (date.today() - req_d).days
    except Exception:
        days_passed = "—"

    t = dict(r)
    t["days_passed"] = days_passed
    t["request_date"] = human_date(t["request_date"])
    t["created_at"] = datetime.fromisoformat(t["created_at"]).strftime("%d/%m/%Y %H:%M")

    return render_template("ticket_detail.html", t=t)


@app.route("/tickets/<int:ticket_id>/close", methods=["POST"])
@login_required
def close_ticket(ticket_id: int):
    iga_case = request.form.get("iga_case_number", "").strip() or None
    iga_link = request.form.get("iga_link", "").strip() or None

    conn = db_connect()
    cur = conn.cursor()

    data = query_tickets(conn, q=str(ticket_id))
    if not data:
        flash("Ticket no encontrado.", "danger")
        conn.close()
        return redirect(url_for("search"))
    t = data[0]

    cur.execute(
        "UPDATE tickets SET status='Cerrado', iga_case_number=?, iga_link=?, updated_at=? WHERE id=?",
        (iga_case, iga_link, datetime.now().isoformat(timespec='seconds'), ticket_id),
    )
    conn.commit()

    # ------- Notificaciones por email (cierre) -------
    try:
        subject = f"[Portal Ingeniería] Ticket #{ticket_id} CERRADO — {t['site_name']}"
        recipients = [t['creator_email']]
        if t.get('assignee_email'):
            recipients.append(t['assignee_email'])
        cc_list = [email.strip() for email in MAIL_CC_ON_CLOSE.split(',') if email.strip()]
        attachments = []
        if ENABLE_CLOSE_ATTACH_PDF and t.get('pdf_filename'):
            pdf_path = UPLOAD_FOLDER / t['pdf_filename']
            if pdf_path.exists():
                attachments.append(str(pdf_path))
        ticket_url = url_for('ticket_detail', ticket_id=ticket_id, _external=True)
        body_html = f"""
        <h3>El ticket fue cerrado (Completado)</h3>
        <p><b>Ticket:</b> #{ticket_id}<br>
        <b>Sitio:</b> {t['site_name']}<br>
        <b>Asignado a:</b> {t.get('assignee_name') or '—'}<br>
        <b>N° Caso {EXTERNAL_SYSTEM_NAME}:</b> {iga_case or 'No informado'}<br>
        <b>Link {EXTERNAL_SYSTEM_NAME}:</b> {f'<a href="{iga_link}">Abrir link</a>' if iga_link else 'No informado'}</p>
        <p><a href="{ticket_url}">Ver detalles del ticket</a></p>
        """
        send_mail(subject, to=recipients, cc=cc_list, body_html=body_html, attachments=attachments)
    except Exception as e:
        logger.warning(f"[MAIL] Error envío cierre #{ticket_id}: {e}")
        flash(f"Ticket #{ticket_id} cerrado, pero hubo un error al enviar la notificación por email.", "warning")
    else:
        flash(f"Ticket #{ticket_id} cerrado con éxito. Se envió una notificación.", "success")

    conn.close()
    return redirect(url_for("ticket_detail", ticket_id=ticket_id))



@app.post("/tickets/<int:ticket_id>/delete")
@login_required
def delete_ticket(ticket_id: int):
    """Elimina un ticket y su PDF asociado. Protegido por password de administrador."""
    password = request.form.get("admin_password", "")
    if password != ADMIN_PASSWORD:
        flash("Password admin incorrecta.", "warning")
        return redirect(url_for("ticket_detail", ticket_id=ticket_id))

    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT pdf_filename, site_name FROM tickets WHERE id=?", (ticket_id,))
    row = cur.fetchone()

    if not row:
        conn.close()
        flash("Ticket no encontrado.", "warning")
        return redirect(url_for("home"))

    pdf_filename = row["pdf_filename"]
    site_name = row["site_name"]

    cur.execute("DELETE FROM tickets WHERE id=?", (ticket_id,))
    conn.commit()
    conn.close()

    # Borrar archivo PDF si existe
    if pdf_filename:
        pdf_path = UPLOAD_FOLDER / pdf_filename
        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except Exception as e:
            logger.warning(f"[DELETE] No se pudo borrar el archivo {pdf_path}: {e}")

    flash(f"Ticket #{ticket_id} ({site_name}) eliminado definitivamente.", "success")
    return redirect(url_for("home"))

@app.route("/search")
@login_required
def search():
    q = request.args.get("q", "").strip()
    status = request.args.get("status") or None
    priority = request.args.get("priority") or None
    assignee_id = request.args.get("assignee_id") or None

    conn = db_connect()
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM modernization_types ORDER BY name ASC")
    modernization_types = [dict(row) for row in cur.fetchall()]
    cur.execute("SELECT id, name FROM assignees ORDER BY name ASC")
    assignees = [dict(row) for row in cur.fetchall()]

    results = query_tickets(conn, q=q, status=status, priority=priority, assignee_id=assignee_id)
    for r in results:
        r["created_at"] = datetime.fromisoformat(r["created_at"]).strftime("%d/%m/%Y %H:%M")

    conn.close()

    return render_template("search.html", results=results, priorities=PRIORITIES, assignees=assignees)


@app.route('/uploads/<path:filename>')
@login_required
def download_pdf(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


# ---------- Admin: Tipos ----------
@app.route('/admin/types', methods=['GET', 'POST'])
@login_required
def admin_types():
    conn = db_connect()
    cur = conn.cursor()

    if request.method == 'POST':
        password = request.form.get('password', '')
        new_type = (request.form.get('new_type') or '').strip()
        if password != ADMIN_PASSWORD:
            flash('Password incorrecto.', 'warning')
        else:
            if new_type:
                try:
                    cur.execute('INSERT INTO modernization_types(name) VALUES (?)', (new_type,))
                    conn.commit()
                    flash('Tipo agregado.', 'success')
                except sqlite3.IntegrityError:
                    flash('Ese tipo ya existe.', 'info')
            else:
                flash('Indicá un nombre de tipo.', 'warning')

    cur.execute('SELECT id, name FROM modernization_types ORDER BY name ASC')
    types = [dict(row) for row in cur.fetchall()]
    conn.close()
    return render_template('admin_types.html', modernization_types=types)


@app.post('/admin/types/delete/<int:type_id>')
@login_required
def delete_type(type_id: int):
    conn = db_connect()
    cur = conn.cursor()
    password = request.form.get('password', '')
    if password != ADMIN_PASSWORD:
        conn.close()
        flash('Password incorrecto.', 'warning')
        return redirect(url_for('admin_types'))

    cur.execute('DELETE FROM modernization_types WHERE id=?', (type_id,))
    conn.commit()
    conn.close()
    flash('Tipo eliminado.', 'success')
    return redirect(url_for('admin_types'))


# ---------- Admin: Responsables ----------
@app.route('/admin/assignees', methods=['GET','POST'])
@login_required
def admin_assignees():
    conn = db_connect()
    cur = conn.cursor()

    if request.method == 'POST':
        password = request.form.get('password','')
        name = (request.form.get('name') or '').strip()
        email = (request.form.get('email') or '').strip()
        if password != ADMIN_PASSWORD:
            flash('Password incorrecto.', 'warning')
        else:
            if name:
                try:
                    cur.execute('INSERT INTO assignees(name, email) VALUES(?, ?) ON CONFLICT(name) DO UPDATE SET email=excluded.email', (name, email))
                except sqlite3.OperationalError:
                    cur.execute('SELECT id FROM assignees WHERE name=?', (name,))
                    row = cur.fetchone()
                    if row:
                        cur.execute('UPDATE assignees SET email=? WHERE id=?', (email, row['id']))
                    else:
                        cur.execute('INSERT INTO assignees(name, email) VALUES(?, ?)', (name, email))
                conn.commit()
                flash('Responsable agregado/actualizado.', 'success')
            else:
                flash('Indicá el nombre.', 'warning')

    cur.execute('SELECT id, name, email FROM assignees ORDER BY name ASC')
    rows = cur.fetchall()
    assignees = [dict(r) for r in rows]
    conn.close()
    return render_template('admin_assignees.html', assignees=assignees)


@app.post('/admin/assignees/delete/<int:assignee_id>')
@login_required
def delete_assignee(assignee_id: int):
    conn = db_connect()
    cur = conn.cursor()
    password = request.form.get('password','')
    if password != ADMIN_PASSWORD:
        conn.close()
        flash('Password incorrecto.', 'warning')
        return redirect(url_for('admin_assignees'))
    cur.execute('DELETE FROM assignees WHERE id=?', (assignee_id,))
    conn.commit()
    conn.close()
    flash('Responsable eliminado.', 'success')
    return redirect(url_for('admin_assignees'))


# ---------- Exportaciones ----------

def _rows_for_export(filters: dict):
    conn = db_connect()
    rows = query_tickets(conn, **filters)
    conn.close()
    out = []
    for r in rows:
        out.append({
            'id': r['id'],
            'site_name': r['site_name'],
            'modernization_type': r.get('modernization_type_name'),
            'request_date': r['request_date'],
            'priority': r['priority'],
            'assignee': r.get('assignee_name'),
            'assignee_email': r.get('assignee_email'),
            'creator_email': r.get('creator_email'),
            'iga_case_number': r.get('iga_case_number'),
            'iga_link': r.get('iga_link'),
            'status': r['status'],
            'created_at': r['created_at'],
            'updated_at': r['updated_at'],
        })
    return out

@app.route('/export.csv')
@login_required
def export_csv():
    filters = {
        'q': request.args.get('q') or None,
        'status': request.args.get('status') or None,
        'priority': request.args.get('priority') or None,
        'assignee_id': request.args.get('assignee_id') or None,
    }
    rows = _rows_for_export(filters)
    si = StringIO()
    if rows:
        writer = csv.DictWriter(si, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    else:
        writer = csv.writer(si)
        writer.writerow(["Sin datos"])
    resp = make_response(si.getvalue())
    resp.headers['Content-Type'] = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = f"attachment; filename=\"tickets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv\""
    return resp

@app.route('/export.xlsx')
@login_required
def export_xlsx():
    filters = {
        'q': request.args.get('q') or None,
        'status': request.args.get('status') or None,
        'priority': request.args.get('priority') or None,
        'assignee_id': request.args.get('assignee_id') or None,
    }
    rows = _rows_for_export(filters)
    try:
        import pandas as pd
    except Exception:
        flash('Para exportar a Excel instalá pandas y openpyxl: <code>pip install pandas openpyxl</code>. Se descargará CSV.', 'warning')
        return redirect(url_for('export_csv', **request.args))

    df = pd.DataFrame(rows)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Tickets')
    resp = make_response(bio.getvalue())
    resp.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    resp.headers['Content-Disposition'] = f"attachment; filename=\"tickets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx\""
    return resp

# ---------- Debug de correo y logs (opcional) ----------
@app.get('/debug/mail-test')
@login_required
def debug_mail_test():
    if not os.getenv('ENABLE_DEBUG_MAIL'):
        return ("Debug mail deshabilitado. Setea ENABLE_DEBUG_MAIL=1", 403)
    to = request.args.get('to') or os.getenv('DEBUG_TO') or os.getenv('MAIL_FROM')
    if not to:
        return ("Falta parámetro 'to' o variable DEBUG_TO/MAIL_FROM", 400)
    body = """
    <h3>Prueba de correo</h3>
    <p>Si ves este mensaje, el envío por Outlook/SMTP está funcionando.</p>
    """
    send_mail("[Portal Ingeniería] Prueba de correo", to, body)
    return f"Enviado a {to}"

@app.get('/debug/log')
@login_required
def debug_log():
    if not os.getenv('ENABLE_DEBUG_MAIL'):
        return ("Debug log deshabilitado. Setea ENABLE_DEBUG_MAIL=1", 403)
    try:
        return send_from_directory(str(BASE_DIR), 'portal.log', as_attachment=True)
    except Exception as e:
        return (f'No se pudo leer portal.log: {e}', 500)

# ------------------------------
# Inicialización
# ------------------------------
def bootstrap_db():
    """Crea la base y carga semillas si hace falta."""
    init_db()
    conn = db_connect()
    cur = conn.cursor()
    # Tipos default
    cur.execute('SELECT COUNT(*) FROM modernization_types')
    if cur.fetchone()[0] == 0:
        cur.executemany('INSERT INTO modernization_types(name) VALUES (?)', [
            ("Swap 4G→5G",),
            ("Cambio AAU",),
            ("Sectorización",),
        ])
    # Responsables default
    cur.execute('SELECT COUNT(*) FROM assignees')
    if cur.fetchone()[0] == 0:
        cur.executemany('INSERT INTO assignees(name, email) VALUES (?, ?)', [
            ("Andres Martinez", "andres.martinez@telecom.com.ar"),
            ("Evangelina Ortiz", "evangelina.ortiz@telecom.com.ar"),
            ("Juan Herrero", "juan.herrero@telecom.com.ar"),
        ])
    conn.commit()
    conn.close()


# Ejecutar siempre que se importe el módulo (local y en Render)
bootstrap_db()


if __name__ == "__main__":
    # Solo para ejecución local directa
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5006")), debug=True, threaded=False, use_reloader=False)
