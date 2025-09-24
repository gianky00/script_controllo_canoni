from app import db
from datetime import datetime

class Setting(db.Model):
    __tablename__ = 'settings'
    key = db.Column(db.String(128), primary_key=True)
    value = db.Column(db.String(256))

    def __repr__(self):
        return f'<Setting {self.key}>'

class Task(db.Model):
    __tablename__ = 'tasks'
    id = db.Column(db.String(36), primary_key=True)
    name = db.Column(db.String(128), nullable=False)
    script_path = db.Column(db.String(256), nullable=False)
    schedule_data = db.Column(db.Text, nullable=False) # JSON string
    enabled = db.Column(db.Boolean, default=True, nullable=False)
    retry_count = db.Column(db.Integer, default=0)
    retry_delay = db.Column(db.Integer, default=5)
    notify_on_success = db.Column(db.Boolean, default=False)
    notify_on_failure = db.Column(db.Boolean, default=True)
    runs = db.relationship('Run', backref='task', lazy='dynamic', cascade="all, delete-orphan")

    def __repr__(self):
        return f'<Task {self.name}>'

class Run(db.Model):
    __tablename__ = 'runs'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    task_id = db.Column(db.String(36), db.ForeignKey('tasks.id'), nullable=False)
    start_time = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    end_time = db.Column(db.DateTime)
    duration_seconds = db.Column(db.Float)
    exit_code = db.Column(db.Integer)
    log_output = db.Column(db.Text)

    def __repr__(self):
        return f'<Run {self.id} for task {self.task_id}>'

class Timbratura(db.Model):
    __tablename__ = 'timbrature'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    sito = db.Column(db.String(128))
    reparto = db.Column(db.String(128))
    data = db.Column(db.Date)
    nome = db.Column(db.String(128))
    cognome = db.Column(db.String(128))
    ingresso = db.Column(db.Time)
    uscita = db.Column(db.Time)
    ingresso_contabile = db.Column(db.Time)
    uscita_contabile = db.Column(db.Time)
    ore_contabili = db.Column(db.Float)
    avvisi_sistema = db.Column(db.String(256))
    note_utente = db.Column(db.Text)

    def __repr__(self):
        return f'<Timbratura {self.nome} {self.cognome} on {self.data}>'

class Oda(db.Model):
    __tablename__ = 'oda'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero_oda = db.Column(db.String(128), nullable=False)
    posizione_oda = db.Column(db.String(128))

    def __repr__(self):
        return f'<Oda {self.numero_oda}>'
