from flask import Blueprint, render_template, jsonify, request, flash, redirect, url_for, current_app
from app import db, scheduler
from app.models import Task, Setting, Timbratura, Oda
from app.tasks import run_scarica_timbrature, run_scarica_canoni
import json

bp = Blueprint('main', __name__)

def sync_scheduler_from_db():
    """
    Synchronizes the running scheduler with the tasks defined in the database.
    Removes jobs that are no longer in the DB or are disabled.
    Adds or updates jobs that are in the DB and enabled.
    """
    print("Syncing scheduler with database...")
    db_tasks = Task.query.all()
    db_task_ids = {task.id for task in db_tasks if task.enabled}

    # Remove jobs from scheduler that are no longer in the DB or are disabled
    for job in scheduler.get_jobs():
        if job.id not in db_task_ids:
            scheduler.remove_job(job.id)
            print(f"Removed job: {job.id}")

    # Add or update jobs
    for task in db_tasks:
        if task.enabled:
            if task.id in TASK_REGISTRY:
                schedule_params = json.loads(task.schedule_data)
                scheduler.add_job(
                    func=TASK_REGISTRY[task.id],
                    trigger='cron',
                    id=task.id,
                    name=task.name,
                    replace_existing=True,
                    **schedule_params
                )
                print(f"Scheduled/Updated job: {task.name}")
            else:
                print(f"Warning: Task function for '{task.id}' not found in TASK_REGISTRY.")

# Mapping from task ID to function
TASK_REGISTRY = {
    'scaricaTimbratureIsab': run_scarica_timbrature,
    'scaricaTScanoni': run_scarica_canoni,
}

@bp.route('/')
def index():
    return render_template('index.html', title='Dashboard')

@bp.route('/data')
def data_viewer():
    # In a real application, you'd want to paginate this query
    records = Timbratura.query.all()
    return render_template('data_viewer.html', title='Visualizzatore Timbrature', records=records)

@bp.route('/scheduler')
def scheduler_page():
    tasks = Task.query.all()
    task_list = []
    for task in tasks:
        job = scheduler.get_job(task.id)
        task_list.append({
            'id': task.id,
            'name': task.name,
            'schedule_data': json.loads(task.schedule_data),
            'enabled': task.enabled,
            'next_run_time': job.next_run_time.strftime('%Y-%m-%d %H:%M:%S') if job and job.next_run_time else 'N/A'
        })
    return render_template('scheduler.html', title='Task Scheduler', tasks=task_list)

@bp.route('/settings', methods=['GET', 'POST'])
def settings_page():
    # Define all keys that should be in the settings page
    required_keys = [
        'LOGIN_URL', 'USERNAME', 'PASSWORD', 'FORNITORE_DA_SELEZIONARE',
        'DOWNLOAD_DIR', 'DIR_SPOSTAMENTO_TS', 'PERCORSO_FILE_MACRO', 'ESEGUIRE_MACRO',
        'DATA_DA_INSERIRE'
    ]

    if request.method == 'POST':
        from app.security import encrypt_password
        for key in required_keys:
            value = request.form.get(key, '')
            setting = Setting.query.get(key)
            if not setting:
                setting = Setting(key=key)

            if key == 'PASSWORD':
                # Don't update the password if the field is submitted empty
                if value:
                    setting.value = encrypt_password(value)
            else:
                setting.value = value

            db.session.add(setting)
        db.session.commit()
        flash('Impostazioni salvate con successo!', 'success')
        return redirect(url_for('main.settings_page'))

    settings = Setting.query.all()
    settings_dict = {s.key: s.value for s in settings}

    # Ensure all required keys are present for the template
    for key in required_keys:
        if key not in settings_dict:
            settings_dict[key] = ''

    return render_template('settings.html', title='Impostazioni', settings=settings_dict)

@bp.route('/oda', methods=['GET', 'POST'])
def oda_management():
    if request.method == 'POST':
        numero_oda = request.form.get('numero_oda')
        posizione_oda = request.form.get('posizione_oda')
        if numero_oda:
            new_oda = Oda(numero_oda=numero_oda, posizione_oda=posizione_oda)
            db.session.add(new_oda)
            db.session.commit()
            flash('OdA aggiunto con successo!', 'success')
        else:
            flash('Il Numero OdA è obbligatorio.', 'danger')
        return redirect(url_for('main.oda_management'))

    odas = Oda.query.all()
    return render_template('oda_management.html', title='Gestione OdA', odas=odas)

@bp.route('/oda/delete/<int:oda_id>', methods=['POST'])
def delete_oda(oda_id):
    oda_to_delete = Oda.query.get_or_404(oda_id)
    db.session.delete(oda_to_delete)
    db.session.commit()
    flash('OdA eliminato con successo.', 'success')
    return redirect(url_for('main.oda_management'))

@bp.route('/tasks', methods=['GET'])
def get_tasks():
    tasks = Task.query.all()
    return jsonify([
        {
            'id': task.id,
            'name': task.name,
            'schedule_data': json.loads(task.schedule_data),
            'enabled': task.enabled,
            'next_run_time': scheduler.get_job(task.id).next_run_time.isoformat() if scheduler.get_job(task.id) else None
        } for task in tasks
    ])

@bp.route('/tasks/run/<task_id>', methods=['POST'])
def run_task_now(task_id):
    task_func = TASK_REGISTRY.get(task_id)
    if task_func:
        scheduler.add_job(func=task_func, id=f"{task_id}_manual_{json.dumps(request.json)}", trigger='date')
        flash(f"Task '{task_id}' has been triggered to run immediately.", "success")
    else:
        flash(f"Task '{task_id}' not found.", "danger")
    return redirect(url_for('main.scheduler_page'))

@bp.route('/task/new', methods=['GET', 'POST'])
def new_task():
    if request.method == 'POST':
        task_id = request.form['id']
        if Task.query.get(task_id):
            flash('Task with this ID already exists.', 'danger')
        else:
            new_task = Task(
                id=task_id,
                name=request.form['name'],
                script_path=request.form['script_path'],
                schedule_data=request.form['schedule_data'],
                enabled='enabled' in request.form
            )
            db.session.add(new_task)
            db.session.commit()
            sync_scheduler_from_db()
            flash('New task created!', 'success')
            return redirect(url_for('main.scheduler_page'))
    return render_template('task_form.html', title='New Task')

@bp.route('/task/edit/<task_id>', methods=['GET', 'POST'])
def edit_task(task_id):
    task = Task.query.get_or_404(task_id)
    if request.method == 'POST':
        task.name = request.form['name']
        task.script_path = request.form['script_path']
        task.schedule_data = request.form['schedule_data']
        task.enabled = 'enabled' in request.form
        db.session.commit()
        sync_scheduler_from_db()
        flash('Task updated!', 'success')
        return redirect(url_for('main.scheduler_page'))

    # For GET request, pre-populate the form
    return render_template('task_form.html', title='Edit Task', task=task)

@bp.route('/task/delete/<task_id>', methods=['POST'])
def delete_task(task_id):
    task = Task.query.get_or_404(task_id)
    db.session.delete(task)
    db.session.commit()
    sync_scheduler_from_db()
    flash('Task deleted!', 'success')
    return redirect(url_for('main.scheduler_page'))
