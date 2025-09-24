import json
from app import db, models
from app.main import sync_scheduler_from_db
import click
from flask.cli import with_appcontext

@click.command('init-db')
@with_appcontext
def init_db_command():
    """Creates the database tables and seeds initial data."""
    db.create_all()
    print("Initialized the database. All tables should be created now.")

    # Seed initial tasks
    print("Seeding initial tasks...")
    tasks_to_create = [
        {
            'id': 'scaricaTScanoni',
            'name': 'Scarica Canoni TS',
            'script_path': 'scaricaTScanoni.py', # This is just a reference
            'schedule_data': json.dumps({'day_of_week': 'mon-fri', 'hour': 2, 'minute': 0}),
            'enabled': True
        },
        {
            'id': 'scaricaTimbratureIsab',
            'name': 'Scarica Timbrature ISAB',
            'script_path': 'scaricaTimbratureIsab.py', # This is just a reference
            'schedule_data': json.dumps({'day_of_week': 'tue,thu', 'hour': 4, 'minute': 30}),
            'enabled': True
        }
    ]
    for task_data in tasks_to_create:
        task = models.Task.query.get(task_data['id'])
        if not task:
            task = models.Task(**task_data)
            db.session.add(task)
    db.session.commit()
    print("Initial tasks seeded.")

    # Sync the scheduler with the newly created tasks
    print("Syncing scheduler with database...")
    sync_scheduler_from_db()
    print("Scheduler synced.")

def register_commands(app):
    app.cli.add_command(init_db_command)
