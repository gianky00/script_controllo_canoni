from app import create_app, db
from app.models import Setting
from commands import register_commands

app = create_app()
register_commands(app)

@app.shell_context_processor
def make_shell_context():
    return {'db': db, 'Setting': Setting}

if __name__ == '__main__':
    app.run(debug=True)
