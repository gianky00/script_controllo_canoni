#!/bin/bash
set -e

# Set Flask environment variables
export FLASK_APP="app:create_app"
export FLASK_RUN_HOST=0.0.0.0
export FLASK_ENV=production
unset FLASK_DEBUG

# Clean up previous runs
echo "--- Cleaning up old files ---"
rm -rf instance/ logs/ flask.log

# Initialize the database
echo "--- Initializing database ---"
flask init-db

# Run the application
echo "--- Starting Flask server ---"
flask run > flask.log 2>&1 &
