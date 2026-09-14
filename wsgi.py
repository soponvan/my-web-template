import sys
from pathlib import Path

# Add the project directory to the Python path
project_dir = str(Path(__file__).parent)
if project_dir not in sys.path:
    sys.path.insert(0, project_dir)

# Import the Flask app
from app import app

# This is the WSGI app object that will be served by PythonAnywhere
application = app

if __name__ == "__main__":
    app.run()
