import sys
import os

# Tambahkan path project ke sys.path
path = os.path.dirname(os.path.abspath(__file__))
if path not in sys.path:
    sys.path.insert(0, path)

from app import app, init_db

init_db()

application = app
