import os


bind = f"{os.getenv('BIND_HOST', '127.0.0.1')}:{os.getenv('PORT', '8000')}"
# Solver runs execute in background threads and keep coordination state in-process,
# so the safe default is a single Gunicorn worker unless the execution model changes.
workers = int(os.getenv("GUNICORN_WORKERS", "1"))
timeout = int(os.getenv("GUNICORN_TIMEOUT", "300"))
graceful_timeout = 30
accesslog = "-"
errorlog = "-"
capture_output = True
