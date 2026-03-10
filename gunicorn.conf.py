import multiprocessing
import os


bind = f"{os.getenv('BIND_HOST', '127.0.0.1')}:{os.getenv('PORT', '8000')}"
workers = int(os.getenv("GUNICORN_WORKERS", str(max(2, min(4, multiprocessing.cpu_count())))))
timeout = int(os.getenv("GUNICORN_TIMEOUT", "300"))
graceful_timeout = 30
accesslog = "-"
errorlog = "-"
capture_output = True
