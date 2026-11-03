# gunicorn.conf.py
import multiprocessing

bind = "0.0.0.0:10000"
workers = 2
worker_class = "gthread"
threads = 4
timeout = 180
keepalive = 5
max_requests = 250
max_requests_jitter = 50
accesslog = "-"
errorlog = "-"
loglevel = "info"
