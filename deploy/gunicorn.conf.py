"""Gunicorn settings for the supported nuLigaHelper production topology."""

bind = "127.0.0.1:8080"
workers = 1
worker_class = "sync"
timeout = 30
graceful_timeout = 30
keepalive = 5
limit_request_line = 4094
limit_request_fields = 100
limit_request_field_size = 8190
accesslog = "-"
errorlog = "-"
capture_output = True

