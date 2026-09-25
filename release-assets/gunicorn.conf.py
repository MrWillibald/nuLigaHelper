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


# Route-level application events are logged without raw URLs or query strings.
access_log_format = '%(m)s %(s)s'
logconfig_dict = {
    "version": 1, "disable_existing_loggers": False,
    "formatters": {"safe": {"()": "production_logging.JournalFormatter",
                              "format": "%(asctime)s level=%(levelname)s %(message)s"}},
    "handlers": {"journal": {"class": "logging.StreamHandler", "formatter": "safe",
                                "stream": "ext://sys.stderr"}},
    "root": {"level": "INFO", "handlers": ["journal"]},
    "loggers": {name: {"level": "INFO", "handlers": ["journal"], "propagate": False}
                for name in ("gunicorn.error", "gunicorn.access")},
}
