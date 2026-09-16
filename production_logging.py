"""Production-only journal formatting: allow identifiers, never raw payloads."""
import logging
import re
import sys


class JournalFormatter(logging.Formatter):
    def format(self, record):
        # Format a fresh record so other handlers never inherit mutations.
        event = 'diagnostic'
        message = record.getMessage()
        if record.name == 'nuligahelper.operations' and re.fullmatch(
                r'operation=[a-z_]+ outcome=(?:success|failure|started|unavailable) reason=[A-Za-z_]+', message):
            safe = message
        else:
            match = re.match(r'(auth_abuse_[a-z_]+)\b', message)
            if record.name == 'nuligahelper.security' and match:
                event = match[1]
            component = ('security' if record.name == 'nuligahelper.security' else
                         'web' if record.name.startswith(('webapp', 'gunicorn')) else 'application')
            safe = f'operation={component} event={event}'
        clone = logging.LogRecord('nuligahelper', record.levelno, '', 0, safe, (), None)
        clone.created, clone.msecs = record.created, record.msecs
        return super().format(clone)


def configure():
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(JournalFormatter('%(asctime)s level=%(levelname)s %(message)s'))
    root = logging.getLogger()
    root.handlers[:] = [handler]
    root.setLevel(logging.INFO)
    # Flask installs its own handler only if none exists on initialization.
    for name in ('webapp', 'nuligahelper.security'):
        logger = logging.getLogger(name)
        logger.handlers.clear()
        logger.propagate = True
