"""systemd ExecStartPre: fail by setting name without disclosing configuration."""
import os
import sys
import common
import production as p


def check():
    for name in ('NULIGAHELPER_SECRET', 'NULIGAHELPER_CONFIG', 'NULIGAHELPER_DB',
                 'NULIGAHELPER_STATE_DIR', 'NULIGAHELPER_OPERATIONS'):
        if not os.environ.get(name):
            return name
    try:
        p.text(os.environ['NULIGAHELPER_SECRET'], 'NULIGAHELPER_SECRET')
    except p.ConfigurationError:
        return 'NULIGAHELPER_SECRET'
    try:
        data = p.validate_operations(p.read_json(os.environ['NULIGAHELPER_OPERATIONS']))
    except Exception:
        return 'NULIGAHELPER_OPERATIONS'
    for name, key in [('NULIGAHELPER_CONFIG','config_file'), ('NULIGAHELPER_DB','database'),
                      ('NULIGAHELPER_STATE_DIR','state_dir'), ('NULIGAHELPER_TRUSTED_HOSTS','hostname')]:
        if os.environ.get(name) != data[key]:
            return name
    if os.environ.get('TZ') != 'UTC':
        return 'TZ'
    try:
        cfg = common.load_config()
        for section, names in {'email': ('smtpserver', 'mail_ID', 'mail_password'),
                               'twilio': ('twilio_sid', 'twilio_token', 'twilio_service_ID'),
                               'dropbox': ('dropbox_token', 'dropbox_folder')}.items():
            for name in names:
                p.text(cfg['club'][section][name], 'NULIGAHELPER_CONFIG')
    except ValueError as error:
        name = str(error)
        return name if name in {item for fields in common.PROVIDER_ENV.values() for item in fields.values()} else 'NULIGAHELPER_CONFIG'
    except Exception:
        return 'NULIGAHELPER_CONFIG'
    return None


if __name__ == '__main__':
    error = check()
    if error:
        print(error, file=sys.stderr)
        raise SystemExit(1)
