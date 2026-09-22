from __future__ import annotations

from logging.config import fileConfig

from alembic import context

import db


config = context.config
if config.config_file_name:
    fileConfig(config.config_file_name, disable_existing_loggers=False)

target_metadata = db.Base.metadata


def run_migrations_offline() -> None:
    raise RuntimeError(
        "Offline SQL generation is not supported; use the guarded project commands."
    )


def run_migrations_online() -> None:
    connection = config.attributes.get("connection")
    if connection is None:
        raise RuntimeError(
            "Alembic requires a programmatically supplied database connection."
        )

    context.configure(
        connection=connection,
        target_metadata=target_metadata,
        render_as_batch=True,
        compare_type=True,
        compare_server_default=True,
    )
    with context.begin_transaction():
        context.run_migrations()


if context.is_offline_mode():
    run_migrations_offline()
else:
    run_migrations_online()
