"""Mark the reviewed single-team schema as the adoption baseline."""

revision = "0001_current_schema_baseline"
down_revision = None
branch_labels = None
depends_on = None


def upgrade() -> None:
    # Existing databases are stamped here only after schema preflight. Fresh
    # databases are created from current ORM metadata and stamped at head.
    pass


def downgrade() -> None:
    raise RuntimeError("Production rollback restores the pre-migration snapshot.")
