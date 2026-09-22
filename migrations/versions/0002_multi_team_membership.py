"""Replace singular person team fields with a membership association.

Revision ID: 0002_multi_team_membership
Revises: 0001_current_schema_baseline
"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "0002_multi_team_membership"
down_revision = "0001_current_schema_baseline"
branch_labels = None
depends_on = None


def _count(bind, table: str) -> int:
    return bind.execute(sa.text(f'SELECT COUNT(*) FROM "{table}"')).scalar_one()


def upgrade() -> None:
    bind = op.get_bind()
    retained_counts = {
        table: _count(bind, table)
        for table in (
            "persons",
            "assignments",
            "assignment_audit",
            "auth_tokens",
        )
    }
    expected_memberships = bind.execute(
        sa.text(
            "SELECT COUNT(*) FROM ("
            "SELECT id AS person_id, team_id FROM persons WHERE team_id IS NOT NULL "
            "UNION "
            "SELECT id AS person_id, desired_team_id FROM persons "
            "WHERE desired_team_id IS NOT NULL)"
        )
    ).scalar_one()

    op.create_table(
        "person_teams",
        sa.Column(
            "person_id",
            sa.Integer(),
            sa.ForeignKey("persons.id", ondelete="CASCADE"),
            primary_key=True,
        ),
        sa.Column(
            "team_id",
            sa.Integer(),
            sa.ForeignKey("teams.id", ondelete="CASCADE"),
            primary_key=True,
        ),
    )
    bind.execute(
        sa.text(
            "INSERT INTO person_teams (person_id, team_id) "
            "SELECT id, team_id FROM persons WHERE team_id IS NOT NULL "
            "UNION "
            "SELECT id, desired_team_id FROM persons WHERE desired_team_id IS NOT NULL"
        )
    )
    if _count(bind, "person_teams") != expected_memberships:
        raise RuntimeError("Copied membership count does not match the legacy union")

    invalid_mv = bind.execute(
        sa.text(
            "SELECT teams.id FROM teams "
            "LEFT JOIN person_teams ON person_teams.person_id = teams.mv_person_id "
            "AND person_teams.team_id = teams.id "
            "WHERE teams.mv_person_id IS NOT NULL AND person_teams.person_id IS NULL"
        )
    ).fetchall()
    if invalid_mv:
        raise RuntimeError(f"MV appointments without qualifying membership: {invalid_mv!r}")

    with op.batch_alter_table("persons", recreate="always") as batch:
        batch.drop_column("desired_team_id")
        batch.drop_column("team_id")

    for table, expected in retained_counts.items():
        actual = _count(bind, table)
        if actual != expected:
            raise RuntimeError(
                f"Retained row count changed for {table}: {expected} -> {actual}"
            )
    columns = {
        row[1] for row in bind.exec_driver_sql("PRAGMA table_info(persons)").fetchall()
    }
    if {"team_id", "desired_team_id"} & columns:
        raise RuntimeError("Obsolete person team columns remain after migration")


def downgrade() -> None:
    raise RuntimeError(
        "The membership migration is restored from its offline snapshot, not downgraded in place."
    )
