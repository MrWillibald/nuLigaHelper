"""Offline birth-date upgrades preserve every legacy identity and relationship."""

from datetime import date
from pathlib import Path
import sqlite3
import tempfile

import helpers as h
import db
import schema_migrations
from schema_fixtures import LEGACY_SCHEMA, create_legacy_database, create_versioned_database


def _older_head(path):
    engine = create_versioned_database(path, "0003_game_day_task_blocks")
    engine.dispose()
    with sqlite3.connect(path) as connection:
        people = [
            (1, "Active admin", db.ACCOUNT_ACTIVE, True, "admin@example.test", None),
            (2, "Inactive", db.ACCOUNT_INACTIVE, False, None, "+491700000002"),
            (3, "Pending", db.ACCOUNT_REGISTERED, False, "pending@example.test", None),
            (4, "Verified", db.ACCOUNT_VERIFIED, False, None, "+491700000004"),
            (5, "Contactless", db.ACCOUNT_ACTIVE, False, None, None),
            (6, "MV", db.ACCOUNT_ACTIVE, False, "mv@example.test", "+491700000006"),
        ]
        connection.executemany(
            "INSERT INTO persons (id, name, account_status, is_admin, email, phone) "
            "VALUES (?, ?, ?, ?, ?, ?)", people
        )
        connection.executemany("INSERT INTO person_teams VALUES (?, 1)",
                               [(person[0],) for person in people])
        connection.execute("UPDATE teams SET mv_person_id=6 WHERE id=1")
        connection.execute("INSERT INTO games (id, season_year, game_nr, date, ak) "
                           "VALUES (10, 2026, '9001', '01.11.2026', 'BL M')")
        connection.executemany(
            "INSERT INTO assignments (id, game_id, person_id, role, slot) VALUES (?, 10, ?, ?, ?)",
            [(20, 1, db.ROLE_SALE, 0), (21, 5, db.ROLE_SALE, 1),
             (22, 6, db.ROLE_TIMEKEEPER, 0)]
        )
        connection.execute("INSERT INTO day_blocks (id, season_year, date, phase) "
                           "VALUES (30, 2026, '01.11.2026', 'preparation')")
        connection.execute("INSERT INTO block_assignments VALUES (40, 30, 5, 0)")
        connection.execute(
            "INSERT INTO assignment_audit "
            "(id, changed_at, actor_tier, action, affected_person_id, role, slot, "
            "actor_name, affected_person_name, block_id, block_snapshot) "
            "VALUES (50, '2026-01-01', 'system', 'claim', 5, 'Vorbereitung', 0, "
            "'System', 'Contactless', 30, '01.11.2026 | Vorbereitung')"
        )
        connection.execute(
            "INSERT INTO auth_tokens "
            "(nonce, code, purpose, person_id, issued_at, expires_at) "
            "VALUES ('synthetic-nonce', '123456', 'register', 3, "
            "'2026-01-01', '2026-01-01 00:15')"
        )
    return db.make_engine(str(path))


def _values(path):
    with sqlite3.connect(path) as connection:
        values = {}
        for table in ("persons", "teams", "person_teams", "games", "assignments",
                      "assignment_audit", "auth_tokens", "day_blocks", "block_assignments"):
            columns = [row[1] for row in connection.execute(f'PRAGMA table_info("{table}")')
                       if row[1] not in {"birth_date", "cake_quantity", "delivery_time"}]
            selected = ", ".join(f'"{column}"' for column in columns)
            clause = " WHERE phase <> 'cake_delivery'" if table == "day_blocks" else ""
            values[table] = sorted(connection.execute(
                f'SELECT {selected} FROM "{table}"{clause}'
            ).fetchall(), key=repr)
        return values


def test_birth_date_revision_extends_single_head_and_fresh_database_has_no_drift():
    assert schema_migrations.head_revisions() == (schema_migrations.HEAD_REVISION,)
    engine = h.make_engine()
    assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
    with h.Session(engine) as session:
        person = db.Person(name="Date-only", birth_date=date(1980, 1, 1))
        session.add(person)
        session.commit()
        session.expire(person)
        assert person.birth_date == date(1980, 1, 1)


def test_guarded_old_head_migration_preserves_all_statuses_contacts_memberships_and_history():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "old-head.db"
        engine = _older_head(path)
        before = _values(path)
        result = schema_migrations.migrate_to_head(path, engine)
        assert result.previous_state.revision == "0003_game_day_task_blocks"
        assert result.backup_path is not None and result.backup_path.exists()
        assert _values(path) == before, "the additive migration must preserve all prior values"
        assert schema_migrations.inspect_schema(result.backup_path).kind == "versioned"
        assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
        with h.Session(engine) as session:
            people = db.get_all_persons(session)
            assert all(person.birth_date is None for person in session.query(db.Person))
            game = session.query(db.Game).one()
            status = db.staffing_status(game)
            assert {item["code"] for item in status["deficiencies"]} == {
                "unknown_birth_date", "missing_adult_seller"
            }
            assert len(game.assignments) == 3
            assert len(people) == 3
        engine.dispose()


def test_unversioned_baseline_upgrade_keeps_unknown_dates_and_legacy_membership_union():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "baseline.db"
        create_legacy_database(path)
        with sqlite3.connect(path) as connection:
            connection.execute("INSERT INTO teams (id, name) VALUES (2, 'Youth')")
            connection.execute(
                "INSERT INTO persons (id, name, team_id, desired_team_id, is_admin, account_status) "
                "VALUES (1, 'Pending legacy', 1, 2, 0, 'registered')"
            )
            connection.execute(
                "INSERT INTO games (id, season_year, game_nr, date, ak) "
                "VALUES (1, 2026, '9001', '01.11.2026', 'BL M')"
            )
            connection.execute(
                "INSERT INTO assignments (id, game_id, person_id, role, slot) "
                "VALUES (1, 1, 1, 'Reinigung', 0)"
            )
        engine = db.make_engine(str(path))
        result = schema_migrations.migrate_to_head(path, engine)
        assert result.revision == schema_migrations.HEAD_REVISION
        with h.Session(engine) as session:
            person = session.get(db.Person, 1)
            assert person.birth_date is None
            assert person.account_status == db.ACCOUNT_REGISTERED
            assert {team.id for team in person.teams} == {1, 2}
            assert session.get(db.Assignment, 1).role == db.ROLE_SUPPORT
        with sqlite3.connect(path) as connection:
            columns = {row[1]: row for row in connection.execute("PRAGMA table_info(persons)")}
            assert columns["birth_date"][2:5] == ("DATE", 0, None)
        engine.dispose()


def test_faithful_baseline_upgrade_has_no_model_drift_through_the_entire_revision_chain():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "faithful-baseline.db"
        # Match the original ORM's client-side defaults and table-level unique
        # clauses, while keeping the separately supported manual fixture above.
        schema = LEGACY_SCHEMA.replace(" DEFAULT 0", "").replace(" DEFAULT 'active'", "")
        schema = schema.replace("nonce VARCHAR(120) NOT NULL UNIQUE", "nonce VARCHAR(120) NOT NULL")
        schema = schema.replace("expires_at DATETIME NOT NULL, used_at DATETIME\n);",
                                "expires_at DATETIME NOT NULL, used_at DATETIME,\n    UNIQUE (nonce)\n);")
        schema = schema.replace("UNIQUE (\n        action, dimension, subject_digest, channel, window_started_at\n    )",
                                "UNIQUE (action, dimension, subject_digest, channel, window_started_at)")
        with sqlite3.connect(path) as connection:
            connection.executescript(schema)
        assert schema_migrations.inspect_schema(path).kind == "baseline"
        engine = db.make_engine(str(path))
        schema_migrations.migrate_to_head(path, engine)
        assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
        engine.dispose()


def test_baseline_unreviewed_default_drift_fails_postflight_with_retained_backup():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "unexpected-default.db"
        # The existing fingerprint recognizes this table shape. Postflight
        # must still reject an unexpected historical default, never skip drift.
        with sqlite3.connect(path) as connection:
            connection.executescript(LEGACY_SCHEMA.replace("DEFAULT 'active'", "DEFAULT 'inactive'"))
        assert schema_migrations.inspect_schema(path).kind == "baseline"
        engine = db.make_engine(str(path))
        try:
            schema_migrations.migrate_to_head(path, engine)
        except schema_migrations.SchemaMigrationError as exc:
            assert "unreviewed model metadata drift" in str(exc)
            assert "Retained backup:" in str(exc)
            assert len(list(path.parent.glob("unexpected-default.db.pre-schema-*.db"))) == 1
        else:
            raise AssertionError("only unchanged recognized historical defaults can pass postflight")
        engine.dispose()


def test_birth_date_source_near_miss_fails_closed_before_upgrade():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "near-miss.db"
        engine = _older_head(path)
        with sqlite3.connect(path) as connection:
            connection.execute("ALTER TABLE persons ADD COLUMN unrelated TEXT")
        try:
            schema_migrations.migrate_to_head(path, engine)
        except schema_migrations.SchemaMigrationError as exc:
            assert "does not match" in str(exc)
        else:
            raise AssertionError("a version label must not authorize an unrecognized source schema")
        assert schema_migrations.inspect_schema(path).revision == "0003_game_day_task_blocks"
        engine.dispose()


def test_manually_stamped_head_without_birth_date_is_refused_at_startup_and_noop_migration():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "forged-head.db"
        engine = _older_head(path)
        with sqlite3.connect(path) as connection:
            connection.execute("UPDATE alembic_version SET version_num=?",
                               (schema_migrations.HEAD_REVISION,))
        assert schema_migrations.inspect_schema(path).kind == "head"
        before = path.read_bytes()
        for call, error_type in (
            (lambda: schema_migrations.migrate_to_head(path, engine), schema_migrations.SchemaMigrationError),
            (lambda: db.verify_db(engine), db.SQLiteInitializationError),
        ):
            try:
                call()
            except error_type as exc:
                assert "invalid birth-date storage" in str(exc)
            else:
                raise AssertionError("the revision label must not substitute for an actual birth-date column")
        assert path.read_bytes() == before, "refusal must not repair or mutate a forged current schema"
        assert list(path.parent.glob("forged-head.db.pre-schema-*.db")) == []
        engine.dispose()


def test_current_head_with_extra_column_is_refused_without_schema_mutation():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "drifted-head.db"
        engine = db.make_engine(str(path))
        db.initialize_db(engine)
        engine.dispose()
        with sqlite3.connect(path) as connection:
            connection.execute("ALTER TABLE persons ADD COLUMN unexpected TEXT")
        engine = db.make_engine(str(path))
        before = path.read_bytes()
        for call, error_type in (
            (lambda: schema_migrations.migrate_to_head(path, engine), schema_migrations.SchemaMigrationError),
            (lambda: db.verify_db(engine), db.SQLiteInitializationError),
        ):
            try:
                call()
            except error_type as exc:
                assert "unreviewed model metadata drift" in str(exc)
            else:
                raise AssertionError("unexpected current-schema columns must fail closed")
        assert path.read_bytes() == before
        engine.dispose()


def test_genuine_current_head_migration_is_verified_readonly_and_creates_no_snapshot():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "head.db"
        engine = db.make_engine(str(path))
        db.initialize_db(engine)
        engine.dispose()
        before = path.read_bytes()
        engine = db.make_engine(str(path))
        result = schema_migrations.migrate_to_head(path, engine)
        assert result.backup_path is None
        assert result.previous_state.kind == "head"
        assert path.read_bytes() == before
        assert list(path.parent.glob("head.db.pre-schema-*.db")) == []
        db.verify_db(engine)
        engine.dispose()


def test_retained_data_postflight_failure_reports_valid_backup_without_private_values():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "postflight.db"
        engine = _older_head(path)
        original = schema_migrations._retained_data

        def mismatch(source, columns=None):
            selected, fingerprints = original(source, columns)
            if columns is not None:
                fingerprints["persons"] = "synthetic-changed-data"
            return selected, fingerprints

        schema_migrations._retained_data = mismatch
        try:
            try:
                schema_migrations.migrate_to_head(path, engine)
            except schema_migrations.SchemaMigrationError as exc:
                message = str(exc)
                assert "Retained-data postflight failed" in message
                assert "Retained backup:" in message
                assert "admin@example.test" not in message and "1980-01-01" not in message
                snapshots = list(path.parent.glob("postflight.db.pre-schema-*.db"))
                assert len(snapshots) == 1
                assert schema_migrations.inspect_schema(snapshots[0]).revision == "0003_game_day_task_blocks"
            else:
                raise AssertionError("retained-data mismatch must never report a successful upgrade")
        finally:
            schema_migrations._retained_data = original
            engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
