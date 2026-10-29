# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Structure of the db package: the facade keeps every name callers
# used before the split, the ORM metadata stays complete and the
# shared duty clock is patched in one place.
# ---------------------------------------------------------------

import ast
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfo

import helpers as h
import db

DB_DIR = Path(h.PROJECT_DIR) / "db"
# Import order of the submodules; a module may only import from earlier ones.
LAYERS = ["roles", "errors", "models", "people", "storage", "games", "day_blocks", "visibility",
          "clock", "audit", "staffing", "slots", "locks", "duties", "block_duties", "sync"]

# Every non-module name of the former single-file db module (2026-10-29),
# except the `annotations` future flag.
FACADE_NAMES = """
    ACCOUNT_ACTIVE ACCOUNT_INACTIVE ACCOUNT_REGISTERED ACCOUNT_REJECTED ACCOUNT_STATUSES
    ACCOUNT_VERIFIED AgeEligibilityError Assignment AssignmentAudit
    AssignmentTemporarilyUnavailableError AuthAbuseCounter AuthToken
    BASELINE_REQUIRED_POSITIONS BLOCK_CAKE_DELIVERY BLOCK_CLEANUP BLOCK_PHASES
    BLOCK_PHASE_LABELS BLOCK_PREPARATION BLOCK_SLOT_COUNT Base BlockAssignment Boolean
    CakeConfigurationConflictError CheckConstraint Column DEFAULT_DB_PATH
    DatabaseMigrationRequiredError Date DateTime DayBlock DeclarativeBase DutyCutoff
    DutyLockedError ForeignKey GAME_DAY_ROLES GAME_IDENTITY_SCHEMA_VERSION Game
    GameEvent Index Integer IntegrityError Mapped OperationalError Person ROLE_CASH
    ROLE_CLEANING ROLE_SALE ROLE_SECRETARY ROLE_SECURITY ROLE_SLOT_COUNT ROLE_TIMEKEEPER
    RefereeEvent SQLITE_BUSY_TIMEOUT_MS SQLITE_SYNCHRONOUS_FULL SQLITE_TIMEOUT_SECONDS
    SQLiteInitializationError SUPPORT_TEAM_NAME ScheduleVisibilitySetting Session
    ShiftEvent SlotConflictError String SyncEvents TASK_DESCRIPTIONS Table Team
    UniqueConstraint ZoneInfo _audit_assignment _audit_block_assignment
    _authorize_duty_write _block_snapshot _claim_lock_time _configure_sqlite_connection
    _cutoff_time _engine_database_path _establish_sqlite_runtime _game_snapshot
    _is_sqlite_contention _lock_day_block _lock_game _replacement_change
    _slot_assignment _sqlite_pragma_value _stored_block_slot _stored_slot
    _valid_cake_delivery_time _validate_block_slot _validate_database
    _validate_retirement _validate_slot _wait_for_assignment_retry
    approve_person assign_person block_capacity block_is_configured block_slot_label
    calculated_block_time cancel_block_replacement cancel_replacement
    change_managed_team_membership claim_block_slot claim_eligibility claim_slot
    classify_game_category configure_cake_block create_engine database_ready dataclass
    date datetime deactivate_person delete delete_person duty_cutoff_at duty_is_locked
    duty_now enforce_due_duty_locks field game_display_name game_sort_key
    get_all_person_records get_all_persons get_all_teams get_day_blocks
    get_games_on_date get_or_create_person get_or_create_team get_support_team
    get_team_members has_team init_db initialize_db inspect_sqlite_runtime is_spielfest
    lock_day_before_batch lock_due_duties make_engine make_session_factory mapped_column
    membership_label membership_team_ids membership_team_names membership_teams
    missing_slots offered_positions person_label person_teams position_label
    reactivate_person reconcile_day_blocks register_person relationship
    release_block_slot release_slot replace_person_teams request_block_replacement
    request_replacement required_positions resolve_db_path role_label
    schedule_date_visible schedule_visibility_boundary select selectinload sessionmaker
    set_role_assignments set_team_mv staffing_status sync_games takeover_block_slot
    takeover_slot timedelta unassign_person update update_schedule_visibility
    validate_birth_date verify_db verify_person
""".split()

TABLES = {
    "assignment_audit", "assignments", "auth_abuse_counters", "auth_tokens",
    "block_assignments", "day_blocks", "duty_cutoffs", "games", "person_teams", "persons",
    "schedule_visibility_setting", "teams",
}


def test_facade_keeps_every_former_name():
    missing = [name for name in FACADE_NAMES if not hasattr(db, name)]
    assert not missing, f"callers and tests use these db names: {missing}"


def test_metadata_contains_every_table():
    assert set(db.Base.metadata.tables) == TABLES, (
        "Alembic drift checks and initialization need the complete model metadata"
    )



def test_patched_duty_clock_reaches_lock_decisions():
    engine = h.make_engine()
    berlin = ZoneInfo("Europe/Berlin")
    with h.Session(engine) as session:
        db.sync_games(session, h.sample_games(), h.SEASON)
        session.commit()
        game = session.query(db.Game).order_by(db.Game.id).first()
        cutoff = db.duty_cutoff_at(game)
        before = cutoff.replace(hour=8, minute=59)
        with patch("db.clock.duty_now", return_value=before):
            assert db.duty_now() == before, "the facade clock must delegate to db.clock"
            assert not db.duty_is_locked(session, game), "before 09:00 the duty is open"
        with patch("db.clock.duty_now", return_value=cutoff.astimezone(berlin)):
            assert db.duty_is_locked(session, game), "at 09:00 the day before the duty locks"


def test_modules_read_patchable_values_through_their_module():
    for path in sorted(DB_DIR.glob("*.py")):
        source = path.read_text()
        assert "from db.clock import duty_now" not in source and "import duty_now" not in source, (
            f"{path.name} must call clock.duty_now() so db.clock.duty_now patches apply"
        )
        if path.stem != "storage":
            assert "import SQLITE_TIMEOUT_SECONDS" not in source, (
                f"{path.name} must read storage.SQLITE_TIMEOUT_SECONDS at call time"
            )
    for path in sorted((Path(h.PROJECT_DIR) / "test").glob("*.py")):
        if path.name == Path(__file__).name:
            continue
        source = path.read_text()
        assert 'patch("db.duty_now"' not in source and "patch('db.duty_now'" not in source, (
            f"{path.name}: patch db.clock.duty_now; the facade delegates to it"
        )


def test_submodules_import_only_lower_layers():
    for name in LAYERS:
        tree = ast.parse((DB_DIR / f"{name}.py").read_text())
        typing_only = {id(sub) for node in ast.walk(tree) if isinstance(node, ast.If)
                       and isinstance(node.test, ast.Name) and node.test.id == "TYPE_CHECKING"
                       for sub in ast.walk(node)}
        for node in ast.walk(tree):
            if id(node) in typing_only:
                continue  # annotation-only imports never run
            if isinstance(node, ast.ImportFrom) and node.module and node.module.startswith("db."):
                target = node.module.split(".")[1]
            elif isinstance(node, ast.ImportFrom) and node.module == "db":
                target = None
                for alias in node.names:
                    assert LAYERS.index(alias.name) < LAYERS.index(name), (
                        f"db.{name} imports higher layer {alias.name}"
                    )
                continue
            else:
                continue
            assert LAYERS.index(target) < LAYERS.index(name), (
                f"db.{name} imports db.{target}; keep the layer order acyclic"
            )


if __name__ == "__main__":
    h.run_all(dict(globals()))
