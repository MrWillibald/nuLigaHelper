# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Domain layer facade. Callers keep using `import db` and db.<name>;
# the implementation lives in topic modules, lowest layer first:
#   roles, errors, models      constants, exceptions, ORM models
#   people, storage, games     persons/teams, SQLite engine, game helpers
#   day_blocks, visibility     dated blocks, overview date boundary
#   clock, audit, staffing     cutoff clock, audit snapshots, staffing
#   slots, locks               shared slot guards, day-before locks
#   duties, block_duties       duty state machine, block/cake writes
#   sync                       merging scraped games
# Business rules belong here, not in nuligahelper.web or nuligahelper.cli. Tests
# patch db.clock.duty_now and db.storage.SQLITE_TIMEOUT_SECONDS.
# ---------------------------------------------------------------

# Former single-module imports stay available as db.<name> (e.g. db.Session,
# db.select); callers and tests rely on them.
import logging
import re
import secrets
import time
from nuligahelper.db import contacts
from nuligahelper.db import age_eligibility
from nuligahelper.db.task_labels import role_label
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
from sqlalchemy import Boolean, CheckConstraint, Column, Date, DateTime, Index, Integer, String, ForeignKey, Table, UniqueConstraint, create_engine, delete, event, select, update
from sqlalchemy.exc import IntegrityError, OperationalError
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    Session,
    mapped_column,
    relationship,
    selectinload,
    sessionmaker,
)

from nuligahelper.db import clock, storage  # patch db.clock.duty_now / db.storage.SQLITE_TIMEOUT_SECONDS
from nuligahelper.db.roles import (
    ROLE_TIMEKEEPER,
    ROLE_SECRETARY,
    ROLE_SALE,
    ROLE_SECURITY,
    ROLE_CASH,
    ROLE_CLEANING,
    ROLE_SLOT_COUNT,
    GAME_DAY_ROLES,
    BASELINE_REQUIRED_POSITIONS,
    offered_positions,
    required_positions,
    position_label,
    BLOCK_PREPARATION,
    BLOCK_CAKE_DELIVERY,
    BLOCK_CLEANUP,
    BLOCK_PHASES,
    BLOCK_SLOT_COUNT,
    BLOCK_PHASE_LABELS,
    TASK_DESCRIPTIONS,
)
from nuligahelper.db.errors import (
    SQLiteInitializationError,
    DatabaseMigrationRequiredError,
    AssignmentTemporarilyUnavailableError,
    SlotConflictError,
    DutyLockedError,
    CakeConfigurationConflictError,
    AgeEligibilityError,
)
from nuligahelper.db.models import (
    SUPPORT_TEAM_NAME,
    ACCOUNT_REGISTERED,
    ACCOUNT_VERIFIED,
    ACCOUNT_ACTIVE,
    ACCOUNT_REJECTED,
    ACCOUNT_INACTIVE,
    ACCOUNT_STATUSES,
    Base,
    ScheduleVisibilitySetting,
    person_teams,
    Team,
    Person,
    Game,
    Assignment,
    DayBlock,
    BlockAssignment,
    DutyCutoff,
    AssignmentAudit,
    AuthToken,
    AuthAbuseCounter,
)
from nuligahelper.db.people import (
    get_support_team,
    get_or_create_team,
    get_all_teams,
    get_or_create_person,
    get_all_persons,
    get_all_person_records,
    membership_teams,
    membership_team_ids,
    membership_team_names,
    membership_label,
    person_label,
    has_team,
    replace_person_teams,
    change_managed_team_membership,
    get_team_members,
    register_person,
    verify_person,
    approve_person,
    validate_birth_date,
    delete_person,
    deactivate_person,
    reactivate_person,
    set_team_mv,
)
from nuligahelper.db.storage import (
    DEFAULT_DB_PATH,
    SQLITE_TIMEOUT_SECONDS,
    SQLITE_BUSY_TIMEOUT_MS,
    SQLITE_SYNCHRONOUS_FULL,
    GAME_IDENTITY_SCHEMA_VERSION,
    resolve_db_path,
    _sqlite_pragma_value,
    _configure_sqlite_connection,
    inspect_sqlite_runtime,
    make_engine,
    _establish_sqlite_runtime,
    _validate_database,
    _engine_database_path,
    initialize_db,
    verify_db,
    init_db,
    make_session_factory,
    _is_sqlite_contention,
    _wait_for_assignment_retry,
    database_ready,
)
from nuligahelper.db.games import (
    get_games_on_date,
    is_spielfest,
    game_display_name,
    game_sort_key,
    classify_game_category,
)
from nuligahelper.db.day_blocks import (
    get_day_blocks,
    _valid_cake_delivery_time,
    block_is_configured,
    block_capacity,
    block_slot_label,
    calculated_block_time,
    reconcile_day_blocks,
)
from nuligahelper.db.visibility import (
    schedule_visibility_boundary,
    schedule_date_visible,
    update_schedule_visibility,
)
from nuligahelper.db.clock import (
    _cutoff_time,
    duty_cutoff_at,
    duty_is_locked,
    _claim_lock_time,
)
from nuligahelper.db.audit import (
    _game_snapshot,
    _audit_assignment,
    _block_snapshot,
    _audit_block_assignment,
)
from nuligahelper.db.staffing import (
    missing_slots,
    claim_eligibility,
    staffing_status,
)
from nuligahelper.db.slots import (
    _stored_slot,
    _validate_slot,
    _slot_assignment,
    _lock_game,
    _authorize_duty_write,
    _validate_retirement,
    _validate_block_slot,
    _lock_day_block,
    _stored_block_slot,
)
from nuligahelper.db.locks import (
    enforce_due_duty_locks,
    lock_due_duties,
    lock_day_before_batch,
)
from nuligahelper.db.duties import (
    _replacement_change,
    request_replacement,
    request_block_replacement,
    cancel_replacement,
    cancel_block_replacement,
    takeover_slot,
    takeover_block_slot,
    claim_slot,
    release_slot,
    set_role_assignments,
    assign_person,
    unassign_person,
)
from nuligahelper.db.block_duties import (
    configure_cake_block,
    claim_block_slot,
    release_block_slot,
)
from nuligahelper.db.sync import (
    ShiftEvent,
    RefereeEvent,
    GameEvent,
    SyncEvents,
    sync_games,
)


def duty_now():
    """Delegate to db.clock.duty_now so a patched test clock applies everywhere."""
    return clock.duty_now()
