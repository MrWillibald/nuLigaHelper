# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Database layer: SQLAlchemy models, session handling, game sync
# and small query helpers used by notifier and (later) web UI
# ---------------------------------------------------------------

from __future__ import annotations

import logging
import re
import time

import contact_validation as contacts
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

from sqlalchemy import Boolean, CheckConstraint, Column, DateTime, Index, Integer, String, ForeignKey, Table, UniqueConstraint, create_engine, delete, event, select
from sqlalchemy.exc import IntegrityError, OperationalError
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    Session,
    mapped_column,
    relationship,
    sessionmaker,
)

# ---------------------------------------------------------------------------
# Role labels (also used as task names in notifications)
# ---------------------------------------------------------------------------

ROLE_MV = "MV"
ROLE_TIMEKEEPER = "Zeitnehmer"
ROLE_SECRETARY = "Sekretär"
ROLE_SALE = "Verkauf"
ROLE_SECURITY = "Ordnungsdienst"
ROLE_SUPPORT = "Unterstützung"
# Compatibility alias for callers which import the old constant name. Stored and
# displayed current assignments always use ROLE_SUPPORT.
ROLE_CLEANING = ROLE_SUPPORT

# Standard receiver order for game-day notifications
GAME_DAY_ROLES = [
    ROLE_TIMEKEEPER,
    ROLE_SECRETARY,
    ROLE_SALE,
    ROLE_SALE,
    ROLE_SECURITY,
    ROLE_SUPPORT,
]

# Slot count per role; keys define the display order for open-task lists
ROLE_SLOT_COUNT = {
    ROLE_TIMEKEEPER: 1,
    ROLE_SECRETARY: 1,
    ROLE_SALE: 2,
    ROLE_SECURITY: 1,
    ROLE_SUPPORT: 1,
}

# Optional duties stay assignable and notifiable but do not make a game incomplete.
REQUIRED_ROLE_SLOT_COUNT = {
    role: slots for role, slots in ROLE_SLOT_COUNT.items() if role != ROLE_SUPPORT
}

BLOCK_PREPARATION = "preparation"
BLOCK_CLEANUP = "cleanup"
BLOCK_PHASES = (BLOCK_PREPARATION, BLOCK_CLEANUP)
BLOCK_SLOT_COUNT = 3
BLOCK_PHASE_LABELS = {
    BLOCK_PREPARATION: "Vorbereitung",
    BLOCK_CLEANUP: "Aufräumen",
}

DEFAULT_DB_PATH = "nuliga_helper.db"
SQLITE_TIMEOUT_SECONDS = 5.0
SQLITE_BUSY_TIMEOUT_MS = 5000
SQLITE_SYNCHRONOUS_FULL = 2
GAME_IDENTITY_SCHEMA_VERSION = 2


class SQLiteInitializationError(RuntimeError):
    """Raised when the configured SQLite runtime invariants cannot be established."""


class DatabaseMigrationRequiredError(SQLiteInitializationError):
    """Raised when a legacy game-identity schema requires explicit migration."""


class AssignmentTemporarilyUnavailableError(RuntimeError):
    """Raised when an assignment CAS cannot finish within the lock deadline."""

    def __init__(self):
        super().__init__(
            "Die Datenbank ist vorübergehend ausgelastet. Bitte versuche es erneut."
        )


def resolve_db_path(db_path: str) -> str:
    """Resolve a relative DB path against the project directory."""
    import os

    if os.path.isabs(db_path):
        return db_path
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), db_path)


def _sqlite_pragma_value(dbapi_connection, pragma: str):
    cursor = dbapi_connection.cursor()
    try:
        cursor.execute(f"PRAGMA {pragma}")
        row = cursor.fetchone()
        return row[0] if row else None
    finally:
        cursor.close()


def _configure_sqlite_connection(dbapi_connection, _connection_record=None) -> None:
    cursor = dbapi_connection.cursor()
    try:
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.execute(f"PRAGMA busy_timeout={SQLITE_BUSY_TIMEOUT_MS}")
        cursor.execute("PRAGMA synchronous=FULL")
    finally:
        cursor.close()

    actual = {
        "foreign_keys": _sqlite_pragma_value(dbapi_connection, "foreign_keys"),
        "busy_timeout": _sqlite_pragma_value(dbapi_connection, "busy_timeout"),
        "synchronous": _sqlite_pragma_value(dbapi_connection, "synchronous"),
    }
    expected = {
        "foreign_keys": 1,
        "busy_timeout": SQLITE_BUSY_TIMEOUT_MS,
        "synchronous": SQLITE_SYNCHRONOUS_FULL,
    }
    mismatches = [
        f"{name}={actual[name]!r} (expected {expected[name]!r})"
        for name in expected
        if actual[name] != expected[name]
    ]
    if mismatches:
        raise SQLiteInitializationError(
            "SQLite connection safety settings could not be verified: "
            + ", ".join(mismatches)
        )


def inspect_sqlite_runtime(connection) -> dict[str, int | str | None]:
    """Return the SQLite settings used by startup checks and focused tests."""
    return {
        "foreign_keys": connection.exec_driver_sql(
            "PRAGMA foreign_keys"
        ).scalar_one_or_none(),
        "busy_timeout": connection.exec_driver_sql(
            "PRAGMA busy_timeout"
        ).scalar_one_or_none(),
        "synchronous": connection.exec_driver_sql(
            "PRAGMA synchronous"
        ).scalar_one_or_none(),
        "journal_mode": connection.exec_driver_sql(
            "PRAGMA journal_mode"
        ).scalar_one_or_none(),
    }


def make_engine(db_path: str = DEFAULT_DB_PATH):
    """Create an engine with the supported SQLite runtime profile."""
    database_url = (
        "sqlite:///:memory:"
        if db_path == ":memory:"
        else f"sqlite:///{resolve_db_path(db_path)}"
    )
    engine = create_engine(
        database_url,
        connect_args={"timeout": SQLITE_TIMEOUT_SECONDS},
    )
    event.listen(engine, "connect", _configure_sqlite_connection)
    return engine


SUPPORT_TEAM_NAME = "Supporter"

ACCOUNT_REGISTERED = "registered"
ACCOUNT_VERIFIED = "verified"
ACCOUNT_ACTIVE = "active"
ACCOUNT_REJECTED = "rejected"
ACCOUNT_INACTIVE = "inactive"
ACCOUNT_STATUSES = {
    ACCOUNT_REGISTERED,
    ACCOUNT_VERIFIED,
    ACCOUNT_ACTIVE,
    ACCOUNT_REJECTED,
    ACCOUNT_INACTIVE,
}


def _establish_sqlite_runtime(engine) -> None:
    """Establish and verify the writable SQLite runtime profile."""
    try:
        with engine.connect() as connection:
            _configure_sqlite_connection(connection.connection.driver_connection)
            mode = connection.exec_driver_sql("PRAGMA journal_mode=WAL").scalar_one()
            if str(mode).lower() != "wal":
                raise SQLiteInitializationError(
                    "SQLite startup could not establish WAL journal mode "
                    f"(database reported {mode!r}). Ensure the database is writable "
                    "and stored on a local filesystem."
                )
            settings = inspect_sqlite_runtime(connection)
            expected = {
                "foreign_keys": 1,
                "busy_timeout": SQLITE_BUSY_TIMEOUT_MS,
                "synchronous": SQLITE_SYNCHRONOUS_FULL,
            }
            mismatches = [
                f"{name}={settings[name]!r} (expected {value!r})"
                for name, value in expected.items()
                if settings[name] != value
            ]
            if mismatches:
                raise SQLiteInitializationError(
                    "SQLite startup safety checks failed: " + ", ".join(mismatches)
                )
    except SQLiteInitializationError:
        raise
    except OperationalError as exc:
        raise SQLiteInitializationError(
            "SQLite startup failed while establishing and verifying WAL mode. "
            "Ensure the database is writable, not held by another startup, and "
            "stored on a local filesystem."
        ) from exc



def _validate_database(engine) -> None:
    """Run post-initialization/runtime relational integrity checks."""

    try:
        with engine.connect() as connection:
            violations = connection.exec_driver_sql(
                "PRAGMA foreign_key_check"
            ).fetchall()
    except OperationalError as exc:
        raise SQLiteInitializationError(
            "SQLite startup could not run foreign_key_check. Ensure the database "
            "is readable and not locked, then retry."
        ) from exc
    if violations:
        details = "; ".join(
            f"table={table!r}, rowid={rowid!r}, parent={parent!r}, fk_index={fk_id!r}"
            for table, rowid, parent, fk_id in violations
        )
        raise SQLiteInitializationError(
            "SQLite foreign_key_check failed: " + details
            + ". Repair the listed relationships or restore a valid backup."
        )


def _engine_database_path(engine) -> str | None:
    database = engine.url.database
    if not database or database == ":memory:":
        return None
    return resolve_db_path(database)


def initialize_db(engine) -> None:
    """Create and stamp a new database; refuse every populated target."""
    import schema_migrations

    database_path = _engine_database_path(engine)
    if database_path is not None:
        state = schema_migrations.inspect_schema(database_path)
        if state.kind != "empty":
            raise SQLiteInitializationError(
                f"Database initialization requires an absent or empty target; "
                f"found schema state {state.kind!r}."
            )

    _establish_sqlite_runtime(engine)
    Base.metadata.create_all(engine)
    schema_migrations.stamp_head(engine)
    with engine.begin() as connection:
        connection.exec_driver_sql(
            f"PRAGMA user_version={GAME_IDENTITY_SCHEMA_VERSION}"
        )
    _validate_database(engine)

    session = Session(engine)
    try:
        if get_support_team(session) is None:
            session.add(Team(name=SUPPORT_TEAM_NAME, is_support=True))
            session.commit()
            logging.info(f"Support team '{SUPPORT_TEAM_NAME}' created")
    finally:
        session.close()

    if database_path is not None:
        state = schema_migrations.inspect_schema(database_path)
        if state.kind != "head":
            raise SQLiteInitializationError(
                f"Fresh database validation expected Alembic head, found {state.kind!r}."
            )


def verify_db(engine) -> None:
    """Verify an existing head database without creating or upgrading schema."""
    import schema_migrations

    database_path = _engine_database_path(engine)
    if database_path is None:
        raise SQLiteInitializationError(
            "Runtime verification requires a file-backed SQLite database."
        )
    state = schema_migrations.inspect_schema(database_path)
    if state.kind == "legacy_game_identity":
        raise DatabaseMigrationRequiredError(
            "Die Datenbank verwendet die alte Spielidentität. Stoppe Webdienst "
            "und Tagesjob und führe 'manage_db.py migrate-game-identity "
            "--confirm-stopped' aus."
        )
    if state.kind == "invalid":
        details = "; ".join(state.details)
        raise SQLiteInitializationError(
            "SQLite foreign_key_check failed: " + details
            + ". Repair the listed relationships or restore a valid backup."
        )
    if state.kind != "head":
        details = "; ".join(state.details)
        suffix = f" Details: {details}" if details else ""
        raise SQLiteInitializationError(
            f"Database schema is {state.kind!r}, not the required Alembic head. "
            "Stop database users and run 'manage_db.py migrate-schema "
            f"--confirm-stopped'.{suffix}"
        )
    _establish_sqlite_runtime(engine)
    _validate_database(engine)


# Backward-compatible name for explicit test/bootstrap callers. Runtime entry
# points deliberately use verify_db() instead.
init_db = initialize_db


def get_support_team(session: Session) -> Team | None:
    return session.scalars(select(Team).where(Team.is_support.is_(True))).first()


def get_or_create_team(session: Session, name: str) -> Team:
    name = name.strip()
    existing = get_support_team(session)
    if existing is not None and existing.name.lower() == name.lower():
        return existing
    team = session.scalars(select(Team).where(Team.name == name)).first()
    if team is None:
        team = Team(name=name)
        session.add(team)
        session.flush()
    return team


def get_all_teams(session: Session) -> list[Team]:
    """All teams ordered by name."""
    return list(
        session.scalars(select(Team).order_by(Team.is_support.desc(), Team.name))
    )


def make_session_factory(engine):
    return sessionmaker(bind=engine, expire_on_commit=False)


# ---------------------------------------------------------------------------
# ORM models
# ---------------------------------------------------------------------------


class Base(DeclarativeBase):
    pass


person_teams = Table(
    "person_teams",
    Base.metadata,
    Column("person_id", ForeignKey("persons.id", ondelete="CASCADE"), primary_key=True),
    Column("team_id", ForeignKey("teams.id", ondelete="CASCADE"), primary_key=True),
)


class Team(Base):
    """A club team from the game plan or the general support team."""

    __tablename__ = "teams"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(120), unique=True)
    is_support: Mapped[bool] = mapped_column(default=False)
    mv_person_id: Mapped[int | None] = mapped_column(
        ForeignKey("persons.id"), nullable=True
    )

    persons: Mapped[list["Person"]] = relationship(
        secondary=person_teams,
        back_populates="teams",
        order_by="Person.name, Person.id",
    )
    games: Mapped[list["Game"]] = relationship(back_populates="team")
    mv_person: Mapped["Person | None"] = relationship(foreign_keys=[mv_person_id])

    def __repr__(self):
        return f"<Team {self.name!r}{' (support)' if self.is_support else ''}>"


class Person(Base):
    """A club member that can be assigned tasks for home games."""

    __tablename__ = "persons"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(120))
    email: Mapped[str | None] = mapped_column(String(200), nullable=True, unique=True)
    phone: Mapped[str | None] = mapped_column(String(60), nullable=True, unique=True)
    is_admin: Mapped[bool] = mapped_column(Boolean, default=False)
    account_status: Mapped[str] = mapped_column(String(20), default=ACCOUNT_ACTIVE)
    registered_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    rejected_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    verified_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    approved_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    teams: Mapped[list[Team]] = relationship(
        secondary=person_teams,
        back_populates="persons",
        order_by="Team.name, Team.id",
    )
    assignments: Mapped[list["Assignment"]] = relationship(back_populates="person")
    block_assignments: Mapped[list["BlockAssignment"]] = relationship(
        back_populates="person"
    )

    def __repr__(self):
        return f"<Person {self.name!r}>"


class Game(Base):
    """A single home game as scraped from nuLiga."""

    __tablename__ = "games"
    __table_args__ = (
        UniqueConstraint("season_year", "game_nr", name="uq_season_game_nr"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    season_year: Mapped[int] = mapped_column(Integer)
    game_nr: Mapped[str] = mapped_column(String(100))

    day: Mapped[str | None] = mapped_column(String(20), nullable=True)
    date: Mapped[str | None] = mapped_column(String(20), nullable=True)
    time: Mapped[str | None] = mapped_column(String(30), nullable=True)
    hall: Mapped[int | None] = mapped_column(Integer, nullable=True)
    ak: Mapped[str | None] = mapped_column(String(10), nullable=True)
    home: Mapped[str | None] = mapped_column(String(120), nullable=True)
    guest: Mapped[str | None] = mapped_column(String(120), nullable=True)
    score: Mapped[str | None] = mapped_column(String(120), nullable=True)

    # Team providing the game judges (managed by the club, not scraped).
    # Legacy free-text column `jteam` is kept for old databases; new code
    # uses the `team` relationship.
    jteam: Mapped[str | None] = mapped_column(String(120), nullable=True)
    team_id: Mapped[int | None] = mapped_column(ForeignKey("teams.id"), nullable=True)

    team: Mapped[Team | None] = relationship(back_populates="games")
    assignments: Mapped[list["Assignment"]] = relationship(
        back_populates="game", cascade="all, delete-orphan"
    )

    @property
    def judge_team_name(self) -> str | None:
        """Display name of the responsible team (new relation or legacy text)."""
        if self.team is not None:
            return self.team.name
        return self.jteam

    def assignment_by_role(self, role: str) -> "Assignment | None":
        for a in self.assignments:
            if a.role == role:
                return a
        return None

    def assignments_by_role(self, role: str) -> list["Assignment"]:
        return sorted(
            (a for a in self.assignments if a.role == role), key=lambda a: a.slot
        )

    def receivers_for_roles(self, roles: list[str]) -> list[Person]:
        """Return persons for the given role sequence (duplicates included)."""
        result = []
        for role in roles:
            a = self.assignment_by_role(role)
            if a is not None:
                result.append(a.person)
        return result

    def __repr__(self):
        return f"<Game {self.season_year}/{self.game_nr} {self.date} {self.time}>"


class Assignment(Base):
    """A person assigned to a game for a specific task/role."""

    __tablename__ = "assignments"
    # A person can hold at most one task per game, no matter the role.
    __table_args__ = (
        UniqueConstraint("game_id", "person_id", name="uq_game_person"),
        UniqueConstraint("game_id", "role", "slot", name="uq_game_role_slot"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    game_id: Mapped[int] = mapped_column(ForeignKey("games.id"))
    person_id: Mapped[int] = mapped_column(ForeignKey("persons.id"))
    role: Mapped[str] = mapped_column(String(40))
    slot: Mapped[int] = mapped_column(Integer, default=0)

    game: Mapped[Game] = relationship(back_populates="assignments")
    person: Mapped[Person] = relationship(back_populates="assignments")

    def __repr__(self):
        return (
            f"<Assignment game={self.game_id} person={self.person_id} "
            f"role={self.role!r} slot={self.slot}>"
        )


class DayBlock(Base):
    """One automatic preparation or cleanup block for a home-game date."""

    __tablename__ = "day_blocks"
    __table_args__ = (
        UniqueConstraint(
            "season_year", "date", "phase", name="uq_day_block_season_date_phase"
        ),
        CheckConstraint(
            "phase IN ('preparation', 'cleanup')", name="ck_day_block_phase"
        ),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    season_year: Mapped[int] = mapped_column(Integer)
    date: Mapped[str] = mapped_column(String(20))
    phase: Mapped[str] = mapped_column(String(20))
    assignments: Mapped[list["BlockAssignment"]] = relationship(
        back_populates="block", cascade="all, delete-orphan"
    )

    @property
    def label(self) -> str:
        return BLOCK_PHASE_LABELS[self.phase]

    def assignment_for_slot(self, slot: int) -> "BlockAssignment | None":
        return next((item for item in self.assignments if item.slot == slot), None)


class BlockAssignment(Base):
    """A person assigned to one slot of a day block."""

    __tablename__ = "block_assignments"
    __table_args__ = (
        UniqueConstraint("block_id", "person_id", name="uq_block_person"),
        UniqueConstraint("block_id", "slot", name="uq_block_slot"),
        CheckConstraint("slot >= 0 AND slot < 3", name="ck_block_assignment_slot"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    block_id: Mapped[int] = mapped_column(ForeignKey("day_blocks.id"))
    person_id: Mapped[int] = mapped_column(ForeignKey("persons.id"))
    slot: Mapped[int] = mapped_column(Integer)

    block: Mapped[DayBlock] = relationship(back_populates="assignments")
    person: Mapped[Person] = relationship(back_populates="block_assignments")


class AssignmentAudit(Base):
    """Append-only snapshot of one assignment mutation."""

    __tablename__ = "assignment_audit"
    __table_args__ = (
        CheckConstraint(
            "(game_snapshot IS NOT NULL AND block_snapshot IS NULL) OR "
            "(game_snapshot IS NULL AND block_snapshot IS NOT NULL)",
            name="ck_assignment_audit_one_target",
        ),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    changed_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    actor_person_id: Mapped[int | None] = mapped_column(
        ForeignKey("persons.id"), nullable=True
    )
    actor_tier: Mapped[str] = mapped_column(String(20))
    action: Mapped[str] = mapped_column(String(20))
    affected_person_id: Mapped[int | None] = mapped_column(
        ForeignKey("persons.id"), nullable=True
    )
    game_id: Mapped[int | None] = mapped_column(ForeignKey("games.id"), nullable=True)
    block_id: Mapped[int | None] = mapped_column(
        ForeignKey("day_blocks.id", ondelete="SET NULL"), nullable=True
    )
    role: Mapped[str] = mapped_column(String(40))
    slot: Mapped[int] = mapped_column(Integer)
    actor_name: Mapped[str] = mapped_column(String(120))
    affected_person_name: Mapped[str] = mapped_column(String(120))
    game_snapshot: Mapped[str | None] = mapped_column(String(300), nullable=True)
    block_snapshot: Mapped[str | None] = mapped_column(String(300), nullable=True)


class AuthToken(Base):
    """Single-use nonce backing an auth code or a transitional legacy link."""

    __tablename__ = "auth_tokens"

    id: Mapped[int] = mapped_column(primary_key=True)
    nonce: Mapped[str] = mapped_column(String(120), unique=True)
    code: Mapped[str | None] = mapped_column(String(12), nullable=True)
    purpose: Mapped[str] = mapped_column(String(30))
    person_id: Mapped[int] = mapped_column(ForeignKey("persons.id"))
    issued_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    expires_at: Mapped[datetime] = mapped_column(DateTime)
    used_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    person: Mapped[Person] = relationship()


class AuthAbuseCounter(Base):
    """Opaque fixed-window authentication abuse counter."""

    __tablename__ = "auth_abuse_counters"
    __table_args__ = (
        UniqueConstraint(
            "action", "dimension", "subject_digest", "channel",
            "window_started_at", name="uq_auth_abuse_bucket",
        ),
        Index("ix_auth_abuse_expires_at", "expires_at"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    action: Mapped[str] = mapped_column(String(40))
    dimension: Mapped[str] = mapped_column(String(20))
    subject_digest: Mapped[str] = mapped_column(String(64))
    channel: Mapped[str] = mapped_column(String(10))
    window_started_at: Mapped[datetime] = mapped_column(DateTime)
    count: Mapped[int] = mapped_column(Integer)
    expires_at: Mapped[datetime] = mapped_column(DateTime)


# ---------------------------------------------------------------------------
# Sync events
# ---------------------------------------------------------------------------


@dataclass
class ShiftEvent:
    game_id: int
    game_nr: str
    old_date: str
    old_time: str
    new_date: str
    new_time: str


@dataclass
class RefereeEvent:
    game_id: int
    game_nr: str
    date: str
    time: str


@dataclass
class GameEvent:
    game_id: int
    game_nr: str
    ak: str


@dataclass
class SyncEvents:
    shifts: list[ShiftEvent] = field(default_factory=list)
    referee_alerts: list[RefereeEvent] = field(default_factory=list)
    new_games: list[GameEvent] = field(default_factory=list)
    removed_games: list[GameEvent] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Sync scraped games into the database
# ---------------------------------------------------------------------------


def sync_games(session: Session, scraped: list[dict], season_year: int) -> SyncEvents:
    """
    Merge freshly scraped game data into the database.

    - New games are inserted.
    - Existing games get their scraped fields updated; date/time changes and
      newly missing referees ("§77") are reported via the returned events.
    - Games no longer present in the scrape stay untouched but are logged.

    `scraped` items must be dicts with keys:
    day, date, time, hall, game_nr, ak, home, guest, score
    """
    events = SyncEvents()
    game_numbers = [str(rec.get("game_nr", "")).strip() for rec in scraped]
    if any(not number for number in game_numbers):
        raise ValueError("Jedes Spiel benötigt eine Spielnummer.")
    duplicate_keys = sorted({key for key in game_numbers if game_numbers.count(key) > 1})
    if duplicate_keys:
        conflicts = [rec for rec in scraped if str(rec.get("game_nr", "")).strip() in duplicate_keys]
        raise ValueError(
            f"Mehrdeutige Spielidentität {duplicate_keys}: {conflicts!r}"
        )
    existing = {
        g.game_nr: g
        for g in session.scalars(select(Game).where(Game.season_year == season_year))
    }
    new_games: list[Game] = []

    for rec in scraped:
        rec = {**rec, "game_nr": str(rec["game_nr"]).strip()}
        game_nr = rec["game_nr"]
        game = existing.get(game_nr)

        if game is None:
            game = Game(season_year=season_year, **rec)
            session.add(game)
            new_games.append(game)
            logging.info(f"New game {game_nr} added to database")
            continue

        if (game.date != rec["date"]) or (game.time != rec["time"]):
            events.shifts.append(
                ShiftEvent(
                    game_id=game.id,
                    game_nr=game_nr,
                    old_date=game.date or "",
                    old_time=game.time or "",
                    new_date=rec["date"],
                    new_time=rec["time"],
                )
            )
        if ("§77" in rec["score"]) and ("§77" not in (game.score or "")):
            events.referee_alerts.append(
                RefereeEvent(
                    game_id=game.id,
                    game_nr=game_nr,
                    date=rec["date"],
                    time=rec["time"],
                )
            )

        for f in ("day", "date", "time", "hall", "ak", "home", "guest", "score"):
            setattr(game, f, rec[f])

    scraped_keys = set(game_numbers)
    for game_nr in sorted(set(existing) - scraped_keys):
        game = existing[game_nr]
        events.removed_games.append(
            GameEvent(game.id, game.game_nr, game.ak or "")
        )
        logging.warning(f"Game {game.game_nr} not contained in online plan anymore")

    # Teams mirror the scraped age classes ("ak") so they are always available
    for ak in sorted({rec["ak"] for rec in scraped if rec["ak"]}):
        get_or_create_team(session, ak)

    session.flush()
    reconcile_day_blocks(
        session,
        season_year,
        {
            str(record.get("date") or "").strip()
            for record in scraped
            if str(record.get("date") or "").strip()
        },
    )
    events.new_games.extend(
        GameEvent(game.id, game.game_nr, game.ak or "")
        for game in new_games
    )
    session.commit()
    logging.info(
        f"Sync completed: {len(events.new_games)} new, {len(events.shifts)} shifted, "
        f"{len(events.referee_alerts)} referee alerts"
    )
    return events


# ---------------------------------------------------------------------------
# Query helpers
# ---------------------------------------------------------------------------


def get_games_on_date(session: Session, date: str) -> list[Game]:
    """Return all games on the given date string (format dd.mm.yyyy), ordered by time."""
    return list(
        session.scalars(select(Game).where(Game.date == date).order_by(Game.time, Game.game_nr))
    )


def get_day_blocks(
    session: Session, season_year: int, block_date: str
) -> list[DayBlock]:
    """Return a date's blocks in preparation/cleanup order."""
    blocks = list(
        session.scalars(
            select(DayBlock).where(
                DayBlock.season_year == season_year, DayBlock.date == block_date
            )
        )
    )
    order = {phase: index for index, phase in enumerate(BLOCK_PHASES)}
    return sorted(blocks, key=lambda block: (order.get(block.phase, 99), block.id))


def block_slot_label(block_or_phase: DayBlock | str, slot: int) -> str:
    phase = block_or_phase.phase if isinstance(block_or_phase, DayBlock) else block_or_phase
    if phase not in BLOCK_PHASE_LABELS or not 0 <= slot < BLOCK_SLOT_COUNT:
        raise ValueError("Ungültiger Aufgabenplatz.")
    return f"{BLOCK_PHASE_LABELS[phase]} {slot + 1}"


def calculated_block_time(
    block: DayBlock, games: list[Game] | None = None
) -> datetime | None:
    """Calculate a block boundary time, retaining adjacent calendar dates."""
    if games is None:
        session = Session.object_session(block)
        if session is None:
            return None
        games = get_games_on_date(session, block.date)
        games = [game for game in games if game.season_year == block.season_year]
    try:
        base_date = datetime.strptime(block.date, "%d.%m.%Y")
    except (TypeError, ValueError):
        return None
    candidates = []
    for game in games:
        token = (game.time or "").split()
        match = re.fullmatch(r"(\d{1,2}):(\d{2})", token[0] if token else "")
        if not match:
            continue
        hour, minute = map(int, match.groups())
        if hour > 23 or minute > 59:
            continue
        candidates.append(base_date.replace(hour=hour, minute=minute))
    if not candidates:
        return None
    if block.phase == BLOCK_PREPARATION:
        return min(candidates) - timedelta(minutes=90)
    if block.phase == BLOCK_CLEANUP:
        return max(candidates) + timedelta(minutes=60)
    return None


def is_spielfest(game_or_age_group: Game | str | None) -> bool:
    """Return whether a game/age group represents a collapsed SPF event."""
    age_group = (
        game_or_age_group.ak
        if isinstance(game_or_age_group, Game)
        else game_or_age_group
    )
    return bool(re.search(r"(?<!\w)spf(?!\w)", age_group or "", re.IGNORECASE))


def game_display_name(game: Game) -> str:
    """Human-facing event name without exposing synthetic matchup placeholders."""
    if is_spielfest(game):
        return f"Spielfest {game.ak or ''}".strip()
    return f"{game.home or '?'} – {game.guest or '?'}"


def game_sort_key(game: "Game") -> tuple:
    """
    Chronological sort key for games with German dd.mm.yyyy date strings
    (plain string ordering would sort by day-of-month first).
    """
    try:
        d = datetime.strptime(game.date or "", "%d.%m.%Y").date()
    except ValueError:
        d = date.max
    raw_time = (game.time or "").split()
    time_parts = raw_time[0].split(":") if raw_time else []
    if len(time_parts) == 2 and all(p.isdigit() for p in time_parts):
        time_key = (int(time_parts[0]), int(time_parts[1]))
    else:
        time_key = (99, 99)
    number_key = (
        (0, int(game.game_nr))
        if str(game.game_nr).isdigit()
        else (1, str(game.game_nr).casefold())
    )
    return (d, time_key, number_key, game.id or 0)


def get_or_create_person(session: Session, name: str, email: str | None = None,
                         phone: str | None = None,
                         teams: list[Team] | tuple[Team, ...] | None = None) -> Person:
    """Seed a person by name; application identity must use ``Person.id``."""
    name = name.strip()
    person = session.scalars(select(Person).where(Person.name == name)).first()
    if person is None:
        person = Person(name=name, email=email, phone=phone)
        session.add(person)
        session.flush()
        logging.info(f"Person '{name}' created")
    else:
        if email is not None:
            person.email = email
        if phone is not None:
            person.phone = phone
    if teams is not None:
        replace_person_teams(session, person, [team.id for team in teams])
    return person


def get_all_persons(session: Session) -> list[Person]:
    """Return approved, active roster members in display order."""
    return list(
        session.scalars(
            select(Person)
            .where(Person.account_status == ACCOUNT_ACTIVE)
            .order_by(Person.name, Person.id)
        )
    )


def get_all_person_records(session: Session) -> list[Person]:
    """Return every person, including registrations and inactive records."""
    return list(session.scalars(select(Person).order_by(Person.name, Person.id)))


def membership_teams(person: Person) -> tuple[Team, ...]:
    """Return memberships in a stable, presentation-safe order."""
    return tuple(sorted(person.teams, key=lambda team: (team.name.casefold(), team.id)))


def membership_team_ids(person: Person) -> tuple[int, ...]:
    return tuple(team.id for team in membership_teams(person))


def membership_team_names(person: Person) -> tuple[str, ...]:
    return tuple(team.name for team in membership_teams(person))


def membership_label(person: Person, empty: str = "ohne Team") -> str:
    names = membership_team_names(person)
    return ", ".join(names) if names else empty


def person_label(person: Person) -> str:
    return f"{person.name} · {membership_label(person)}"


def has_team(person: Person, team_or_id: Team | int | None) -> bool:
    if team_or_id is None:
        return False
    team_id = team_or_id.id if isinstance(team_or_id, Team) else team_or_id
    return any(team.id == team_id for team in person.teams)


def replace_person_teams(
    session: Session, person: Person, team_ids: list[int] | tuple[int, ...] | set[int]
) -> None:
    """Atomically replace memberships after validating the complete requested set."""
    requested = set(team_ids)
    if any(not isinstance(team_id, int) for team_id in requested):
        raise ValueError("Ungültige Mannschaftsauswahl.")
    teams = list(
        session.scalars(select(Team).where(Team.id.in_(requested)).order_by(Team.name, Team.id))
    ) if requested else []
    if {team.id for team in teams} != requested:
        raise ValueError("Mindestens eine ausgewählte Mannschaft existiert nicht.")
    removed_ids = {team.id for team in person.teams} - requested
    for team in session.scalars(
        select(Team).where(
            Team.mv_person_id == person.id,
            Team.id.in_(removed_ids),
        )
    ) if removed_ids else ():
        team.mv_person_id = None
    person.teams = teams
    session.flush()


def change_managed_team_membership(
    session: Session,
    mv: Person,
    person: Person,
    team: Team,
    *,
    add: bool,
) -> None:
    """Change only one roster membership under an appointed MV's authority."""
    if team.mv_person_id != mv.id or mv.account_status != ACCOUNT_ACTIVE:
        raise ValueError("Diese Mannschaft darf nicht verwaltet werden.")
    if person.account_status != ACCOUNT_ACTIVE:
        raise ValueError("Nur aktive Personen können Mannschaften zugeordnet werden.")
    if not add and person.id == mv.id and has_team(mv, team):
        raise ValueError("Ein MV kann die eigene qualifizierende Mitgliedschaft nicht entfernen.")
    existing = {membership.id: membership for membership in person.teams}
    if add:
        existing[team.id] = team
    else:
        existing.pop(team.id, None)
    person.teams = sorted(existing.values(), key=lambda item: (item.name.casefold(), item.id))
    session.flush()


def get_team_members(session: Session, team: Team) -> list[Person]:
    """Return assignable members of one team."""
    return list(
        session.scalars(
            select(Person)
            .where(
                Person.teams.any(Team.id == team.id),
                Person.account_status == ACCOUNT_ACTIVE,
            )
            .order_by(Person.name, Person.id)
        )
    )


def register_person(
    session: Session,
    name: str,
    teams: list[Team] | tuple[Team, ...],
    email: str | None = None,
    phone: str | None = None,
) -> Person:
    """Create an unverified registration with one or two canonical contacts."""
    email = contacts.normalize_email(email)
    phone = contacts.normalize_phone(phone)
    if not email and not phone:
        raise ValueError("Mindestens eine Kontaktmöglichkeit ist erforderlich.")
    if not teams:
        raise ValueError("Bitte wähle mindestens eine Mannschaft.")
    person = Person(
        name=name.strip(),
        email=email,
        phone=phone,
        account_status=ACCOUNT_REGISTERED,
        registered_at=datetime.now(),
    )
    session.add(person)
    session.flush()
    replace_person_teams(session, person, [team.id for team in teams])
    return person


def verify_person(session: Session, person: Person, at: datetime | None = None) -> None:
    if person.account_status != ACCOUNT_REGISTERED:
        raise ValueError("Die Registrierung kann nicht verifiziert werden.")
    person.account_status = ACCOUNT_VERIFIED
    person.verified_at = at or datetime.now()
    session.flush()


def approve_person(session: Session, person: Person, at: datetime | None = None) -> None:
    if person.account_status != ACCOUNT_VERIFIED or not person.teams:
        raise ValueError("Die Registrierung kann nicht freigegeben werden.")
    person.account_status = ACCOUNT_ACTIVE
    person.approved_at = at or datetime.now()
    session.flush()


class SlotConflictError(ValueError):
    def __init__(self, current_person_id: int | None):
        super().__init__("Der Aufgabenplatz wurde zwischenzeitlich geändert.")
        self.current_person_id = current_person_id


def _is_sqlite_contention(exc: OperationalError) -> bool:
    original = exc.orig
    error_code = getattr(original, "sqlite_errorcode", None)
    if isinstance(error_code, int) and error_code & 0xFF in (5, 6):
        return True
    message = str(original).lower()
    return "database is locked" in message or "database is busy" in message


def _wait_for_assignment_retry(deadline: float, exc: OperationalError) -> None:
    if not _is_sqlite_contention(exc):
        raise exc
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise AssignmentTemporarilyUnavailableError() from exc
    time.sleep(min(0.05, remaining))


def _stored_slot(
    session: Session, game_id: int, role: str, slot: int
) -> Assignment | None:
    return session.scalars(
        select(Assignment).where(
            Assignment.game_id == game_id,
            Assignment.role == role,
            Assignment.slot == slot,
        )
    ).first()


def _validate_slot(role: str, slot: int) -> None:
    if role not in ROLE_SLOT_COUNT or slot < 0 or slot >= ROLE_SLOT_COUNT[role]:
        raise ValueError("Ungültiger Aufgabenplatz.")


def _slot_assignment(game: Game, role: str, slot: int) -> Assignment | None:
    return next(
        (a for a in game.assignments if a.role == role and a.slot == slot), None
    )


def _game_snapshot(game: Game) -> str:
    return (
        f"{game.game_nr} | {game.date or ''} {game.time or ''} | "
        f"{game_display_name(game)}"
    )


def _audit_assignment(
    session: Session,
    assignment: Assignment,
    action: str,
    actor: Person | None,
    actor_tier: str,
) -> None:
    session.add(
        AssignmentAudit(
            actor_person_id=actor.id if actor else None,
            actor_tier=actor_tier,
            action=action,
            affected_person_id=assignment.person_id,
            game_id=assignment.game_id,
            role=assignment.role,
            slot=assignment.slot,
            actor_name=actor.name if actor else "System",
            affected_person_name=assignment.person.name,
            game_snapshot=_game_snapshot(assignment.game),
        )
    )


def claim_slot(
    session: Session,
    game: Game,
    role: str,
    slot: int,
    expected_person_id: int | None,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> Assignment:
    """Claim one slot using bounded, freshly revalidated compare-and-swap."""
    _validate_slot(role, slot)
    game_id = game.id
    person_id = person.id
    actor_id = actor.id if actor else None
    deadline = time.monotonic() + SQLITE_TIMEOUT_SECONDS

    while True:
        try:
            current = _stored_slot(session, game_id, role, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is not None:
                if current.person_id == person_id:
                    return current
                raise SlotConflictError(current.person_id)

            stored_game = session.get(Game, game_id)
            stored_person = session.get(Person, person_id)
            if stored_game is None or stored_person is None:
                raise ValueError("Spiel oder Person wurde nicht gefunden.")
            if stored_person.account_status != ACCOUNT_ACTIVE:
                raise ValueError("Diese Person kann nicht eingeteilt werden.")
            other = session.scalars(
                select(Assignment).where(
                    Assignment.game_id == game_id,
                    Assignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                raise ValueError(
                    f"{stored_person.name} ist für dieses Spiel bereits als "
                    f"'{other.role}' eingeteilt."
                )

            stored_actor = session.get(Person, actor_id) if actor_id else None
            assignment = Assignment(
                game=stored_game,
                person=stored_person,
                role=role,
                slot=slot,
            )
            session.add(assignment)
            session.flush()
            _audit_assignment(
                session, assignment, "claim", stored_actor, actor_tier
            )
            session.flush()
            return assignment
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError as exc:
            session.rollback()
            current = _stored_slot(session, game_id, role, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id) from None
            other = session.scalars(
                select(Assignment).where(
                    Assignment.game_id == game_id,
                    Assignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                stored_person = session.get(Person, person_id)
                name = stored_person.name if stored_person else "Die Person"
                raise ValueError(
                    f"{name} ist für dieses Spiel bereits als '{other.role}' eingeteilt."
                ) from None
            raise


def release_slot(
    session: Session,
    game: Game,
    role: str,
    slot: int,
    expected_person_id: int | None,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> Person | None:
    """Release one slot using bounded, freshly revalidated compare-and-swap."""
    _validate_slot(role, slot)
    game_id = game.id
    actor_id = actor.id if actor else None
    actor_name = actor.name if actor else "System"
    deadline = time.monotonic() + SQLITE_TIMEOUT_SECONDS

    while True:
        try:
            current = _stored_slot(session, game_id, role, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is None:
                return None

            person = current.person
            audit = AssignmentAudit(
                actor_person_id=actor_id,
                actor_tier=actor_tier,
                action="release",
                affected_person_id=current.person_id,
                game_id=current.game_id,
                role=current.role,
                slot=current.slot,
                actor_name=actor_name,
                affected_person_name=person.name,
                game_snapshot=_game_snapshot(current.game),
            )
            result = session.execute(
                delete(Assignment).where(
                    Assignment.id == current.id,
                    Assignment.game_id == game_id,
                    Assignment.role == role,
                    Assignment.slot == slot,
                    Assignment.person_id == expected_person_id,
                ).execution_options(synchronize_session=False)
            )
            if getattr(result, "rowcount", None) != 1:
                session.rollback()
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise AssignmentTemporarilyUnavailableError()
                time.sleep(min(0.05, remaining))
                continue
            session.add(audit)
            session.flush()
            session.expire(current.game, ["assignments"])
            return person
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError:
            session.rollback()
            raise


def _validate_block_slot(block: DayBlock, slot: int) -> None:
    if block.phase not in BLOCK_PHASES or not 0 <= slot < BLOCK_SLOT_COUNT:
        raise ValueError("Ungültiger Aufgabenplatz.")


def _stored_block_slot(
    session: Session, block_id: int, slot: int
) -> BlockAssignment | None:
    return session.scalars(
        select(BlockAssignment).where(
            BlockAssignment.block_id == block_id,
            BlockAssignment.slot == slot,
        )
    ).first()


def _block_snapshot(block: DayBlock) -> str:
    return f"{block.date} | {BLOCK_PHASE_LABELS.get(block.phase, block.phase)}"


def _audit_block_assignment(
    session: Session,
    assignment: BlockAssignment,
    action: str,
    actor: Person | None,
    actor_tier: str,
) -> None:
    session.add(
        AssignmentAudit(
            actor_person_id=actor.id if actor else None,
            actor_tier=actor_tier,
            action=action,
            affected_person_id=assignment.person_id,
            block_id=assignment.block_id,
            role=assignment.block.label,
            slot=assignment.slot,
            actor_name=actor.name if actor else "System",
            affected_person_name=assignment.person.name,
            block_snapshot=_block_snapshot(assignment.block),
        )
    )


def claim_block_slot(
    session: Session,
    block: DayBlock,
    slot: int,
    expected_person_id: int | None,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> BlockAssignment:
    """Claim one day-block slot with bounded compare-and-swap semantics."""
    _validate_block_slot(block, slot)
    block_id, person_id = block.id, person.id
    actor_id = actor.id if actor else None
    deadline = time.monotonic() + SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            current = _stored_block_slot(session, block_id, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is not None:
                if current.person_id == person_id:
                    return current
                raise SlotConflictError(current.person_id)
            stored_block = session.get(DayBlock, block_id)
            stored_person = session.get(Person, person_id)
            if stored_block is None or stored_person is None:
                raise ValueError("Tagesblock oder Person wurde nicht gefunden.")
            if stored_person.account_status != ACCOUNT_ACTIVE:
                raise ValueError("Diese Person kann nicht eingeteilt werden.")
            other = session.scalars(
                select(BlockAssignment).where(
                    BlockAssignment.block_id == block_id,
                    BlockAssignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                raise ValueError(
                    f"{stored_person.name} ist für diesen Tagesblock bereits als "
                    f"'{stored_block.label}' eingeteilt."
                )
            assignment = BlockAssignment(
                block=stored_block, person=stored_person, slot=slot
            )
            session.add(assignment)
            session.flush()
            _audit_block_assignment(
                session,
                assignment,
                "claim",
                session.get(Person, actor_id) if actor_id else None,
                actor_tier,
            )
            session.flush()
            return assignment
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError:
            session.rollback()
            current = _stored_block_slot(session, block_id, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id) from None
            other = session.scalars(
                select(BlockAssignment).where(
                    BlockAssignment.block_id == block_id,
                    BlockAssignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                stored_person = session.get(Person, person_id)
                raise ValueError(
                    f"{stored_person.name if stored_person else 'Die Person'} ist für "
                    "diesen Tagesblock bereits eingeteilt."
                ) from None
            raise


def release_block_slot(
    session: Session,
    block: DayBlock,
    slot: int,
    expected_person_id: int | None,
    actor: Person | None = None,
    actor_tier: str = "system",
    action: str = "release",
) -> Person | None:
    """Release one day-block slot with bounded compare-and-swap semantics."""
    _validate_block_slot(block, slot)
    block_id = block.id
    actor_id = actor.id if actor else None
    deadline = time.monotonic() + SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            current = _stored_block_slot(session, block_id, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is None:
                return None
            person = current.person
            audit = AssignmentAudit(
                actor_person_id=actor_id,
                actor_tier=actor_tier,
                action=action,
                affected_person_id=current.person_id,
                block_id=current.block_id,
                role=current.block.label,
                slot=current.slot,
                actor_name=actor.name if actor else "System",
                affected_person_name=person.name,
                block_snapshot=_block_snapshot(current.block),
            )
            result = session.execute(
                delete(BlockAssignment).where(
                    BlockAssignment.id == current.id,
                    BlockAssignment.block_id == block_id,
                    BlockAssignment.slot == slot,
                    BlockAssignment.person_id == expected_person_id,
                ).execution_options(synchronize_session=False)
            )
            if getattr(result, "rowcount", None) != 1:
                session.rollback()
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise AssignmentTemporarilyUnavailableError()
                time.sleep(min(0.05, remaining))
                continue
            session.add(audit)
            session.flush()
            session.expire(current.block, ["assignments"])
            return person
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError:
            session.rollback()
            raise


def reconcile_day_blocks(
    session: Session, season_year: int, active_dates: set[str]
) -> None:
    """Ensure two blocks per active date and audit removal for vanished dates."""
    existing = list(
        session.scalars(select(DayBlock).where(DayBlock.season_year == season_year))
    )
    by_key = {(block.date, block.phase): block for block in existing}
    for block_date in sorted(active_dates):
        for phase in BLOCK_PHASES:
            if (block_date, phase) not in by_key:
                session.add(
                    DayBlock(season_year=season_year, date=block_date, phase=phase)
                )
    session.flush()
    for block in existing:
        if block.date in active_dates:
            continue
        removed = 0
        for assignment in list(block.assignments):
            release_block_slot(
                session,
                block,
                assignment.slot,
                assignment.person_id,
                actor_tier="system",
                action="remove",
            )
            removed += 1
        logging.info(
            "day_block_removed season_year=%s date=%s phase=%s removed_assignments=%s",
            block.season_year,
            block.date,
            block.phase,
            removed,
        )
        session.delete(block)
    session.flush()


def set_role_assignments(session: Session, game: Game, role: str,
                          person_ids: list[int], actor: Person | None = None,
                          actor_tier: str = "system") -> None:
    """Replace all assignments of `role` on `game` with the given slot order."""
    _validate_slot(role, 0)
    if len(person_ids) > ROLE_SLOT_COUNT[role]:
        raise ValueError("Zu viele Personen für diese Aufgabe.")
    # Release changed slots first so moving a person between slots does not
    # temporarily violate the one-task-per-game constraint.
    for slot in range(ROLE_SLOT_COUNT[role]):
        current = _slot_assignment(game, role, slot)
        desired_id = person_ids[slot] if slot < len(person_ids) else None
        if current is not None and current.person_id != desired_id:
            release_slot(
                session, game, role, slot, current.person_id, actor, actor_tier
            )
    for slot in range(ROLE_SLOT_COUNT[role]):
        current = _slot_assignment(game, role, slot)
        desired_id = person_ids[slot] if slot < len(person_ids) else None
        if desired_id is None or (
            current is not None and current.person_id == desired_id
        ):
            continue
        person_id = desired_id
        person = session.get(Person, person_id)
        if person is not None:
            claim_slot(session, game, role, slot, None, person, actor, actor_tier)
    session.commit()


def delete_person(
    session: Session,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> None:
    """Delete a person together with all their assignments and MV roles."""
    for team in session.scalars(
        select(Team).where(Team.mv_person_id == person.id)
    ):
        team.mv_person_id = None
    for assignment in list(person.assignments):
        _audit_assignment(session, assignment, "unassign", actor, actor_tier)
        assignment.game.assignments.remove(assignment)
    for assignment in list(person.block_assignments):
        _audit_block_assignment(session, assignment, "unassign", actor, actor_tier)
        assignment.block.assignments.remove(assignment)
    session.flush()
    for entry in session.scalars(
        select(AssignmentAudit).where(AssignmentAudit.actor_person_id == person.id)
    ):
        entry.actor_person_id = None
    for entry in session.scalars(
        select(AssignmentAudit).where(AssignmentAudit.affected_person_id == person.id)
    ):
        entry.affected_person_id = None
    session.delete(person)
    session.commit()


def deactivate_person(
    session: Session,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "admin",
) -> None:
    """Deactivate a person, retaining history while freeing future assignments."""
    import common

    today = common.effective_today()
    for team in session.scalars(select(Team).where(Team.mv_person_id == person.id)):
        team.mv_person_id = None
    for assignment in list(person.assignments):
        try:
            game_date = datetime.strptime(assignment.game.date or "", "%d.%m.%Y").date()
        except ValueError:
            game_date = date.max
        if game_date >= today:
            release_slot(
                session,
                assignment.game,
                assignment.role,
                assignment.slot,
                person.id,
                actor,
                actor_tier,
            )
    for assignment in list(person.block_assignments):
        try:
            block_date = datetime.strptime(assignment.block.date, "%d.%m.%Y").date()
        except ValueError:
            block_date = date.max
        if block_date >= today:
            release_block_slot(
                session,
                assignment.block,
                assignment.slot,
                person.id,
                actor,
                actor_tier,
            )
    person.account_status = ACCOUNT_INACTIVE
    session.commit()


def reactivate_person(session: Session, person: Person) -> None:
    if person.account_status != ACCOUNT_INACTIVE:
        raise ValueError("Die Person ist nicht deaktiviert.")
    person.account_status = ACCOUNT_ACTIVE
    session.commit()


def set_team_mv(session: Session, team: Team, person: Person | None) -> None:
    """
    Make *person* the single Mannschaftsverantwortlicher of *team*.
    The person must already be a member of the team; None clears the role.
    """
    if person is not None and not has_team(person, team):
        raise ValueError(f"{person.name} ist kein Mitglied von {team.name}")
    if person is not None and person.account_status != ACCOUNT_ACTIVE:
        raise ValueError(f"{person.name} ist nicht aktiv")
    team.mv_person_id = person.id if person is not None else None
    session.commit()


def missing_slots(game: Game) -> dict[str, int]:
    """Open task slots of a game as {role: missing_count}, in display order."""
    result = {}
    for role, slots in REQUIRED_ROLE_SLOT_COUNT.items():
        missing = max(0, slots - len(game.assignments_by_role(role)))
        if missing:
            result[role] = missing
    return result


def assign_person(
    session: Session,
    game: Game,
    person: Person,
    role: str,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> Assignment:
    """Assign a person to a game with the given role (idempotent).

    A person may only ever hold one task per game – assigning them to
    another role of the same game raises ValueError.
    """
    existing = next(
        (a for a in game.assignments if a.person_id == person.id and a.role == role), None
    )
    if existing is not None:
        return existing
    other = next(
        (a for a in game.assignments if a.person_id == person.id and a.role != role), None
    )
    if other is not None:
        raise ValueError(
            f"{person.name} ist für dieses Spiel bereits als '{other.role}' eingeteilt."
        )
    slot = next(
        (
            index
            for index in range(ROLE_SLOT_COUNT.get(role, 0))
            if _slot_assignment(game, role, index) is None
        ),
        None,
    )
    if slot is None:
        raise ValueError("Für diese Aufgabe ist kein Platz mehr frei.")
    return claim_slot(session, game, role, slot, None, person, actor, actor_tier)


def unassign_person(
    session: Session,
    game: Game,
    person: Person,
    role: str,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> bool:
    """Remove an assignment; returns True if something was removed."""
    assignment = next(
        (a for a in game.assignments if a.person_id == person.id and a.role == role), None
    )
    if assignment is None:
        return False
    release_slot(
        session, game, role, assignment.slot, person.id, actor, actor_tier
    )
    return True


def database_ready(database_path: str) -> bool:
    """Read the existing SQLite schema without creating a missing database."""
    import sqlite3
    from pathlib import Path
    try:
        uri = Path(database_path).resolve().as_uri() + '?mode=ro'
        with sqlite3.connect(uri, uri=True, timeout=1) as connection:
            connection.execute('PRAGMA query_only=ON')
            # Resolve essential schema even for a freshly initialized empty roster.
            connection.execute('SELECT id FROM games LIMIT 0')
            connection.execute('SELECT id FROM persons LIMIT 0')
            connection.execute('SELECT id FROM auth_tokens LIMIT 0')
        return True
    except (sqlite3.Error, OSError, ValueError):
        return False
