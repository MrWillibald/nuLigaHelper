# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# SQLAlchemy ORM models and account/team constants.
# ---------------------------------------------------------------

from __future__ import annotations

from datetime import date, datetime

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    Column,
    Date,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    Table,
    UniqueConstraint,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

from nuligahelper.db.roles import BLOCK_PHASE_LABELS


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


class Base(DeclarativeBase):
    pass


class ScheduleVisibilitySetting(Base):
    """One global, presentation-only boundary; seeded only by explicit setup."""

    __tablename__ = "schedule_visibility_setting"
    __table_args__ = (CheckConstraint("id = 1", name="ck_schedule_visibility_singleton"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    hide_from: Mapped[date | None] = mapped_column(Date, nullable=True)


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
    # Null represents honestly unknown legacy data, never an inferred age.
    birth_date: Mapped[date | None] = mapped_column(Date, nullable=True)
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

    def assignment_by_role(self, role: str, slot: int = 0) -> "Assignment | None":
        for a in self.assignments:
            if a.role == role and a.slot == slot:
                return a
        return None

    def assignments_by_role(self, role: str) -> list["Assignment"]:
        return sorted(
            (a for a in self.assignments if a.role == role), key=lambda a: a.slot
        )

    def receivers_for_roles(self, roles: list[str]) -> list[Person]:
        """Return each actual occupant once in role and position order."""
        result = []
        for role in dict.fromkeys(roles):
            result.extend(a.person for a in self.assignments_by_role(role))
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
    replacement_token: Mapped[str | None] = mapped_column(String(64), nullable=True)
    locked_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    game: Mapped[Game] = relationship(back_populates="assignments")
    person: Mapped[Person] = relationship(back_populates="assignments")

    def __repr__(self):
        return (
            f"<Assignment game={self.game_id} person={self.person_id} "
            f"role={self.role!r} slot={self.slot}>"
        )


class DayBlock(Base):
    """One automatic dated task block, with configurable cake delivery."""

    __tablename__ = "day_blocks"
    __table_args__ = (
        UniqueConstraint(
            "season_year", "date", "phase", name="uq_day_block_season_date_phase"
        ),
        CheckConstraint(
            "phase IN ('preparation', 'cake_delivery', 'cleanup')",
            name="ck_day_block_phase",
        ),
        CheckConstraint(
            "cake_quantity IS NULL OR "
            "(typeof(cake_quantity) = 'integer' AND cake_quantity >= 0)",
            name="ck_day_block_cake_quantity",
        ),
        CheckConstraint(
            "phase = 'cake_delivery' OR "
            "(cake_quantity IS NULL AND delivery_time IS NULL)",
            name="ck_day_block_cake_metadata",
        ),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    season_year: Mapped[int] = mapped_column(Integer)
    date: Mapped[str] = mapped_column(String(20))
    phase: Mapped[str] = mapped_column(String(20))
    cake_quantity: Mapped[int | None] = mapped_column(Integer, nullable=True)
    delivery_time: Mapped[str | None] = mapped_column(String(5), nullable=True)
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
        CheckConstraint("slot >= 0", name="ck_block_assignment_slot"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    block_id: Mapped[int] = mapped_column(ForeignKey("day_blocks.id"))
    person_id: Mapped[int] = mapped_column(ForeignKey("persons.id"))
    slot: Mapped[int] = mapped_column(Integer)
    replacement_token: Mapped[str | None] = mapped_column(String(64), nullable=True)
    locked_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    block: Mapped[DayBlock] = relationship(back_populates="assignments")
    person: Mapped[Person] = relationship(back_populates="block_assignments")


class DutyCutoff(Base):
    """Durable actual day-before reminder event, including then-empty positions."""

    __tablename__ = "duty_cutoffs"
    __table_args__ = (
        UniqueConstraint("season_year", "date", name="uq_duty_cutoff_season_date"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    season_year: Mapped[int] = mapped_column(Integer)
    date: Mapped[str] = mapped_column(String(20))
    locked_at: Mapped[datetime] = mapped_column(DateTime)


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
