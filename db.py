# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Database layer: SQLAlchemy models, session handling, game sync
# and small query helpers used by notifier and (later) web UI
# ---------------------------------------------------------------

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date, datetime

from sqlalchemy import Integer, String, ForeignKey, UniqueConstraint, create_engine, select
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
ROLE_CLEANING = "Reinigung"

# Standard receiver order for game-day notifications
GAME_DAY_ROLES = [
    ROLE_TIMEKEEPER,
    ROLE_SECRETARY,
    ROLE_SALE,
    ROLE_SALE,
    ROLE_SECURITY,
    ROLE_CLEANING,
]

DEFAULT_DB_PATH = "nuliga_helper.db"


def resolve_db_path(db_path: str) -> str:
    """Resolve a relative DB path against the project directory."""
    import os

    if os.path.isabs(db_path):
        return db_path
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), db_path)


def make_engine(db_path: str = DEFAULT_DB_PATH):
    """Create the SQLAlchemy engine for the SQLite database."""
    return create_engine(f"sqlite:///{resolve_db_path(db_path)}")


SUPPORT_TEAM_NAME = "Supporter"


def init_db(engine) -> None:
    """Create all tables and make sure the support team exists."""
    Base.metadata.create_all(engine)
    session = Session(engine)
    try:
        if get_support_team(session) is None:
            session.add(Team(name=SUPPORT_TEAM_NAME, is_support=True))
            session.commit()
            logging.info(f"Support team '{SUPPORT_TEAM_NAME}' created")
    finally:
        session.close()


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


class Team(Base):
    """A club team from the game plan or the general support team."""

    __tablename__ = "teams"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(120), unique=True)
    is_support: Mapped[bool] = mapped_column(default=False)

    persons: Mapped[list["Person"]] = relationship(back_populates="team")
    games: Mapped[list["Game"]] = relationship(back_populates="team")

    def __repr__(self):
        return f"<Team {self.name!r}{' (support)' if self.is_support else ''}>"


class Person(Base):
    """A club member that can be assigned tasks for home games."""

    __tablename__ = "persons"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(120), unique=True)
    email: Mapped[str | None] = mapped_column(String(200), nullable=True)
    phone: Mapped[str | None] = mapped_column(String(60), nullable=True)
    team_id: Mapped[int | None] = mapped_column(ForeignKey("teams.id"), nullable=True)

    team: Mapped[Team | None] = relationship(back_populates="persons")
    assignments: Mapped[list["Assignment"]] = relationship(back_populates="person")

    def __repr__(self):
        return f"<Person {self.name!r}>"


class Game(Base):
    """A single home game as scraped from nuLiga."""

    __tablename__ = "games"
    __table_args__ = (UniqueConstraint("season_year", "game_nr", name="uq_season_gamenr"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    season_year: Mapped[int] = mapped_column(Integer)
    game_nr: Mapped[int] = mapped_column(Integer)

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
        return [a for a in self.assignments if a.role == role]

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
    __table_args__ = (
        UniqueConstraint("game_id", "person_id", "role", name="uq_game_person_role"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    game_id: Mapped[int] = mapped_column(ForeignKey("games.id"))
    person_id: Mapped[int] = mapped_column(ForeignKey("persons.id"))
    role: Mapped[str] = mapped_column(String(40))

    game: Mapped[Game] = relationship(back_populates="assignments")
    person: Mapped[Person] = relationship(back_populates="assignments")

    def __repr__(self):
        return f"<Assignment game={self.game_id} person={self.person_id} role={self.role!r}>"


# ---------------------------------------------------------------------------
# Sync events
# ---------------------------------------------------------------------------


@dataclass
class ShiftEvent:
    game_nr: int
    old_date: str
    old_time: str
    new_date: str
    new_time: str


@dataclass
class RefereeEvent:
    game_nr: int
    date: str
    time: str


@dataclass
class SyncEvents:
    shifts: list[ShiftEvent] = field(default_factory=list)
    referee_alerts: list[RefereeEvent] = field(default_factory=list)
    new_games: list[int] = field(default_factory=list)
    removed_games: list[int] = field(default_factory=list)


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
    existing = {
        g.game_nr: g
        for g in session.scalars(select(Game).where(Game.season_year == season_year))
    }

    for rec in scraped:
        game_nr = rec["game_nr"]
        game = existing.get(game_nr)

        if game is None:
            game = Game(season_year=season_year, **rec)
            session.add(game)
            events.new_games.append(game_nr)
            logging.info(f"New game {game_nr} added to database")
            continue

        if (game.date != rec["date"]) or (game.time != rec["time"]):
            events.shifts.append(
                ShiftEvent(
                    game_nr=game_nr,
                    old_date=game.date or "",
                    old_time=game.time or "",
                    new_date=rec["date"],
                    new_time=rec["time"],
                )
            )
        if ("§77" in rec["score"]) and ("§77" not in (game.score or "")):
            events.referee_alerts.append(
                RefereeEvent(game_nr=game_nr, date=rec["date"], time=rec["time"])
            )

        for f in ("day", "date", "time", "hall", "ak", "home", "guest", "score"):
            setattr(game, f, rec[f])

    scraped_nrs = {rec["game_nr"] for rec in scraped}
    for game_nr in sorted(set(existing) - scraped_nrs):
        events.removed_games.append(game_nr)
        logging.warning(f"Game {game_nr} not contained in online plan anymore")

    # Teams mirror the scraped age classes ("ak") so they are always available
    for ak in sorted({rec["ak"] for rec in scraped if rec["ak"]}):
        get_or_create_team(session, ak)

    session.commit()
    logging.info(
        f"Sync completed: {len(events.new_games)} new, {len(events.shifts)} shifted, "
        f"{len(events.referee_alerts)} referee alerts"
    )
    return events


# ---------------------------------------------------------------------------
# Query helpers
# ---------------------------------------------------------------------------


def get_game(session: Session, season_year: int, game_nr: int) -> Game | None:
    return session.scalars(
        select(Game).where(Game.season_year == season_year, Game.game_nr == game_nr)
    ).first()


def get_games_on_date(session: Session, date: str) -> list[Game]:
    """Return all games on the given date string (format dd.mm.yyyy), ordered by time."""
    return list(
        session.scalars(select(Game).where(Game.date == date).order_by(Game.time, Game.game_nr))
    )


def game_sort_key(game: "Game") -> tuple:
    """
    Chronological sort key for games with German dd.mm.yyyy date strings
    (plain string ordering would sort by day-of-month first).
    """
    try:
        d = datetime.strptime(game.date or "", "%d.%m.%Y").date()
    except ValueError:
        d = date.max
    time_parts = (game.time or "").split()[0].split(":")
    if len(time_parts) == 2 and all(p.isdigit() for p in time_parts):
        time_key = (int(time_parts[0]), int(time_parts[1]))
    else:
        time_key = (99, 99)
    return (d, time_key, game.game_nr)


def get_or_create_person(session: Session, name: str, email: str | None = None,
                         phone: str | None = None) -> Person:
    """Return the person with the given name, creating them if necessary."""
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
    return person


def get_all_persons(session: Session) -> list[Person]:
    return list(
        session.scalars(
            select(Person).order_by(Person.name)
        )
    )


def set_role_assignments(session: Session, game: Game, role: str,
                         person_ids: list[int]) -> None:
    """Replace all assignments of `role` on `game` with the given slot order."""
    for assignment in [a for a in game.assignments if a.role == role]:
        game.assignments.remove(assignment)
    session.flush()
    for person_id in person_ids:
        person = session.get(Person, person_id)
        if person is not None:
            game.assignments.append(Assignment(person=person, role=role))
    session.commit()


def delete_person(session: Session, person: Person) -> None:
    """Delete a person together with all their assignments."""
    for assignment in list(person.assignments):
        session.delete(assignment)
    session.delete(person)
    session.commit()


def assign_person(session: Session, game: Game, person: Person, role: str) -> Assignment:
    """Assign a person to a game with the given role (idempotent)."""
    existing = next(
        (a for a in game.assignments if a.person_id == person.id and a.role == role), None
    )
    if existing is not None:
        return existing
    assignment = Assignment(game=game, person=person, role=role)
    session.add(assignment)
    session.flush()
    return assignment


def unassign_person(session: Session, game: Game, person: Person, role: str) -> bool:
    """Remove an assignment; returns True if something was removed."""
    assignment = next(
        (a for a in game.assignments if a.person_id == person.id and a.role == role), None
    )
    if assignment is None:
        return False
    session.delete(assignment)
    session.flush()
    return True
