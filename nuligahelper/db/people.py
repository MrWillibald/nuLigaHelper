# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Persons, teams, memberships, registration and account maintenance.
# ---------------------------------------------------------------

from __future__ import annotations

import logging
from datetime import date, datetime

from sqlalchemy import select, update
from sqlalchemy.orm import selectinload, Session

from nuligahelper.db import age_eligibility
from nuligahelper.db import contacts

from nuligahelper.db.models import (
    ACCOUNT_ACTIVE,
    ACCOUNT_INACTIVE,
    ACCOUNT_REGISTERED,
    ACCOUNT_VERIFIED,
    Assignment,
    AssignmentAudit,
    BlockAssignment,
    DayBlock,
    Game,
    Person,
    Team,
)


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


def get_or_create_person(session: Session, name: str, email: str | None = None,
                         phone: str | None = None,
                         teams: list[Team] | tuple[Team, ...] | None = None,
                         *, birth_date: date | str | None = None) -> Person:
    """Seed a person by name; application identity must use ``Person.id``."""
    name = name.strip()
    if birth_date is not None:
        birth_date = validate_birth_date(birth_date)
    person = session.scalars(select(Person).where(Person.name == name)).first()
    if person is None:
        person = Person(name=name, email=email, phone=phone, birth_date=birth_date)
        session.add(person)
        session.flush()
        logging.info(f"Person '{name}' created")
    else:
        if email is not None:
            person.email = email
        if phone is not None:
            person.phone = phone
        if birth_date is not None:
            person.birth_date = birth_date
    if teams is not None:
        replace_person_teams(session, person, [team.id for team in teams])
    return person


def get_all_persons(session: Session) -> list[Person]:
    """Return approved, active roster members in display order."""
    return list(
        session.scalars(
            select(Person)
            .where(Person.account_status == ACCOUNT_ACTIVE)
            .options(selectinload(Person.teams))
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
    *,
    birth_date: date | str,
) -> Person:
    """Create an unverified registration with one or two canonical contacts."""
    birth_date = validate_birth_date(birth_date)
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
        birth_date=birth_date,
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


validate_birth_date = age_eligibility.validate_birth_date


def delete_person(
    session: Session,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> None:
    """Delete an erroneous unassigned record while retaining audit snapshots."""
    with session.no_autoflush:
        session.execute(update(Person).where(Person.id == person.id).values(id=Person.id))
    session.flush()
    session.expire(person, ["assignments", "block_assignments"])
    if person.assignments or person.block_assignments:
        raise ValueError("Die Person hat verbindliche Einteilungen und kann nicht gelöscht werden.")
    for team in session.scalars(select(Team).where(Team.mv_person_id == person.id)):
        team.mv_person_id = None
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
    session: Session, person: Person, actor: Person | None = None,
    actor_tier: str = "admin",
) -> None:
    """Refuse cancellation of current/future commitments; preserve past history."""
    from nuligahelper import common
    with session.no_autoflush:
        session.execute(update(Person).where(Person.id == person.id).values(id=Person.id))
    session.flush()
    session.expire(person, ["assignments", "block_assignments"])
    today = common.effective_today()
    # Read scalar dates after serialization: cached containers may predate a
    # separately committed reschedule from the past into the current season.
    duty_dates = list(session.scalars(select(Game.date).join(
        Assignment, Assignment.game_id == Game.id
    ).where(Assignment.person_id == person.id)))
    duty_dates.extend(session.scalars(select(DayBlock.date).join(
        BlockAssignment, BlockAssignment.block_id == DayBlock.id
    ).where(BlockAssignment.person_id == person.id)))
    if any((age_eligibility.parse_game_date(value) or date.max) >= today
           for value in duty_dates):
        raise ValueError("Die Person hat verbindliche Einteilungen und kann nicht deaktiviert werden.")
    for team in session.scalars(select(Team).where(Team.mv_person_id == person.id)):
        team.mv_person_id = None
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
