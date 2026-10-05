"""Shared calendar and game-category rules without exposing private dates."""

from __future__ import annotations

import re
from datetime import date, datetime

import common


ADULT = "adult"
YOUTH = "youth"
UNKNOWN = "unknown"

# One local catalog until the message-template change supplies a central one.
MESSAGES = {
    "unknown_birth_date": "{role}: Das Geburtsdatum fehlt; die Altersberechtigung ist ungeklärt.",
    "underage": "{role}: Am Spieltag sind mindestens {minimum_age} Jahre erforderlich.",
    "unresolved_game_date": "{role}: Das Spieldatum ist ungeklärt; die Altersberechtigung kann nicht geprüft werden.",
    "unresolved_category": "{role}: Die Altersklasse ist ungeklärt; das Mindestalter kann nicht bestimmt werden.",
    "missing_adult_seller": "Verkauf: Mindestens eine eingeteilte Person muss am Spieltag 18 Jahre oder älter sein.",
}


def validate_birth_date(value: date | str | None) -> date:
    """Accept a calendar date, requiring known, nonfuture date-only input."""
    if isinstance(value, datetime):
        raise ValueError("Bitte gib ein gültiges Geburtsdatum ohne Uhrzeit an.")
    if isinstance(value, date):
        parsed = value
    elif isinstance(value, str) and value.strip():
        value = value.strip()
        try:
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
                parsed = date.fromisoformat(value)
            elif re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", value):
                parsed = datetime.strptime(value, "%d.%m.%Y").date()
            else:
                raise ValueError
        except ValueError:
            raise ValueError("Bitte gib ein gültiges Geburtsdatum an (JJJJ-MM-TT).") from None
    else:
        raise ValueError("Bitte gib ein Geburtsdatum an.")
    if parsed > common.effective_today():
        raise ValueError("Das Geburtsdatum darf nicht in der Zukunft liegen.")
    return parsed


def classify_game_category(age_class: str | None) -> str:
    """Recognize explicit M/F, mA–mE/wA–wE and SPF tokens, independent of league."""
    if not isinstance(age_class, str) or not age_class.strip():
        return UNKNOWN
    # A full token matters: e.g. malformed 'mF', 'M1' or 'SPFoo' is not a class.
    tokens = re.findall(r"(?<!\w)(?:[mw][A-E]|[MF]|GE)(?!\w)", age_class)
    if "GE" in tokens:
        return UNKNOWN
    adult = any(token in {"M", "F"} for token in tokens)
    youth = any(re.fullmatch(r"[mw][A-E]", token) for token in tokens)
    youth = youth or bool(re.search(r"(?<!\w)SPF(?!\w)", age_class, re.I))
    if adult == youth:
        return UNKNOWN
    return ADULT if adult else YOUTH


def parse_game_date(value: str | None) -> date | None:
    try:
        return datetime.strptime(value or "", "%d.%m.%Y").date()
    except (ValueError, TypeError):
        return None


def completed_age(birth_date: date, game_date: date) -> int:
    """Completed years; 29 February advances on 1 March in non-leap years."""
    return game_date.year - birth_date.year - (
        (game_date.month, game_date.day) < (birth_date.month, birth_date.day)
    )


def deficiency(role: str, code: str, *, minimum_age: int | None = None,
               slot: int | None = None) -> dict:
    result = {"role": role, "code": code}
    if minimum_age is not None:
        result["minimum_age"] = minimum_age
    if slot is not None:
        result["slot"] = slot
    result["message"] = MESSAGES[code].format(role=role, minimum_age=minimum_age)
    return result


def timing_eligibility(game, role: str, person, slot: int = 0) -> dict | None:
    """Return a public-safe refusal for an individual timing duty."""
    if role not in {"Zeitnehmer", "Sekretär"}:
        return None
    category = classify_game_category(game.ak)
    if category == UNKNOWN:
        return deficiency(role, "unresolved_category", slot=slot)
    minimum = 14 if category == YOUTH else (18 if role == "Zeitnehmer" else 16)
    game_date = parse_game_date(game.date)
    if game_date is None:
        return deficiency(role, "unresolved_game_date", minimum_age=minimum, slot=slot)
    if person.birth_date is None:
        return deficiency(role, "unknown_birth_date", minimum_age=minimum, slot=slot)
    if completed_age(person.birth_date, game_date) < minimum:
        return deficiency(role, "underage", minimum_age=minimum, slot=slot)
    return None


def sale_coverage(game, sellers) -> dict | None:
    """Only explicitly assigned sellers can prove adult coverage."""
    game_date = parse_game_date(game.date)
    if game_date is None:
        return deficiency("Verkauf", "unresolved_game_date", minimum_age=18)
    if any(person.birth_date is not None and
           completed_age(person.birth_date, game_date) >= 18 for person in sellers):
        return None
    return deficiency("Verkauf", "missing_adult_seller", minimum_age=18)
