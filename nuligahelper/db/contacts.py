"""Server-side validation and canonicalization for person contact data."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Iterable

import phonenumbers
from email_validator import EmailNotValidError, validate_email


class ContactValidationError(ValueError):
    """A German, field-specific contact validation error."""

    def __init__(self, field_name: str, message: str):
        super().__init__(message)
        self.field_name = field_name
        self.message = message


def normalize_email(value: str | None, *, required: bool = False) -> str | None:
    """Return a stable e-mail representation without making DNS requests."""
    candidate = (value or "").strip()
    if not candidate:
        if required:
            raise ContactValidationError(
                "email", "Bitte gib eine E-Mail-Adresse ein."
            )
        return None
    try:
        result = validate_email(
            candidate, check_deliverability=False, test_environment=True
        )
    except EmailNotValidError as exc:
        raise ContactValidationError(
            "email", "Bitte gib eine gültige E-Mail-Adresse ein."
        ) from exc
    return result.normalized.casefold()


def normalize_calling_code(value: str | None) -> str:
    """Validate and normalize a calling code such as +49."""
    digits = (value or "").strip().replace(" ", "")
    if digits.startswith("00"):
        digits = "+" + digits[2:]
    elif digits and not digits.startswith("+"):
        digits = "+" + digits
    if not digits.startswith("+") or not digits[1:].isdigit():
        raise ContactValidationError(
            "country_code", "Bitte wähle eine gültige Ländervorwahl."
        )
    country_code = int(digits[1:])
    if country_code not in phonenumbers.COUNTRY_CODE_TO_REGION_CODE:
        raise ContactValidationError(
            "country_code", "Bitte wähle eine gültige Ländervorwahl."
        )
    return f"+{country_code}"


def normalize_phone(
    value: str | None,
    calling_code: str | None = None,
    *,
    required: bool = False,
) -> str | None:
    """Parse a phone number and return E.164, checking an optional prefix."""
    candidate = (value or "").strip()
    if not candidate:
        if required:
            raise ContactValidationError(
                "phone", "Bitte gib eine Telefonnummer ein."
            )
        return None

    selected = normalize_calling_code(calling_code) if calling_code else None
    explicit_international = candidate.startswith("+") or candidate.startswith("00")
    try:
        if explicit_international:
            international = "+" + candidate[2:] if candidate.startswith("00") else candidate
            parsed = phonenumbers.parse(international, None)
        elif selected:
            country_code = int(selected[1:])
            regions = phonenumbers.region_codes_for_country_code(country_code)
            region = next((item for item in regions if item != "001"), None)
            if region is None:
                raise phonenumbers.NumberParseException(
                    phonenumbers.NumberParseException.INVALID_COUNTRY_CODE,
                    "No geographic region for country code",
                )
            parsed = phonenumbers.parse(candidate, region)
        else:
            parsed = phonenumbers.parse(candidate, None)
    except phonenumbers.NumberParseException as exc:
        raise ContactValidationError(
            "phone", "Bitte gib eine gültige Telefonnummer ein."
        ) from exc

    if selected and parsed.country_code != int(selected[1:]):
        raise ContactValidationError(
            "phone",
            "Die Telefonnummer passt nicht zur gewählten Ländervorwahl.",
        )
    if not phonenumbers.is_possible_number(parsed) or not phonenumbers.is_valid_number(parsed):
        raise ContactValidationError(
            "phone", "Bitte gib eine gültige Telefonnummer ein."
        )
    return phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)


def mask_contact(channel: str, value: str) -> str:
    """Return a useful destination hint without echoing the full contact."""
    if channel == "email":
        local, domain = value.split("@", 1)
        shown = local[:1]
        return f"{shown}{'•' * max(3, len(local) - 1)}@{domain}"
    return f"{'•' * max(4, len(value) - 4)}{value[-4:]}"


@dataclass
class ContactPreflightReport:
    invalid: list[dict] = field(default_factory=list)
    changed: list[dict] = field(default_factory=list)
    collisions: list[dict] = field(default_factory=list)

    @property
    def clean(self) -> bool:
        return not (self.invalid or self.changed or self.collisions)


def analyze_existing_contacts(people: Iterable[object]) -> ContactPreflightReport:
    """Inspect existing values without mutating the supplied ORM records."""
    report = ContactPreflightReport()
    normalized_values: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for person in people:
        identity = {"person_id": person.id, "name": person.name}
        for field_name, normalizer in (
            ("email", normalize_email),
            ("phone", normalize_phone),
        ):
            original = getattr(person, field_name)
            if not original:
                continue
            try:
                normalized = normalizer(original, required=True)
            except ContactValidationError as exc:
                report.invalid.append({
                    **identity,
                    "field": field_name,
                    "value": original,
                    "error": exc.message,
                })
                continue
            entry = {**identity, "field": field_name, "value": original}
            normalized_values[(field_name, normalized)].append(entry)
            if original != normalized:
                report.changed.append({**entry, "normalized": normalized})
    for (field_name, normalized), entries in sorted(normalized_values.items()):
        if len(entries) > 1:
            report.collisions.append({
                "field": field_name,
                "normalized": normalized,
                "people": entries,
            })
    return report
