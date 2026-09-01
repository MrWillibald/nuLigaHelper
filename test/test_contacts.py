"""Offline tests for shared contact validation and preflight reporting."""

from types import SimpleNamespace

import helpers  # adds the project root for standalone execution
import contact_validation as contacts


def _error(function, *args, **kwargs):
    try:
        function(*args, **kwargs)
    except contacts.ContactValidationError as exc:
        return exc
    raise AssertionError("invalid contact must raise ContactValidationError")


def test_email_normalization_and_validation_are_offline_and_stable():
    assert contacts.normalize_email("  HELPER@Example.COM ") == "helper@example.com"
    assert contacts.normalize_email("") is None
    assert _error(contacts.normalize_email, "kein-kontakt", required=True).field_name == "email"
    assert _error(contacts.normalize_email, "", required=True).field_name == "email"


def test_phone_normalization_accepts_common_format_variations():
    expected = "+491701234567"
    assert contacts.normalize_phone("0170 1234567", "+49", required=True) == expected
    assert contacts.normalize_phone("170-123 45 67", "49", required=True) == expected
    assert contacts.normalize_phone("+49 (170) 1234567", "+49", required=True) == expected
    assert contacts.normalize_phone(expected, required=True) == expected
    assert contacts.normalize_phone("") is None


def test_phone_rejects_invalid_values_and_mismatched_prefixes():
    assert _error(
        contacts.normalize_phone, "123", "+49", required=True
    ).field_name == "phone"
    mismatch = _error(
        contacts.normalize_phone, "+44 20 8366 1177", "+49", required=True
    )
    assert "Ländervorwahl" in mismatch.message
    assert _error(
        contacts.normalize_phone, "0170 1234567", "+999", required=True
    ).field_name == "country_code"


def test_contact_masking_keeps_only_a_destination_hint():
    assert contacts.mask_contact("email", "helper@example.com") == "h•••••@example.com"
    masked_phone = contacts.mask_contact("sms", "+491701234567")
    assert masked_phone.endswith("4567") and "123" not in masked_phone


def test_preflight_reports_invalid_changed_and_colliding_without_mutation():
    people = [
        SimpleNamespace(id=1, name="A", email=" HELPER@Example.COM ", phone=None),
        SimpleNamespace(id=2, name="B", email="helper@example.com", phone="not-a-number"),
        SimpleNamespace(id=3, name="C", email=None, phone="+49 170 1234567"),
        SimpleNamespace(id=4, name="D", email=None, phone="+491701234567"),
    ]
    before = [(person.email, person.phone) for person in people]

    report = contacts.analyze_existing_contacts(people)

    assert [(item["person_id"], item["field"]) for item in report.invalid] == [(2, "phone")]
    assert {(item["person_id"], item["field"]) for item in report.changed} == {
        (1, "email"), (3, "phone")
    }
    assert {(item["field"], item["normalized"]) for item in report.collisions} == {
        ("email", "helper@example.com"),
        ("phone", "+491701234567"),
    }
    assert [(person.email, person.phone) for person in people] == before


if __name__ == "__main__":
    helpers.run_all(dict(globals()))
