"""Offline catalog contracts and distinguishable synthetic rendered examples."""

from dataclasses import FrozenInstanceError, replace
from unittest.mock import patch

import helpers as h
import messages


def _context(key):
    return {field: f"VALUE_{field.upper()}" for field in messages.CATALOG[key].fields}


def _refusal(call, required_words):
    try:
        call()
    except messages.MessageRenderError as exc:
        diagnostic = str(exc)
        for word in required_words:
            assert word in diagnostic, "render refusal must identify the key or field contract"
        return diagnostic
    raise AssertionError("invalid context produced a dispatchable message")


def test_every_catalog_entry_has_documented_fields_and_a_rendered_example():
    for key, template in messages.CATALOG.items():
        assert template.purpose and template.caller, f"{key} needs an inventory purpose and caller"
        assert set(template.fields) <= messages.FIELD_DESCRIPTIONS.keys(), f"{key} has undocumented fields"
        context = _context(key)
        if isinstance(template, messages.FragmentTemplate):
            combined = messages.render_fragment(key, **context)
        else:
            result = messages.render(key, **context)
            combined = result.subject + result.email + (result.sms or "")
            if template.shared_sms:
                assert result.email == result.sms, f"{key} must retain intentional shared-channel copy"
            elif template.sms is None:
                assert result.sms is None, f"{key} is email-only"
        for field, value in context.items():
            assert value in combined, f"{key} must place {field} in its intended rendered content"


def test_referee_values_occupy_semantic_positions_in_both_channels():
    message = messages.render(
        "referee.missing", recipient_name="Recipient R", age_class="Age Class A",
        game_date="31.12.2042", game_start_time="23:17",
        notified_person_names="Notified N, Notified M",
    )
    for body in (message.email, message.sms):
        assert "Hallo Recipient R" in body
        assert "Heimspiel der Age Class A am 31.12.2042 um 23:17" in body, \
            "age class, date and event time must no longer shift positional arguments"
        assert "Personen werden benachrichtigt: Notified N, Notified M" in body


def test_game_and_block_times_are_explicit_and_use_current_task_labels():
    game = messages.render("game.day_before", **{
        **_context("game.day_before"), "task_label": "Aktueller Spiel-Dienst",
        "game_start_time": "14:13", "game_date": "25.12.2042",
    })
    assert "Das Spiel beginnt um 14:13" in game.email
    assert "Aktueller Spiel-Dienst" in game.subject and "Aktueller Spiel-Dienst" in game.sms
    block = messages.render("block.preparation.weekly", **{
        **_context("block.preparation.weekly"), "task_label": "Aktuelle Vorbereitung",
        "block_meeting_time": "12:13", "block_date": "25.12.2042",
    })
    assert "Treffpunkt: 12:13" in block.email and "Treffpunkt: 12:13" in block.sms
    assert "Aktuelle Vorbereitung" in block.subject
    assert "Spiel beginnt" not in block.email, "block meeting time must not become a match start"


def test_missing_extra_and_unknown_contexts_refuse_without_values_in_diagnostics():
    secret = "123456-secret-test-value"
    diagnostic = _refusal(
        lambda: messages.render("auth.login_code", auth_code=secret),
        ("auth.login_code", "recipient_name"),
    )
    assert secret not in diagnostic
    diagnostic = _refusal(
        lambda: messages.render("auth.login_code", recipient_name="Fixture Person", auth_code=secret, unexpected=secret),
        ("auth.login_code", "unexpected"),
    )
    assert secret not in diagnostic and "Fixture Person" not in diagnostic
    _refusal(lambda: messages.render("missing.key"), ("missing.key",))
    _refusal(lambda: messages.render_fragment("missing.fragment"), ("missing.fragment",))
    _refusal(lambda: messages.render("block.time_unset"), ("block.time_unset",))
    _refusal(lambda: messages.render_fragment("game.new"), ("game.new",))
    _refusal(lambda: messages.render_fragment("task.cake"), ("task_label",))


def test_context_values_remain_literal_plain_text_and_messages_are_immutable():
    name = "Fixture {auth_code} {recipient_name.__class__}"
    rendered = messages.render("auth.login_code", recipient_name=name, auth_code="641203")
    assert name in rendered.email and name in rendered.sms, "braces in display values must stay literal"
    assert rendered.email.count("641203") == 1, "a value must not create another substitution"
    try:
        rendered.subject = "Changed"
    except FrozenInstanceError:
        pass
    else:
        raise AssertionError("rendered message must be immutable before transport receives it")
    _refusal(
        lambda: messages.render("auth.login_code", recipient_name=object(), auth_code="641203"),
        ("recipient_name",),
    )


def test_subject_line_breaks_and_invalid_template_expressions_refuse_before_dispatch():
    for separator in ("\n", "\r", "\r\n", "\u2028"):
        context = {**_context("game.day_before"), "task_label": "Fixture" + separator + "Bcc: hidden"}
        diagnostic = _refusal(lambda: messages.render("game.day_before", **context), ("game.day_before",))
        assert "hidden" not in diagnostic and "Bcc" not in diagnostic
    original = messages.CATALOG["auth.login_code"]
    for invalid_email in ("{recipient_name.__class__} {auth_code}", "{recipient_name!r} {auth_code}", "{recipient_name:>30} {auth_code}", "{recipient_name} {auth_code} {unexpected}", "{broken"):
        with patch.dict(messages.CATALOG, {"auth.login_code": replace(original, email=invalid_email)}):
            diagnostic = _refusal(
                lambda: messages.render("auth.login_code", recipient_name="Secret Display", auth_code="641203"),
                ("auth.login_code",),
            )
            assert "Secret Display" not in diagnostic and "641203" not in diagnostic
    with patch.dict(messages.CATALOG, {"auth.login_code": replace(original, subject="Invalid\nHeader")}):
        _refusal(lambda: messages.render("auth.login_code", **_context("auth.login_code")), ("auth.login_code",))


def test_dormant_templates_and_removed_definitions_have_explicit_dispositions():
    assert messages.CATALOG["newspaper.article"].status == "dormant"
    assert set(messages.UNUSED_LEGACY_KEYS) == {"mailEarlyTask", "textEarlyTask"}
    mapped = [legacy for entry in messages.CATALOG.values() if isinstance(entry, messages.MessageTemplate) for legacy in entry.legacy_keys]
    assert len(mapped) == len(set(mapped)), "every legacy text definition has one catalog destination"
    assert messages.LEGACY_RECIPIENT_KEY not in mapped, "recipient metadata must not be treated as copy"


if __name__ == "__main__":
    h.run_all(dict(globals()))
