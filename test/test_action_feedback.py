"""Shared HTML feedback and progressive enhancement, with synthetic accounts only."""
import contextlib
import re
from datetime import date
from pathlib import Path

from lxml import html

import helpers as h
import db
from test_auth import (
    _capture_messages, _restore_messages, _csrf, _challenge, _code,
    _request_login, _confirm_login,
)
from test_person_birth_dates import _roster, _signed_client


@contextlib.contextmanager
def _messages():
    messages, originals = _capture_messages()
    try:
        yield messages
    finally:
        _restore_messages(originals)


def _entries(response):
    page = html.fromstring(response.get_data(as_text=True))
    return page.xpath('//*[@id="feedback"]/*[@data-feedback-message]')


def _result(response, severity, text):
    entries = _entries(response)
    assert len(entries) == 1, "a logical action must appear once in the shared region"
    assert entries[0].get("data-severity") == severity
    assert text in entries[0].text_content()
    assert not entries[0].xpath('.//button'), "server fallback must contain no inert dismiss controls"
    return entries[0]


def test_shared_fallback_escapes_text_on_schedule_management_and_auth_pages():
    app, engine, ids = _roster()
    client = _signed_client(app, ids, "admin")
    for path in ("/", "/personen", "/login", "/registrieren"):
        with client.session_transaction() as session:
            session['_flashes'] = [('success', '<img src=x onerror=alert(1)>')]
        response = client.get(path)
        entry = _result(response, 'success', '<img src=x onerror=alert(1)>')
        assert not entry.xpath('.//img'), "display names must never become markup"
        assert not _entries(client.get(path)), "server flash replayed after consumption"
        assert 'class="form-message' not in response.get_data(as_text=True)
    style = Path(h.PROJECT_DIR, 'static/style.css').read_text()
    assert '.feedback-enhanced{position:fixed' in style
    assert '.feedback{margin:' in style, "no-JS region must remain in document flow"
    assert '@media(prefers-reduced-motion:reduce)' in style


def test_management_results_fields_and_private_values_remain_separate():
    app, engine, ids = _roster()
    admin = _signed_client(app, ids, 'admin')
    created = admin.post('/personen/add', data=h.csrf_data({
        'name': '<b>Feedback Helper</b>', 'birth_date': '1990-01-01',
        'team_ids': [ids['own_team']],
    }), follow_redirects=True)
    entry = _result(created, 'success', 'wurde angelegt')
    assert '<b>Feedback Helper</b>' in entry.text_content() and not entry.xpath('.//b')
    with h.Session(engine) as session:
        person_id = session.query(db.Person).filter_by(name='<b>Feedback Helper</b>').one().id
    for action, text in [('edit', 'Daten gespeichert'), ('teams', 'Mannschaften'),
                         ('deactivate', 'deaktiviert'), ('reactivate', 'reaktiviert'),
                         ('delete', 'gelöscht')]:
        response = admin.post(f'/personen/{person_id}/{action}', data=h.csrf_data({
            'name': '<b>Feedback Helper</b>', 'birth_date': '1990-01-01', 'team_ids': [ids['own_team']],
        }), follow_redirects=True)
        _result(response, 'warning' if action == 'edit' else 'success', text)
    invalid = admin.post('/personen/add', data=h.csrf_data({
        'name': 'Invalid', 'email': 'invalid', 'phone': '123', 'birth_date': '9999-01-01',
        'team_ids': [ids['own_team']],
    }))
    assert invalid.status_code == 400 and not _entries(invalid)
    page = html.fromstring(invalid.get_data(as_text=True))
    for field in ('email', 'phone', 'birth_date'):
        assert page.xpath(f'//*[@id="person-new-{field}-error"]')
        assert page.xpath(f'//input[@name="{field}" and @aria-invalid="true"]')
    member = _signed_client(app, ids, 'member')
    refused = member.post(f'/personen/{ids["other"]}/edit', data=h.csrf_data({
        'email': 'invalid', 'birth_date': '1999-12-31',
    }))
    assert refused.status_code == 403 and not refused.is_json
    _result(refused, 'error', 'Berechtigung')
    assert '1999-12-31' not in refused.get_data(as_text=True)
    invalid_self = member.post(f'/personen/{ids["member"]}/edit', data=h.csrf_data({
        'birth_date': '9999-01-01', 'email': 'invalid',
    }))
    assert invalid_self.status_code == 400
    page = html.fromstring(invalid_self.get_data(as_text=True))
    assert page.xpath(f'//*[@id="person-{ids["member"]}-birth_date-error"]')
    assert '1999-06-07' not in invalid_self.get_data(as_text=True), "foreign birth date leaked through validation"
    script = Path(h.PROJECT_DIR, 'static/app.js').read_text()
    assert 'wirklich löschen?' in script and 'Alle Diensteinträge werden entfernt' in script


def test_registration_decisions_and_form_refusals_keep_rights_and_html_contract():
    app, engine, ids = _roster()
    admin = _signed_client(app, ids, 'admin')
    mv = _signed_client(app, ids, 'mv')
    forbidden = mv.post(f'/personen/{ids["mv"]}/teams', data=h.csrf_data())
    assert forbidden.status_code == 403 and not forbidden.is_json
    _result(forbidden, 'error', 'MV')
    malformed = admin.post(f'/registrierungen/{ids["pending"]}/invalid', data=h.csrf_data())
    assert malformed.status_code == 400 and not malformed.is_json
    _result(malformed, 'error', 'Entscheidung')
    with h.Session(engine) as session:
        session.get(db.Person, ids['pending']).teams = [session.get(db.Team, ids['own_team'])]
        session.commit()
    with _messages():
        approved = admin.post(f'/registrierungen/{ids["pending"]}/approve', data=h.csrf_data(), follow_redirects=True)
    _result(approved, 'success', 'freigegeben')
    with h.Session(engine) as session:
        rejected = db.Person(name='Rejected Feedback', account_status=db.ACCOUNT_VERIFIED)
        session.add(rejected); session.commit(); rejected_id = rejected.id
    response = admin.post(f'/registrierungen/{rejected_id}/reject', data=h.csrf_data(), follow_redirects=True)
    _result(response, 'success', 'abgelehnt')
    assert not _entries(admin.get('/personen')), 'decision confirmation replayed'
    api = mv.post('/api/teams/999999/mv', json={'person_id': None}, headers=h.csrf_headers())
    assert api.status_code == 403 and api.get_json()['ok'] is False
    csrf = admin.post('/personen/add', data={})
    assert csrf.status_code == 403 and not csrf.is_json


def test_generic_code_feedback_and_login_logout_are_one_use_without_javascript():
    app, engine, ids = _roster()
    with h.Session(engine) as session:
        session.get(db.Person, ids['member']).email = 'feedback@example.test'
        inactive = db.Person(name='Inactive Feedback', email='inactive-feedback@example.test', account_status=db.ACCOUNT_INACTIVE)
        session.add(inactive); session.commit()
    with _messages() as messages:
        client = app.test_client(); csrf = _csrf(client)
        known = _request_login(client, csrf, contact='feedback@example.test')
        request_text = _result(known, 'info', 'Falls die Angaben bekannt sind').text_content()
        challenge = _challenge(known); code = _code(messages[-1])
        for address in ('unknown-feedback@example.test', 'inactive-feedback@example.test'):
            response = _request_login(client, csrf, contact=address)
            assert _result(response, 'info', 'Falls die Angaben bekannt sind').text_content() == request_text
        # Repeated requests exceed the contact limit; the same informational shape remains.
        for _ in range(12):
            response = _request_login(client, csrf, contact='inactive-feedback@example.test')
            assert _result(response, 'info', 'Falls die Angaben bekannt sind').text_content() == request_text
        response = _confirm_login(client, csrf, challenge, code)
        assert response.status_code == 302
        _result(client.get(response.location), 'success', 'angemeldet')
        assert not _entries(client.get('/'))
        with client.session_transaction() as session:
            session['_flashes'] = [('error', 'obsolete')]
            csrf = session['csrf_token']
        logout = client.post('/logout', data={'csrf_token': csrf}, follow_redirects=True)
        _result(logout, 'success', 'abgemeldet')
        assert 'obsolete' not in logout.get_data(as_text=True)
        assert 'data-auth-boundary="true"' in logout.get_data(as_text=True)
        with client.session_transaction() as session:
            assert 'person_id' not in session
    page = html.fromstring(known.get_data(as_text=True))
    assert page.xpath('//form[@method="post"]//input[@name="challenge"]')
    assert page.xpath('//input[@name="code" and not(@disabled)]')
    assert '15 Minuten' in page.text_content()
    assert 'class="auth-message"' not in known.get_data(as_text=True)
    invalid = client.post('/login', data={'csrf_token': _csrf(client), 'email': 'invalid', 'channel': 'email'})
    assert 'id="email-error"' in invalid.get_data(as_text=True) and not _entries(invalid)


def test_contact_verification_confirms_action_but_preserves_approval_status():
    app, engine, ids = _roster()
    client = app.test_client()
    with _messages() as messages:
        requested = client.post('/registrieren', data={
            'csrf_token': _csrf(client, '/registrieren'), 'name': 'Verified Feedback',
            'email': 'verified-feedback@example.test', 'channel': 'email',
            'birth_date': '2001-02-03', 'team_ids': [ids['own_team']], 'consent': 'yes',
        })
        _result(requested, 'info', 'Falls die Angaben verwendet werden können')
        with client.session_transaction() as session:
            csrf = session['csrf_token']
        completed = client.post('/registrieren', data={
            'csrf_token': csrf, 'action': 'confirm_code', 'challenge': _challenge(requested), 'code': _code(messages[-1]),
        }, follow_redirects=True)
    _result(completed, 'success', 'Kontakt bestätigt')
    assert 'Deine Registrierung wartet auf Freigabe.' in completed.get_data(as_text=True)
    assert client.get('/personen').status_code == 302
    status = client.get('/registrierung/status')
    assert not _entries(status) and 'wartet auf Freigabe' in status.get_data(as_text=True)


if __name__ == '__main__':
    h.run_all(dict(globals()))
