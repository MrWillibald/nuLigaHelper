"""Visibility is shared presentation scope, independent of binding duties."""

from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
import os
from pathlib import Path
import tempfile
from threading import Barrier, Event
import time
from unittest.mock import patch
from zoneinfo import ZoneInfo

from lxml import html
from sqlalchemy import select

import helpers as h
import db
import notifier
import webapp

TODAY = date(2026, 10, 24)
BOUNDARY = date(2026, 12, 1)
SETTING_URL = '/spielplan/sichtbarkeit'


def site():
    path = Path(h._TEST_DIR) / f'visibility-{next(tempfile._get_candidate_names())}.db'
    engine = db.make_engine(str(path))
    db.initialize_db(engine)
    with patch.dict(os.environ, {'NULIGAHELPER_DB': str(path)}):
        app = webapp.create_app()
    with h.Session(engine) as session:
        support = db.get_support_team(session)
        people = {role: db.Person(name=f'Fixture {role}', birth_date=h.ADULT_BIRTH_DATE,
                                  email=f'{role}@example.test', teams=[support],
                                  is_admin=role == 'admin',
                                  account_status=db.ACCOUNT_VERIFIED if role == 'pending' else db.ACCOUNT_ACTIVE)
                  for role in ('admin', 'mv', 'member', 'replacement', 'pending', 'private')}
        people['young'] = db.Person(name='Fixture young', birth_date=date(2015, 1, 1))
        session.add_all(people.values())
        session.flush()
        support.mv_person = people['mv']
        base = h.sample_games()[0]
        games = [{**base, 'game_nr': str(9100 + index), 'date': value, 'guest': f'Opponent {index}'}
                 for index, value in enumerate(('01.09.2026', '30.11.2026', '01.12.2026', '01.01.2027', 'invalid'))]
        games += [{**base, 'game_nr': '9200', 'date': '30.11.2026', 'time': '20:00', 'ak': 'BL mD', 'guest': 'Late opponent'}]
        db.sync_games(session, games, h.SEASON)
        stored = {game.game_nr: game for game in session.query(db.Game).filter_by(season_year=h.SEASON).all()}
        for game in stored.values():
            game.team = support
        db.claim_slot(session, stored['9102'], db.ROLE_SALE, 0, None, people['member'])
        db.request_replacement(session, stored['9102'], db.ROLE_SALE, 0, people['member'].id)
        # Retained youth Kasse remains part of season-wide personal counts.
        session.add(db.Assignment(game=stored['9200'], person=people['member'], role=db.ROLE_CASH, slot=0))
        # A saved age deficiency remains visible only in outstanding scope.
        session.add(db.Assignment(game=stored['9101'], person=people['young'], role=db.ROLE_SECURITY, slot=0))
        cake = next(b for b in db.get_day_blocks(session, h.SEASON, '01.12.2026') if b.phase == db.BLOCK_CAKE_DELIVERY)
        db.configure_cake_block(session, cake, '12:00', 2, None, None)
        db.claim_block_slot(session, cake, 0, None, people['member'])
        session.commit()
        ids = {key: value.id for key, value in people.items()}
        ids.update({key: value.id for key, value in stored.items()})
        ids['cake'] = cake.id
        ids['support'] = support.id
        ids['playing'] = session.query(db.Team).filter_by(name='BL F').one().id
    return app, engine, ids


def client(app, person_id=None):
    browser = app.test_client()
    if person_id:
        h.sign_in(browser, person_id)
    return browser


def snapshot(engine, include_setting=False):
    with engine.connect() as connection:
        return {table.name: [tuple(row) for row in connection.execute(select(table))]
                for table in db.Base.metadata.sorted_tables
                if include_setting or table.name != 'schedule_visibility_setting'}


def set_boundary(engine, ids, value='2026-12-01', expected=''):
    with h.Session(engine) as session, patch('common.effective_today', return_value=TODAY):
        result = db.update_schedule_visibility(session, actor_id=ids['admin'], hide_from=value,
                                              expected_hide_from=expected)
        session.commit()
        return result


def page(browser, path='/', today=TODAY):
    with patch('common.effective_today', return_value=today):
        response = browser.get(path)
    assert response.status_code == 200, (path, response.status_code)
    return html.fromstring(response.data)


def save(browser, value='2026-12-01', expected='', **extra):
    with patch('common.effective_today', return_value=TODAY):
        return browser.post(SETTING_URL, data=h.csrf_data(
            {'hide_from': value, 'expected_hide_from': expected, **extra}), follow_redirects=True)


def test_date_policy_uses_calendar_dates_and_keeps_unknown_dates():
    for boundary, visible, hidden in (
        (date(2026, 12, 1), '30.11.2026', '01.12.2026'),
        (date(2027, 1, 1), '31.12.2026', '01.01.2027'),
        (date(2028, 3, 1), '29.02.2028', '01.03.2028'),
    ):
        assert db.schedule_date_visible(visible, boundary)
        assert not db.schedule_date_visible(hidden, boundary)
        assert not db.schedule_date_visible('02.03.2028', boundary)
        for invalid in ('', None, 'invalid', '29.02.2027'):
            assert db.schedule_date_visible(invalid, boundary)
        assert db.schedule_date_visible(hidden, None)
    with patch('common.effective_today', return_value=date(2027, 8, 1)):
        assert not db.schedule_date_visible('01.08.2027', BOUNDARY), 'fixed boundary never rolls to a new season'


def test_setting_is_durable_validated_and_allows_clear_and_expired_noop():
    _app, engine, ids = site()
    before = snapshot(engine)
    assert set_boundary(engine, ids)
    for invalid in ('2026-10-24', '2026-10-23', '2026-02-30', 'bad', '20261201'):
        with h.Session(engine) as session, patch('common.effective_today', return_value=TODAY):
            try:
                db.update_schedule_visibility(session, actor_id=ids['admin'], hide_from=invalid,
                                              expected_hide_from='2026-12-01')
            except ValueError:
                session.rollback()
            else:
                raise AssertionError('changed dates require a canonical valid future date')
            assert db.schedule_visibility_boundary(session) == BOUNDARY
    with h.Session(engine) as session, patch('common.effective_today', return_value=date(2027, 8, 1)):
        assert not db.update_schedule_visibility(session, actor_id=ids['admin'], hide_from='2026-12-01', expected_hide_from='2026-12-01')
        session.commit()
        assert db.update_schedule_visibility(session, actor_id=ids['admin'], hide_from='', expected_hide_from='2026-12-01')
        session.commit()
        assert db.schedule_visibility_boundary(session) is None
    assert snapshot(engine) == before, 'settings must not touch duty, lock, request or audit state'


def test_setting_form_prg_validation_conflict_clear_and_no_javascript():
    app, engine, ids = site()
    admin = client(app, ids['admin'])
    initial = page(admin)
    assert initial.xpath('//label[@for="hide-from"]/text()') == ['Heimspieltage ausblenden ab']
    assert initial.xpath('//input[@name="expected_hide_from"]/@value') == ['']
    before = snapshot(engine)
    response = save(admin, 'bad')
    assert response.status_code == 200 and response.history[0].status_code == 303
    tree = html.fromstring(response.data)
    assert tree.xpath('//input[@id="hide-from"]/@aria-invalid') == ['true']
    assert 'gültiges Datum' in tree.text_content()
    assert 'Sichtbarkeit gespeichert.' in save(admin).get_data(as_text=True)
    assert page(client(app)).xpath('//div[contains(@class,"visibility-scope")]')
    stale = save(admin, '2027-01-01', '')
    assert 'zwischenzeitlich geändert' in stale.get_data(as_text=True)
    assert html.fromstring(stale.data).xpath('//input[@name="expected_hide_from"]/@value') == ['2026-12-01']
    assert 'Sichtbarkeit gespeichert.' in save(admin, '', '2026-12-01').get_data(as_text=True)
    assert snapshot(engine) == before


def test_tiers_csrf_missing_fields_and_fresh_admin_checks_refuse_without_writes():
    app, engine, ids = site()
    before = snapshot(engine, True)
    for role in (None, 'member', 'mv', 'pending'):
        browser = client(app, ids[role] if role else None)
        if not role:
            page(browser)
        response = save(browser)
        assert response.status_code in (200, 403), 'authentication redirects follow existing account policy'
        assert snapshot(engine, True) == before
    admin = client(app, ids['admin'])
    for data in ({'hide_from': '2026-12-01', 'expected_hide_from': ''},
                 h.csrf_data({'hide_from': '2026-12-01', 'expected_hide_from': ''}, token='wrong')):
        assert admin.post(SETTING_URL, data=data).status_code == 403
    for data in ({'hide_from': '2026-12-01'}, {'expected_hide_from': ''}):
        result = admin.post(SETTING_URL, data=h.csrf_data(data), follow_redirects=True)
        assert 'unvollständig' in result.get_data(as_text=True)
        assert snapshot(engine, True) == before
    # The writer must refresh a previously loaded admin even inside one session.
    with h.Session(engine) as reader:
        cached_admin = reader.get(db.Person, ids['admin'])
        with h.Session(engine) as writer:
            writer.get(db.Person, ids['admin']).is_admin = False
            writer.commit()
        assert cached_admin.is_admin, 'hold the old identity-map record until the serialized recheck'
        try:
            db.update_schedule_visibility(reader, actor_id=ids['admin'], hide_from='2026-12-01', expected_hide_from='')
        except ValueError as exc:
            assert 'aktive Administratoren' in str(exc)
            reader.rollback()
        else:
            raise AssertionError('a cached administrator record must not authorize writes')
    assert admin.post(SETTING_URL, data=h.csrf_data({'hide_from': '2026-12-01', 'expected_hide_from': ''})).status_code == 403
    with h.Session(engine) as session:
        assert db.schedule_visibility_boundary(session) is None
        session.get(db.Person, ids['admin']).account_status = db.ACCOUNT_INACTIVE
        session.commit()
    assert admin.post(SETTING_URL, data=h.csrf_data()).status_code == 302


def test_concurrent_writers_and_stale_wal_snapshot_cannot_overwrite_new_boundary():
    _app, engine, ids = site()
    barrier = Barrier(2)

    def write(value):
        with h.Session(engine) as session, patch('common.effective_today', return_value=TODAY):
            session.connection().exec_driver_sql('BEGIN')
            assert db.schedule_visibility_boundary(session) is None
            barrier.wait(timeout=5)
            try:
                db.update_schedule_visibility(session, actor_id=ids['admin'], hide_from=value, expected_hide_from='')
                session.commit()
                return 'saved'
            except ValueError as exc:
                session.rollback()
                assert 'zwischenzeitlich' in str(exc)
                return 'conflict'

    with ThreadPoolExecutor(max_workers=2) as pool:
        outcomes = list(pool.map(write, ('2026-12-01', '2027-01-01')))
    assert sorted(outcomes) == ['conflict', 'saved']
    assert snapshot(engine)['assignment_audit'], 'existing audit evidence remains present'


def test_setting_write_contention_is_bounded_and_does_not_report_success():
    app, engine, ids = site()
    admin = client(app, ids['admin'])
    started = Event()
    before = snapshot(engine, True)
    with engine.connect() as locker:
        locker.exec_driver_sql('BEGIN IMMEDIATE')
        def request_save():
            started.set()
            return save(admin)
        start = time.monotonic()
        with ThreadPoolExecutor(max_workers=1) as pool:
            future = pool.submit(request_save)
            assert started.wait(2)
            response = future.result(timeout=db.SQLITE_TIMEOUT_SECONDS + 4)
        assert time.monotonic() - start < db.SQLITE_TIMEOUT_SECONDS + 4
        text = response.get_data(as_text=True)
        assert 'Sichtbarkeit gespeichert.' not in text and 'Sichtbarkeit unverändert.' not in text
        assert 'erneut' in text or 'vorübergehend' in text
        locker.rollback()
    assert snapshot(engine, True) == before


def test_schedule_hides_complete_dates_before_filters_but_keeps_invalid_dates_and_times():
    app, engine, ids = site()
    set_boundary(engine, ids)
    guest = client(app)
    before = snapshot(engine, True)
    tree = page(guest)
    text = tree.text_content()
    assert 'Opponent 1' in text and 'Opponent 4' in text
    assert 'Opponent 2' not in text and 'Opponent 3' not in text
    assert 'Dezember 2026' not in text and 'Januar 2027' not in text and 'Ohne Datum' in text
    assert tree.xpath('//select[@name="date"]/option/@value') == ['', '01.09.2026', '30.11.2026', 'invalid']
    for phase in db.BLOCK_PHASES:
        blocks = tree.xpath(f'//details[contains(@class,"task-block-{phase}")]')
        assert len(blocks) == 3, 'visible dates and unresolved dates retain each block phase'
    hidden = page(guest, '/?show_hidden=1&date=01.12.2026')
    assert 'Unbekannter Spieltag' in hidden.text_content() and 'Keine Spiele oder Tagesdienste' in hidden.text_content()
    assert not hidden.xpath('//details[contains(@class,"game-card")]')
    assert not hidden.xpath('//input[@name="hide_from"]')
    for query in ('/?date=30.11.2026&playing_team=team-' + str(ids['playing']),
                  '/?person=fixture%20young&date=30.11.2026',
                  f'/?responsible_team=team-{ids["support"]}&date=30.11.2026&person=fixture%20young'):
        filtered = page(guest, query)
        assert 'Opponent 1' in filtered.text_content() and 'Opponent 2' not in filtered.text_content()
        if 'playing_team' in query:
            assert '16:00' in filtered.text_content() and '21:00' in filtered.text_content(), 'block times use first and last full-date games'
    for private in ('@example.test', '1990-01-01', '2015-01-01', 'person_id', 'Fixture private'):
        assert private not in html.tostring(tree, encoding='unicode')
    assert snapshot(engine, True) == before, 'filter and copied-override GETs must be read-only'


def test_admin_override_navigation_filters_clear_and_revoked_rights():
    app, engine, ids = site()
    set_boundary(engine, ids)
    for role in (None, 'member', 'mv'):
        browser = client(app, ids[role] if role else None)
        tree = page(browser, '/?show_hidden=1')
        assert 'Opponent 2' not in tree.text_content()
        assert 'Ausgeblendete Spieltage anzeigen' not in tree.text_content()
        if role:
            assert '01.12.2026' not in page(browser, '/statistik?show_hidden=1').xpath('//td[@data-label="Termin"]/text()')
    admin = client(app, ids['admin'])
    assert 'Opponent 2' not in page(admin).text_content(), 'admins initially use shared scope'
    tree = page(admin, '/?show_hidden=1&person=Fixture')
    assert tree.xpath('//form[@aria-label="Spielplan filtern"]/input[@name="show_hidden"]/@value') == ['1']
    clear = tree.xpath('//a[text()="Filter löschen"]/@href')
    assert clear and all('show_hidden=1' in href for href in clear)
    assert tree.xpath('//nav//a[text()="Statistik"]/@href') == ['/statistik?show_hidden=1']
    assert 'Opponent 2' in page(admin, '/?show_hidden=1').text_content()
    stats = page(admin, '/statistik?show_hidden=1')
    assert stats.xpath('//nav//a[text()="Heimspielplan"]/@href') == ['/?show_hidden=1']
    with h.Session(engine) as session:
        session.get(db.Person, ids['admin']).is_admin = False
        session.commit()
    assert 'Opponent 2' not in page(admin, '/?show_hidden=1').text_content()


def test_retained_boundary_scopes_past_days_empty_view_and_new_season_without_reset():
    app, engine, ids = site()
    # Unknown dates stay visible, so remove only this synthetic erroneous empty record.
    with h.Session(engine) as session:
        session.delete(session.get(db.Game, ids['9104']))
        session.commit()
    set_boundary(engine, ids, '2026-10-25')
    guest, admin = client(app), client(app, ids['admin'])
    tree = page(guest, today=date(2027, 3, 1))
    assert tree.xpath('//details[@class="past-details"]/summary/text()') == ['Vergangene Spieltage (1)']
    expanded = page(admin, '/?show_hidden=1', date(2027, 3, 1))
    assert expanded.xpath('//details[@class="past-details"]/summary/text()') == ['Vergangene Spieltage (4)']
    selected = page(admin, '/?show_hidden=1&date=01.12.2026', date(2027, 3, 1))
    assert selected.xpath('//details[@class="past-details"]/@open'), 'selected past day must start expanded'
    with h.Session(engine) as session:
        # Hiding the sole past date is a permitted unchanged saved-value lifecycle.
        session.get(db.ScheduleVisibilitySetting, 1).hide_from = date(2026, 8, 1)
        session.commit()
    for browser in (guest, admin):
        empty = page(browser)
        assert 'Im angezeigten Zeitraum sind keine Spiele' in empty.text_content()
        assert 'Führe zuerst' not in empty.text_content()
    assert page(admin).xpath('//form[@aria-label="Sichtbarkeit der Heimspieltage"]')
    with h.Session(engine) as session:
        db.sync_games(session, [{**h.sample_games()[0], 'game_nr': '9999', 'date': '01.08.2027',
                                 'guest': 'Next season fixture'}], 2027)
        session.commit()
    assert 'Next season fixture' not in page(guest, today=date(2027, 8, 1)).text_content()
    with h.Session(engine) as session:
        assert db.schedule_visibility_boundary(session) == date(2026, 8, 1)


def test_statistics_outstanding_scope_counts_and_complete_season_totals():
    app, engine, ids = site()
    member = client(app, ids['member'])
    admin = client(app, ids['admin'])
    before = page(member, '/statistik')
    set_boundary(engine, ids)
    after = page(member, '/statistik')
    def sections(tree):
        return tree.xpath('//details[contains(@class,"statistics-disclosure")]')
    for index in (0, 1):
        assert html.tostring(sections(before)[index]) == html.tostring(sections(after)[index]), 'team shares, totals, retained roles and personal counts stay season-wide'
    gaps = sections(after)[2]
    assert '(6 betroffene Spiele oder Tagesblöcke)' in gaps.text_content(), 'two valid games, one invalid game and three blocks count once each'
    assert 'Altersanforderung' in gaps.text_content() and 'Admin-Einrichtung erforderlich' in gaps.text_content()
    assert '01.12.2026' not in gaps.text_content() and '01.01.2027' not in gaps.text_content()
    assert 'gesamte Saison' in after.text_content()
    full = sections(page(admin, '/statistik?show_hidden=1'))[2]
    assert html.tostring(full) == html.tostring(sections(before)[2])
    assert '01.09.2026' not in full.text_content(), 'full-date mode still excludes past issues'
    with h.Session(engine) as session:
        session.get(db.ScheduleVisibilitySetting, 1).hide_from = date(2026, 8, 1)
        session.delete(session.get(db.Game, ids['9104']))
        session.commit()
    empty = sections(page(member, '/statistik'))[2].text_content()
    assert '(0 betroffene Spiele oder Tagesblöcke)' in empty
    assert 'Im angezeigten Zeitraum gibt es keine offenen Dienste' in empty
    assert 'vollständig besetzt' not in empty
    assert client(app).get('/statistik?show_hidden=1').status_code == 302
    for private in ('1990-01-01', '2015-01-01', '@example.test'):
        assert private not in html.tostring(after, encoding='unicode')


def test_hidden_direct_candidates_claims_takeovers_and_html_confirmation_keep_existing_guards():
    app, engine, ids = site()
    set_boundary(engine, ids)
    member, replacement = client(app, ids['member']), client(app, ids['replacement'])
    game_id = ids['9102']
    before = snapshot(engine, True)
    candidates = replacement.get(f'/api/games/{game_id}/candidates')
    assert candidates.status_code == 200 and 'slots' in candidates.json
    assert snapshot(engine, True) == before
    with h.Session(engine) as session:
        token = session.query(db.Assignment).filter_by(game_id=game_id, role=db.ROLE_SALE, slot=0).one().replacement_token
    payload = {'game_id': game_id, 'role': db.ROLE_SALE, 'slot': 0, 'person_id': ids['replacement'],
               'expected_person_id': ids['member'], 'expected_request_token': token}
    response = replacement.post('/api/assignment/takeover', json=payload, headers=h.csrf_headers())
    assert response.status_code == 200 and response.json['ok'], 'hidden status must not prohibit an eligible transfer'
    assert member.post('/api/assignment/takeover', json={**payload, 'person_id': ids['member']}, headers=h.csrf_headers()).status_code == 409
    target = {'game_id': game_id, 'role': db.ROLE_SECURITY, 'slot': 0, 'person_id': ids['member'], 'expected_person_id': ''}
    before = snapshot(engine, True)
    review = member.post('/dienst', data=h.csrf_data({**target, 'decision': 'review'}))
    assert review.status_code == 200 and 'Bestätigen' in review.get_data(as_text=True)
    assert member.post('/dienst', data=h.csrf_data({**target, 'decision': 'cancel'})).status_code == 302
    assert snapshot(engine, True) == before, 'hidden-date review/cancel never writes'
    assert member.post('/dienst', data=h.csrf_data({**target, 'decision': 'confirm', 'action': 'claim'})).status_code == 302
    young = client(app, ids['young'])
    refused = young.post('/api/assignment/claim', json={'game_id': ids['9103'], 'role': db.ROLE_SECURITY, 'slot': 0,
                       'person_id': ids['young'], 'expected_person_id': None}, headers=h.csrf_headers())
    assert refused.status_code == 400, 'hidden claims still enforce age'
    with patch('db.duty_now', return_value=datetime(2026, 11, 30, 9, tzinfo=ZoneInfo('Europe/Berlin'))):
        locked = member.post('/api/assignment/request-replacement', json={'game_id': game_id, 'role': db.ROLE_SECURITY,
                      'slot': 0, 'expected_person_id': ids['member']}, headers=h.csrf_headers())
    assert locked.status_code == 400 and 'gesperrt' in locked.json['error'], 'hidden duties retain their cutoff'
    assert client(app).get(f'/api/games/{game_id}/candidates').status_code == 401


def test_hidden_block_and_game_json_claims_and_native_takeover_preserve_capacity_and_tokens():
    app, engine, ids = site()
    set_boundary(engine, ids)
    mv, member = client(app, ids['mv']), client(app, ids['member'])
    admin = client(app, ids['admin'])
    game = {'game_id': ids['9103'], 'role': db.ROLE_SECURITY, 'slot': 0,
            'person_id': ids['mv'], 'expected_person_id': None}
    response = mv.post('/api/assignment/claim', json=game, headers=h.csrf_headers())
    assert response.status_code == 200 and response.json['ok']
    assert mv.post('/api/assignment/claim', json=game, headers=h.csrf_headers()).status_code == 409
    second = {**game, 'role': db.ROLE_SECRETARY}
    assert mv.post('/api/assignment/claim', json=second, headers=h.csrf_headers()).status_code == 400
    with h.Session(engine) as session:
        preparation = next(b for b in db.get_day_blocks(session, h.SEASON, '01.12.2026')
                           if b.phase == db.BLOCK_PREPARATION)
        block_id = preparation.id
    assert mv.get(f'/api/blocks/{block_id}/candidates').status_code == 200
    claim = {'block_id': block_id, 'slot': 0, 'person_id': ids['member'], 'expected_person_id': None}
    assert member.post('/api/block-assignment/claim', json=claim, headers=h.csrf_headers()).json['ok']
    assert member.post('/api/block-assignment/claim', json={**claim, 'slot': 1}, headers=h.csrf_headers()).status_code == 400
    request = {'block_id': block_id, 'slot': 0, 'expected_person_id': ids['member']}
    assert member.post('/api/block-assignment/request-replacement', json=request, headers=h.csrf_headers()).json['ok']
    with h.Session(engine) as session:
        token = session.query(db.BlockAssignment).filter_by(block_id=block_id, slot=0).one().replacement_token
    takeover = {**request, 'person_id': ids['admin'], 'expected_request_token': token, 'action': 'takeover'}
    before = snapshot(engine, True)
    assert admin.post('/dienst', data=h.csrf_data({**takeover, 'decision': 'review'})).status_code == 200
    assert snapshot(engine, True) == before
    assert admin.post('/dienst', data=h.csrf_data({**takeover, 'decision': 'confirm'})).status_code == 302
    assert admin.post('/dienst', data=h.csrf_data({**takeover, 'decision': 'confirm'})).status_code == 409
    with h.Session(engine) as session:
        assignment = session.query(db.BlockAssignment).filter_by(block_id=block_id, slot=0).one()
        assert assignment.person_id == ids['admin'] and assignment.replacement_token is None
        for action in ('takeover_in', 'takeover_out'):
            assert session.query(db.AssignmentAudit).filter_by(block_id=block_id, action=action).count() == 1
        # Admin and MV rights form a union for this setting as well.
        session.get(db.Team, ids['support']).mv_person_id = ids['admin']
        session.commit()
    assert 'Sichtbarkeit gespeichert.' in save(admin, '2027-01-01', '2026-12-01').get_data(as_text=True)


def test_setting_and_preview_do_not_lock_hidden_duties_but_notifier_and_due_lock_still_do():
    from test_notifier import TextRecorder
    app, engine, ids = site()
    before = snapshot(engine)
    due = datetime(2026, 11, 30, 9, tzinfo=ZoneInfo('Europe/Berlin'))
    with patch('db.duty_now', return_value=due), \
         patch.object(notifier.Notifier, 'send_Mail', side_effect=AssertionError('unexpected mail')), \
         patch.object(notifier.Notifier, 'send_SMS', side_effect=AssertionError('unexpected SMS')):
        admin = client(app, ids['admin'])
        assert 'Sichtbarkeit gespeichert.' in save(admin).get_data(as_text=True)
        for path in ('/', '/?show_hidden=1', '/?date=01.12.2026', '/statistik', '/statistik?show_hidden=1'):
            page(admin, path)
        assert snapshot(engine) == before, 'settings and preview must never materialize due locks'
    with h.Session(engine) as session:
        sender = notifier.Notifier(h.load_club_config(), session, h.SEASON)
        recorder = TextRecorder()
        sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
        assert sender.notify_pre('01.12.2026') == 1, 'hidden game duties still notify'
        assert sender.notify_cakes_early('01.12.2026') == 1, 'hidden cake duties still notify'
        assert len(recorder.mails) == 2
        with patch('db.duty_now', return_value=due):
            db.lock_due_duties(session)
            session.commit()
        duty = session.query(db.Assignment).filter_by(game_id=ids['9102']).one()
        assert duty.locked_at is not None and duty.replacement_token is None
        assert session.query(db.AssignmentAudit).filter_by(game_id=ids['9102'], action='lock').count() == 1


if __name__ == '__main__':
    h.run_all(dict(globals()))
