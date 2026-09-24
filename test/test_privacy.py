"""Synthetic policy gates and transactional privacy cleanup regression tests."""
import copy
from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
from unittest.mock import patch

import helpers as h
import db
import privacy
import production as p

NOW = datetime(2026, 9, 16, tzinfo=timezone.utc)
NAIVE = NOW.replace(tzinfo=None)


def policy():
    return {
        'schema_version': 2,
        'version': 'synthetic-v1',
        'cleanup': {**{key: 3600 for key in privacy.RULES}, 'batch_size': 2},
    }


def test_policy_requires_valid_retention_periods():
    data = policy()
    assert privacy.validate_cleanup(data)
    for key in privacy.RULES:
        bad = copy.deepcopy(data)
        bad['cleanup'][key] = None
        try: privacy.validate_cleanup(bad)
        except p.ConfigurationError: pass
        else: raise AssertionError('missing retention accepted: ' + key)


def seed():
    engine = h.make_engine()
    with h.Session(engine) as session:
        support = db.get_support_team(session)
        for status in ('registered','verified','rejected','active','inactive'):
            person = db.Person(name=status,account_status=status,registered_at=NAIVE-timedelta(days=2),
                verified_at=NAIVE-timedelta(days=2),rejected_at=NAIVE-timedelta(days=2),
                approved_at=NAIVE-timedelta(days=1) if status in {'active','inactive'} else None)
            if status == 'verified':
                person.teams = [support]
            session.add(person)
        active = db.Person(name='Audit owner', account_status='active')
        session.add(active)
        session.flush()
        session.add(db.AssignmentAudit(actor_person_id=active.id,affected_person_id=active.id,
            actor_tier='admin',action='claim',role=db.ROLE_SALE,slot=0,
            actor_name='Audit owner',affected_person_name='Audit owner',game_snapshot='Synthetic game'))
        for index, age in enumerate([7201,7202,7203,3600,3599]):
            session.add(db.AuthToken(nonce='synthetic-'+str(index),code='123456',purpose='login',
                person_id=active.id,issued_at=NAIVE-timedelta(days=2),expires_at=NAIVE-timedelta(seconds=age)))
            session.add(db.AuthAbuseCounter(action='login_request',dimension='contact',subject_digest=str(index)*64,
                channel='email',window_started_at=NAIVE-timedelta(days=2),count=1,
                expires_at=NAIVE-timedelta(seconds=age)))
        session.commit()
    return engine


def snapshots(engine):
    with engine.connect() as connection:
        return {table.name: list(connection.execute(table.select()).all()) for table in
                (db.Person.__table__,db.person_teams,db.AuthToken.__table__,db.AuthAbuseCounter.__table__,db.AssignmentAudit.__table__)}


def test_cleanup_is_bounded_idempotent_and_preserves_audits_and_approved_people():
    engine = seed()
    data = policy()
    before = snapshots(engine)
    preview = privacy.cleanup(engine,data,now=NOW)
    assert snapshots(engine) == before, 'preview changed records'
    assert preview['counts']['auth_tokens'] == preview['counts']['abuse_counters'] == 2
    assert all(preview['counts'][name] == 1 for name in ('registered','verified','rejected'))
    applied = privacy.cleanup(engine,data,now=NOW,apply=True)
    assert applied['counts'] == preview['counts']
    after = snapshots(engine)
    assert after['assignment_audit'] == before['assignment_audit']
    with h.Session(engine) as session:
        assert {x.name for x in session.query(db.Person)} == {'active','inactive','Audit owner'}
    second = privacy.cleanup(engine,data,now=NOW,apply=True)
    assert second['counts']['auth_tokens'] == second['counts']['abuse_counters'] == 1
    assert not any(privacy.cleanup(engine,data,now=NOW,apply=True)['counts'].values())
    assert snapshots(engine)['assignment_audit'] == before['assignment_audit']
    with h.Session(engine) as session:
        assert session.query(db.AuthToken).count() == 2, 'exact boundary and newer tokens must survive'
    output = json.dumps(applied)
    for canary in ('123456','Audit owner','synthetic-0'):
        assert canary not in output


def test_interruption_rolls_back_all_cleanup_classes():
    engine = seed()
    before = snapshots(engine)
    with patch('sqlalchemy.engine.Connection.commit',side_effect=OSError('synthetic interruption')):
        try: privacy.cleanup(engine,policy(),now=NOW,apply=True)
        except OSError: pass
        else: raise AssertionError('interruption did not propagate')
    assert snapshots(engine) == before


def test_registration_referenced_by_audit_is_preserved_without_audit_edits():
    engine = seed()
    with h.Session(engine) as session:
        person = session.query(db.Person).filter_by(account_status='rejected').one()
        session.add(db.AssignmentAudit(actor_person_id=None,affected_person_id=person.id,
            actor_tier='system',action='release',role=db.ROLE_SALE,slot=0,
            actor_name='System',affected_person_name='Historic name',game_snapshot='Synthetic game'))
        session.commit()
    before = snapshots(engine)['assignment_audit']
    assert privacy.cleanup(engine,policy(),now=NOW,apply=True)['counts']['rejected'] == 0
    assert snapshots(engine)['assignment_audit'] == before


if __name__ == '__main__':
    h.run_all(globals())
