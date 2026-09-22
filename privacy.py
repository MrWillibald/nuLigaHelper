"""Bounded, aggregate-only cleanup under configured retention periods.

Production processes run in UTC. Cutoffs never shorten the abuse service's
security expiry; routine cleanup never updates assignment audit records.
"""
from datetime import datetime, timedelta, timezone
import argparse
import json
import os
from pathlib import Path

from sqlalchemy import select, delete, func, exists, or_

import db
import auth_abuse
import production as p


RULES = ('auth_grace_seconds', 'abuse_grace_seconds', 'registered_seconds',
         'verified_seconds', 'rejected_seconds')


def validate_cleanup(policy):
    p.validate_policy(policy)
    rules = policy.get('cleanup', {})
    p.require(isinstance(rules, dict), 'cleanup')
    for name in RULES:
        p.number(rules.get(name), 'cleanup.' + name, 0 if 'grace' in name else 1)
    p.require(type(rules.get('batch_size')) is int and 1 <= rules['batch_size'] <= 1000, 'cleanup.batch_size')
    return rules


def cleanup(engine, policy, *, apply=False, now=None):
    rules = validate_cleanup(policy)
    stamp = now or datetime.now(timezone.utc)
    p.require(stamp.tzinfo is not None, 'cleanup.now')
    now_naive = stamp.astimezone(timezone.utc).replace(tzinfo=None)
    limit = rules['batch_size']
    cutoffs = {key: now_naive - timedelta(seconds=rules[key]) for key in RULES}
    counts = {}
    with engine.connect() as connection:
        # A single bounded write transaction prevents approvals racing cleanup.
        # Preview performs SELECTs only and rolls back without mutations.
        if apply:
            connection.exec_driver_sql('BEGIN IMMEDIATE')
        counts['abuse_counters'] = auth_abuse.cleanup_records(connection, cutoffs['abuse_grace_seconds'], limit, apply=apply)
        for label, table, cutoff in [('auth_tokens', db.AuthToken.__table__, cutoffs['auth_grace_seconds'])]:
            ids = list(connection.scalars(select(table.c.id).where(table.c.expires_at < cutoff)
                                           .order_by(table.c.expires_at, table.c.id).limit(limit)))
            counts[label] = len(ids)
            if apply and ids:
                connection.execute(delete(table).where(table.c.id.in_(ids)))
        persons = db.Person.__table__
        assignments = db.Assignment.__table__
        audits = db.AssignmentAudit.__table__
        teams = db.Team.__table__
        for status, column in [('registered', persons.c.registered_at),
                               ('verified', persons.c.verified_at), ('rejected', persons.c.rejected_at)]:
            predicate = (
                (persons.c.account_status == status) & (persons.c.approved_at.is_(None)) &
                (persons.c.is_admin.is_(False)) &
                (column < cutoffs[status + '_seconds']) &
                ~exists(select(assignments.c.id).where(assignments.c.person_id == persons.c.id)) &
                ~exists(select(teams.c.id).where(teams.c.mv_person_id == persons.c.id)) &
                ~exists(select(audits.c.id).where(or_(audits.c.actor_person_id == persons.c.id,
                                                     audits.c.affected_person_id == persons.c.id))))
            ids = list(connection.scalars(select(persons.c.id).where(predicate).order_by(column, persons.c.id).limit(limit)))
            counts[status] = len(ids)
            if apply and ids:
                connection.execute(delete(db.AuthToken).where(db.AuthToken.person_id.in_(ids)))
                connection.execute(delete(persons).where(persons.c.id.in_(ids)))
        if apply:
            connection.commit()
        else:
            connection.rollback()
    return {'policy_version': policy['version'], 'applied': apply, 'completed_at': stamp.isoformat(),
            'cutoffs': {k:v.isoformat() + '+00:00' for k,v in cutoffs.items()}, 'counts': counts}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--apply', action='store_true')
    parser.add_argument('--record-preview', action='store_true')
    parser.add_argument('--policy', default=os.environ.get('NULIGAHELPER_POLICY'))
    args = parser.parse_args(argv)
    try:
        policy = p.read_json(args.policy)
        path = os.environ.get('NULIGAHELPER_DB')
        p.require(bool(path) and db.database_ready(path), 'NULIGAHELPER_DB')
        receipt = Path(os.environ.get('NULIGAHELPER_STATE_DIR', '.')) / 'cleanup-preview.json'
        if args.apply:
            p.require(p.read_json(receipt).get('policy_sha256') == p.digest(policy), 'cleanup.preview_required')
        engine = db.make_engine(path)
        try:
            result = cleanup(engine, policy, apply=args.apply)
        finally:
            engine.dispose()
        print(json.dumps(result, sort_keys=True))
        if args.record_preview and not args.apply:
            import tempfile
            fd, temp = tempfile.mkstemp(prefix='.preview-', dir=receipt.parent)
            try:
                with os.fdopen(fd, 'w') as stream:
                    json.dump({'policy_sha256': p.digest(policy), 'preview': result}, stream)
                    stream.flush()
                    os.fsync(stream.fileno())
                os.replace(temp, receipt)
            finally:
                Path(temp).unlink(missing_ok=True)
        if args.apply:
            p.write_success('cleanup')
        return 0
    except Exception as error:
        print('cleanup=failed reason=' + (str(error) if isinstance(error,p.ConfigurationError) else type(error).__name__))
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
