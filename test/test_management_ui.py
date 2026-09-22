"""Person-management permissions, filtering and tier-specific presentation."""

import os
import tempfile

import helpers as h
import db
import webapp


_previous_db = os.environ["NULIGAHELPER_DB"]
_db_path = os.path.join(
    h._TEST_DIR, f"management-{next(tempfile._get_candidate_names())}.db"
)
os.environ["NULIGAHELPER_DB"] = _db_path
try:
    db.initialize_db(db.make_engine(_db_path))
    app = webapp.create_app()
finally:
    os.environ["NULIGAHELPER_DB"] = _previous_db
ENGINE = db.make_engine(_db_path)

with h.Session(ENGINE) as session:
    h.sync_sample_games(session)
    regular_teams = [team for team in db.get_all_teams(session) if not team.is_support]
    own_team, second_team, other_team = regular_teams[:3]
    support = db.get_support_team(session)
    admin = db.Person(
        name="Admin Person", email="admin@management.test", teams=[support], is_admin=True
    )
    mv = db.Person(name="Multi MV", email="mv@management.test", teams=[own_team, second_team])
    member = db.Person(name="Visible Member", email="member@management.test", teams=[own_team])
    other = db.Person(name="Other Active", email="private@management.test", teams=[other_team])
    inactive = db.Person(
        name="Hidden Inactive", teams=[other_team], account_status=db.ACCOUNT_INACTIVE
    )
    pending_own = db.Person(
        name="Pending Own", teams=[own_team], account_status=db.ACCOUNT_VERIFIED
    )
    pending_second = db.Person(
        name="Pending Second", teams=[second_team],
        account_status=db.ACCOUNT_VERIFIED,
    )
    pending_other = db.Person(
        name="Pending Other", teams=[other_team],
        account_status=db.ACCOUNT_VERIFIED,
    )
    pending_support = db.Person(
        name="Pending Support", teams=[support],
        account_status=db.ACCOUNT_VERIFIED,
    )
    session.add_all([
        admin, mv, member, other, inactive, pending_own, pending_second,
        pending_other, pending_support,
    ])
    session.flush()
    own_team.mv_person_id = mv.id
    second_team.mv_person_id = mv.id
    session.commit()
    IDS = {
        "admin": admin.id,
        "mv": mv.id,
        "member": member.id,
        "other": other.id,
        "inactive": inactive.id,
        "pending_own": pending_own.id,
        "pending_second": pending_second.id,
        "pending_other": pending_other.id,
        "pending_support": pending_support.id,
        "own_team": own_team.id,
        "second_team": second_team.id,
        "other_team": other_team.id,
        "support": support.id,
    }


def _client(person_key):
    client = app.test_client()
    token = h.sign_in(client, IDS[person_key])
    return client, token


def test_01_member_filters_only_the_visible_roster_without_contact_leaks():
    client, token = _client("member")
    page = client.get("/personen").get_data(as_text=True)
    assert "member@management.test" in page
    assert "private@management.test" not in page
    assert 'id="new-user-card"' not in page
    assert 'id="pending-registration-card"' not in page
    assert 'id="mv-assignment-card"' not in page
    assert client.get("/audit").status_code == 403
    assert client.post("/personen/add", data=h.csrf_data({
        "name": "Forbidden", "team_id": IDS["own_team"],
    }, token)).status_code == 403
    assert client.post(
        f"/personen/{IDS['member']}/teams",
        data=h.csrf_data({"team_ids": [IDS["second_team"]]}, token),
    ).status_code == 403

    by_name = client.get("/personen?name=other").get_data(as_text=True)
    assert "Other Active" in by_name and "Visible Member" not in by_name
    assert "private@management.test" not in by_name
    by_team = client.get(
        f"/personen?team_id={IDS['other_team']}"
    ).get_data(as_text=True)
    assert "Other Active" in by_team and "Visible Member" not in by_team
    for team_id in (IDS["own_team"], IDS["second_team"]):
        multi_team = client.get(f"/personen?team_id={team_id}").get_data(as_text=True)
        assert "Multi MV" in multi_team
    forged_status = client.get(
        "/personen?name=Hidden&status=inactive"
    ).get_data(as_text=True)
    assert "Hidden Inactive" not in forged_status


def test_02_mv_can_create_contactless_people_for_every_managed_team_only():
    client, token = _client("mv")
    page = client.get("/personen").get_data(as_text=True)
    assert 'id="new-user-card"' in page
    assert 'id="pending-registration-card"' not in page
    assert 'id="mv-assignment-card"' not in page
    assert client.get("/audit").status_code == 403
    form = page[page.index('id="new-user-card"'):page.index("</form>", page.index('id="new-user-card"'))]
    assert f'value="{IDS["own_team"]}"' in form
    assert f'value="{IDS["second_team"]}"' in form
    assert f'value="{IDS["other_team"]}"' not in form
    assert f'value="{IDS["support"]}"' not in form

    for name, team_key in (("Created Own", "own_team"), ("Created Second", "second_team")):
        response = client.post("/personen/add", data=h.csrf_data({
            "name": name, "team_id": IDS[team_key], "email": "", "phone": "",
        }, token))
        assert response.status_code == 302
    refused = client.post("/personen/add", data=h.csrf_data({
        "name": "Forged Other", "team_id": IDS["other_team"],
    }, token))
    assert refused.status_code == 403
    with h.Session(ENGINE) as session:
        created = session.query(db.Person).filter(
            db.Person.name.in_(["Created Own", "Created Second"])
        ).order_by(db.Person.name).all()
        assert len(created) == 2
        assert {db.membership_team_ids(person) for person in created} == {
            (IDS["own_team"],), (IDS["second_team"],),
        }
        assert all(
            person.account_status == db.ACCOUNT_ACTIVE
            and person.email is None and person.phone is None
            for person in created
        )
        assert session.query(db.Person).filter_by(name="Forged Other").first() is None


def test_02b_mv_changes_only_managed_active_rosters_and_cannot_remove_self():
    client, token = _client("mv")
    page = client.get("/personen").get_data(as_text=True)
    assert 'class="person-team-badge"' in page
    assert f'id="team-dialog-{IDS["other"]}"' in page
    assert 'data-team-dialog-open=' in page
    added = client.post(
        f"/personen/{IDS['other']}/teams",
        data=h.csrf_data({
            "team_ids": [IDS["own_team"], IDS["second_team"]],
        }, token),
    )
    assert added.status_code == 302
    with h.Session(ENGINE) as session:
        other = session.get(db.Person, IDS["other"])
        assert set(db.membership_team_ids(other)) == {
            IDS["other_team"], IDS["own_team"], IDS["second_team"]
        }

    removed = client.post(
        f"/personen/{IDS['other']}/teams",
        data=h.csrf_data(token=token),
    )
    assert removed.status_code == 302
    assert client.post(
        f"/personen/{IDS['mv']}/teams",
        data=h.csrf_data(token=token),
    ).status_code == 403
    assert client.post(
        f"/personen/{IDS['inactive']}/teams/{IDS['own_team']}/add",
        data=h.csrf_data(token=token),
    ).status_code == 403
    assert client.post(
        f"/personen/{IDS['member']}/teams/{IDS['other_team']}/add",
        data=h.csrf_data(token=token),
    ).status_code == 403


def test_03_registration_decisions_are_admin_only():
    mv_client, mv_token = _client("mv")
    page = mv_client.get("/personen").get_data(as_text=True)
    assert "Pending Own" not in page and "Pending Second" not in page
    assert "Pending Other" not in page and "Pending Support" not in page
    assert mv_client.post(
        f"/registrierungen/{IDS['pending_other']}/approve",
        data=h.csrf_data(token=mv_token),
    ).status_code == 403
    assert mv_client.post(
        f"/registrierungen/{IDS['pending_support']}/approve",
        data=h.csrf_data(token=mv_token),
    ).status_code == 403
    assert mv_client.post(
        f"/registrierungen/{IDS['pending_own']}/approve",
        data=h.csrf_data(token=mv_token),
    ).status_code == 403
    assert mv_client.post(
        f"/registrierungen/{IDS['pending_second']}/reject",
        data=h.csrf_data(token=mv_token),
    ).status_code == 403

    admin_client, admin_token = _client("admin")
    admin_page = admin_client.get("/personen").get_data(as_text=True)
    assert all(name in admin_page for name in (
        "Pending Own", "Pending Second", "Pending Other", "Pending Support"
    ))
    assert admin_client.post(
        f"/registrierungen/{IDS['pending_own']}/approve",
        data=h.csrf_data(token=admin_token),
    ).status_code == 302
    assert admin_client.post(
        f"/registrierungen/{IDS['pending_second']}/reject",
        data=h.csrf_data(token=admin_token),
    ).status_code == 302
    assert admin_client.post(
        f"/registrierungen/{IDS['pending_support']}/approve",
        data=h.csrf_data(token=admin_token),
    ).status_code == 302
    with h.Session(ENGINE) as session:
        assert session.get(db.Person, IDS["pending_own"]).account_status == db.ACCOUNT_ACTIVE
        assert session.get(db.Person, IDS["pending_second"]).account_status == db.ACCOUNT_REJECTED
        assert session.get(db.Person, IDS["pending_other"]).account_status == db.ACCOUNT_VERIFIED
        assert session.get(db.Person, IDS["pending_support"]).account_status == db.ACCOUNT_ACTIVE


def test_04_admin_has_all_management_cards_and_status_filtering():
    client, token = _client("admin")
    page = client.get("/personen").get_data(as_text=True)
    assert 'id="new-user-card"' in page
    assert 'id="pending-registration-card"' in page
    assert 'id="mv-assignment-card"' in page
    assert 'id="new-user-team-dialog"' in page
    assert 'data-new-team-badges' in page
    assert 'data-team-picker-apply' in page
    assert '<select id="new-team" name="team_ids"' not in page
    assert '<select name="team_ids" multiple' not in page[
        page.index('<div class="people-grid persons-grid">'):
        page.index('<div class="management-divider">')
    ]
    assert page.count('class="person-team-badge"') >= 1
    assert 'class="team-membership-dialog"' in page
    assert page.count(f'value="{IDS["mv"]}"') >= 2
    inactive = client.get("/personen?status=inactive").get_data(as_text=True)
    inactive_roster = inactive[
        inactive.index('<div class="people-grid persons-grid">'):
        inactive.index('<div class="management-divider">')
    ]
    assert "Hidden Inactive" in inactive_roster and "Other Active" not in inactive_roster
    assert '<option value="inactive" selected>' in inactive
    active = client.get("/personen?status=active").get_data(as_text=True)
    active_roster = active[
        active.index('<div class="people-grid persons-grid">'):
        active.index('<div class="management-divider">')
    ]
    assert "Other Active" in active_roster and "Hidden Inactive" not in active_roster
    created = client.post("/personen/add", data=h.csrf_data({
        "name": "Admin Other", "team_ids": [IDS["other_team"], IDS["support"]],
    }, token))
    assert created.status_code == 302
    assert client.post(
        f"/personen/{IDS['member']}/teams",
        data=h.csrf_data({
            "team_ids": [IDS["own_team"], IDS["second_team"]],
        }, token),
    ).status_code == 302
    with h.Session(ENGINE) as session:
        person = session.query(db.Person).filter_by(name="Admin Other").one()
        assert set(db.membership_team_ids(person)) == {IDS["other_team"], IDS["support"]}
        member = session.get(db.Person, IDS["member"])
        assert set(db.membership_team_ids(member)) == {
            IDS["own_team"], IDS["second_team"]
        }

    assert client.post(
        f"/personen/{IDS['member']}/teams",
        data=h.csrf_data(token=token),
    ).status_code == 302
    with h.Session(ENGINE) as session:
        assert db.membership_team_ids(session.get(db.Person, IDS["member"])) == ()


if __name__ == "__main__":
    h.run_all(dict(globals()))
