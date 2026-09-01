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
        name="Admin Person", email="admin@management.test", team=support, is_admin=True
    )
    mv = db.Person(name="Multi MV", email="mv@management.test", team=own_team)
    member = db.Person(name="Visible Member", email="member@management.test", team=own_team)
    other = db.Person(name="Other Active", email="private@management.test", team=other_team)
    inactive = db.Person(
        name="Hidden Inactive", team=other_team, account_status=db.ACCOUNT_INACTIVE
    )
    pending_own = db.Person(
        name="Pending Own", desired_team=own_team, account_status=db.ACCOUNT_VERIFIED
    )
    pending_second = db.Person(
        name="Pending Second", desired_team=second_team,
        account_status=db.ACCOUNT_VERIFIED,
    )
    pending_other = db.Person(
        name="Pending Other", desired_team=other_team,
        account_status=db.ACCOUNT_VERIFIED,
    )
    pending_support = db.Person(
        name="Pending Support", desired_team=support,
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

    by_name = client.get("/personen?name=other").get_data(as_text=True)
    assert "Other Active" in by_name and "Visible Member" not in by_name
    assert "private@management.test" not in by_name
    by_team = client.get(
        f"/personen?team_id={IDS['other_team']}"
    ).get_data(as_text=True)
    assert "Other Active" in by_team and "Visible Member" not in by_team
    forged_status = client.get(
        "/personen?name=Hidden&status=inactive"
    ).get_data(as_text=True)
    assert "Hidden Inactive" not in forged_status


def test_02_mv_can_create_contactless_people_for_every_managed_team_only():
    client, token = _client("mv")
    page = client.get("/personen").get_data(as_text=True)
    assert 'id="new-user-card"' in page
    assert 'id="pending-registration-card"' in page
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
        assert {person.team_id for person in created} == {
            IDS["own_team"], IDS["second_team"],
        }
        assert all(
            person.account_status == db.ACCOUNT_ACTIVE
            and person.email is None and person.phone is None
            for person in created
        )
        assert session.query(db.Person).filter_by(name="Forged Other").first() is None


def test_03_registration_decisions_follow_mv_and_admin_team_scope():
    mv_client, mv_token = _client("mv")
    page = mv_client.get("/personen").get_data(as_text=True)
    assert "Pending Own" in page and "Pending Second" in page
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
    ).status_code == 302
    assert mv_client.post(
        f"/registrierungen/{IDS['pending_second']}/reject",
        data=h.csrf_data(token=mv_token),
    ).status_code == 302

    admin_client, admin_token = _client("admin")
    admin_page = admin_client.get("/personen").get_data(as_text=True)
    assert "Pending Other" in admin_page and "Pending Support" in admin_page
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
        "name": "Admin Other", "team_id": IDS["other_team"],
    }, token))
    assert created.status_code == 302
    with h.Session(ENGINE) as session:
        assert session.query(db.Person).filter_by(name="Admin Other").one().team_id == IDS["other_team"]


if __name__ == "__main__":
    h.run_all(dict(globals()))
