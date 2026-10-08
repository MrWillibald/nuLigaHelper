"""Responsive roster presentation preserves identity, filters and private edit recovery."""

from datetime import date

from lxml import html

import helpers as h
import db
from test_auth import _new_app


def _roster():
    app, engine = _new_app()
    with h.Session(engine) as session:
        h.sync_sample_games(session)
        own_team, other_team = [team for team in db.get_all_teams(session) if not team.is_support][:2]
        support = db.get_support_team(session)
        people = {
            "admin": db.Person(name="Mobile Admin", is_admin=True, birth_date=date(1981, 2, 3)),
            "mv": db.Person(name="Mobile MV", teams=[own_team], birth_date=date(1982, 3, 4)),
            "member": db.Person(name="Alex Shared", teams=[own_team, support],
                                email="mobile-self@example.test", phone="+491701234567",
                                birth_date=date(2000, 4, 5)),
            "duplicate": db.Person(name="Alex Shared", teams=[own_team, other_team],
                                   email="mobile-private@example.test", phone="+491701234568"),
            "other": db.Person(name="Mobile Other", teams=[other_team],
                               email="mobile-other@example.test", birth_date=date(1999, 6, 7)),
            "inactive": db.Person(name="Mobile Inactive", teams=[own_team],
                                  account_status=db.ACCOUNT_INACTIVE),
            "pending": db.Person(name="Mobile Pending", teams=[own_team],
                                 account_status=db.ACCOUNT_VERIFIED),
        }
        session.add_all(people.values())
        session.flush()
        own_team.mv_person_id = people["mv"].id
        session.commit()
        ids = {key: person.id for key, person in people.items()}
        ids.update(own_team=own_team.id, other_team=other_team.id, support=support.id)
        names = {team.id: team.name for team in (own_team, other_team, support)}
    return app, engine, ids, names


def _client(app, ids, role):
    client = app.test_client()
    h.sign_in(client, ids[role])
    return client


def _page(client, path="/personen"):
    response = client.get(path)
    assert response.status_code == 200
    return html.fromstring(response.get_data(as_text=True))


def _card(page, person_id):
    return page.get_element_by_id(f"person-{person_id}")


def _maintenance(card):
    return card.xpath('.//details[@data-responsive-disclosure]')


def test_roster_headers_keep_complete_memberships_and_status_outside_authorized_actions():
    app, _, ids, names = _roster()
    page = _page(_client(app, ids, "admin"))
    introduction = page.xpath('//details[contains(@class, "introduction-disclosure")]')[0]
    assert introduction.get("open") is not None and introduction.get("data-desktop-open") == "true"
    assert introduction.xpath('./summary')[0].text_content() == "Hinweise"
    missing_notice = page.xpath('//a[contains(@href, "birth_date=missing")]')[0]
    assert not missing_notice.xpath('ancestor::details'), "the actionable date follow-up stays outside hints"
    for key, teams in (("member", (ids["own_team"], ids["support"])),
                       ("duplicate", (ids["own_team"], ids["other_team"]))):
        card = _card(page, ids[key])
        header = card.xpath('./div[@class="person-card-body"]/h3')[0]
        assert header.text_content() == "Alex Shared" and not header.xpath('ancestor::details')
        badges = card.xpath('.//span[@class="person-team-badge"]')
        assert {badge.text_content() for badge in badges} == {names[team] for team in teams}
        assert all(not badge.xpath('ancestor::details') for badge in badges)
        status = card.xpath('.//span[contains(@class, "status-badge")]')[0]
        assert status.text_content().strip() == "Aktiv" and not status.xpath('ancestor::details')
        maintenance = _maintenance(card)[0]
        assert maintenance.get("id") == f"person-maintenance-{ids[key]}"
        assert maintenance.get("open") is None and maintenance.get("data-desktop-open") == "false"
        assert maintenance.xpath('./summary')[0].text_content() == "Bearbeitung"
        assert card.xpath('.//dialog') and not card.xpath('.//details//dialog'), (
            "membership dialogs remain outside maintenance disclosure lifecycle"
        )
    assert page.get_element_by_id("mv-assignment-card").sourceline < page.xpath('//div[@class="roster-heading"]')[0].sourceline
    assert all(not item.get("data-game") and not item.get("data-block")
               for item in page.xpath('//details[@data-responsive-disclosure]'))


def test_roster_effective_filter_summary_is_outside_form_and_retains_combined_selections():
    app, _, ids, names = _roster()
    client = _client(app, ids, "admin")
    page = _page(client, f"/personen?name=alex&team_id={ids['own_team']}&status=active&birth_date=missing")
    summary = page.xpath('//div[@class="active-filter-summary"]')[0]
    assert not summary.xpath('ancestor::details')
    assert all(value in summary.text_content() for value in
               ("Name: alex", f"Mannschaft: {names[ids['own_team']]}", "Status: Aktiv", "Geburtsdatum: Fehlt"))
    assert summary.xpath('./a/@href') == ["/personen"]
    assert page.get_element_by_id("person-name-filter").get("value") == "alex"
    assert page.xpath('//select[@name="team_id"]/option[@selected]/@value') == [str(ids["own_team"])]
    assert page.xpath('//select[@name="status"]/option[@selected]/@value') == ["active"]
    assert page.xpath('//select[@name="birth_date"]/option[@selected]/@value') == ["missing"]
    assert page.xpath('//article[contains(@class, "person-card")]/@id') == [f"person-{ids['duplicate']}"]
    filter_details = page.xpath('//form[@class="roster-filter-card"]/ancestor::details')[0]
    assert filter_details.get("open") is not None and filter_details.get("data-desktop-open") == "true"
    assert page.xpath('//form[@class="roster-filter-card"]/@method') == ["get"]
    escaped = _page(client, "/personen?name=%3Cscript%3Ealert(1)%3C%2Fscript%3E")
    assert "<script>alert(1)</script>" in escaped.xpath('//div[@class="active-filter-summary"]')[0].text_content()
    assert not escaped.xpath('//div[@class="active-filter-summary"]//script'), "entered names remain escaped text"
    unknown = _page(client, "/personen?team_id=987654321")
    unknown_summary = unknown.xpath('//div[@class="active-filter-summary"]')[0].text_content()
    assert "Unbekannte Mannschaft" in unknown_summary and "987654321" not in unknown_summary
    assert unknown.xpath('//select[@name="team_id"]/option[@selected]/@value') == ["987654321"]
    assert "Keine passenden Personen gefunden." in unknown.text_content()
    assert not unknown.xpath('//article[contains(@class, "person-card")]')


def test_member_and_mv_summaries_ignore_admin_filters_and_keep_other_profiles_absent():
    app, _, ids, _ = _roster()
    for role in ("member", "mv"):
        client = _client(app, ids, role)
        page = _page(client, "/personen?status=inactive&birth_date=missing")
        response = html.tostring(page, encoding="unicode")
        assert not page.xpath('//div[@class="active-filter-summary"]'), "forged admin filters cannot appear active"
        assert "Mobile Inactive" not in response and "Mobile Pending" not in response
        assert "mobile-private@example.test" not in response and "+491701234568" not in response
        assert "mobile-other@example.test" not in response and "1999-06-07" not in response
        assert not page.xpath('//select[@name="status" or @name="birth_date"]')
        other = _card(page, ids["other"])
        if role == "member":
            assert not _maintenance(other), "an ordinary member gets no empty edit disclosure for others"
            assert len(page.xpath('//details[contains(@class, "person-maintenance")]')) == 1
        else:
            maintenance = _maintenance(other)[0]
            assert maintenance.xpath('.//button[@data-team-dialog-open]')
            assert not maintenance.xpath('.//form[@class="person-edit-form"]')
            assert not other.xpath('.//input[@name="email" or @name="phone" or @name="birth_date"]')
            assert not other.xpath('.//details//dialog')
        assert client.post(f"/personen/{ids['other']}/edit", data=h.csrf_data({
            "name": "Forbidden private edit", "birth_date": "1990-01-01",
        })).status_code == 403
    assert app.test_client().get("/personen").status_code == 302


def test_invalid_self_edit_reopens_only_affected_form_and_preserves_submitted_values():
    app, engine, ids, _ = _roster()
    client = _client(app, ids, "member")
    values = {"name": "Changed <Name>", "birth_date": "9999-01-01",
              "email": "bad-address", "phone": "123"}
    response = client.post(f"/personen/{ids['member']}/edit", data=h.csrf_data(values))
    assert response.status_code == 400
    page = html.fromstring(response.get_data(as_text=True))
    card = _card(page, ids["member"])
    maintenance = _maintenance(card)[0]
    assert maintenance.get("open") is not None and maintenance.get("data-disclosure-error") == "true"
    assert page.xpath('//details[@data-disclosure-error]/@id') == [f"person-maintenance-{ids['member']}"]
    assert card.xpath('./div[@class="person-card-body"]/h3/text()') == ["Alex Shared"], "visible identity uses saved record"
    for field, value in values.items():
        assert maintenance.xpath(f'.//input[@name="{field}"]/@value') == [value]
    for field in ("birth_date", "email", "phone"):
        assert maintenance.xpath(f'.//input[@name="{field}"]/@aria-invalid') == ["true"]
        assert maintenance.xpath(f'.//*[@id="person-{ids["member"]}-{field}-error" and @role="alert"]')
    assert "mobile-private@example.test" not in response.get_data(as_text=True)
    with h.Session(engine) as session:
        saved = session.get(db.Person, ids["member"])
        assert saved.name == "Alex Shared" and saved.birth_date == date(2000, 4, 5)
        assert saved.email == "mobile-self@example.test" and saved.phone == "+491701234567"


def test_admin_edit_failure_preserves_missing_fields_and_unrelated_entry_defaults():
    app, engine, ids, _ = _roster()
    client = _client(app, ids, "admin")
    response = client.post(f"/personen/{ids['other']}/edit", data=h.csrf_data({
        "name": "Attempted admin edit", "email": "mobile-self@example.test", "phone": "",
    }))
    assert response.status_code == 400
    page = html.fromstring(response.get_data(as_text=True))
    affected = _maintenance(_card(page, ids["other"]))[0]
    assert affected.get("data-disclosure-error") == "true" and affected.get("open") is not None
    assert affected.xpath('.//input[@name="name"]/@value') == ["Attempted admin edit"]
    assert affected.xpath('.//input[@name="email"]/@value') == ["mobile-self@example.test"]
    assert affected.xpath('.//input[@name="phone"]/@value') == [""]
    assert affected.xpath('.//input[@name="birth_date"]/@value') == ["1999-06-07"], "an omitted date keeps its saved value"
    unrelated = _maintenance(_card(page, ids["member"]))[0]
    assert unrelated.get("data-disclosure-error") is None and unrelated.get("data-desktop-open") == "false"
    assert unrelated.get("open") is None, "only the affected validation-error form starts open"
    assert unrelated.xpath('.//input[@name="name"]/@value') == ["Alex Shared"]
    with h.Session(engine) as session:
        saved = session.get(db.Person, ids["other"])
        assert saved.name == "Mobile Other" and saved.email == "mobile-other@example.test"


def test_native_maintenance_fallbacks_reuse_authorized_csrf_posts_without_nested_forms():
    app, _, ids, _ = _roster()
    admin_page = _page(_client(app, ids, "admin"))
    card = _card(admin_page, ids["duplicate"])
    maintenance = _maintenance(card)[0]
    membership = maintenance.xpath(f'.//noscript/form[@action="/personen/{ids["duplicate"]}/teams"]')[0]
    assert membership.get("method") == "post" and membership.xpath('./input[@name="csrf_token"]')
    assert len(membership.xpath('.//input[@name="team_ids"]')) >= 2
    deletion = maintenance.xpath(f'.//noscript/form[@action="/personen/{ids["duplicate"]}/delete"]')[0]
    assert deletion.get("method") == "post" and deletion.xpath('./input[@name="csrf_token"]')
    assert deletion.xpath('.//input[@type="checkbox" and @required]')
    assert "Diensteinträge" in deletion.text_content() and "deaktivieren" in deletion.text_content()
    assert not admin_page.xpath('//form//form'), "native submission forms remain independent"
    mv_page = _page(_client(app, ids, "mv"))
    own_fallback = _maintenance(_card(mv_page, ids["mv"]))[0].xpath('.//noscript//button')[0]
    assert own_fallback.get("disabled") is not None, "MV fallback cannot remove qualifying own membership"
    assert not mv_page.xpath('//noscript//form[contains(@action, "/delete")]')


if __name__ == "__main__":
    h.run_all(dict(globals()))
