"""Shared assertions for parser-time disclosure enhancement in rendered pages."""


def assert_early_disclosure_markup(page):
    """Verify scripts precede bodies while native fallback metadata stays intact."""
    disclosures = page.xpath('//details[@data-responsive-disclosure]')
    assert disclosures, "the synthetic page must exercise responsive disclosures"
    asset_url = "/static/disclosures.js"
    preload = page.xpath('//head/link[@rel="preload" and @as="script"]')
    assert [link.get("href") for link in preload] == [asset_url], (
        "preload must request the same same-origin URL as every early invocation"
    )
    scripts = page.xpath('//script[@src="/static/disclosures.js"]')
    assert len(scripts) == len(disclosures), "every responsive panel needs exactly one early invocation"
    assert not page.xpath('//script[not(@src)]'), "the existing script policy permits no inline scripts"
    for disclosure in disclosures:
        children = [child for child in disclosure if isinstance(child.tag, str)]
        assert [child.tag for child in children[:3]] == ["summary", "script", "div"], (
            "the parser must execute the initializer immediately after the summary, before any body"
        )
        script, content = children[1:3]
        assert script.get("src") == asset_url and script.getparent() is disclosure, (
            "document.currentScript must identify the enclosing disclosure"
        )
        assert not any(name in script.attrib for name in ("async", "defer", "type", "nomodule")), (
            "early initialization must use a synchronous classic external script"
        )
        assert not script.text or not script.text.strip(), "each invocation must contain no inline code"
        assert "disclosure-content" in content.get("class", "").split()
        desktop_open = disclosure.get("data-desktop-open")
        assert desktop_open in ("true", "false"), "every panel keeps its explicit desktop default"
        error = disclosure.get("data-disclosure-error")
        assert error in (None, "true"), "the server error marker is authoritative before fields exist"
        assert (disclosure.get("open") is not None) == (desktop_open == "true" or error == "true"), (
            "server markup must retain expanded introduction/filter/team-statistics and error fallbacks"
        )
        if disclosure.xpath('.//*[@aria-invalid="true"]'):
            assert error == "true", "invalid maintenance fields need an early server error marker"
