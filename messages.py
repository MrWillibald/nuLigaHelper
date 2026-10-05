"""Authoritative German notification copy and explicit named plain-text contracts.

The catalog contains active notifications and dormant newspaper content. Transport,
recipient selection and scheduling stay with the callers named on each entry.
Fields describe the exact context accepted by render()/render_fragment(); their
semantic meanings are documented in FIELD_DESCRIPTIONS. Values are substituted once,
without expression evaluation. Errors report keys and field names, never values.
"""

from dataclasses import dataclass
from string import Formatter


class MessageRenderError(ValueError):
    """An unknown key, invalid template or incomplete plain-text context."""


@dataclass(frozen=True)
class RenderedMessage:
    subject: str
    email: str
    sms: str | None


@dataclass(frozen=True)
class MessageTemplate:
    fields: tuple[str, ...]
    subject: str
    email: str
    sms: str | None = None
    shared_sms: bool = False
    purpose: str = ""
    caller: str = ""
    status: str = "active"
    legacy_keys: tuple[str, ...] = ()


@dataclass(frozen=True)
class FragmentTemplate:
    fields: tuple[str, ...]
    text: str
    purpose: str
    caller: str
    status: str = "active"


FIELD_DESCRIPTIONS = {
    "recipient_name": "Name of the person receiving this notification.",
    "task_label": "Current semantic duty or block label, without position numbers.",
    "game_date": "Scheduled event date as German dd.mm.yyyy text.",
    "block_date": "Date of the home-game day owning a preparation, cleanup or cake block.",
    "game_start_time": "Scheduled event start time, not a block meeting time.",
    "age_class": "Scraped age class; Spielfest retains its full event label.",
    "age_class_display_label": "Newspaper display name derived from the scraped age class.",
    "home_team_name": "Ordinary game's home team.",
    "away_team_name": "Ordinary game's away team.",
    "responsible_team_name": "The team's name whose MV receives the follow-up.",
    "timekeeper_name": "Saved timekeeper name, or staffing.position_unassigned when vacant.",
    "timekeeper_task_label": "Current display label for the timekeeper duty.",
    "secretary_name": "Saved secretary name, or staffing.position_unassigned when vacant.",
    "secretary_task_label": "Current display label for the secretary duty.",
    "age_eligibility_feedback": "Rendered staffing.deficiencies fragment, empty if resolved.",
    "age_eligibility_reasons": "Joined domain eligibility explanations; no dates of birth or exact ages.",
    "partner_names": "Names of the other preparation volunteers, or the partner_none fragment.",
    "block_meeting_time": "Calculated block meeting/delivery time, including a differing calendar date.",
    "block_meeting_date": "Actual calendar date of a block meeting outside the home-game date.",
    "old_game_date": "Event's previously scheduled date.",
    "old_game_start_time": "Event's previously scheduled start time.",
    "new_game_date": "Event's newly scheduled date.",
    "new_game_start_time": "Event's newly scheduled start time.",
    "notified_person_names": "Names of all selected referee-alert recipients.",
    "article_date": "Requested newspaper publication date.",
    "game_weekday": "Display weekday of the advertised home-game date.",
    "schedule_text": "Joined rendered newspaper schedule rows.",
    "selected_team_names": "All selected registration teams joined by the domain membership display helper.",
    "registrant_name": "Name of the verified person waiting for administrative approval.",
    "auth_code": "Six-digit, purpose-bound authentication code; never included in diagnostics.",
    "affected_component_names": "Sorted names of affected operational components.",
    "occurred_at": "Operator-alert occurrence time formatted as an ISO UTC timestamp.",
    "runbook_reference": "Reference to the existing operational incident instructions.",
}

_GAME_FIELDS = ("recipient_name", "game_date", "task_label", "age_class", "home_team_name", "away_team_name", "game_start_time")
_SPF_FIELDS = ("recipient_name", "game_date", "task_label", "age_class", "game_start_time")
_MV_FIELDS = ("recipient_name", "responsible_team_name", "game_date", "timekeeper_name", "secretary_name", "timekeeper_task_label", "secretary_task_label", "age_class", "game_start_time", "age_eligibility_feedback")
_BLOCK_FIELDS = ("recipient_name", "block_date", "task_label", "block_meeting_time")
_SHIFT_FIELDS = ("recipient_name", "task_label", "age_class", "old_game_date", "old_game_start_time", "new_game_date", "new_game_start_time")

CATALOG: dict[str, MessageTemplate | FragmentTemplate] = {
    "game.day_before": MessageTemplate(
        _GAME_FIELDS,
        'Benachrichtigung Dienst {task_label}',
        'Hallo {recipient_name},\n\ndu bist morgen ({game_date}) als {task_label} beim Spiel '
        'der {age_class} {home_team_name} gegen {away_team_name} eingeteilt.\nDas Spiel '
        'beginnt um {game_start_time}. Bitte sei mindestens eine Stunde vor Spielbeginn in '
        'der Halle.\n\nViele Grüße\nTuS Raubling Handball',
        'Hallo {recipient_name},\ndu bist morgen ({game_date}) als {task_label} beim Spiel '
        'der {age_class} eingeteilt. Das Spiel beginnt um {game_start_time}. Bitte sei '
        'mindestens eine Stunde vor Spielbeginn in der Halle.\nVG TuS Handball',
        purpose='Day-before reminder for every occupied ordinary-game duty.',
        caller='Notifier.notify_game_day',
        legacy_keys=('mailTask', 'textTask'),
    ),
    "game.weekly": MessageTemplate(
        _GAME_FIELDS,
        'Benachrichtigung Dienst {task_label}',
        'Hallo {recipient_name},\n\ndu bist am nächsten Heimspieltag ({game_date}) als '
        '{task_label} beim Spiel der {age_class} {home_team_name} gegen {away_team_name} '
        'eingeteilt. \nDas Spiel beginnt um {game_start_time}. Bitte sei mindestens eine '
        'Stunde vor Spielbeginn in der Halle.\n\nViele Grüße\nTuS Raubling Handball',
        'Hallo {recipient_name},\ndu bist am nächsten Heimspieltag ({game_date}) als '
        '{task_label} beim Spiel der {age_class} eingeteilt. Das Spiel beginnt um '
        '{game_start_time}. Bitte sei mindestens eine Stunde vor Spielbeginn in der '
        'Halle.\nVG TuS Handball',
        purpose='One-week reminder for every occupied ordinary-game duty.',
        caller='Notifier.notify_pre',
        legacy_keys=('mailPreTask', 'textPreTask'),
    ),
    "spielfest.day_before": MessageTemplate(
        _SPF_FIELDS,
        'Benachrichtigung Dienst {task_label}',
        'Hallo {recipient_name},\n\ndu bist morgen ({game_date}) für den Dienst {task_label} '
        'beim Spielfest {age_class} eingeteilt. Das Spielfest beginnt um {game_start_time}.'
        ' Bitte sei rechtzeitig in der Halle.\n\nViele Grüße\nTuS Raubling Handball',
        shared_sms=True,
        purpose='Day-before Spielfest reminder without invented opponents.',
        caller='Notifier.notify_game_day',
    ),
    "spielfest.weekly": MessageTemplate(
        _SPF_FIELDS,
        'Benachrichtigung Dienst {task_label}',
        'Hallo {recipient_name},\n\ndu bist nächste Woche ({game_date}) für den Dienst '
        '{task_label} beim Spielfest {age_class} eingeteilt. Das Spielfest beginnt um '
        '{game_start_time}. Bitte sei rechtzeitig in der Halle.\n\nViele Grüße\n TuS'
        ' Raubling Handball',
        shared_sms=True,
        purpose='One-week Spielfest reminder without invented opponents.',
        caller='Notifier.notify_pre',
    ),
    "mv.game.day_before": MessageTemplate(
        _MV_FIELDS + ("home_team_name", "away_team_name"),
        'Benachrichtigung Verantwortlich',
        'Hallo {recipient_name},\n\nmorgen ({game_date}) ist deine Mannschaft '
        '{responsible_team_name} für die Dienste beim Spiel {home_team_name} gegen '
        '{away_team_name} ({age_class}) verantwortlich. Das Spiel beginnt um '
        '{game_start_time} in der Halle.\n\n{timekeeper_task_label}: {timekeeper_name}\n'
        '{secretary_task_label}: {secretary_name}\n\nBitte prüfe die offenen Dienste im '
        'Heimspielplan. Die eingeteilten Helfer erhalten ebenfalls eine Benachrichtigung.\n\n'
        'Die Spieler erhalten ebenfalls eine Benachrichtigung.\n\nViele Grüße\n'
        'TuS Raubling Handball{age_eligibility_feedback}',
        'Hallo {recipient_name},\nmorgen ({game_date}) ist deine Mannschaft '
        '{responsible_team_name} für die Dienste beim Spiel der {age_class} verantwortlich.'
        ' Spielbeginn: {game_start_time}. {timekeeper_task_label}: {timekeeper_name}; '
        '{secretary_task_label}: {secretary_name}. Bitte prüfe die offenen Dienste im '
        'Heimspielplan.\nVG TuS Handball{age_eligibility_feedback}',
        purpose='MV follow-up while game duties are vacant or age eligibility is unresolved.',
        caller='Notifier.notify_game_day',
        legacy_keys=('mailMVSubject', 'mailMV', 'textMV'),
    ),
    "mv.spielfest.day_before": MessageTemplate(
        _MV_FIELDS,
        'Benachrichtigung Verantwortlich',
        'Hallo {recipient_name},\n\nmorgen ({game_date}) ist deine Mannschaft '
        '{responsible_team_name} für die Dienste beim Spielfest {age_class} verantwortlich.'
        ' Das Spielfest beginnt um {game_start_time}.\n\n{timekeeper_task_label}: '
        '{timekeeper_name}\n{secretary_task_label}: {secretary_name}\n\nBitte prüfe die '
        'offenen Dienste im Heimspielplan.\n\nDie Spieler erhalten ebenfalls eine '
        'Benachrichtigung.\n\nViele Grüße\nTuS Raubling Handball{age_eligibility_feedback}',
        shared_sms=True,
        purpose='MV follow-up for incomplete Spielfest staffing.',
        caller='Notifier.notify_game_day',
    ),
    "block.day_before": MessageTemplate(
        _BLOCK_FIELDS,
        'Benachrichtigung Dienst {task_label}',
        'Hallo {recipient_name},\n\nmorgen ({block_date}) bist du für {task_label} eingeteilt. Zeit '
        'für deinen Dienst: {block_meeting_time}.\n\nViele Grüße\nTuS Raubling Handball',
        'Hallo {recipient_name}, morgen ({block_date}) bist du für {task_label} eingeteilt. '
        'Dienstzeit: {block_meeting_time}.\nVG TuS Handball',
        purpose='Day-before reminder for every occupied day-block position.',
        caller='Notifier.notify_blocks_day_before',
        legacy_keys=('mailBlockTask', 'textBlockTask'),
    ),
    "block.weekly": MessageTemplate(
        _BLOCK_FIELDS,
        'Benachrichtigung Dienst {task_label}',
        'Hallo {recipient_name},\n\nnächste Woche ({block_date}) bist du für {task_label} eingeteilt. '
        'Zeit für deinen Dienst: {block_meeting_time}.\n\nViele Grüße\nTuS Raubling Handball',
        'Hallo {recipient_name}, nächste Woche ({block_date}) bist du für {task_label} eingeteilt. '
        'Dienstzeit: {block_meeting_time}.\nVG TuS Handball',
        purpose='One-week reminder for cleanup and cake delivery.',
        caller='Notifier._notify_blocks_early',
        legacy_keys=('mailBlockPreTask', 'textBlockPreTask'),
    ),
    "block.preparation.weekly": MessageTemplate(
        _BLOCK_FIELDS + ("partner_names",),
        'Vorbereitung Dienst {task_label}',
        'Hallo {recipient_name},\n\nam nächsten Heimspieltag ({block_date}) bist du für '
        '{task_label} eingeteilt. Du arbeitest mit {partner_names} zusammen. Treffpunkt: '
        '{block_meeting_time}.\n\nViele Grüße\nTuS Raubling Handabll',
        'Hallo {recipient_name}, am nächsten Heimspieltag ({block_date}) bist du für '
        '{task_label} mit {partner_names} eingetilt. Treffpunkt: {block_meeting_time}.\nVG TuS Handball',
        purpose='One-week preparation reminder using the existing preparation sender.',
        caller='Notifier.notify_service_early',
        legacy_keys=('mailPreparationTask', 'textPreparationTask'),
    ),
    "game.shifted": MessageTemplate(
        _SHIFT_FIELDS + ("home_team_name", "away_team_name"),
        'Benachrichtigung Verschiebung Dienst {task_label}',
        'Hallo {recipient_name},\n\ndein Dienst {task_label} beim Spiel der {age_class} '
        '{home_team_name} gegen {away_team_name} vom {old_game_date} {old_game_start_time} '
        'verschiebt sich auf {new_game_date} {new_game_start_time}. Bitte suche dir '
        'selbstständig Ersatz, falls du an diesem Termin keine Zeit hast.\n\nViele Grüße\n'
        'TuS Raubling Handball',
        'Hallo {recipient_name},\ndein Dienst {task_label} vom {old_game_date} '
        '{old_game_start_time} verschiebt sich auf {new_game_date} {new_game_start_time}. '
        'Bitte suche dir selbstständig Ersatz, falls du an diesem Termin keine Zeit hast.\n'
        'VG TuS Handball',
        purpose='Ordinary-game rescheduling notice for saved helpers.',
        caller='Notifier.notify_shifts',
        legacy_keys=('mailShifted', 'textShifted'),
    ),
    "spielfest.shifted": MessageTemplate(
        _SHIFT_FIELDS,
        'Benachrichtigung Verschiebung Dienst {task_label}',
        'Hallo {recipient_name},\n\ndein Dienst {task_label} beim Spielfest {age_class} wurde'
        ' verschoben: nicht mehr am {old_game_date} um {old_game_start_time}, sondern am '
        '{new_game_date} um {new_game_start_time}.\n\nViele Grüße\nTuS Raubling Handball',
        shared_sms=True,
        purpose='Spielfest rescheduling notice without invented opponents.',
        caller='Notifier.notify_shifts',
    ),
    "referee.missing": MessageTemplate(
        ("recipient_name", "age_class", "game_date", "game_start_time", "notified_person_names"),
        'Benachrichtigung fehlender Schiedsrichter',
        'Hallo {recipient_name},\nfür das Heimspiel der {age_class} am {game_date} um '
        '{game_start_time} wird vom BHV kein Schiedsrichter nominiert! Bitte für einen '
        'Schiedsrichter sorgen.\nFolgende Personen werden benachrichtigt: '
        '{notified_person_names}\nVG TuS Handball',
        shared_sms=True,
        purpose='Missing-referee alert with corrected semantic field placement.',
        caller='Notifier._notify_missing_referee',
        legacy_keys=('mailRefCoordSubject', 'mailRefCoord'),
    ),
    "game.new": MessageTemplate(
        (),
        'Benachrichtigung Fehler',
        'Hallo Manu,\n\nbeim Update des Heimspielplans ist ein Fehler aufgetreten, die '
        'Spielnummern stimmen nicht überein. Bitte überprüfe den Fehler.\n\nViele Grüße TuS Handball',
        purpose='Admin information about newly scraped games.',
        caller='Notifier.notify_new_games',
        legacy_keys=('mailErrorSubject', 'mailError'),
    ),
    "auth.registration_code": MessageTemplate(
        ("recipient_name", "auth_code"), "Registrierung nuLigaHelper",
        "Hallo {recipient_name},\n{auth_code} ist dein Code für die Registrierung bei "
        "nuLigaHelper des TuS Raubling Handball. Er gilt 15 Minuten.\nVG TuS Handball",
        shared_sms=True,
        purpose="Purpose-bound registration code through the explicitly selected contact route.",
        caller="webapp._send_challenge",
    ),
    "auth.login_code": MessageTemplate(
        ("recipient_name", "auth_code"), "Anmeldung nuLigaHelper",
        "Hallo {recipient_name},\n{auth_code} ist dein Code für die Anmeldung bei "
        "nuLigaHelper des TuS Raubling Handball. Er gilt 15 Minuten.\nVG TuS Handball",
        shared_sms=True,
        purpose="Purpose-bound login code through the explicitly selected contact route.",
        caller="webapp._send_challenge",
    ),
    "auth.existing_account": MessageTemplate(
        (), "Registrierung nuLigaHelper", "Für diesen Kontakt besteht bereits ein Konto. "
        "Bitte nutze die Anmeldung.", shared_sms=True,
        purpose="Existing-account notice through the explicitly selected contact route.",
        caller="webapp.register",
    ),
    "account.approval_request": MessageTemplate(
        ("recipient_name", "registrant_name", "selected_team_names"), "Neue Registrierung",
        "Hallo {recipient_name},\n{registrant_name} hat den Kontakt bestätigt und "
        "wartet auf Freigabe für: {selected_team_names}.\n"
        "Bitte prüfe die Registrierung unter \"Helfer verwalten\".\nVG TuS Handball",
        "Hallo {recipient_name},\nneue Registrierung von {registrant_name} für "
        "{selected_team_names}. Bitte unter \"Helfer verwalten\" prüfen.\nVG TuS Handball",
        purpose="Best-effort approval request to active administrators, listing all chosen teams.",
        caller="webapp._notify_registration_approvers",
    ),
    "account.welcome": MessageTemplate(
        ("recipient_name",), "Registrierung freigegeben",
        "Hallo {recipient_name},\n\nherzlich Willkommen beim nuLigaHelper des TuS Raubling "
        "Handball!\nDeine Registrierung wurde freigegeben. Du kannst dich jetzt "
        "anmelden und offene Dienste im Heimspielplan übernehmen.\n\nViele Grüße\nTuS Raubling "
        "Handball",
        "Hallo {recipient_name},\nherzlich Willkommen beim nuLigaHelper des TuS Raubling "
        "Handball! Deine Registrierung wurde freigegeben. Melde dich an und übernimm "
        "offene Dienste im Heimspielplan.\nVG TuS Handball",
        purpose="Best-effort welcome after a verified registration is approved.",
        caller="webapp._notify_registration_approved",
    ),
    "operations.alert": MessageTemplate(
        ("affected_component_names", "occurred_at", "runbook_reference"), "nuLigaHelper: Betriebsmeldung",
        "Komponenten: {affected_component_names}\nZeit: {occurred_at}\nAnleitung: {runbook_reference}\n",
        purpose="Dedicated operator-only SMTP alert with occurrence time and runbook.",
        caller="operations.alert",
    ),
    "newspaper.article": MessageTemplate(
        ("article_date", "game_weekday", "game_date", "schedule_text"),
        'Veranstaltungshinweis für Gemeindeanzeiger',
        'Hallo Frau Neuner,\n\nhier der Veranstaltungshinweis zu unserem Spieltag für den '
        'Gemeindeanzeiger am {article_date}:\n\nHandball: nächster Heimspieltag\nAm '
        '{game_weekday}, den {game_date}, finden folgende Begegnungen in der Sporthalle des'
        ' TuS Raubling statt:\n{schedule_text}\nDie Handballer des TuS Raubling freuen sich '
        'auf euren Besuch!\n\nMit freundlichen Grüßen\nTuS Raubling Handball',
        purpose='Retained newspaper article with club-specific copy; daily call remains disabled.',
        caller='Notifier.send_article (main.py call remains disabled)',
        status='dormant',
        legacy_keys=('mailNewspaperSubject', 'mailNewspaper'),
    ),
    "staffing.position_unassigned": FragmentTemplate((), "noch offen", "Vacant timing position in MV follow-up.", "Notifier.notify_game_day"),
    "staffing.deficiencies": FragmentTemplate(
        ("age_eligibility_reasons",), "\n\nAltersanforderungen offen: {age_eligibility_reasons}",
        "MV follow-up suffix from non-private eligibility explanations.", "Notifier.notify_game_day",
    ),
    "block.time_unset": FragmentTemplate((), "noch offen", "Unknown block meeting/delivery time.", "Notifier._block_time_text"),
    "block.time_cross_date": FragmentTemplate(("block_meeting_time", "block_meeting_date"), "{block_meeting_time} am {block_meeting_date}", "Block task time outside the home-game calendar date.", "Notifier._block_time_text"),
    "preparation.partner_none": FragmentTemplate((), "niemandem", "No other preparation volunteer assigned.", "Notifier.notify_service_early"),
    "task.cake": FragmentTemplate(("task_label",), "{task_label} (ein Kuchen)", "Current cake-delivery label and per-position quantity.", "Notifier._block_reminder_task"),
    "newspaper.game_row": FragmentTemplate(("game_start_time", "age_class_display_label", "home_team_name", "away_team_name"), "{game_start_time} {age_class_display_label} {home_team_name} - {away_team_name}\n", "Ordinary newspaper schedule row.", "Notifier.send_article", status="dormant"),
    "newspaper.minis_row": FragmentTemplate(("game_start_time",), "Ab {game_start_time} Spielfest der Minis\n", "First legacy MI tournament schedule row.", "Notifier.send_article", status="dormant"),
    "newspaper.e_youth_row": FragmentTemplate(("game_start_time",), "Ab {game_start_time} Turnier der gemischten E-Jugend\n", "First legacy GE tournament schedule row.", "Notifier.send_article", status="dormant"),
    "newspaper.women": FragmentTemplate((), "Damen", "Legacy F newspaper display label.", "Notifier.send_article", status="dormant"),
    "newspaper.men": FragmentTemplate((), "Herren", "Legacy M newspaper display label.", "Notifier.send_article", status="dormant"),
}

# These definitions existed only in the committed configuration example. No
# caller reads them; active early preparation uses block.preparation.weekly.
UNUSED_LEGACY_KEYS = ("mailEarlyTask", "textEarlyTask")
LEGACY_RECIPIENT_KEY = "mailRefCoordTargets"

_FORMATTER = Formatter()
_LINE_BREAKS = "\r\n\v\f\x1c\x1d\x1e\x85\u2028\u2029"


def _template_fields(key: str, text: str) -> set[str]:
    try:
        parsed = tuple(_FORMATTER.parse(text))
    except (TypeError, ValueError):
        raise MessageRenderError(f"Invalid template for {key}") from None
    fields = set()
    for _, field, spec, conversion in parsed:
        if field is None:
            continue
        if not field.isidentifier() or spec or conversion:
            raise MessageRenderError(f"Invalid named placeholder for {key}")
        fields.add(field)
    return fields


def _validated_context(key: str, template, context: dict) -> dict[str, str]:
    texts = ((template.text,) if isinstance(template, FragmentTemplate) else
             (template.subject, template.email) + ((template.sms,) if template.sms is not None else ()))
    declared = set(template.fields)
    referenced = set().union(*(_template_fields(key, text) for text in texts))
    if declared != referenced or len(declared) != len(template.fields):
        raise MessageRenderError(f"Invalid field declaration for {key}")
    if not isinstance(template, FragmentTemplate) and template.shared_sms and template.sms is not None:
        raise MessageRenderError(f"Ambiguous SMS variant for {key}")
    missing = declared - context.keys()
    unknown = context.keys() - declared
    if missing or unknown:
        details = []
        if missing:
            details.append("missing fields: " + ", ".join(sorted(missing)))
        if unknown:
            details.append("unknown fields: " + ", ".join(sorted(unknown)))
        raise MessageRenderError(f"Invalid context for {key}; " + "; ".join(details))
    invalid = [name for name, value in context.items() if type(value) not in (str, int)]
    if invalid:
        raise MessageRenderError(f"Invalid plain-text fields for {key}: " + ", ".join(sorted(invalid)))
    return {name: str(value) for name, value in context.items()}


def render(key: str, **context) -> RenderedMessage:
    """Render a complete message, validating all channels before dispatch."""
    template = CATALOG.get(key)
    if not isinstance(template, MessageTemplate):
        raise MessageRenderError(f"Unknown message key: {key}")
    values = _validated_context(key, template, context)
    subject = template.subject.format_map(values)
    if any(character in subject for character in _LINE_BREAKS):
        raise MessageRenderError(f"Invalid subject line break for {key}")
    email = template.email.format_map(values)
    sms = email if template.shared_sms else (template.sms.format_map(values) if template.sms is not None else None)
    return RenderedMessage(subject, email, sms)


def render_fragment(key: str, **context) -> str:
    """Render a catalog fragment using the same exact context contract."""
    template = CATALOG.get(key)
    if not isinstance(template, FragmentTemplate):
        raise MessageRenderError(f"Unknown fragment key: {key}")
    return template.text.format_map(_validated_context(key, template, context))
