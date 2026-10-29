# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Duty roles, slot counts, day-block phases, labels and task descriptions.
# ---------------------------------------------------------------

from __future__ import annotations

from typing import TYPE_CHECKING

from nuligahelper.db import age_eligibility
from nuligahelper.db.task_labels import role_label

if TYPE_CHECKING:
    from nuligahelper.db.models import Game


# Stable assignment keys; use role_label() for current human-readable names.
ROLE_TIMEKEEPER = "Zeitnehmer"
ROLE_SECRETARY = "Sekretär"
ROLE_SALE = "Verkauf"
ROLE_SECURITY = "Ordnungsdienst"
ROLE_CASH = "Kasse"
ROLE_CLEANING = "Wischer"

# Slot count per role; keys define the display order for open-task lists
ROLE_SLOT_COUNT = {
    ROLE_TIMEKEEPER: 1,
    ROLE_SECRETARY: 1,
    ROLE_SALE: 2,
    ROLE_SECURITY: 1,
    ROLE_CASH: 1,
    ROLE_CLEANING: 1,
}

# Semantic roles occur once; repeated positions are resolved from assignments.
GAME_DAY_ROLES = list(ROLE_SLOT_COUNT)
BASELINE_REQUIRED_POSITIONS = (
    (ROLE_TIMEKEEPER, 0), (ROLE_SECRETARY, 0),
    (ROLE_SALE, 0), (ROLE_SALE, 1), (ROLE_SECURITY, 0),
)


def offered_positions(game: Game) -> tuple[tuple[str, int], ...]:
    """Adult games offer seven duties; youth/unknown classes keep five."""
    if age_eligibility.classify_game_category(game.ak) == age_eligibility.ADULT:
        return tuple((role, slot) for role, count in ROLE_SLOT_COUNT.items()
                     for slot in range(count))
    return BASELINE_REQUIRED_POSITIONS


def required_positions(game: Game) -> tuple[tuple[str, int], ...]:
    """Every offered game position is required; retained removed duties are not."""
    return offered_positions(game)


def position_label(role: str, slot: int) -> str:
    """Number positions only when the semantic role has multiple places."""
    label = role_label(role)
    return f"{label} {slot + 1}" if ROLE_SLOT_COUNT.get(role, 0) > 1 else label


BLOCK_PREPARATION = "preparation"
BLOCK_CAKE_DELIVERY = "cake_delivery"
BLOCK_CLEANUP = "cleanup"
BLOCK_PHASES = (BLOCK_PREPARATION, BLOCK_CAKE_DELIVERY, BLOCK_CLEANUP)
# Preparation and cleanup retain their fixed capacity. Cake capacity is saved
# per dated block and must be read through block_capacity().
BLOCK_SLOT_COUNT = 3
BLOCK_PHASE_LABELS = {
    BLOCK_PREPARATION: "Vorbereitung",
    BLOCK_CAKE_DELIVERY: "Kuchenlieferung",
    BLOCK_CLEANUP: "Aufräumen",
}

# Public task guidance, keyed by semantic role rather than numbered positions.
# Keep coverage in sync with ROLE_SLOT_COUNT and BLOCK_PHASES.
TASK_DESCRIPTIONS = {
    ROLE_TIMEKEEPER: (
        "Mindestalter am Spieltag: 14 Jahre bei Jugendspielen einschließlich SPF, 16 Jahre bei Erwachsenenspielen. "
        "Mindestens eine Stunde vor dem angesetzten Spielbeginn am Spielort sein. "
        "An den technischen Besprechungen mit den Schiedsrichtern spätestens 30 Minuten "
        "vor dem Spiel und direkt nach dem Spiel teilnehmen. "
        "Spieluhr und Anzeigetafel bedienen, Zwei-Minuten-Strafen verwalten und "
        "korrekte Spielerwechsel sicherstellen. "
        "Eine offizielle Zeitnehmer-Schulung ist nicht erforderlich."
    ),
    ROLE_SECRETARY: (
        "Mindestalter am Spieltag: 14 Jahre bei Jugendspielen einschließlich SPF, 16 Jahre bei Erwachsenenspielen. "
        "Mindestens eine Stunde vor dem angesetzten Spielbeginn am Spielort sein. "
        "An den technischen Besprechungen mit den Schiedsrichtern spätestens 30 Minuten "
        "vor dem Spiel und direkt nach dem Spiel teilnehmen. "
        "nuScore für den elektronischen Spielbericht bedienen und korrekte Spielerwechsel "
        "sicherstellen. Eine offizielle Sekretär-Schulung ist nicht erforderlich."
    ),
    ROLE_SALE: (
        "Mindestens eine Stunde vor dem angesetzten Spielbeginn am Spielort sein. "
        "Bei Jugendspielen einschließlich SPF gelten im Verkauf keine Altersanforderungen. "
        "Bei Erwachsenenspielen und ungeklärten Altersklassen muss im vollständig besetzten "
        "Verkauf mindestens eine eingeteilte Person am Spieltag 18 Jahre oder älter sein. "
        "Bei der Vorbereitung von Speisen und Getränken für den "
        "Verkaufsstand helfen und diese während des Spiels verkaufen."
    ),
    ROLE_SECURITY: (
        "Mindestalter am Spieltag: 18 Jahre bei allen Spielen; ein bekanntes Geburtsdatum ist erforderlich. "
        "Mindestens eine Stunde vor dem angesetzten Spielbeginn am Spielort sein. "
        "An den technischen Besprechungen mit den Schiedsrichtern spätestens 30 Minuten "
        "vor dem Spiel und direkt nach dem Spiel teilnehmen. Für einen sicheren Spielablauf "
        "sorgen. Zuschauer, die sich unangemessen verhalten oder Spieler bzw. Schiedsrichter "
        "beleidigen, vom Spielort verweisen. Während des Spiels sichtbar durch ein Abzeichen "
        "oder eine Weste gekennzeichnet sein."
    ),
    ROLE_CASH: (
        "Mindestens eine Stunde vor dem angesetzten Spielbeginn am Spielort sein. "
        "Am Eingang Zuschauerkarten für die Erwachsenenspiele verkaufen und den "
        "Ordner als zweiter Ordner unterstützen. Eine Teilnahme an den technischen "
        "Besprechungen mit den Schiedsrichtern ist nicht erforderlich."
    ),
    ROLE_CLEANING: (
        "Als Wischer während des Spiels bei Bedarf einzelne Stellen des Spielfelds "
        "wischen. In Spielfeldnähe bereithalten, jedoch nicht auf der Auswechselbank. "
        "Eine Teilnahme an den technischen Besprechungen mit den Schiedsrichtern ist "
        "nicht erforderlich. Die Schiedsrichter nehmen die Anwesenheit des Wischers "
        "zur Kenntnis."
    ),
    BLOCK_PREPARATION: (
        "Zur angegebenen Uhrzeit am Spielort sein. "
        "Bei der Vorbereitung des Spielfelds und des Verkaufsstands helfen."
    ),
    BLOCK_CAKE_DELIVERY: (
        "Zur angegebenen Lieferzeit einen Kuchen zum Verkaufsstand bringen. "
        "Ein Platz entspricht einem Kuchen."
    ),
    BLOCK_CLEANUP: (
        "Zur angegebenen Uhrzeit am Spielort sein. "
        "Beim Abbau des Verkaufsstands und bei der Reinigung des Spielfelds helfen."
    ),
}
