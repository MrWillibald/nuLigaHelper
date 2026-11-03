# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Required positions, vacancies and age-eligibility checks per game.
# ---------------------------------------------------------------

from __future__ import annotations

from typing import TYPE_CHECKING

from nuligahelper.db import age_eligibility

from nuligahelper.db.roles import (
    ROLE_SALE,
    ROLE_SECRETARY,
    ROLE_SECURITY,
    ROLE_SLOT_COUNT,
    ROLE_TIMEKEEPER,
    offered_positions,
    required_positions,
)
from nuligahelper.db.games import classify_game_category

if TYPE_CHECKING:
    from nuligahelper.db.models import Game, Person


def missing_slots(game: Game) -> dict[str, int]:
    """Open task slots of a game as {role: missing_count}, in display order."""
    result = {}
    occupied = {(assignment.role, assignment.slot) for assignment in game.assignments}
    for role, slot in required_positions(game):
        if (role, slot) not in occupied:
            result[role] = result.get(role, 0) + 1
    return result


def claim_eligibility(game: Game, role: str, person: Person, slot: int = 0) -> dict | None:
    """Eligibility of a hypothetical claim; authorization and CAS stay separate."""
    if (role, slot) not in offered_positions(game):
        return {"role": role, "slot": slot, "code": "duty_not_offered",
                "message": "Diese Aufgabe wird für dieses Spiel nicht angeboten."}
    reason = age_eligibility.individual_eligibility(game, role, person, slot)
    if reason is not None or role != ROLE_SALE:
        return reason
    if classify_game_category(game.ak) == age_eligibility.YOUTH:
        return None
    if age_eligibility.parse_game_date(game.date) is None:
        return age_eligibility.deficiency(
            ROLE_SALE, "unresolved_game_date", minimum_age=18
        )
    sellers = [assignment.person for assignment in game.assignments_by_role(ROLE_SALE)
               if assignment.slot != slot]
    sellers.append(person)
    # Incremental staffing may leave coverage outstanding. The final claim
    # must establish coverage; no individual minimum applies to the other seller.
    if len(sellers) < ROLE_SLOT_COUNT[ROLE_SALE]:
        return None
    return age_eligibility.sale_coverage(game, sellers)


def staffing_status(game: Game) -> dict:
    """Keep physical vacancies distinct from invalid or unresolved eligibility."""
    vacancies = missing_slots(game)
    deficiencies = []
    for role in (ROLE_TIMEKEEPER, ROLE_SECRETARY, ROLE_SECURITY):
        for assignment in game.assignments_by_role(role):
            reason = age_eligibility.individual_eligibility(
                game, role, assignment.person, assignment.slot
            )
            if reason is not None:
                deficiencies.append(reason)
    sale_reason = age_eligibility.sale_coverage(
        game, [assignment.person for assignment in game.assignments_by_role(ROLE_SALE)]
    )
    if sale_reason is not None:
        deficiencies.append(sale_reason)
    required = required_positions(game)
    occupied = {(assignment.role, assignment.slot) for assignment in game.assignments}
    category = classify_game_category(game.ak)
    return {"vacancies": vacancies, "deficiencies": deficiencies,
            "category": category,
            "classification_unresolved": category == age_eligibility.UNKNOWN,
            "required_filled": sum(position in occupied for position in required),
            "required_total": len(required),
            "complete": not vacancies and not deficiencies}
