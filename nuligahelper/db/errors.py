# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Exceptions raised by the domain layer.
# ---------------------------------------------------------------

from __future__ import annotations


class SQLiteInitializationError(RuntimeError):
    """Raised when the configured SQLite runtime invariants cannot be established."""


class DatabaseMigrationRequiredError(SQLiteInitializationError):
    """Raised when a legacy game-identity schema must be converted by an earlier release."""


class AssignmentTemporarilyUnavailableError(RuntimeError):
    """Raised when an assignment CAS cannot finish within the lock deadline."""

    def __init__(self):
        super().__init__(
            "Die Datenbank ist vorübergehend ausgelastet. Bitte versuche es erneut."
        )


class SlotConflictError(ValueError):
    def __init__(self, current_person_id: int | None):
        super().__init__("Der Aufgabenplatz wurde zwischenzeitlich geändert.")
        self.current_person_id = current_person_id


class DutyLockedError(ValueError):
    """The saved reminder event made this occupied commitment immutable."""

    def __init__(self):
        super().__init__("Diese Einteilung ist seit der Vortagserinnerung verbindlich gesperrt.")


class CakeConfigurationConflictError(ValueError):
    """A stale settings edit, with the currently saved public configuration."""

    def __init__(self, current_delivery_time: str | None, current_cake_quantity: int | None):
        super().__init__("Die Kuchenlieferung wurde zwischenzeitlich geändert.")
        self.current_delivery_time = current_delivery_time
        self.current_cake_quantity = current_cake_quantity


class AgeEligibilityError(ValueError):
    """A public-safe structured eligibility refusal, without a person's age."""

    def __init__(self, reason: dict):
        super().__init__(reason["message"])
        self.reason = reason
