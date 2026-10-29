"""Human-readable task names, independent of stored role identifiers."""


def role_label(role: str) -> str:
    """Resolve current wording without changing assignment or transport keys."""
    return "Ordner" if role == "Ordnungsdienst" else role
