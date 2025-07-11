"""Utility helpers for modules."""

def normalize_name(name: str) -> str:
    """Return a whitespace-normalized version of *name*.

    Whitespace characters like tabs, newlines, full width and non‑breaking spaces
    are replaced with regular spaces. Consecutive spaces collapse into one.
    """
    if name is None:
        name = ""
    else:
        name = str(name)

    for ch in ("\u3000", "\u00A0", "\t", "\r", "\n"):
        name = name.replace(ch, " ")
    name = " ".join(name.split())
    return name.strip()
