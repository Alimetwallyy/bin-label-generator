# app/utils.py
import re
from typing import Optional, Dict

# A robust regex that attempts to capture components of many bay id formats.
# Adjust pattern if your real formats differ.
BAY_REGEX = re.compile(
    r"""
    ^\s*
    (?:BAY[-_\s]?)?                    # optional leading BAY or BAY-
    (?P<aisle>\d{1,4})                 # aisle or first numeric block
    (?:[-_\s](?P<section>\d{1,4}))?    # optional section
    (?:[-_\s](?P<number>\d{1,4}))?     # trailing numeric ID (prefer last)
    \s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)


def parse_bay_id(bay_id: str) -> Optional[Dict[str, str]]:
    """Parse a bay identifier into components.

    Returns dict with keys 'aisle', 'section', 'number', 'raw' or None if no match.
    """
    if not bay_id or not bay_id.strip():
        return None
    m = BAY_REGEX.match(bay_id.strip())
    if not m:
        return None
    data = m.groupdict()
    # If number missing, try to extract trailing 3-digit substring
    if data.get("number") is None:
        found = re.findall(r"(\d{3})\b", bay_id)
        if found:
            data["number"] = found[-1]
    return {"raw": bay_id.strip(), "aisle": data.get("aisle") or "", "section": data.get("section") or "", "number": data.get("number") or ""}


def normalize_bay_id(bay: str) -> str:
    """Make a consistent canonical bay ID for duplicate comparison (uppercase, normalized dashes)."""
    b = bay.strip().upper()
    # normalize different dashes to standard hyphen
    b = re.sub(r"[–—_ ]+", "-", b)
    return b
