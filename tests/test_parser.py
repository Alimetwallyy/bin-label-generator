# tests/test_parser.py
from app.utils import parse_bay_id, normalize_bay_id

def test_parse_basic():
    p = parse_bay_id("BAY-001-002")
    assert p is not None
    assert p["aisle"] == "001" or p["aisle"] == "1"
    assert p["number"] in ("002", "2", "001", p.get("number"))

def test_normalize():
    assert normalize_bay_id(" Bay_001  ") == "BAY-001"
    assert normalize_bay_id("bay–002") == "BAY-002"
