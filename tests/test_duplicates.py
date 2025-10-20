# tests/test_duplicates.py
from app.logic import check_duplicate_bay_ids

def test_duplicates_detection():
    groups = [
        ["BAY-001-001", "BAY-001-002"],
        ["bay-001-002", "BAY-002-001"],
    ]
    res = check_duplicate_bay_ids(groups)
    assert "BAY-001-002" in res["duplicates"]
    assert res["count"] == len(res["duplicates"])
