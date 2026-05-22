"""Offline behavior tests for outlook.py pure logic.

No network, no credentials — imports the script's pure helper functions and
asserts their behavior. Run directly:

    python3 tests/test_offline.py

Exits non-zero if any check fails.
"""
import importlib.util
import sys
from pathlib import Path

_SCRIPT = Path(__file__).resolve().parent.parent / "scripts" / "outlook.py"
spec = importlib.util.spec_from_file_location("outlook", _SCRIPT)
o = importlib.util.module_from_spec(spec)
spec.loader.exec_module(o)

fails = []


def check(name, cond):
    print(("PASS " if cond else "FAIL ") + name)
    if not cond:
        fails.append(name)


def raises(fn):
    try:
        fn()
        return False
    except SystemExit:
        return True


# --- _txid: deterministic, <=40 hex
t1 = o._txid("Dentist", "2026-05-20T10:00:00", "2026-05-20T11:00:00", "calid")
t2 = o._txid("Dentist", "2026-05-20T10:00:00", "2026-05-20T11:00:00", "calid")
t3 = o._txid("Dentist", "2026-05-20T10:00:00", "2026-05-20T11:00:00", "OTHER")
check("txid deterministic", t1 == t2)
check("txid varies by content", t1 != t3)
check("txid is 40 chars", len(t1) == 40)

# --- naive ISO validation
check("rejects Z suffix", raises(lambda: o._validate_naive_iso("2026-05-20T10:00:00Z", "start")))
check("rejects +02:00 offset", raises(lambda: o._validate_naive_iso("2026-05-20T10:00:00+02:00", "start")))
check("accepts naive ISO", o._validate_naive_iso("2026-05-20T10:00:00", "start") is None)

# --- range validation
check("rejects end<=start", raises(lambda: o._validate_range("2026-05-20T11:00:00", "2026-05-20T10:00:00")))
check("accepts end>start", o._validate_range("2026-05-20T10:00:00", "2026-05-20T11:00:00") is None)

# --- recurrence: yearly birthday
r = o._build_recurrence("yearly", "2026-04-15")
check("yearly type", r["pattern"]["type"] == "absoluteYearly")
check("yearly dayOfMonth", r["pattern"]["dayOfMonth"] == 15)
check("yearly month", r["pattern"]["month"] == 4)
check("yearly noEnd", r["range"]["type"] == "noEnd")

# --- weekly with days + until
r = o._build_recurrence("weekly", "2026-05-18", days_raw="mon,thu", until="2026-12-31")
check("weekly days en", r["pattern"]["daysOfWeek"] == ["monday", "thursday"])
check("weekly endDate", r["range"]["type"] == "endDate" and r["range"]["endDate"] == "2026-12-31")

# --- weekly German abbreviations
r = o._build_recurrence("weekly", "2026-05-18", days_raw="mo,di,mi,do,fr,sa,so")
check("weekly days de", r["pattern"]["daysOfWeek"] == ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"])

# --- monthly with count
r = o._build_recurrence("monthly", "2026-06-01", count=12)
check("monthly type", r["pattern"]["type"] == "absoluteMonthly" and r["pattern"]["dayOfMonth"] == 1)
check("monthly numbered", r["range"]["type"] == "numbered" and r["range"]["numberOfOccurrences"] == 12)

# --- daily interval
r = o._build_recurrence("daily", "2026-05-20", interval=3)
check("daily interval", r["pattern"]["type"] == "daily" and r["pattern"]["interval"] == 3)

# --- mutually exclusive until+count
check("until+count mutually exclusive", raises(lambda: o._build_recurrence("daily", "2026-05-20", until="2026-12-31", count=5)))

# --- unknown recurrence/weekday
check("unknown recurrence rejected", raises(lambda: o._build_recurrence("hourly", "2026-05-20")))
check("unknown weekday rejected", raises(lambda: o._normalize_weekdays("funday", None)))

# --- _looks_like_id
check("id detected", o._looks_like_id("AAMk" + "x" * 60))
check("name not id", not o._looks_like_id("Work"))

# --- no recurrence
check("no recurrence returns None", o._build_recurrence(None, "2026-05-20") is None)

print()
print("RESULT:", "ALL PASS" if not fails else f"{len(fails)} FAILED: {fails}")
sys.exit(1 if fails else 0)
