import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

ROSTER_FILE = "TJ4生产部&生产设备部人员清单，07-31-2026.xlsx"
BONUS_FILE = "2026年07月份安全质量奖核算数据.xlsx"


def _read(name):
    path = os.path.join(ROOT, name)
    if not os.path.exists(path):
        pytest.skip(f"缺少样本文件 {name}")
    with open(path, "rb") as handle:
        return handle.read()


@pytest.fixture(scope="session")
def roster_bytes():
    return _read(ROSTER_FILE)


@pytest.fixture(scope="session")
def bonus_bytes():
    return _read(BONUS_FILE)


@pytest.fixture(scope="session")
def roster(roster_bytes):
    from tj4tools.roster import parse_roster

    return parse_roster(roster_bytes, ROSTER_FILE)


@pytest.fixture(scope="session")
def bonus(bonus_bytes):
    from tj4tools.roster import parse_bonus

    return parse_bonus(bonus_bytes, BONUS_FILE)


@pytest.fixture(scope="session")
def roster_with_interns(roster_bytes):
    from tj4tools.roster import parse_roster

    return parse_roster(roster_bytes, ROSTER_FILE, include_interns=True)


@pytest.fixture(scope="session")
def result(roster, bonus):
    from tj4tools.roster import reconcile

    return reconcile(roster, bonus)
