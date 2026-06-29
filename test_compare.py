"""Isolated tests for compare.py's annuity engine — so we can trust it.

Two parts:
  - EngineVsBank: compare.py's reconstructed initial plan must match Nordea's
    published plan (nordea_payplan_data.json), monthly then yearly — same
    rounding-envelope tolerances as test_payplan.py.
  - Invariants: the engine amortises to zero, principal sums to the loan, a fixed
    payment equal to the annuity reproduces re-annuitisation, the window returns
    the principal at the uploan, and a downpayment lowers the bank payment.

Run:  .venv/bin/python -m unittest -v test_compare
"""
import json
import unittest
from datetime import date
from pathlib import Path

import compare

HERE = Path(__file__).resolve().parent
FIRST_DUE = date(2026, 7, 15)
TOL = 1.5                      # per-term total/avdrag/rente (whole-krone rounding)
_HAS_DATA = ((HERE / "nordea_payplan_data.json").exists()
             and (HERE / "nordea_loan_detail.json").exists())
_SKIP = "no downloaded data — run ./fetch.sh first"


def bal_tol(d):               # accumulated payment-rounding envelope (~1 kr/elapsed month)
    return 1.0 * ((d.year - FIRST_DUE.year) * 12 + (d.month - FIRST_DUE.month)) + 3.0


def _plan_a():
    A = compare.load_plan(None)
    return A, compare.schedule(A["principal"], A["rate"], A["terms"], A["first_due"],
                               A["fee"], A["factor"], A["tpy"])


@unittest.skipUnless(_HAS_DATA, _SKIP)
class EngineVsBank(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.A, cls.plan = _plan_a()
        cls.by_date = {r["date"]: r for r in cls.plan}
        plans = json.loads((HERE / "nordea_payplan_data.json").read_text())["pay_plans"]
        cls.monthly = [p for p in plans if not p["due_date"].endswith("-12-31")]
        cls.yearly = [p for p in plans if p["due_date"].endswith("-12-31")]

    def test_monthly_matches_bank(self):
        for p in self.monthly:
            d = date.fromisoformat(p["due_date"])
            r = self.by_date[d]
            self.assertAlmostEqual(p["total"], r["total"], delta=TOL, msg=f"{d} total")
            self.assertAlmostEqual(p["amount"], r["principal"], delta=TOL, msg=f"{d} avdrag")
            self.assertAlmostEqual(p["interest"], r["interest"], delta=TOL, msg=f"{d} rente")
            self.assertAlmostEqual(p["loan_balance"], r["balance"], delta=bal_tol(d),
                                   msg=f"{d} restgjeld")

    def test_yearly_matches_bank(self):
        for p in self.yearly:
            y = int(p["due_date"][:4])
            months = [r for r in self.plan if r["date"].year == y]
            year_end = max(months, key=lambda r: r["date"])
            tol = bal_tol(year_end["date"])
            self.assertAlmostEqual(p["total"], sum(r["total"] for r in months), delta=tol,
                                   msg=f"{y} total")
            self.assertAlmostEqual(p["amount"], sum(r["principal"] for r in months), delta=tol,
                                   msg=f"{y} avdrag")
            self.assertAlmostEqual(p["loan_balance"], year_end["balance"], delta=tol,
                                   msg=f"{y} restgjeld")


@unittest.skipUnless(_HAS_DATA, _SKIP)
class Invariants(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.A, cls.plan = _plan_a()

    def test_amortises_to_zero(self):
        self.assertLess(self.plan[-1]["balance"], 1.0)
        self.assertEqual(self.plan[-1]["date"], date(2056, 6, 15))

    def test_principal_sums_to_opening(self):
        total_principal = sum(r["principal"] for r in self.plan)
        self.assertAlmostEqual(total_principal, self.A["principal"], delta=5.0)

    def test_fixed_payment_reproduces_reannuitisation(self):
        # with no partial first term (factor=1), paying the annuity as a constant
        # amount must reproduce the re-annuitised schedule exactly.
        A = self.A
        keep = compare.pmt(A["rate"] / A["tpy"], A["terms"], A["principal"])
        args = (A["principal"], A["rate"], A["terms"], A["first_due"], A["fee"], 1.0, A["tpy"])
        base = compare.schedule(*args)
        yours = compare.schedule(*args, fixed=keep)
        self.assertEqual(len(base), len(yours))
        for b, y in zip(base, yours):
            self.assertAlmostEqual(b["balance"], y["balance"], delta=0.01, msg=str(b["date"]))

    def test_window_returns_principal_at_uploan(self):
        A = self.A
        amt, up_idx = 1_005_000.0, 48
        keep = compare.pmt(A["rate"] / A["tpy"], A["terms"], A["principal"])
        yours = compare.schedule(A["principal"], A["rate"], A["terms"], A["first_due"],
                                 A["fee"], A["factor"], A["tpy"], fixed=keep,
                                 adj={0: -amt, up_idx: +amt})
        jump = yours[up_idx]["balance"] - yours[up_idx - 1]["balance"]
        self.assertGreater(jump, amt * 0.9)        # balance jumps back up by ~the uploan
        self.assertLess(jump, amt)                 # minus that term's principal

    def test_downpayment_lowers_bank_payment(self):
        A = self.A
        initial = compare.schedule(A["principal"], A["rate"], A["terms"], A["first_due"],
                                   A["fee"], A["factor"], A["tpy"])[1]["total"]
        reduced = compare.schedule(A["principal"] - 1_005_000, A["rate"], A["terms"],
                                   A["first_due"], A["fee"], A["factor"], A["tpy"])[1]["total"]
        self.assertLess(reduced, initial)


if __name__ == "__main__":
    unittest.main(verbosity=2)
