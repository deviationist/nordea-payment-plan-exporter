#!/usr/bin/env python3
"""Refresh the raw Nordea JSON sources by driving a real login and intercepting
the bank's own authenticated API responses (direct fetches 401 — the SPA holds a
bearer token, so we let the app make the calls and capture the responses).

Automated: navigate, fill Fødselsnummer (.env SSN) and BankID password
(.env BANKID_PWD). Manual: you confirm on your phone / enter the OTP — the script
waits for that. Credentials are read from .env and never printed or logged.

Saves an immutable, timestamped snapshot under captures/ (gitignored) — never
overwrites a prior capture, never touches the committed baseline. Render it with
`./build.sh --capture <timestamp>`.

    ./fetch.sh        # bootstraps a playwright venv and runs this
"""
import argparse
import hashlib
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

# a realistic UA helps avoid trivial headless fingerprinting
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

HERE = Path(__file__).resolve().parent

# endpoint patterns (note: list must not match the detail/pay-plans URLs)
PAT_PAY = re.compile(r"/loans-v1/loans/\d+/pay-plans(\?|$)")
PAT_DETAIL = re.compile(r"/loans-v1/loans/\d+(\?|$)")
PAT_LIST = re.compile(r"/loans-v1/loans(\?|$)")


def load_env(path=HERE / ".env"):
    env = {}
    for line in path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env


def dismiss_cookies(page):
    """The Nordea landing page shows a cookie overlay (#CookieReportsPanel) that
    intercepts the 'Logg inn' click. Accept it if we can, else remove the overlay."""
    page.wait_for_timeout(1200)
    try:
        page.get_by_role(
            "button", name=re.compile(r"tillat alle|godta alle|aksepter|accept all", re.I)
        ).first.click(timeout=3000)
        return
    except Exception:
        pass
    page.evaluate(
        "() => { const p = document.getElementById('CookieReportsPanel'); if (p) p.remove(); }")


def main():
    ap = argparse.ArgumentParser(
        description="Refresh the Nordea JSON sources via a real login.")
    ap.add_argument("--headed", action="store_true",
                    help="run with a visible browser window (default: headless; "
                         "use if Nordea ever blocks the headless browser)")
    args = ap.parse_args()

    env = load_env()
    ssn, pwd = env.get("SSN"), env.get("BANKID_PWD")
    if not ssn or not pwd:
        sys.exit("Missing SSN or BANKID_PWD in .env")

    captured = {}

    def on_response(resp):
        url = resp.url
        try:
            if PAT_PAY.search(url) and resp.ok:
                captured["pay"] = resp.json()
            elif PAT_DETAIL.search(url) and not PAT_PAY.search(url) and resp.ok:
                captured["detail"] = resp.json()
            elif PAT_LIST.search(url) and resp.ok:
                captured["list"] = resp.json()
        except Exception:
            pass

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=not args.headed, args=["--disable-blink-features=AutomationControlled"])
        page = browser.new_context(user_agent=UA).new_page()
        page.on("response", on_response)

        # --- login (recorded flow) ---
        page.goto("https://netbank.nordea.no/")
        dismiss_cookies(page)
        page.get_by_role("button", name="Logg inn").click()
        page.get_by_role("textbox", name="Fødselsnummer").fill(ssn)   # .env SSN
        page.keyboard.press("Enter")
        page.get_by_role("button", name="Bekreft innlogging").click(timeout=60000)
        frame = page.frame_locator('iframe[title="BankID"]')
        frame.get_by_role("textbox", name="Ditt BankID-passord").fill(pwd, timeout=60000)
        frame.get_by_role("button", name="Neste").click()

        print(">>> Confirm on your phone / enter the OTP in the browser window …",
              flush=True)
        page.wait_for_url("**/overview/**", timeout=240000)   # waits for you
        print(">>> Logged in. Capturing loan data …", flush=True)

        # --- pick the loan from the API data, then deep-link (no DOM selector) ---
        # The overview load fires /loans-v1/loans (intercepted into captured["list"]).
        # Choose the loan from that structured data; the /loans/details/<hash> route
        # is sha256(loan_id), so we deep-link straight to its pay-plans page, which
        # fires BOTH the detail and pay-plans calls.
        page.goto("https://netbank.nordea.no/overview/")
        page.wait_for_timeout(2000)

        with_plan = [l for l in captured.get("list", {}).get("loans", [])
                     if l.get("has_repayment_plan")]
        want = env.get("LOAN_ID")
        if want:                                     # explicit choice (disambiguates >1 loan)
            chosen = next((l for l in with_plan if l.get("loan_id") == want), None)
            if not chosen:
                sys.exit(f"LOAN_ID {want} not found among loans with a repayment plan")
        else:                                        # else the single mortgage with a plan
            pool = [l for l in with_plan if l.get("group") == "mortgage"] or with_plan
            if len(pool) != 1:
                found = ", ".join(f'{l.get("loan_id")} ({l.get("group")})' for l in pool) or "none"
                sys.exit(f"Could not pick a single loan — set LOAN_ID in .env. Candidates: {found}")
            chosen = pool[0]

        loan_id = chosen["loan_id"]
        loan_hash = hashlib.sha256(loan_id.encode()).hexdigest()   # = the /details/<hash> route
        page.goto(f"https://netbank.nordea.no/loans/details/{loan_hash}/pay-plans")
        page.wait_for_timeout(2500)                  # fires detail + pay-plans (intercepted)

        browser.close()

    missing = {"pay", "detail"} - captured.keys()
    if missing:
        sys.exit(f"Did not capture: {missing} (got {sorted(captured)})")

    # Write the canonical downloaded files (what ./build.sh reads by default) AND an
    # immutable, timestamped snapshot under captures/ for pre/post comparison.
    # All of these are gitignored — personal data never enters version control.
    ts = datetime.now().strftime("%Y-%m-%dT%H%M%S")
    capdir = HERE / "captures"
    capdir.mkdir(exist_ok=True)
    written = []
    for key, stem, canonical in (("pay", "payplan", "nordea_payplan_data.json"),
                                 ("detail", "loan-detail", "nordea_loan_detail.json")):
        blob = json.dumps(captured[key], ensure_ascii=False) + "\n"
        (HERE / canonical).write_text(blob, encoding="utf-8")          # latest (default build)
        snap = capdir / f"{stem}-{ts}.json"
        if snap.exists():
            sys.exit(f"Refusing to overwrite existing capture {snap}")
        snap.write_text(blob, encoding="utf-8")                        # immutable archive
        written.append(snap.relative_to(HERE))

    bal = captured["detail"].get("amount", {}).get("balance")
    print(f"\nDownloaded (loan balance {bal}) → nordea_payplan_data.json + "
          f"nordea_loan_detail.json", flush=True)
    print(f"Archived snapshot {ts}:", flush=True)
    for p in written:
        print(f"  {p}", flush=True)
    print(f"\n  build latest:    ./build.sh", flush=True)
    print(f"  build snapshot:  ./build.sh --capture {ts}", flush=True)


if __name__ == "__main__":
    main()
