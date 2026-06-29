# nordea-payment-plan-exporter

Turns a Nordea mortgage **nedbetalingsplan** (amortization plan) into a
spreadsheet — including the per-term breakdown (Avdrag / Rente / Gebyrer)
that the bank hides behind the "Beskrivelse" link — and builds a **dynamic,
formula-driven Excel model** you can re-run at any interest rate.

## Why this exists (the rationale)

The mortgage is a **floating-rate annuity loan** (nedbetalingslån, not a rammelån).
This models a **temporary extra downpayment**: a lump sum applied to the loan for a
fixed window, then withdrawn again at the end (via an *opplåning* / loan increase,
since you can't freely redraw from a nedbetalingslån). The amount and window dates
are configured in `.env` (see *Config*).

The strategy this tool supports:

1. **Apply the lump sum** → the mortgage balance drops, so the bank charges less
   interest each month.
2. **Keep paying the original terminbeløp anyway.** Since interest is now lower,
   more of that unchanged payment goes to **principal** — the "freed-up" interest
   becomes extra downpayment. That head-start is the **forsprang**.
3. **At the window end**, the lump is withdrawn again via opplåning. It leaves,
   but the **interest you saved while it was deployed stays as principal** — a
   permanent head-start that clears the loan earlier.

**Why a custom tool rather than the bank's plan?** When the downpayment lands,
Nordea *recalculates* to a **lower** required payment and its system only holds
*its* plan — it won't maintain "keep paying the old amount" for you. So this
spreadsheet is the **authoritative record of the original payment plan you intend
to keep paying**, plus the tracker for how the temporary downpayment runs you ahead
of it. Forsprang is always measured **against the initial plan**.

Two phases: **estimate now** (re-annuitised from the loan parameters) → **hard
comparison later** (diff the bank's real pre- and post-downpayment exports). All
personal values (amount, dates, login) live in `.env`, never in the code.

## Config (`.env`)

Copy `.env.example` → `.env` (gitignored) and fill in:

| Key | Used by | Meaning |
|---|---|---|
| `SSN`, `BANKID_PWD` | `fetch.sh` | Nordea login (OTP/phone stays manual) |
| `DOWNPAYMENT` | `build.sh`, `compare.sh` | temporary downpayment in kr (e.g. `500000`; empty/0 = initial plan only) |
| `DOWN_DATE` | `build.sh`, `compare.sh` | when it's paid, `YYYY-MM-DD` (default: first term) |
| `UP_DATE` | `build.sh`, `compare.sh` | uploan repayment date, `YYYY-MM-DD` (e.g. `2030-01-01`) |

Nothing personal is hardcoded; CLI flags (`--downpayment`, `--down-date`,
`--up-date`) override `.env` for one-offs.

## Data source

The numbers come from two Nordea Nettbank JSON endpoints, downloaded by `fetch.sh`
into two **gitignored** files — personal data never enters version control:

- `GET /api/dbf/ca/loans-v1/loans/<id>/pay-plans` → `nordea_payplan_data.json`
- `GET /api/dbf/ca/loans-v1/loans/<id>` → `nordea_loan_detail.json`

`build_sheet.py` derives the loan facts from `nordea_loan_detail.json` (nothing
hand-typed). On a clean checkout: `cp .env.example .env` (fill it in) →
`./fetch.sh` to download → `./build.sh`. Re-run `./fetch.sh` any time to refresh.

## Usage

```sh
./build.sh          # build both CSV and XLSX
./build.sh csv      # CSV only
./build.sh xlsx     # XLSX only
```

`build.sh` bootstraps a local `.venv` on first run (installs `requirements.txt`)
and then runs `build_sheet.py --format <both|csv|xlsx>`. No Nordea login needed —
it reads the saved JSON sources. You can also call the script directly in an
existing venv: `python build_sheet.py --format both`.

Produces:

- **`nordea-nedbetalingsplan.xlsx`** — three sheets:
  - **Modell (dynamisk)** — inputs + a full 360-month annuity schedule, entirely
    formula-driven. Yellow cells are editable. Each month's nominal rate (col C)
    carries forward, so changing one cell re-annuitises every later month
    (payment, interest, balance, payoff date) — like a real floating-rate loan.
    Includes the **temporary downpayment window** scenario (cols L–Q, below).
  - **Lånedetaljer** — static loan facts.
  - **Bankens plan (original)** — the bank's own published plan, as extracted.
- **`nordea-nedbetalingsplan.csv`** — flat snapshot of the dynamic monthly
  schedule (computed values; one row per month).

## Notes

- **Nominell vs. effektiv rente:** the schedule applies the **nominell** rate
  (e.g. 4,99 %) to the balance each month. **Effektiv rente** is a computed output
  = nominal + monthly compounding + the per-term fee, **rounded up to 2 decimals**
  (the bank's reporting convention) in cells B11 (compounding only) / B12 (IRR
  incl. fee). It's forward-looking, so it lands a hair under the bank's
  origination figure (e.g. ~5,12 % vs the bank's stated 5,13 %, the gap being the
  short first-term stub and the sunk one-time setup fees).
- The bank's published plan may be **stale** for a floating-rate loan (computed
  at an older rate). The dynamic model is the way to see your real cost today.
- **First term is a partial period.** Nordea bills interest for only part of the
  first month (the loan is disbursed mid-month), so term 1's interest is scaled
  by the factor in cell **B10** (≈ 20/30 days, auto-derived from the bank data).
  The avdrag/principal and balance follow the normal annuity, matching the bank's
  first row to the krone (±rounding).

## Ekstra nedbetaling (temporary downpayment window)

Models a **temporary extra downpayment for a fixed window** — a lump sum applied to
the mortgage at the start and **withdrawn at the end** (via opplåning). Throughout,
you keep paying the original monthly amount, so the lower balance redirects interest
into principal while the lump is deployed; the lump is returned, but the **interest
saved during the window stays as a head-start** that clears the loan early.

Inputs (yellow cells, top of the Modell sheet — all from `.env`, see *Config*):

- `D5` **Beløp** — the lent amount (`.env` `DOWNPAYMENT`).
- `D6` **Startdato** / `D7` **Sluttdato** — the window (`.env` `DOWN_DATE` /
  `UP_DATE`). `D8`/`D9` show the matching term numbers.

Scenario columns (L–Q), alongside the baseline A–K ("before" plan):

- `L` ±-adjustment (−beløp at the start term, +beløp redraw at the end term)
- `M` scenario opening balance · `N` scenario interest
- `O` **Ekstra avdrag pr. mnd** — interest saved that month = extra principal paid
- `P` scenario balance · `Q` **Forsprang** — how far ahead of the original plan

Summary (cells D11–D14): scenario payoff date, months saved, total interest
saved, and **Forsprang etter vindu** (the lasting head-start once the loan is
repaid). Set `DOWNPAYMENT=0` (or `./build.sh --downpayment 0`) for the initial
plan with no scenario.

## Comparison report (`compare.sh`)

A static **initial vs reduced** report (CSV + XLSX → `nordea-sammenligning.*`):

```sh
./compare.sh                                   # estimate of the reduced plan
./compare.sh --down-date 2026-07-15 --up-date 2030-07-15
./compare.sh --post <ts>                       # hard comparison vs a captured reduced plan
```

Columns, month by month: the **initial** plan's payment/balance, the **reduced**
plan's (Nordea's recalculated lower payment), your **top-up** (initial − reduced,
the extra you transfer), your **balance** if you keep paying the initial amount,
and your **forsprang vs the initial plan**. The downpayment is modelled as a **window**
— paid in at `--down-date`, returned via opplåning at `--up-date` — so the
forsprang peaks during the window and settles to the lasting benefit afterward.

The reduced plan is an **annuity estimate** by default; pass `--post <timestamp>`
(a `fetch.sh` capture taken *after* the downpayment lands) to diff against
Nordea's **real** recalculated plan.

## Tests

```sh
./build.sh test          # runs test_payplan + test_compare
```

`test_compare.py` validates `compare.py`'s annuity engine in isolation — its
reconstructed initial plan matches the bank (monthly + yearly), plus invariants
(amortises to zero, principal sums to the loan, a fixed payment reproduces
re-annuitisation, the window returns the principal at the uploan).

`test_payplan.py` rebuilds the workbook and compares our computed schedule (the
cached values a viewer actually shows) against Nordea's own
`nordea_payplan_data.json` — **monthly** for the bank's monthly rows
(15.07.2026 → 15.12.2027), **yearly aggregates + year-end balance** for the rest
(2028 → 2056). It prints the max deviation per metric.

Findings: per-month figures match to the **krone** (max ≈ 1 kr; balance ≈ 7 kr by
end-2027); total interest over the life matches to **0,004 %**. The only larger
gap (~333 kr ≈ 0,006 %) is in the final payoff year and is fully explained — the
bank's steady payment is the whole-krone **29 246** vs our exact annuity
**29 245,58**, so Nordea pays principal down ~0,42 kr/month faster and the
residual lands in the last payment. Tolerances scale with that known
payment-rounding envelope (~1 kr per elapsed month).

A second, **skipped** test (`ScenarioVsBankRecalc`) activates once
`nordea_payplan_after_downpayment.json` (the bank's *recalculated* plan, captured
after the downpayment actually lands) is dropped in — see *Planned*.

## Automated data collection (`fetch.sh`)

```sh
cp .env.example .env     # fill in SSN + BANKID_PWD
./fetch.sh               # headless (default; only the phone-push is manual)
./fetch.sh --headed      # visible window (fallback if Nordea blocks headless)
```

Drives a real Nordea login with Playwright: fills **Fødselsnummer** (`.env:SSN`)
and **BankID password** (`.env:BANKID_PWD`), then waits while **you approve the
BankID push on your phone** (or type the OTP). It then intercepts the bank's own
authenticated API responses — direct `fetch()` of the endpoints returns **401**
(the SPA holds a bearer token), so the script lets the app make the calls and
captures them — discovers the loan via `/loans-v1/loans`.

Each run writes an **immutable, timestamped snapshot** under `captures/`
(gitignored) — e.g. `captures/payplan-2026-07-01T0930.json` +
`loan-detail-2026-07-01T0930.json`. It **never overwrites** a prior capture and
**never touches the committed baseline** (`nordea_*.json`), so a post-downpayment fetch
can't clobber the pre-downpayment plan. Render a snapshot with:

```sh
./build.sh --capture 2026-07-01T0930          # build that snapshot
./build.sh                                     # the committed baseline (default)
```

- Credentials live in `.env` (gitignored, plaintext) and are never printed/logged.
- **Headless by default** — works because the only manual step is the out-of-band
  phone approval (validated end-to-end). Anti-fingerprinting tweaks (realistic UA,
  disabled automation flag) help; if Nordea ever blocks the headless browser, run
  `./fetch.sh --headed`.
- Bootstraps its own `.venv-fetch` (Playwright + Chromium) on first run, separate
  from the build venv so the build path stays dependency-free.

## Planned

- **Scenario vs. bank recalc** — once the downpayment actually lands, capture Nordea's
  *recalculated* plan (run `fetch.sh`, save as `nordea_payplan_after_downpayment.json`)
  to activate the skipped `ScenarioVsBankRecalc` test. Note it only correlates if
  the bank keeps the payment fixed (vs. re-annuitising to a lower one) — compare
  against a *permanent* extra paydown (window with no redraw).

## Privacy

The JSON sources, the README, and the output files contain personal financial
data (IBAN, account numbers, name, balances). `.gitignore` excludes the
`.csv`/`.xlsx` outputs; the JSON sources are committed so the build is
reproducible. Keep the repo private.
