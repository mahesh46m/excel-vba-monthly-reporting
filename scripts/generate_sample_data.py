"""Generate the sample input for the VBA monthly-reporting demo.

All data is synthetic and randomly generated (fixed seed). It does not come
from any real company or employer, and the numbers are not real-world metrics.

Produces:
  data/transactions_2026-01.csv  raw extract the VBA macro imports (defects included)
  data/expected_summary.csv      what the summary sheet should contain, computed
                                 independently in Python so the macro's output
                                 can be checked against it
"""
import csv
import random
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

SEED = 5
DATA = Path(__file__).resolve().parents[1] / "data"

CATEGORIES = ["Hardware", "Software", "Services", "Support", "Training"]
REGIONS = ["North", "South", "East", "West"]
REPS = ["A. Rao", "B. Smith", "C. Lee", "D. Patel", "E. Garcia", "F. Chen"]


def main() -> None:
    rng = random.Random(SEED)
    DATA.mkdir(exist_ok=True)
    rows = []
    tid = 90000
    for day in range(31):
        d = date(2026, 1, 1) + timedelta(days=day)
        for _ in range(rng.randint(4, 9)):
            tid += 1
            rows.append(
                [
                    f"T{tid}",
                    str(d),
                    rng.choice(CATEGORIES),
                    rng.choice(REGIONS),
                    rng.choice(REPS),
                    round(rng.uniform(200, 6000), 2),
                ]
            )
    # Defects the macro must handle:
    dupes = [list(r) for r in rng.sample(rows, 5)]          # 5 exact duplicates
    for r in rng.sample(rows, 4):                            # 4 padded category labels
        r[2] = f"  {r[2]} "
    for r in rng.sample(rows, 3):                            # 3 US-style dates
        y, m, dd = r[1].split("-")
        r[1] = f"{m}/{dd}/{y}"
    rows.extend(dupes)
    rng.shuffle(rows)

    with (DATA / "transactions_2026-01.csv").open("w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["txn_id", "date", "category", "region", "rep", "amount"])
        w.writerows(rows)

    # Expected results after cleaning: dedupe on txn_id, trim labels,
    # normalise dates. Computed here as ground truth for the macro output.
    seen = set()
    by_cat = defaultdict(float)
    by_region = defaultdict(float)
    total = 0.0
    count = 0
    for txn_id, d, cat, region, rep, amount in rows:
        if txn_id in seen:
            continue
        seen.add(txn_id)
        cat = cat.strip()
        if "/" in d:
            m, dd, y = d.split("/")
            d = f"{y}-{int(m):02d}-{int(dd):02d}"
        by_cat[cat] += amount
        by_region[region] += amount
        total += amount
        count += 1

    with (DATA / "expected_summary.csv").open("w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["dimension", "member", "total_amount"])
        for k in sorted(by_cat):
            w.writerow(["category", k, round(by_cat[k], 2)])
        for k in sorted(by_region):
            w.writerow(["region", k, round(by_region[k], 2)])
        w.writerow(["overall", "all", round(total, 2)])

    print(f"rows written: {len(rows)} (incl. 5 injected duplicates)")
    print(f"clean rows: {count}; grand total: {total:,.2f}")
    print("by category:", {k: round(v, 2) for k, v in sorted(by_cat.items())})
    print("by region:", {k: round(v, 2) for k, v in sorted(by_region.items())})


if __name__ == "__main__":
    main()
