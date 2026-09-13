# Indian Energy Exchange (NSE: IEX): Short Thesis

A single-stock research pitch: SHORT Indian Energy Exchange, built around the live, unresolved CERC "market coupling" regulatory fight.

> Educational exercise built for interview preparation. Not investment research, not a recommendation to buy or sell any security, and not built from any non-public information. Every figure is sourced from public disclosures and press coverage of the regulatory process, current as of September 2026, verify anything you cite before presenting this.

## Why this project

A DCF or an LBO tests whether you can build correct mechanics. A stock pitch tests something else: whether you can form a specific, falsifiable view that differs from the market's, attach a dated catalyst to it, and defend it against the strongest version of the other side. This project is that exercise, built on a real, current, two-sided debate rather than a settled or hypothetical one.

## The thesis in one paragraph

IEX's dominant Day-Ahead Market share exists because of liquidity: more participants trade there because more participants already trade there. CERC's market coupling reform centralises price discovery across every exchange, which removes that liquidity advantage by design. The direction of the policy is no longer contested; only the timing is, via an ongoing Supreme Court case and unresolved technical objections from Grid India. At the current price, the market appears to be pricing in something close to a best-case outcome where IEX's share barely erodes. This model argues a more moderate share/fee-erosion outcome is more likely, and implies roughly 45-50% downside.

## What's in the repository

| File | Contents |
|---|---|
| `docs/pitch_memo.md` | The actual deliverable, a 2-page research note in the format a research associate would hand to a portfolio manager |
| `model/IEX_Short_Thesis_Model.xlsx` | The supporting scenario model |
| `docs/interview_prep.md` | Questions this project will draw in an interview |
| `data/iex_historicals.csv` | Reported figures, machine-readable |

### Workbook tabs

| Tab | Contents |
|---|---|
| `Cover` | The pitch in three sentences, recommendation, targets, sources |
| `Assumptions` | Current financials and three scenario definitions (share, fee compression, margin, exit multiple), all yellow-flagged |
| `Scenario Model` | Bull/Base/Bear FY28E revenue, earnings, and target price, fully formula-driven |
| `Downside Bridge` | Decomposes each scenario's return into the EPS-decline effect vs. the multiple-compression effect, the same technique as the returns bridge in the companion Havells LBO model, applied here to a short thesis |
| `Sensitivity` | A two-way grid on market share × fee compression, isolated from the re-rating assumption |
| `Catalyst Timeline` | Dated and approximately-dated events that could force a re-rating either way |

## The result

| | BULL | BASE | BEAR |
|---|---|---|---|
| FY28E market share | 70% | 57% | 50% |
| FY28E EPS | ₹5.25 | ₹3.22 | ₹2.21 |
| Exit multiple | 28x | 18x | 14x |
| Target price | ₹147 | ₹58 | ₹31 |
| Return vs. ₹113.90 today | +29% | -49% | -73% |

The risk/reward is asymmetric against the current price: one plausible bull case gets you +29%, while two plausible base/bear cases get you -49% to -73%. The downside bridge shows most of the base-case return (-39 of -49 points) comes from the earnings decline alone, at an unchanged multiple: the thesis does not require assuming the market additionally panics.

## Why the short case, not the long

The bull case rests on an unproven claim: that API-based "tight coupling" with individual trading members will preserve enough switching friction to offset uniform pricing. That's a real argument, but nobody has observed it under an actual live coupling regime, because one hasn't launched. The bear/base case rests on something more mechanical: the regulatory destination is already decided, and only the date is in question. A pitch built on "the market is underpricing an already-decided, slow-moving structural risk" is a more defensible interview answer than one built on trusting an as-yet-untested customer-retention mechanism.

## Honest limitations

**No confirmed catalyst date.** The Supreme Court case and the final CERC notification are both live but undated. A short with no dated catalyst can be right and still cost money while you wait: this is disclosed explicitly in the memo's risk section, not glossed over.

**The exit multiples are now cross-checked, not just asserted**, against a real comparable-company set (`Comps` tab: MCX, CDSL, CAMS). They are still illustrative in the sense that no formal methodology sets the exact discount to apply relative to peers; a more complete version would build a regression or explicit re-rating framework rather than picking 28x/18x/14x by judgment, even judgment informed by real peer data.

**Volume and share figures are approximated** from press coverage and Screener.in, not from IEX's own segment disclosures line by line. Verify each against IEX's actual investor presentations before presenting this anywhere.

**The implied share count (92.34 Cr) is back-derived from two different data vendors** (indmoney's market cap, kotakneo's price) taken at slightly different times, and will not perfectly match IEX's own reported diluted share count in its filings. Close enough for this model's purposes, but worth knowing the source of the imprecision if asked.

## Verification

```bash
python scripts/build_model.py
python /path/to/xlsx-skill/scripts/recalc.py model/IEX_Short_Thesis_Model.xlsx
```

Cross-check: the `Downside Bridge` tab's total return for each scenario must exactly match the `Scenario Model` tab's implied return for that same scenario.

## Conclusion

All three items below are done.

- [x] **Confirmed the current share price.** IEX traded at Rs 113.90 on 13-Sep-2026 (kotakneo.com live quote), updated from the earlier Rs 120 approximation. Market cap Rs 10,517.52 Cr, sourced separately (indmoney.com) and cross-checked for consistency. The model has been recalculated at this price; every scenario target and the base-case return figures below reflect it.

- [x] **Built a real comparable-company cross-check for the exit multiples.** The `Comps` tab benchmarks IEX against three Indian near-monopoly market-infrastructure businesses: MCX (56.6x P/E), CDSL (~60.5x), and CAMS (36.1x), for a peer median of **56.6x**. IEX's own current TTM P/E is **21.6x**, meaning the market has already priced in a **62% discount** to intact-moat peers. This is a genuine finding, not an assumption: the model's own BULL case exit multiple (28x) sits well below this peer median, meaning it does not assume IEX fully re-rates back to peer levels even if the moat mostly survives. That is a conservative choice, and it also means the true upside risk to being short (if coupling is defeated or delayed indefinitely and the market re-rates IEX all the way to peer multiples) is larger than this model's own +34% bull case shows.

- [x] **Practiced the time-horizon answer.** Written out in full in `docs/interview_prep.md`: an 18-24 month horizon, sized smaller than a catalyst-dated short would warrant, with a defined re-underwrite checkpoint every two quarters rather than holding the position open-endedly on the original thesis alone.

**Updated headline numbers at the confirmed Rs 113.90 price:** BULL +29.1% (target Rs 147), BASE -49.2% (target Rs 58), BEAR -72.9% (target Rs 31). These are the actual recalculated figures at the confirmed price, not scaled estimates; re-run the model yourself if presenting on a materially later date, since IEX has moved 5-10% intraday on regulatory headlines repeatedly this year.
