![vba_of_wall_street](https://github.com/njgeorge000158/The-VBA-of-Wall-Street/assets/137228821/45e17b05-4811-47ca-bb51-13a2eef4ffbd)

----

# **VBA-Driven Stock Trading Analysis: 2018–2020**

----

## **Project Overview**

This project encompassed developing a VBA script to format, summarize, and analyze approximately 2.26 million stock trading records spanning 3,000 tickers across three calendar years — 2018, 2019, and 2020. The primary objective was to transform raw, high-volume trading data into a structured, searchable summary that reveals meaningful performance metrics at both the individual ticker and cross-market levels.

The script processes a year's dataset by iterating over every trading record and aggregating key statistics about each ticker: opening price, closing price, total trading volume, and net price change. From these aggregates, it calculates the percent change and identifies the tickers with the greatest total volume, highest percent gain, and steepest percent loss. This cross-market summary provides an at-a-glance view of standout performers and laggards without requiring manual filtering or pivot analysis.

Beyond its analytical outputs, the script also applyies conditional formatting to highlight positive returns in green and negative returns in red, making performance patterns immediately legible across a dense summary table.

The result is a pipeline that collapses millions of raw rows into a compact, formatted, and analytically rich summary that reduces what would otherwise require hours of manual Excel work into a single automated process.

## **Development and Testing**

Before processing the full dataset, I established a controlled testing environment using a small, alphabetically organized subset of the trading records. Working against a reduced sample eliminated the computational overhead of 2.26 million rows during early development, enabling rapid iteration and debugging without waiting on full-scale execution cycles.

This staged approach was essential for catching logic errors in the aggregation routines, particularly edge cases involving ticker boundaries, where a naïve row-comparison strategy could place the final record of one ticker to the opening of the next. Validating against an alphabetically structured subset made these boundary conditions predictable and easy to inspect, since the expected groupings were immediately obvious from the data itself.

Once the script produced correct, stable results on the test data, I deployed it across three separate Excel workbooks, one per calendar year, each containing the complete trading history for that period. Running the script on the full dataset confirmed that performance held at scale and that the aggregation logic generalized cleanly beyond the controlled test conditions.

## **Script Functionality**

For each calendar year, the script executes two distinct analytical passes over the raw trading data, producing a two-tiered output that moves from granular per-ticker detail to market-wide extremes.

In the first pass, the script traverses every row of the dataset and computes three core metrics for each ticker: absolute yearly price change, percentage price change, and cumulative trading volume. Yearly price change is derived by subtracting the ticker's first recorded opening price from its last recorded closing price, anchoring the calculation to market open and close rather than arbitrary intermediate values. Percent change is then computed from this difference, normalizing returns across tickers regardless of their nominal price scale. Cumulative volume is accrued row by row across all trading sessions for the year. Then, the script writes these three metrics to a structured summary table in the same workbook as the source data, keeping inputs and outputs on the same worksheet for straightforward auditing and cross-reference.

The second pass operates on the completed summary table rather than the raw records, scanning the per-ticker metrics to identify three market-wide highlights: the ticker with the greatest percentage gain, the ticker with the greatest percentage loss, and the ticker with the highest cumulative trading volume. These findings are written to a dedicated table, delivering an at-a-glance snapshot of the year's standout extremes without requiring any manual filtering or sorting.

Together, the two passes produce a self-contained, layered output — precise per-stock metrics that support individual analysis alongside a cross-market summary that captures the year's defining performers in a single view.

## **Findings**

Across a dataset of 3,000 tickers spanning three years, statistically meaningful patterns are inherently difficult to surface — the sheer breadth of the universe dilutes any signal that might otherwise stand out. Against that backdrop, the behavior of one ticker, RKS, is striking.

RKS claimed the largest percentage decline in both 2018 and 2019, posting losses of -90.02% and -91.60%, respectively. In 2020, it narrowly missed the bottom position — declining -88.65% while VNG edged it out with a loss of -89.05% — yet still ranked among the year's most severe underperformers. Taken together, RKS sustained catastrophic losses in each of the three years under analysis, with annual declines never rising above -88%, and twice claiming the worst performance in the entire 3,000-ticker universe.

The near-perfect consistency of this underperformance is the dataset's only clearly discernible cross-year pattern, and its persistence demands scrutiny. A loss of this magnitude in a single year might be attributable to an isolated shock — a failed product launch, a regulatory action, or a liquidity event. Losses of this magnitude across three consecutive years suggest something more structural: deteriorating fundamentals, chronic capital erosion, sector-level headwinds, or potentially irregular trading activity that warrants closer examination. Whether the explanation lies in the company's financials, its competitive position, or anomalies in how the stock trades, RKS represents the clearest candidate for deeper investigation that this dataset produces.

## **Conclusion**

This project illustrates what VBA automation makes possible at a scale that manual analysis cannot practically reach. Processing 2.26 million trading records across 3,000 tickers and three calendar years by hand — aggregating opens, closes, and volumes, computing returns, and identifying extremes — would be prohibitively time-consuming and error-prone. The script reduced that workload to a single automated execution per year, producing clean, structured, auditable summaries with no manual intervention beyond the initial data preparation.

The analytical yield was appropriately modest for a dataset of this breadth. With 3,000 tickers, the universe is wide enough that most patterns dissolve into noise, and the year-over-year findings largely reflect that diffuseness. The exception is RKS, whose sustained, near-total losses across all three years — never recovering above -88% in any period — constitute the one signal in this dataset that rises clearly above the baseline. Whether that signal reflects genuine financial distress, structural deterioration, or something more irregular, it is the finding most likely to reward further investigation.

More broadly, the project reinforces a principle that scales beyond this particular dataset: automation does not merely accelerate analysis, it makes certain analyses tractable in the first place. The infrastructure built here — a tested, deployable script capable of processing millions of records reliably — is as significant an output as the findings it produced.

----

## Copyright

Nicholas J. George © 2026. All Rights Reserved.
