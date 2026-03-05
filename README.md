![vba_of_wall_street](https://github.com/njgeorge000158/The-VBA-of-Wall-Street/assets/137228821/45e17b05-4811-47ca-bb51-13a2eef4ffbd)

----

# **VBA-Driven Stock Trading Analysis: 2018–2020**

----

## **Project Overview**

For this project, I developed a VBA script to format, summarize, and analyze approximately 2.26 million stock trading records spanning 3,000 tickers across three years: 2018, 2019, and 2020. The goal was to transform raw, high-volume trading data into a structured and queryable summary, extracting meaningful performance metrics at both the individual stock and cross-market level.

## **Development and Testing**

To ensure accuracy before processing the full dataset, I built and validated the script against a small, alphabetically organized subset of the data. This controlled testing environment allowed for rapid iteration and debugging without the computational overhead of the complete records. Once the script performed reliably on the test data, I deployed it across three separate Excel files — one per year — each containing the full trading history for the period.

## **Script Functionality**

For each year, the script processes the raw stock data and computes three key metrics for every ticker: yearly price change, yearly percent change, and total trading volume. These calculations are written to a structured summary table directly within the same worksheet, keeping the output co-located with its source data for easy reference.

From this per-ticker summary, the script performs a second pass to identify three market-wide highlights — the greatest percentage increase, the greatest percentage decrease, and the highest total volume — and writes these findings to a dedicated summary table. The result is a two-tiered output: granular per-stock metrics alongside an at-a-glance snapshot of market extremes for each year.

## **Findings**

Given the scale of the dataset — 3,000 tickers over three years — clear, repeating patterns are rare, which makes the behavior of one ticker, RKS, particularly noteworthy. RKS recorded the greatest percentage decrease in both 2018 and 2019, posting losses of -90.02% and -91.60%, respectively. In 2020, RKS again ranked among the worst performers with a decline of -88.65%, though it narrowly missed the bottom position, edged out by VNG, which posted a decrease of -89.05%. The near-consistency of RKS's extreme underperformance across all three years is the sole discernible pattern in an otherwise diffuse dataset, and it raises questions worth investigating further — whether through fundamental analysis of the company's financials, examination of sector-wide headwinds, or a review of trading anomalies that might explain such sustained and severe losses.

## **Conclusion**

This project demonstrated the power of VBA automation in handling large-scale financial datasets that would be impractical to process manually. The script reliably condensed 2.26 million records into clean, year-by-year summaries and surfaced performance extremes across the full 3,000-stock universe. While the data yielded limited patterning overall, the persistent underperformance of RKS across the entire three-year window stands as the analysis's most actionable finding — and a compelling candidate for deeper investigation.

----

## Copyright

Nicholas J. George © 2023. All Rights Reserved.
