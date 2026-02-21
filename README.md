# Daybook Ledger Analyser (SGV)

A professional financial tool designed to transform and normalise Tally Daybook exports into structured, analysis-ready Excel reports.

## What this Application Does

This application solves the problem of "messy" accounting exports from Tally. When Tally exports a Daybook, the data is often hierarchical and difficult to analyse in Excel. This tool:

1.  **Normalises Data**: Flattens the hierarchical Tally layout into a clean, row-based format.
2.  **Extracts Cost Centres**: Automatically identifies and pulls out Cost Centre names and amounts from sub-rows.
3.  **Identifies Creditors**: Maps credit entries to their respective ledger heads.
4.  **Cleans Entries**: Automatically skips non-essential lines like GST (SGST/CGST/IGST), Round-off, and TDS entries to focus on core transaction data.
5.  **Professional Export**: Generates a brand-themed Excel report with proper formatting, currency symbols, and colour-coded rows for easy reading.

## Key Features

-   **Intelligent Parsing**: Automatically detects transaction blocks starting from Row 8.
-   **Configurable Ledgers**: Users can define which primary ledgers to track via a dedicated sidebar configuration.
-   **Interactive Dashboard**: Built with Streamlit, providing real-time filtering by Date, Ledger, Voucher Type, and Entry Type.
-   **Corporate Aesthetics**: Designed with a high-end corporate theme (SGV Brand) for professional financial reporting.
-   **Financial Summary**: Instant metrics showing total records, unique transactions, and segregated totals for Cost Centres and Creditors.

