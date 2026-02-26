# S&P Capital IQ Excel Plug-in Formula Reference

Legacy CIQ formula reference (v8.x). All mnemonics work with `=CIQ("TICKER","MNEMONIC")` syntax.

---

## Ticker Format

CIQ uses `EXCHANGE:TICKER` format. Plain tickers (no exchange prefix) also work for unambiguous US names.

### Exchange Prefixes

| Prefix | Exchange | Example |
|--------|----------|---------|
| `NYSE` | New York Stock Exchange | `NYSE:WMS` |
| `NASDAQ` | Nasdaq (generic) | `NASDAQ:CSIQ` |
| `NASDAQGS` | Nasdaq Global Select | `NASDAQGS:NWPX` |
| `NasdaqGM` | Nasdaq Global Market | `NasdaqGM:WLDN` |
| `TSX` | Toronto Stock Exchange | `TSX:CCL.B` |
| `TSXV` | TSX Venture Exchange | `TSXV:AEP` |
| `ASX` | Australian Securities Exchange | (use plain ticker if available) |

### Ticker Suffixes (Canadian)

| Suffix | Meaning | Example |
|--------|---------|---------|
| `.A` | Class A shares | `TSX:TCL.A`, `TSX:HPS.A`, `TSX:HMM.A` |
| `.B` | Class B shares | `TSX:CCL.B`, `TSX:BBD.B` |
| `.UN` | Trust units | `TSX:RPI.UN` |

### Notes
- US tickers generally work without exchange prefix (e.g., `DSGX`, `ROP`, `MANH`)
- Canadian tickers require `TSX:` or `TSXV:` prefix (e.g., `TSX:CCL.B`)
- Australian tickers: use plain ticker without exchange prefix (e.g., `WTC` for WiseTech)
- US OTC tickers (e.g., `CMDXF`, `CNSWF`) work for international companies cross-listed on pink sheets
- Some OTC tickers return `NA` for close price but still provide fundamentals (market cap, revenue, etc.)

---

## CIQ Functions

| Function | Syntax | Use Case |
|----------|--------|----------|
| `CIQ` | `=CIQ(T, Metric)` | Single value lookup |
| `CIQ` | `=CIQ(T, Metric, "IQ_FQ-N")` | Single value for specific period |
| `CIQRANGE` | `=CIQRANGE(T, Metric, Period/D1, D2, ...)` | Range of formulas (each cell gets its own formula) |
| `CIQRANGEV` | `=CIQRANGEV(T, Metric, Period/D1, D2, ...)` | Range of values (data fills as values, more efficient) |
| `CIQRANGEA` | `=CIQRANGEA(T, Metric, ...)` | Range across (horizontal expansion) |
| `CIQAVG` | `=CIQAVG(T, Metric, D1, D2)` | Average over date range |
| `CIQHI` | `=CIQHI(T, Metric, D1, D2)` | High value over date range |
| `CIQLO` | `=CIQLO(T, Metric, D1, D2)` | Low value over date range |
| `CIQPC` | `=CIQPC(T, Metric, D1, D2)` | Percent change over period |
| `CIQGETDATE` | `=CIQGETDATE(IQ_TODAY)` | Today's date |

### CIQRANGE vs CIQRANGEV vs CIQRANGEA

- **CIQRANGE** — range of formulas: each expanded cell contains its own CIQ formula
- **CIQRANGEV** — range of values: data fills as plain values (more efficient, preferred for large pulls)
- **CIQRANGEA** — range across: expands horizontally instead of vertically
- All three accept the same parameter patterns (substitute `CIQRANGEV` or `CIQRANGEA` for `CIQRANGE`)
- **CIQ compat mode (Pro):** `CIQRANGE` and `CIQRANGEV` work; `CIQRANGEA` does NOT (returns function name as string)

### CIQRANGE Syntax by Data Type

**CRITICAL: The parameter pattern varies by data type. Using the wrong pattern returns `(Invalid Time Period)` or `(Invalid Period Type)`.**

| Data Type | Syntax | Example |
|-----------|--------|---------|
| Pricing / Market | `=CIQRANGE(T, Metric, D1, D2, , , , , Label)` | `=CIQRANGE("DSGX","IQ_CLOSEPRICE","2/26/2025","2/26/2026",,,,,,"Price")` |
| Financials | `=CIQRANGE(T, Metric, IQ_FQ-N, , , , , , Label)` | `=CIQRANGE("DSGX","IQ_EBITDA",IQ_FQ-16,,,,,,,"EBITDA")` |
| **Trading Multiples** | `=CIQRANGEV(T, Metric, Period, D1, D2, , , , Label)` | `=CIQRANGEV("DSGX","IQ_TEV_EBITDA",IQ_LTM,"-1Y","2/26/2026",,,,"TEV/EBITDA")` |
| Dividends | `=CIQRANGE(T, Metric, D1, D2)` | |
| News | `=CIQRANGE(T, IQ_NEWS, Start Rank, End Rank)` | |
| Key Dev | `=CIQRANGE(T, IQ_KEY_DEV_ID, D1, D2, Category)` | |
| Quick Comps | `=CIQRANGE(T, IQ_QUICK_COMP, Start Rank, End Rank)` | |

*T = Ticker, D1 = Start Date, D2 = End Date, N = number of periods*

### Relative Period Types (for Multiples and Financials)

| Period | Meaning | Example Use |
|--------|---------|-------------|
| `IQ_LTM` | Last Twelve Months | Trading multiples: `CIQRANGEV(T, "IQ_TEV_EBITDA", IQ_LTM, ...)` |
| `IQ_FY` | Fiscal Year | Annual financial data |
| `IQ_CY` | Calendar Year | Calendar-year aligned data |
| `IQ_FQ` | Fiscal Quarter | Quarterly financial data: `CIQRANGE(T, "IQ_EBITDA", IQ_FQ-12, ...)` |
| `IQ_CQ` | Calendar Quarter | Calendar-quarter aligned data |
| `IQ_FH` | Fiscal Half | Semi-annual financial data |
| `IQ_CH` | Calendar Half | Calendar half-year data |

### Date Parameters

- **Absolute dates:** `"2/26/2025"`, `"12/31/2024"` (M/D/YYYY format)
- **Relative dates:** `"-1Y"` (1 year back), `"-6M"` (6 months), `"-90D"` (90 days)
- **Relative period offsets:** `IQ_FQ-12` (12 fiscal quarters back), `IQ_FY-5` (5 fiscal years back)

### Workbook Layout for CIQRANGE Expansion

- **Formula goes in Row 1** of its column (not Row 2)
- The `Label` parameter becomes the column header
- Data expands downward starting from Row 2
- For financial data, add a separate date column: `=CIQRANGE(T,"IQ_PERIODDATE_BS",IQ_FQ-N,,,,,,,"Date")`
- Market data returns ~252 daily data points per year; multiples return ~198 daily data points per year

---

## Income Statement

| Metric | Mnemonic |
|--------|----------|
| Total Revenues | `IQ_TOTAL_REV` |
| Cost Of Revenues | `IQ_COST_REV` |
| Cost Of Goods Sold | `IQ_COGS` |
| Gross Profit | `IQ_GP` |
| SG&A Expense | `IQ_SGA_SUPPL` |
| R&D Expense | `IQ_RD_EXP` |
| Depreciation & Amortization | `IQ_DA_SUPPL` |
| Amort. of Goodwill & Intangibles | `IQ_GW_INTAN_AMORT` |
| Operating Income | `IQ_OPER_INC` |
| Net Interest Expense | `IQ_NET_INTEREST_EXP` |
| EBT Excl. Unusual Items | `IQ_EBT_EXCL` |
| Total Unusual Items | `IQ_TOTAL_UNUSUAL` |
| EBT Incl. Unusual Items | `IQ_EBT` |
| Income Tax Expense | `IQ_INC_TAX` |
| Earnings from Continuing Ops | `IQ_EARNING_CO` |
| Earnings of Discontinued Ops | `IQ_DO` |
| Extraordinary Items & Acct Changes | `IQ_EXTRA_ACC_ITEMS` |
| Net Income | `IQ_NI` |
| Pref. Dividends & Other Adj. | `IQ_PREF_DIV_OTHER` |
| Merger/Restructuring Costs | `IQ_MERGER_RESTRUCTURE` |
| NI to Common Incl. Extra Items | `IQ_NI_AVAIL_INCL` |
| NI to Common Excl. Extra Items | `IQ_NI_AVAIL_EXCL` |
| Basic EPS (Incl. Extra) | `IQ_BASIC_EPS_INCL` |
| Basic EPS (Excl. Extra) | `IQ_BASIC_EPS_EXCL` |
| Diluted EPS (Incl. Extra) | `IQ_DILUT_EPS_INCL` |
| Diluted EPS (Excl. Extra) | `IQ_DILUT_EPS_EXCL` |
| Weighted Avg. Basic Shares | `IQ_BASIC_WEIGHT` |
| Weighted Avg. Diluted Shares | `IQ_DILUT_WEIGHT` |
| Normalized Basic EPS | `IQ_EPS_NORM` |
| Normalized Diluted EPS | `IQ_DILUT_EPS_NORM` |

### Income Statement - Supplemental

| Metric | Mnemonic |
|--------|----------|
| EBITDA | `IQ_EBITDA` |
| EBITDA (Incl. Equity Inc.) | `IQ_EBITDA_EQ_INC` |
| EBITA | `IQ_EBITA` |
| EBIT | `IQ_EBIT` |
| EBITDAR | `IQ_EBITDAR` |
| Net Rental Expense | `IQ_NET_RENTAL_EXP_FN` |
| Normalized Net Income | `IQ_NI_NORM` |
| Same Store Sales Growth % | `IQ_SAME_STORE` |
| Effective Tax Rate | `IQ_EFFECT_TAX_RATE` |
| Payout Ratio | `IQ_PAYOUT_RATIO` |
| Interest on Long Term Debt | `IQ_INT_EXP_LTD` |
| Total Current Taxes | `IQ_CURR_TAXES` |
| Total Deferred Taxes | `IQ_DEFERRED_TAXES_TOTAL` |

---

## Balance Sheet

| Metric | Mnemonic |
|--------|----------|
| Cash & Equivalents | `IQ_CASH_EQUIV` |
| Short Term Investments | `IQ_ST_INVEST` |
| Total Cash & ST Investments | `IQ_CASH_ST_INVEST` |
| Accounts Receivable | `IQ_AR` |
| Total Receivables | `IQ_TOTAL_RECEIV` |
| Inventory | `IQ_INVENTORY` |
| Total Current Assets | `IQ_TOTAL_CA` |
| Gross PP&E | `IQ_GPPE` |
| Net PP&E | `IQ_NPPE` |
| Long-term Investments | `IQ_LT_INVEST` |
| Total Intangibles (Goodwill + Intangibles) | `IQ_GW_INTAN` |
| Total Assets | `IQ_TOTAL_ASSETS` |
| Accounts Payable | `IQ_AP` |
| Short-term Borrowings | `IQ_ST_DEBT` |
| Current Portion of LT Debt | `IQ_CURRENT_PORT_DEBT` |
| Current Portion of Capital Leases | `IQ_CURRENT_PORT_LEASES` |
| Total Current Liabilities | `IQ_TOTAL_CL` |
| Long-Term Debt | `IQ_LT_DEBT` |
| Capital Leases | `IQ_CAPITAL_LEASES` |
| Minority Interest | `IQ_MINORITY_INTEREST` |
| Total Liabilities | `IQ_TOTAL_LIAB` |
| Total Preferred Equity | `IQ_PREF_EQUITY` |
| Common Stock | `IQ_COMMON` |
| Retained Earnings | `IQ_RE` |
| Treasury Stock | `IQ_TREASURY` |
| Total Common Equity | `IQ_TOTAL_COMMON_EQUITY` |
| Total Equity | `IQ_TOTAL_EQUITY` |
| Total Liabilities & Equity | `IQ_TOTAL_LIAB_EQUITY` |

### Balance Sheet - Supplemental

| Metric | Mnemonic |
|--------|----------|
| Book Value/Share | `IQ_BV_SHARE` |
| Tangible Book Value | `IQ_TBV` |
| Tangible Book Value/Share | `IQ_TBV_SHARE` |
| Total Debt | `IQ_TOTAL_DEBT` |
| Total Current Debt | `IQ_TOTAL_DEBT_CURRENT` |
| Total Non-Current Debt | `IQ_TOTAL_DEBT_NON_CURRENT` |
| Net Debt | `IQ_NET_DEBT` |
| Total Capitalization | `IQ_TOTAL_CAP` |
| Full Time Employees | `IQ_FULL_TIME` |
| Filing Date | `IQ_FILINGDATE_BS` |
| Period Date | `IQ_PERIODDATE_BS` |
| Shares Outstanding (Filing Date) | `IQ_OUTSTANDING_FILING_DATE` |
| Shares Outstanding (BS Date) | `IQ_OUTSTANDING_BS_DATE` |
| Total Shares Out (Filing Date) | `IQ_TOTAL_OUTSTANDING_FILING_DATE` |
| Total Shares Out (BS Date) | `IQ_TOTAL_OUTSTANDING_BS_DATE` |

---

## Cash Flow Statement

| Metric | Mnemonic |
|--------|----------|
| Net Income (CF) | `IQ_NI_CF` |
| Depreciation & Amort. (CF) | `IQ_DA_CF` |
| Asset Writedown & Restructuring | `IQ_ASSET_WRITEDOWN_CF` |
| Stock-Based Compensation | `IQ_STOCK_BASED_CF` |
| Net Cash from Discontinued Ops | `IQ_DO_CF` |
| Change in Accounts Receivable | `IQ_CHANGE_AR` |
| Change in Inventories | `IQ_CHANGE_INVENTORY` |
| Change in Accounts Payable | `IQ_CHANGE_AP` |
| Change in Unearned Revenue | `IQ_CHANGE_UNEARN_REV` |
| Change in Income Taxes | `IQ_CHANGE_INC_TAX` |
| Change in Deferred Taxes | `IQ_CHANGE_DEF_TAX` |
| Cash from Operations | `IQ_CASH_OPER` |
| Capital Expenditure | `IQ_CAPEX` |
| Sale of PP&E | `IQ_SALE_PPE_CF` |
| Cash Acquisitions | `IQ_CASH_ACQUIRE_CF` |
| Divestitures | `IQ_DIVEST_CF` |
| Sale (Purchase) of Intangibles | `IQ_SALE_INTAN_CF` |
| Net Cash from Investments | `IQ_INVEST_SECURITY_CF` |
| Cash from Investing | `IQ_CASH_INVEST` |
| Short Term Debt Issued | `IQ_ST_DEBT_ISSUED` |
| Long-Term Debt Issued | `IQ_LT_DEBT_ISSUED` |
| Total Debt Issued | `IQ_TOTAL_DEBT_ISSUED` |
| Short Term Debt Repaid | `IQ_ST_DEBT_REPAID` |
| Long-Term Debt Repaid | `IQ_LT_DEBT_REPAID` |
| Total Debt Repaid | `IQ_TOTAL_DEBT_REPAID` |
| Issuance of Common Stock | `IQ_COMMON_ISSUED` |
| Repurchase of Common Stock | `IQ_COMMON_REP` |
| Issuance of Preferred Stock | `IQ_PREF_ISSUED` |
| Repurchase of Preferred Stock | `IQ_PREF_REP` |
| Common Dividends Paid | `IQ_COMMON_DIV_CF` |
| Preferred Dividends Paid | `IQ_PREF_DIV_CF` |
| Common and/or Pref. Dividends Paid | `IQ_COMMON_PREF_DIV_CF` |
| Total Dividends Paid | `IQ_TOTAL_DIV_PAID_CF` |
| Cash from Financing | `IQ_CASH_FINAN` |
| Net Change in Cash | `IQ_NET_CHANGE` |

### Cash Flow - Supplemental

| Metric | Mnemonic |
|--------|----------|
| Cash Interest Paid | `IQ_CASH_INTEREST` |
| Cash Taxes Paid | `IQ_CASH_TAXES` |
| Levered Free Cash Flow | `IQ_LEVERED_FCF` |
| Unlevered Free Cash Flow | `IQ_UNLEVERED_FCF` |
| Change in Net Working Capital | `IQ_CHANGE_NET_WORKING_CAPITAL` |
| EBITDA - Capex | `IQ_EBITDA_CAPEX` |
| Net Debt Issued | `IQ_NET_DEBT_ISSUED` |

---

## Market Data

| Metric | Mnemonic |
|--------|----------|
| Last Sale Price | `IQ_LASTSALEPRICE` |
| Close Price | `IQ_CLOSEPRICE` |
| Dividend Adjusted Close Price | `IQ_CLOSEPRICE_ADJ` |
| Pricing Date | `IQ_PRICEDATE` |
| VWAP | `IQ_VWAP` |
| 52 Week High Price | `IQ_YEARHIGH` |
| 52 Week High Date | `IQ_YEARHIGH_DATE` |
| Daily Volume | `IQ_VOLUME` |
| Daily Value Traded | `IQ_VALUE_TRADED` |
| Market Capitalization | `IQ_MARKETCAP` |
| Enterprise Value (TEV) | `IQ_TEV` |
| Shares Outstanding | `IQ_SHARESOUTSTANDING` |
| 5 Year Beta | `IQ_BETA_5YR` |
| 5 Year Price Volatility | `IQ_PRICE_VOL_HIST_5YR` |
| Fund NAV | `IQ_FUND_NAV` |
| Dividend Yield | `IQ_DIVIDEND_YIELD` |
| Latest Annualized Dividend/Share | `IQ_ANNUALIZED_DIVIDEND` |
| Stock Exchange | `IQ_EXCHANGE` |
| Today's Date | `CIQGETDATE(IQ_TODAY)` |

---

## Valuation Multiples

### Trailing (LTM)

| Metric | Mnemonic |
|--------|----------|
| TEV/Total Revenues | `IQ_TEV_TOTAL_REV` |
| TEV/EBITDA | `IQ_TEV_EBITDA` |
| TEV/EBIT | `IQ_TEV_EBIT` |
| TEV/Unlevered FCF | `IQ_TEV_UFCF` |
| Market Cap/Levered FCF | `IQ_MARKET_CAP_LFCF` |
| P/Diluted EPS (Excl. Extra) | `IQ_PE_EXCL` |
| P/BV | `IQ_PBV` |
| P/Tangible BV | `IQ_PTBV` |
| P/Sales | `IQ_PRICE_SALES` |

### Forward (NTM)

| Metric | Mnemonic |
|--------|----------|
| TEV/Forward Total Revenue | `IQ_TEV_TOTAL_REV_FWD` |
| TEV/Forward EBITDA | `IQ_TEV_EBITDA_FWD` |
| TEV/Forward EBIT | `IQ_TEV_EBIT_FWD` |
| P/Forward Diluted EPS (Excl. Extra) | `IQ_PE_EXCL_FWD` |
| PEG Ratio (Forward) | `IQ_PEG_FWD` |
| P/Forward CFPS | `IQ_PRICE_CFPS_FWD` |

---

## Consensus Estimates

| Metric | Mnemonic |
|--------|----------|
| Avg Broker Recommendation (Text) | `IQ_AVG_BROKER_REC` |
| Avg Broker Recommendation (#) | `IQ_AVG_BROKER_REC_NO` |
| Price Target | `IQ_PRICE_TARGET` |
| Est. Annual EPS Growth - 5 Yr | `IQ_EST_EPS_GROWTH_5YR` |
| Primary EPS Estimate | `IQ_EPS_EST` |
| Revenue Estimate | `IQ_REVENUE_EST` |
| EBITDA Estimate | `IQ_EBITDA_EST` |
| EBIT Estimate | `IQ_EBIT_EST` |
| Net Income Estimate | `IQ_NI_EST` |
| Cash Flow Per Share Estimate | `IQ_CFPS_EST` |
| # Analyst Buy Recommendations | `IQ_EST_NUM_BUY` |
| # Analyst Hold Recommendations | `IQ_EST_NUM_HOLD` |
| # Analyst Sell Recommendations | `IQ_EST_NUM_SELL` |
| Estimated Next Earnings Date | `IQ_EST_NEXT_EARNINGS_DATE` |
| Next Earnings Date | `IQ_NEXT_EARNINGS_DATE` |
| EPS Difference (vs Estimate) | `IQ_EST_EPS_DIFF` |
| EPS Surprise % | `IQ_EST_EPS_SURPRISE_PERCENT` |
| Revenue Difference (vs Estimate) | `IQ_EST_REV_DIFF` |
| Revenue Surprise % | `IQ_EST_REV_SURPRISE_PERCENT` |

---

## Ratios

### Profitability

| Metric | Mnemonic |
|--------|----------|
| Return on Assets % | `IQ_RETURN_ASSETS` |
| Return on Equity % | `IQ_RETURN_EQUITY` |
| Gross Margin % | `IQ_GROSS_MARGIN` |
| SG&A Margin % | `IQ_SGA_MARGIN` |
| EBITDA Margin % | `IQ_EBITDA_MARGIN` |
| Net Income Margin % | `IQ_NI_MARGIN` |
| Levered FCF Margin % | `IQ_LFCF_MARGIN` |

### Efficiency

| Metric | Mnemonic |
|--------|----------|
| Accounts Receivable Turnover | `IQ_AR_TURNS` |
| Inventory Turnover | `IQ_INVENTORY_TURNS` |
| Avg Days Sales Outstanding | `IQ_DAYS_SALES_OUT` |
| Avg Days Payable Outstanding | `IQ_DAYS_PAYABLE_OUT` |

### Liquidity

| Metric | Mnemonic |
|--------|----------|
| Current Ratio | `IQ_CURRENT_RATIO` |
| Quick Ratio | `IQ_QUICK_RATIO` |

### Leverage

| Metric | Mnemonic |
|--------|----------|
| Total Debt/Equity | `IQ_TOTAL_DEBT_EQUITY` |
| Total Debt/Capital | `IQ_TOTAL_DEBT_CAPITAL` |
| EBIT/Interest Expense | `IQ_EBIT_INT` |
| Capex as % of Revenues | `IQ_CAPEX_PCT_REV` |
| Total Debt/EBITDA | `IQ_TOTAL_DEBT_EBITDA` |
| Net Debt/EBITDA | `IQ_NET_DEBT_EBITDA` |

---

## Growth Metrics (1 Year)

| Metric | Mnemonic |
|--------|----------|
| Total Revenue Growth % | `IQ_TOTAL_REV_1YR_ANN_GROWTH` |
| Gross Profit Growth % | `IQ_GP_1YR_ANN_GROWTH` |
| EBITDA Growth % | `IQ_EBITDA_1YR_ANN_GROWTH` |
| EBIT Growth % | `IQ_EBIT_1YR_ANN_GROWTH` |
| Net Income Growth % | `IQ_NI_1YR_ANN_GROWTH` |
| Normalized NI Growth % | `IQ_NI_NORM_1YR_ANN_GROWTH` |
| Diluted EPS Growth % | `IQ_EPS_1YR_ANN_GROWTH` |
| Common Equity Growth % | `IQ_COMMON_EQUITY_1YR_ANN_GROWTH` |
| Inventory Growth % | `IQ_INV_1YR_ANN_GROWTH` |
| Total Assets Growth % | `IQ_TOTAL_ASSETS_1YR_ANN_GROWTH` |
| Tangible Book Value Growth % | `IQ_TBV_1YR_ANN_GROWTH` |
| Cash from Operations Growth % | `IQ_CFO_1YR_ANN_GROWTH` |
| Capex Growth % | `IQ_CAPEX_1YR_ANN_GROWTH` |
| Levered FCF Growth % | `IQ_LFCF_1YR_ANN_GROWTH` |
| Unlevered FCF Growth % | `IQ_UFCF_1YR_ANN_GROWTH` |
| Dividend/Share Growth % | `IQ_DPS_1YR_ANN_GROWTH` |

---

## Credit Ratings

| Metric | Mnemonic |
|--------|----------|
| S&P Long-Term Company Rating | `IQ_SP_LC_LT` |
| S&P LT Company Rating Date | `IQ_SP_LC_DATE_LT` |
| S&P Outlook/Credit Watch | `IQ_SP_OUTLOOK_WATCH` |
| S&P Long-Term Security Rating | `IQ_SP_ISSUE_LT` |
| S&P Security Rating Date | `IQ_SP_ISSUE_DATE` |
| S&P Security Rating Action | `IQ_SP_ISSUE_ACTION` |

---

## Capital Structure

| Metric | Mnemonic |
|--------|----------|
| Total Commercial Paper | `IQ_CP` |
| Total Revolving Credit | `IQ_RC` |
| Total Term Loans | `IQ_TERM_LOANS` |
| Total Sr. Bonds & Notes | `IQ_SR_BONDS_NOTES` |
| Total Sub. Bonds & Notes | `IQ_SUB_BONDS_NOTES` |
| Capital Leases (Incl. Current) | `IQ_CAPITAL_LEASES_TOTAL` |
| Other Borrowings | `IQ_OTHER_DEBT` |
| Total Principal Due | `IQ_TOTAL_PRINCIPAL` |
| Total Senior Debt | `IQ_SR_DEBT` |
| Total Senior Debt (% of Total) | `IQ_SR_DEBT_PCT` |
| Total Subordinated Debt | `IQ_TOTAL_SUB_DEBT` |
| Total Convertible Debt | `IQ_CONVERT` |
| Total Short-Term Debt | `IQ_ST_DEBT` |
| Total Short-Term Debt (% of Total) | `IQ_ST_DEBT_PCT` |
| LT Debt (Incl. Capital Leases) | `IQ_LT_DEBT_CAPITAL_LEASES` |
| Senior Debt/EBITDA | `IQ_SR_DEBT_EBITDA` |
| Senior Debt/(EBITDA-Capex) | `IQ_SR_DEBT_EBITDA_CAPEX` |
| Total Sub. Debt/EBITDA | `IQ_TOTAL_SUB_DEBT_EBITDA` |

---

## Fixed Income

| Metric | Mnemonic |
|--------|----------|
| Issue Name | `IQ_ISSUE_NAME` |
| Issuer | `IQ_ISSUER` |
| Issuer Parent Ticker | `IQ_ISSUER_PARENT_TICKER` |
| Issuer Ticker | `IQ_ISSUER_TICKER` |
| Amount Outstanding | `IQ_AMT_OUT` |
| Maturity Date | `IQ_MATURITY_DATE` |
| Offering Coupon | `IQ_OFFER_COUPON` |
| Offering Date | `IQ_OFFER_DATE` |
| Offering Price | `IQ_OFFER_PRICE` |
| Security Type | `IQ_SECURITY_TYPE` |
| Security Level | `IQ_SECURITY_LEVEL` |
| Yield to Worst | `IQ_YTW` |
| YTW Date | `IQ_YTW_DATE` |
| Spread to Worst | `IQ_STW` |
| Offering Yield | `IQ_OFFER_YIELD` |
| Principal Amount | `IQ_PRINCIPAL_AMT` |
| Current Bond Price | `IQ_BOND_PRICE` |
| Callable Feature | `IQ_CALLABLE` |

---

## Company Information

| Metric | Mnemonic |
|--------|----------|
| Company Name | `IQ_COMPANY_NAME` |
| Chinese Short Native Name | `IQ_CHINESE_SHORT_NATIVE_TICKER` |
| Business Description | `IQ_BUSINESS_DESCRIPTION` |
| Headquarters Address | `IQ_COMPANY_ADDRESS` |
| Primary Industry | `IQ_PRIMARY_INDUSTRY` |
| Number of Shareholders | `IQ_NUMBER_SHAREHOLDERS` |

---

## Ownership / Holders

| Metric | Mnemonic |
|--------|----------|
| Institutional Holder | `IQ_INSTITUTIONAL_OWNER` |
| Institutional Holder Total Shares | `IQ_INSTITUTIONAL_SHARES` |
| Insider Holder | `IQ_INSIDER_OWNER` |
| Insider Holder Total Shares | `IQ_INSIDER_SHARES` |
| Holder Name | `IQ_HOLDER_NAME` |
| Holder CIQ ID | `IQ_HOLDER_CIQID` |
| Holder Total Shares | `IQ_HOLDER_SHARES` |
| Mutual Fund Name | `IQ_HOLDER_FUND_NAME` |
| Mutual Fund Shares Held | `IQ_HOLDER_FUND_SHARES` |
| Mutual Fund % of Shares Outstanding | `IQ_HOLDER_FUND_PERCENT` |

---

## Private Equity

| Metric | Mnemonic |
|--------|----------|
| All Direct Investments | `IQ_INVESTMETNS_ALL` |
| All Direct Investments CIQ ID | `IQ_INVESTMETNS_ALL_ID` |
| Fund Families | `IQ_PE_FUND_FAMILIES` |
| Recent Funds | `IQ_RECENT_FUNDS` |
| Fund Name | `IQ_PE_FUND_NAME` |
| Fund Size (USD) | `IQ_PE_FUND_SIZE` |
| Limited Partners | `IQ_LIMITED_PARTNERS` |
| Industries of Interest | `IQ_INVEST_CRITERIA_INDUSTRY` |
| Equity Investment Min (mm-USD) | `IQ_INVEST_CRITERIA_EQUITY_MIN` |

---

## Business Segments

| Metric | Mnemonic |
|--------|----------|
| Business Segment Name | `IQ_BUS_SEG_NAME` |
| Business Segment Description | `IQ_BUS_SEG_DESCRIPTION` |
| Business Segment Revenue | `IQ_BUS_SEG_REV` |
| Business Segment Assets | `IQ_BUS_SEG_ASSETS` |
| Geographic Segment Name | `IQ_GEO_SEG_NAME` |
| Geographic Segment Revenue | `IQ_GEO_SEG_REV` |
| Geographic Segment Assets | `IQ_GEO_SEG_ASSETS` |

---

## Transactions & Identifiers

| Metric | Mnemonic |
|--------|----------|
| Index Constituents | `IQ_CONSTITUENTS` |
| All Transactions | `IQ_TRANSACTION_LIST` |
| M&A Transactions | `IQ_TRANSACTION_LIST_MA` |
| Equity Identifiers | `IQ_EQUITY_LIST` |
| Fixed Income Identifiers | `IQ_FIXED_INCOME_LIST` |
| Bond Identifiers | `IQ_BOND_LIST` |
| Preferred Security Identifiers | `IQ_PREFERRED_LIST` |
| Bank Loan Identifiers | `IQ_BANK_LOAN_LIST` |
| CDS Identifiers | `IQ_CDS_LIST` |

---

## Macroeconomic

| Metric | Mnemonic |
|--------|----------|
| Unemployment Rate | `IQ_UNEMPLOYMENT_RATE` |
| GDP | `IQ_GDP` |
| Real GDP | `IQ_GDP_REAL` |
| Inflation Rate | `IQ_INFLATION_RATE` |
| Budget Spending | `IQ_BUDGET_SPENDING` |
| CPI | `IQ_CPI` |
| Consumer Lending | `IQ_CONSUMER_LENDING` |

---

## Quick Search Index

Common comp table metrics at a glance:

| What You Want | Mnemonic | Notes |
|---------------|----------|-------|
| Stock price | `IQ_CLOSEPRICE` | |
| Market cap | `IQ_MARKETCAP` | |
| Enterprise value | `IQ_TEV` | |
| Revenue (LTM) | `IQ_TOTAL_REV` | |
| EBITDA (LTM) | `IQ_EBITDA` | |
| EBIT (LTM) | `IQ_EBIT` | |
| Net income | `IQ_NI` | |
| EPS (diluted) | `IQ_DILUT_EPS_EXCL` | Excl. extra items |
| Revenue estimate (NTM) | `IQ_REVENUE_EST` | Consensus |
| EBITDA estimate (NTM) | `IQ_EBITDA_EST` | Consensus |
| EPS estimate (NTM) | `IQ_EPS_EST` | Consensus |
| EV/Revenue (LTM) | `IQ_TEV_TOTAL_REV` | Pre-calculated |
| EV/EBITDA (LTM) | `IQ_TEV_EBITDA` | Pre-calculated |
| EV/Revenue (NTM) | `IQ_TEV_TOTAL_REV_FWD` | Pre-calculated |
| EV/EBITDA (NTM) | `IQ_TEV_EBITDA_FWD` | Pre-calculated |
| P/E (LTM) | `IQ_PE_EXCL` | |
| P/E (NTM) | `IQ_PE_EXCL_FWD` | |
| Revenue growth | `IQ_TOTAL_REV_1YR_ANN_GROWTH` | 1-year |
| EBITDA margin | `IQ_EBITDA_MARGIN` | Pre-calculated % |
| Net debt | `IQ_NET_DEBT` | |
| Total debt | `IQ_TOTAL_DEBT` | |
| Shares outstanding | `IQ_SHARESOUTSTANDING` | |
| Dividend yield | `IQ_DIVIDEND_YIELD` | |
| FCF (levered) | `IQ_LEVERED_FCF` | |
| FCF (unlevered) | `IQ_UNLEVERED_FCF` | |
| Capex | `IQ_CAPEX` | |
| Net debt/EBITDA | `IQ_NET_DEBT_EBITDA` | |
| Company name | `IQ_COMPANY_NAME` | |
| Industry | `IQ_PRIMARY_INDUSTRY` | |
