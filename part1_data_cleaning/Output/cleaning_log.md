# Data Cleaning Log
**TASK 9:**   
**Dataset:** raw_orders.xlsx — 932 rows, 21 columns  


---

## 1. Issues Found

### Text Fields
- Leading, trailing, and double internal spaces found across all text columns
- Inconsistent casing found in segment, region, state, city, category, sub_category, ship_mode, payment_status, and order_status (e.g. "CONSUMER", "consumer", "Consumer" all present in the same column)
- Similar category names written differently in some fields

### Date Fields
- Five different date formats found mixed across order_date and ship_date columns:
  - `21 Jul 2024` — text with full month abbreviation
  - `07/27/2024` — MM/DD/YYYY format
  - `28-11-2024` — DD-MM-YYYY format
  - `2024-05-24` — ISO YYYY-MM-DD format
  - `05 Sep 2024` — text with abbreviated month name
- Several dates in MM/DD/YYYY format (e.g. 06/16/2025) were not recognised by Excel as dates due to regional settings and were stored as plain text
- Some ship dates found to be earlier than order dates, which is logically invalid
- Date ambiguity present in certain records where the day and month could not be determined with certainty

### Discount Field
- 18 missing (blank) discount values found
- Negative discount values found (e.g. -0.19, -0.23, -0.09)
- Discount values above the allowed range of 0.5 found (e.g. 0.55, 0.65)
- Discount values stored as percentage strings (e.g. "70%", "85%") instead of decimals

### Duplicates
- 20 exact duplicate rows identified (every column identical)
- 12 order_id groups found with conflicting information across rows

### Missing Values
- region: 26 missing values
- ship_mode: 22 missing values
- discount: 18 missing values

---

## 2. Cleaning Actions Performed

### Text Cleaning
- Applied `=PROPER(TRIM(A2))` to all ten text columns to remove extra spaces and standardize casing
- Used `SUBSTITUTE` in addition to TRIM where double internal spaces were present
- Used Find and Replace after PROPER to fix any specific values that PROPER converted incorrectly (none were found)
- Verified results by checking unique values using column filters after cleaning


### Date Cleaning
- Checked each date cell to determine whether it was stored as a real Excel date (number) or as text
- Used `=IF(ISNUMBER(B2), B2, IFERROR(DATEVALUE(B2),""))` as the base formula for most date formats
- For dates in MM/DD/YYYY format that Excel did not recognise due to regional settings, used the following formula to swap the day and month components:
  `=DATEVALUE(MID(B2,4,2) & "/" & LEFT(B2,2) & "/" & RIGHT(B2,4))`
- Applied this swap only to cells where the middle two digits (day position) were greater than 12, confirming they were in MM/DD/YYYY format
- Formatted all cleaned date columns as DD/MM/YYYY
- Added order_date_flag and ship_date_flag columns to record the status of each date before cleaning
- Added date_validation_flag column to flag ship dates that were earlier than order dates

### Duplicate Handling
- Used Data > Remove Duplicates with all columns selected to remove 20 exact duplicate rows
- Used `=IF(COUNTIF($A$2:$A$932,A2)>1,"Duplicate ID","Unique")` to identify order_ids appearing more than once
- Investigated conflicting duplicate order_ids by filtering and comparing rows side by side
- Did not delete conflicting duplicates, kept all copies and flagged them as "Conflicting Duplicate - Flagged" in the duplicate_flag column

---

## 3. Business Rules Applied

| Rule | Action Taken |
|------|--------------|
| Missing region | Filled as "Unknown" using `=IF(ISBLANK(Q10),"Unknown",Q10)` and flagged as Warning |
| Missing ship_mode | Filled as "Unknown" using `=IF(ISBLANK(AB28),"Unknown",AB28)` and flagged as Warning |
| Missing discount | Treated as 0 in cleaned_discount column where all other sales fields were valid |
| Negative discount | Flagged as "Invalid - Negative" in discount_flag column; excluded from calculated_sales |
| Discount above 0.5 | Flagged as "Invalid - Above Range" in discount_flag column; excluded from calculated_sales |
| Cancelled orders | Excluded from completed sales pivot summaries |
| Failed payments | Excluded from completed sales pivot summaries |
| Refunded orders | Summarized separately in pivot report |
| Ship date before order date | Flagged as "Invalid - Ship Before Order"; shipping_delay_days left blank for these rows |

---

## 4. Calculated Columns Added

| Column | Formula Logic |
|--------|--------------|
| order_date_clean | Parsed and standardized order date in DD/MM/YYYY |
| ship_date_clean | Parsed and standardized ship date in DD/MM/YYYY |
| order_date_flag | Valid / Text - Converted to Date / Invalid - Could Not Parse / Missing |
| ship_date_flag | Same as above applied to ship date |
| date_validation_flag | Valid / Invalid - Ship Before Order / Cannot Check - Date Missing |
| cleaned_discount | 0 if missing; blank if invalid; original value if valid |
| calculated_sales | quantity × unit_price × (1 − cleaned_discount) |
| calculated_profit | calculated_sales − cost |
| profit_margin | calculated_profit ÷ calculated_sales (blank if zero sales) |
| shipping_delay_days | ship_date_clean − order_date_clean in days; blank if dates invalid |
| order_month | Month number extracted from order_date_clean |
| order_year | Year extracted from order_date_clean |
| discount_flag | Valid / Missing / Invalid - Negative / Invalid - Above Range |
| duplicate_flag | Unique / Conflicting Duplicate - Flagged |
| data_quality_flag | Clean / Warning / Invalid |

---

## 5. Records Removed

| Reason | Count |
|--------|-------|
| Exact duplicate rows | 20 |
| **Total removed** | **20** |

No conflicting duplicate records were deleted. All 12 conflicting duplicate order_id groups were retained with a clear flag for business review.

---

## 6. Records Flagged

| Flag | Count |
|------|-------|
| data_quality_flag = Clean | 720 |
| data_quality_flag = Warning | 55 |
| data_quality_flag = Invalid | 137 |
| Conflicting duplicate order_ids | 12 groups flagged |
| Ship date before order date | 12 records |
| Invalid discount values | ~30 records |

---

## 7. Assumptions Made

1. Discount valid range is 0.0 to 0.5 (50%). Any discount above this is considered outside the allowed business range.
2. Missing discount is treated as 0, assuming no discount was applied, provided all other sales fields in that row are populated.
3. For ambiguous date formats, dayfirst=True was assumed as the default (DD/MM/YYYY), consistent with Indian regional settings.
4. Dates in MM/DD/YYYY format were identified by checking whether the middle two digits (day position) exceeded 12, confirming they could not be in DD/MM/YYYY format.
5. Text dates with month names (e.g. "21 Jul 2024") were treated as unambiguous and converted directly.
6. Returned orders are treated similarly to Cancelled orders and excluded from the completed sales summary.
7. Conflicting duplicate records are kept as-is because the correct version cannot be determined without consulting the original source system.

---

## 8. Limitations

- Date ambiguity could not be fully resolved for all records. Some dates in MM/DD/YYYY format where both the day and month are 12 or below (e.g. 06/08/2025) may have been interpreted incorrectly as DD/MM/YYYY. Manual verification is recommended for these records.
- Conflicting duplicate order_ids remain unresolved. A business SME or comparison against the source system is needed to determine which version of each record is correct.
- Unit price values could not be validated as no master price list was available.
- Some calculated_sales values differ from the originally reported sales figures even after discount cleaning. This may indicate upstream rounding differences or data entry errors that cannot be resolved through cleaning alone.
- The cleaning process handles systematic and pattern-based issues. Individual record-level errors such as a wrong customer name or incorrect city require manual spot-checking which was not in scope here.
