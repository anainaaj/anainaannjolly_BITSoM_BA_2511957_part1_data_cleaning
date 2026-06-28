# Part 1: Business Data Cleaning, Validation & Excel Reporting

---

## 1. Problem Summary

A retail company exported order-level sales data from multiple internal systems into a single dataset. Due to the data coming from different sources, the raw file contained a wide range of quality issues including inconsistent text formatting, mixed date formats, duplicate records, missing values, invalid discount entries, and order status inconsistencies.

The objective of this project was to clean and validate the raw dataset, apply business rules to handle problematic records, document all issues found, and produce summary pivot reports that can be used for business review and decision making.

---

## 2. Dataset Description

| Detail | Value |
|--------|-------|
| File name | raw_orders.xlsx |
| Total rows | 932 |
| Total columns | 21 |
| Date range | 2024 to 2025 |

### Columns in the Dataset

| Column | Description |
|--------|-------------|
| order_id | Unique identifier for each order |
| order_date | Date the order was placed |
| ship_date | Date the order was shipped |
| customer_id | Unique identifier for each customer |
| customer_name | Full name of the customer |
| segment | Customer segment (Consumer, Corporate, Small Business, Home Office) |
| region | Geographic region (North, East, South, West) |
| state | State of delivery |
| city | City of delivery |
| category | Product category |
| sub_category | Product sub-category |
| product_name | Name of the product ordered |
| ship_mode | Shipping method used |
| quantity | Number of units ordered |
| unit_price | Price per unit |
| discount | Discount applied to the order |
| sales | Reported sales value |
| cost | Cost of goods |
| profit | Reported profit |
| payment_status | Status of payment (Paid, Refunded, Pending, Failed) |
| order_status | Status of the order (Completed, Cancelled, Returned) |

---

## 3. Tools Used

- **Microsoft Excel** — primary tool for all cleaning, validation, calculated columns, and pivot reports
- Formulas used: TRIM, PROPER, SUBSTITUTE, IF, ISNUMBER, ISBLANK, IFERROR, COUNTIF, COUNTBLANK, DATEVALUE, MID, LEFT, RIGHT, MONTH, YEAR, TEXT, AND, OR, ABS

---

## 4. Cleaning Steps Performed

### Step 1 : Preserve Raw Data
The original file raw_orders.xlsx was kept completely unchanged. All cleaning work was performed in a separate file called cleaned_orders.xlsx to maintain a clean audit trail.

### Step 2 : Clean Text Fields
Applied `=PROPER(TRIM(A2))` across all ten text columns to remove leading, trailing, and internal extra spaces and to standardize casing. Used SUBSTITUTE where double internal spaces were present. Used Find and Replace to fix specific values that PROPER converted incorrectly. Verified results by checking unique values using column filters after each column was cleaned. Converted all formulas to plain values using Paste Special before finalizing.

### Step 3 : Clean and Validate Dates
Identified five different date formats mixed across the order_date and ship_date columns. Used `=IF(ISNUMBER(A2), A2, IFERROR(DATEVALUE(A2),""))` as the base parsing formula. For dates in MM/DD/YYYY format that Excel did not recognise due to regional settings, used a custom formula to swap the day and month components using MID, LEFT, and RIGHT. Formatted all cleaned date columns as DD/MM/YYYY. Added flag columns to record whether each date was valid, converted from text, invalid, or missing. Flagged all ship dates that were earlier than order dates.

### Step 4 : Handle Duplicates
Used Data > Remove Duplicates with all columns selected to identify and remove 20 exact duplicate rows. Used COUNTIF on the order_id column to identify duplicate IDs. Investigated conflicting duplicate records by filtering and comparing rows. Did not delete conflicting duplicates, kept all copies and flagged them clearly in a duplicate_flag column.

### Step 5 : Apply Business Rules
Filled missing region values with "Unknown" and flagged them. Filled missing ship_mode values with "Unknown" and flagged them. Treated missing discount as 0 where other sales fields were valid. Flagged negative and above-range discounts and excluded them from calculations. Excluded cancelled and failed payment orders from sales summaries. Summarized refunded orders separately.

### Step 6 : Create Calculated Columns
Added cleaned_discount, calculated_sales, calculated_profit, profit_margin, shipping_delay_days, order_month, order_year, and data_quality_flag columns. Also added discount_flag, date_flag, date_validation_flag, and duplicate_flag columns for full traceability.

---

## 5. Business Rules Applied

| Rule Area | Action Taken |
|-----------|--------------|
| Missing region | Filled as "Unknown" and flagged as Warning |
| Missing ship_mode | Filled as "Unknown" and flagged as Warning |
| Missing discount | Treated as 0 where all other sales fields are valid |
| Negative discount | Flagged as Invalid; excluded from calculated_sales |
| Discount above 50% | Flagged as Invalid; excluded from calculated_sales |
| Cancelled orders | Excluded from completed sales pivot summaries |
| Failed payments | Excluded from completed sales pivot summaries |
| Refunded orders | Summarized separately in pivot report |
| Ship date before order date | Flagged as Invalid; shipping_delay_days left blank |

---

## 6. Summary of Data Quality Issues Found

| Issue | Count |
|-------|-------|
| Exact duplicate rows removed | 20 |
| Conflicting duplicate order_ids flagged | 12 groups |
| Missing region values | 26 |
| Missing ship_mode values | 22 |
| Missing discount values | 18 |
| Negative discount values | ~15 |
| Discount above 50% | ~15 |
| Ship date before order date | 12 |
| Dates stored as unrecognised text | ~83 |
| Final records flagged as Clean | 720 |
| Final records flagged as Warning | 55 |
| Final records flagged as Invalid | 137 |
| **Final row count after cleaning** | **912** |

---

## 7. Summary of Final Pivot Reports

All pivot reports are saved in **pivot_summary.xlsx**. Sales figures are based on completed orders with successful payments only.

**P1 : Sales and Profit by Region**
Shows total calculated sales and profit broken down by region, sorted by sales from highest to lowest. Helps identify which regions are driving the most revenue.

**P2 : Sales and Profit by Category and Sub-Category**
Nested pivot showing sales and profit grouped first by category and then by sub-category within each category. Sorted by sales within each category group.

**P3 : Order Count by Ship Mode**
Shows how many orders were fulfilled through each shipping method, sorted by order count descending. Useful for understanding logistics distribution.

**P4 : Profit Margin by Customer Segment**
Shows the average profit margin for each customer segment, sorted from highest to lowest margin. Helps identify which segments are most profitable.

**P5 : Refunded, Cancelled, and Failed Orders by Region**
Shows the count of problem orders broken down by region. Useful for identifying regions with higher rates of order issues.

**P6 : Monthly Sales Trend**
Shows total sales and profit grouped by year and month chronologically. Useful for identifying seasonal patterns and growth trends over time.

---

## 8. Key Business Insights

- The majority of records (720 out of 912) passed all quality checks and are ready for analysis.
- 137 records were flagged as Invalid, mostly due to negative discounts, ship dates before order dates, and conflicting duplicate order IDs. These should be investigated with the source systems.
- 12 order_id groups had conflicting information across duplicate rows, suggesting data may have been exported from multiple systems without proper deduplication at source.
- Discounts were the most common source of invalid data, with both negative values and values exceeding 50% found in the dataset. A data entry validation rule should be added at source to prevent this.
- Ship dates appearing before order dates in 12 records suggest either a system clock issue or a data entry error at the time of recording. These records cannot be used for shipping delay analysis.
- Missing region and ship_mode values were filled as Unknown for analysis purposes, but the business should investigate why these fields were not captured at the point of order entry.
- Refunded and cancelled orders are present across all regions, which warrants a deeper investigation into return and cancellation drivers by region and category.

---

## 9. Assumptions and Limitations

### Assumptions
1. The valid discount range is 0.0 to 0.5 (0% to 50%). Any discount above 50% is treated as outside the allowed business range.
2. Missing discount values are treated as 0, assuming no discount was applied, provided all other sales fields in that row are populated.
3. For ambiguous date formats, DD/MM/YYYY was assumed as the default interpretation, consistent with Indian regional settings.
4. Dates in MM/DD/YYYY format were identified by checking whether the middle two digits exceeded 12, confirming they could not be in DD/MM/YYYY format.
5. Returned orders are treated the same as Cancelled orders and excluded from the completed sales summary.
6. Conflicting duplicate records are retained because the correct version cannot be determined without consulting the original source system.

### Limitations
1. Date ambiguity could not be fully resolved for all records. Dates where both the day and month are 12 or below (e.g. 06/08/2025) may have been interpreted in the wrong order. Manual verification is recommended for these records.
2. Conflicting duplicate order_ids remain unresolved. A business SME or comparison against the source system is needed to confirm which version of each record is authoritative.
3. Unit price values could not be validated as no master product price list was available for cross-checking.
4. Some calculated_sales values differ slightly from the originally reported sales figures even after discount cleaning. This may indicate upstream rounding differences or data entry errors that cannot be resolved through cleaning alone.
5. The cleaning process addresses systematic and pattern-based issues. Individual record-level errors such as a wrong customer name or incorrect city require manual spot-checking which was outside the scope of this project.
