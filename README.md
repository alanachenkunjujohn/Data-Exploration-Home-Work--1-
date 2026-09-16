# Data-Exploration-Home-Work--1-
# Excel Data Analysis Workbook

A comprehensive guide and documentation covering basic to intermediate data analysis tasks within Microsoft Excel, including statistical summaries, conditional logic (`IF`), advanced conditional math (`SUMIF`/`COUNTIF`), and text manipulation functions (`LEFT`/`RIGHT`/`MID`).

## 📊 Dataset Reference & Layout

The formulas outlined in this documentation assume your source worksheet contains headers in row 1, with data spanning rows **2 to 35**.

| Column Letter | Header Name | Data Type | Description |
| :--- | :--- | :--- | :--- |
| **A** | Product ID | Text / Code | Unique string in formatting `DD-MMM-CC` |
| **B** | Product Name | Text | General product descriptor |
| **C** | Brand Name | Text | Manufacturer / Brand label |
| **D** | Price (\$) | Numeric | Unit price per individual item |
| **E** | Quantity | Numeric | Total items currently in stock / sold |
| **F** | Category | Text | Department category classification |
| **G** | Total Price | Numeric | Extended total cost calculation |

---

## 🛠️ Step-by-Step Implementation Guide

### 1. Basic Aggregations (Sum, Count, Average)

Use standard mathematical aggregates to evaluate high-level metrics across the dataset.

*   **Total Price of All Products**
    *   **Formula:** `=SUM(D2:D35)`
    *   *Result:* Summarizes the numeric value of every item listed in the Price column.
*   **Total Product Count**
    *   **Formula:** `=COUNTA(B2:B35)` *(or `=COUNT(D2:D35)`)*
    *   *Result:* Counts the total number of populated alphanumeric records present.
*   **Average Product Price**
    *   **Formula:** `=AVERAGE(D2:D35)`
    *   *Result:* Calculates the central arithmetic mean of your inventory cost.

### 2. Boundary Values (Min & Max)

Locate upper and lower data boundaries quickly.

*   **Minimum Price Among Products**
    *   **Formula:** `=MIN(D2:D35)`
    *   *Result:* Scans the field range to isolate the lowest current product price point (`30`).
*   **Maximum Price Among Products**
    *   **Formula:** `=MAX(D2:D35)`
    *   *Result:* Scans the field range to isolate the highest current product price point (`1000`).

### 3. Logical Evaluation (IF Function)

Categorize records dynamically based on specific internal thresholds.

*   **Price Range Classification**
    *   **Formula:** 
        ```excel
        =IF(D2>=500, "High Price", "Standard Price")
        ```
    *   *Instructions:* Input this formula in your first empty column row (e.g., cell `H2`). Press Enter, then double-click the fill handle in the lower-right corner of the cell to drag the conditional classification logic down to row 35.

### 4. Conditional Aggregations (SUMIF & COUNTIF)

Isolate specific segments of data that meet exact criteria configurations.

*   **Summing Specific Categories (Electronics)**
    *   **Formula:** 
        ```excel
        =SUMIF(F2:F35, "Electronics", D2:D35)
        ```
    *   *Result:* Checks range `F2:F35` for matches to "Electronics", then evaluates and adds the matching offsets in range `D2:D35`.
*   **Counting Under Specific Thresholds (< \$100)**
    *   **Formula:** 
        ```excel
        =COUNTIF(D2:D35, "<100")
        ```
    *   *Result:* Tallies only the product rows containing numerical values strictly below 100.

### 5. String Manipulation & Parsing (LEFT, RIGHT, MID)

Extract sub-string attributes safely out of alphanumeric key structures (e.g., extracting parameters from a standard `Product ID` like `28-JAN-US`).

*   **Extract Day (First 2 Characters)**
    *   **Formula:** 
        ```excel
        =LEFT(A2, 2)
        ```
    *   *Result:* Isolates character strings starting from the extreme left boundary (`28`).
*   **Extract Country Code (Last 2 Characters)**
    *   **Formula:** 
        ```excel
        =RIGHT(A2, 2)
        ```
    *   *Result:* Isolates character strings starting from the extreme right boundary (`US`).
*   **Extract Month (4th to 6th Characters)**
    *   **Formula:** 
        ```excel
        =MID(A2, 4, 3)
        ```
    *   *Result:* Identifies position `4` as the anchor point, counting and copying out exactly `3` characters forward (`JAN`).
