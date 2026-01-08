# 🇲🇾 Malaysia Payroll & Bank Transfer System (Shiny App)

## Overview

This Shiny application is a complete **Malaysia payroll processing system** designed to:

- Calculate monthly payroll for employees
- Apply statutory deductions (EPF, SOCSO, EIS, PCB income tax)
- Handle overtime, Sunday work, allowances, commissions, and attendance rules
- Generate **individual Excel payslips** from a predefined template
- Produce a **bank transfer summary** for salary disbursement
- Export payroll and bank files in CSV or Excel format

The application is built for **small to medium organizations** that require transparent, auditable payroll calculations while retaining flexibility in allowances and employer contribution rates.

---

## Key Features

### Payroll Computation
- Basic salary prorated by working days
- Automatic unpaid leave deduction for absences
- Conditional attendance allowance
- Configurable overtime (OT) and Sunday pay multipliers
- Allowance and commission handling
- Individual employer EPF rate per employee
- Accurate SOCSO & EIS contribution bands (2025 tables)
- Monthly PCB (income tax) estimation

### Output Generation
- Final payroll table
- Individual payslip files (`.xlsx`) generated from a template
- Zipped download of all payslips
- Bank transfer summary table
- Exportable bank transfer files (CSV / Excel)

---

## Folder Structure

```
project/
├── app.R
├── employeeslip.xlsx
├── README.md
```

### Required Files
- **employeeslip.xlsx**
  - Excel payslip template
  - Must contain a sheet named `Template`
  - Cell positions must match the write mapping used in the app

---

## Required R Packages

```r
library(shiny)
library(readxl)
library(dplyr)
library(tibble)
library(openxlsx)
library(zip)
```

Install if missing:
```r
install.packages(c("shiny","readxl","dplyr","tibble","openxlsx","zip"))
```

---

## Input Data Format

The app accepts an Excel file with the following columns:

| Column Name | Description |
|------------|-------------|
| STAFF NAME | Employee full name |
| BANK NAME | Bank for salary transfer |
| ACCOUNT NUMBER | Bank account number |
| IDENTIFICATION CARD | NRIC / ID |
| BASIC | Monthly basic salary |
| ANNUAL LEAVE | Annual leave entitlement |
| ABSENCE | Number of unpaid absence days |
| MEDICAL LEAVE | Medical leave days |
| OT | Overtime hours |
| SUNDAY | Number of Sunday workdays |
| ALLOWANCE | Fixed allowance |
| COMMISSION | Sales or performance commission |
| CASH ADVANCE COMPANY | Company cash advance |
| CASH ADVANCE MANAGER | Manager cash advance |
| EMPLOYER EPF RATE | Employer EPF contribution rate |
| MARITAL STATUS | single / married_spouse_working / married_spouse_not_working |
| CHILDREN | Number of dependent children |

Missing columns are automatically created with safe defaults.

---

## Payroll Calculation Logic

### 1. Daily & Leave Calculation
- `Daily Pay = BASIC / Working Days`
- `No Pay Leave = Daily Pay × ABSENCE`
- `Basic – No Pay Leave` used for statutory calculations

### 2. Attendance Allowance
Attendance allowance is **paid only if**:
- No unpaid absence
- Medical leave ≤ 2 days
- Basic salary below threshold

### 3. Overtime (OT)
- OT base can include attendance allowance (optional)
- OT hourly rate calculated from selected base
- `Overtime Pay = OT Hours × OT Hourly Rate`

### 4. Sunday Pay
- Calculated per Sunday worked
- Optional inclusion of attendance allowance in base

### 5. Statutory Wage
Used for SOCSO, EIS, and PCB:
```
Statutory Wage = (Basic – No Pay Leave) + Overtime + Sunday Pay
```

Attendance allowance, allowance, and commission are **excluded** from statutory wage.

---

## Statutory Contributions

### EPF
- Employee EPF rate set globally
- Employer EPF rate can vary per employee
- Calculated on `(Basic – No Pay Leave)`

### SOCSO & EIS
- Calculated based on **Statutory Wage**
- Uses official 2025 contribution tables
- Automatically selects correct wage band

### PCB (Income Tax)
- Annualized calculation
- Applies:
  - Self relief
  - Spouse relief (if applicable)
  - Child relief
  - EPF relief (capped)
- Converted back to monthly PCB

---

## Net & Final Pay

### Nett Pay
```
Nett Pay =
Statutory Wage
− Employee EPF
− Employee SOCSO
− Employee EIS
− Income Tax
```

### Final Pay (Bank Transfer Amount)
```
Final Pay =
Nett Pay
+ Attendance Allowance
+ Allowance
+ Commission
− Cash Advances
```

This is the amount transferred to the employee’s bank account.

---

## Payslip Generation

For each employee:
- Loads `employeeslip.xlsx`
- Writes values to fixed cells (B2, B3, D2, D3, etc.)
- Saves individual payslip
- Compresses all payslips into a ZIP file

### Payslip Cell Mapping (Key Examples)

| Field | Cell |
|------|------|
| Staff Name | B2 |
| ID | B3 |
| Pay Month | D2 |
| Basic Pay | B6 |
| Overtime | B7 |
| Allowance | B10 |
| Commission | B11 |
| Gross Salary | B12 |
| Employee EPF | D6 |
| Total Deductions | D12 |
| Final Pay | D18 |

---

## Bank Transfer Summary

The app generates a clean bank transfer table:

| Name | Bank | Account Number | Total (RM) |
|-----|------|----------------|-----------|

This table can be:
- Viewed in the app
- Downloaded as CSV
- Downloaded as Excel

Suitable for direct upload to bank bulk payment systems.

---

## How to Run

```r
shiny::runApp()
```

Or open `app.R` in RStudio and click **Run App**.

---

## Validation & Safety

- Missing columns auto-filled
- Numeric fields coerced safely
- All monetary values rounded to 2 decimals
- Template existence validated before payslip generation

---

## Disclaimer

This system provides **estimated payroll calculations** and should be validated against official Malaysian statutory requirements before production use.

---

## Author / Maintainer

Payroll system implemented in **R + Shiny**  
Designed for clarity, auditability, and operational payroll workflows.
