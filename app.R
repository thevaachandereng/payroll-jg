library(shiny)
library(readxl)
library(dplyr)
library(tibble)
library(openxlsx)
library(zip)
library(processx)   # LibreOffice
library(blastula)   # email

# =====================================================
# SAMPLE DATA (DEFAULT)
# =====================================================
sample_data <- tribble(
  ~`STAFF NAME`, ~`BANK NAME`, ~`ACCOUNT NUMBER`, ~`IDENTIFICATION CARD`,
  ~BASIC, ~`ANNUAL LEAVE`, ~ABSENCE, ~`MEDICAL LEAVE`, ~OT, ~SUNDAY,
  ~`ALLOWANCE`, ~`COMMISSION`,
  ~`CASH ADVANCE COMPANY`, ~`CASH ADVANCE MANAGER`,
  ~`EMPLOYER EPF RATE`,
  ~`MARITAL STATUS`, ~CHILDREN, ~EMAIL,

  "Ali Ahmad", "Maybank", "123-456-789", "900101-10-1234",
  3500, 1, 0, 0, 10, 4,
  100, 250,
  200, 0,
  0.13,
  "married_spouse_not_working", 2, "thevaasiinen@gmail.com")

  # "Siti Noor", "CIMB", "987-654-321", "920202-08-5678",
  # 4200, 0, 2, 0, 5, 0,
  # 150, 0,
  # 0, 100,
  # 0.13,
  # "single", 0,
  #
  # "Kumar Raj", "RHB", "555-222-888", "880303-14-9988",
  # 5000, 2, 0, 1, 12, 6,
  # 300, 500,
  # 300, 0,
  # 0.13,
  # "married_spouse_working", 1,
  #
  # "Aisyah Rahman", "Hong Leong", "444-111-777", "950505-09-4455",
  # 2800, 0, 1, 0, 0, 0,
  # 80, 0,
  # 0, 0,
  # 0.13,
  # "single", 0,
  #
  # "Daniel Wong", "Public Bank", "333-999-666", "870707-01-3322",
  # 6000, 1, 0, 0, 20, 8,
  # 200, 800,
  # 0, 0,
  # 0.13,
  # "married_spouse_working", 3,
  #
  # "Nur Aina Zulkifli", "AmBank", "222-888-555", "990909-06-7788",
  # 3200, 0, 0, 0, 6, 2,
  # 120, 100,
  # 100, 50,
  # 0.13,
  # "single", 0,
  #
  # "Mohd Firdaus", "BSN", "111-777-444", "850404-02-6611",
  # 3800, 3, 2, 0, 4, 0,
  # 150, 0,
  # 0, 0,
  # 0.13,
  # "married_spouse_not_working", 4,
  #
  # "Rachel Tan", "Maybank", "666-555-333", "930808-12-2244",
  # 4500, 0, 0, 2, 15, 5,
  # 220, 300,
  # 150, 0,
  # 0.13,
  # "single", 0
#)




# =====================================================
# HELPERS
# =====================================================
lookup_band <- function(tbl, wage) {
  tbl %>% filter(wage > wage_low, wage <= wage_high) %>% slice(1)
}

template_path <- "employeeslip.xlsx"

# =====================================================
# SOCSO TABLE 2025 (FIRST CATEGORY - EMPLOYER + EMPLOYEE)
# (Using the SOCSO schedule you provided)
# =====================================================
socso_act4 <- tribble(
  ~wage_low, ~wage_high,
  ~employer_fc, ~employee_fc, ~total_fc,
  ~employer_sc,

  0,    30,   0.40,  0.10,  0.50,  0.30,
  30,    50,   0.70,  0.20,  0.90,  0.50,
  50,    70,   1.10,  0.30,  1.40,  0.80,
  70,   100,   1.50,  0.40,  1.90,  1.10,
  100,   140,   2.10,  0.60,  2.70,  1.50,
  140,   200,   2.95,  0.85,  3.80,  2.10,
  200,   300,   4.35,  1.25,  5.60,  3.10,
  300,   400,   6.15,  1.75,  7.90,  4.40,
  400,   500,   7.85,  2.25, 10.10,  5.60,
  500,   600,   9.65,  2.75, 12.40,  6.90,
  600,   700,  11.35,  3.25, 14.60,  8.10,
  700,   800,  13.15,  3.75, 16.90,  9.40,
  800,   900,  14.85,  4.25, 19.10, 10.60,
  900,  1000,  16.65,  4.75, 21.40, 11.90,

  1000,  1100,  18.35,  5.25, 23.60, 13.10,
  1100,  1200,  20.15,  5.75, 25.90, 14.40,
  1200,  1300,  21.85,  6.25, 28.10, 15.60,
  1300,  1400,  23.65,  6.75, 30.40, 16.90,
  1400,  1500,  25.35,  7.25, 32.60, 18.10,
  1500,  1600,  27.15,  7.75, 34.90, 19.40,
  1600,  1700,  28.85,  8.25, 37.10, 20.60,
  1700,  1800,  30.65,  8.75, 39.40, 21.90,
  1800,  1900,  32.35,  9.25, 41.60, 23.10,
  1900,  2000,  34.15,  9.75, 43.90, 24.40,

  2000,  2100,  35.85, 10.25, 46.10, 25.60,
  2100,  2200,  37.65, 10.75, 48.40, 26.90,
  2200,  2300,  39.35, 11.25, 50.60, 28.10,
  2300,  2400,  41.15, 11.75, 52.90, 29.40,
  2400,  2500,  42.85, 12.25, 55.10, 30.60,
  2500,  2600,  44.65, 12.75, 57.40, 31.90,
  2600,  2700,  46.35, 13.25, 59.60, 33.10,
  2700,  2800,  48.15, 13.75, 61.90, 34.40,
  2800,  2900,  49.85, 14.25, 64.10, 35.60,
  2900,  3000,  51.65, 14.75, 66.40, 36.90,

  3000,  3100,  53.35, 15.25, 68.60, 38.10,
  3100,  3200,  55.15, 15.75, 70.90, 39.40,
  3200,  3300,  56.85, 16.25, 73.10, 40.60,
  3300,  3400,  58.65, 16.75, 75.40, 41.90,
  3400,  3500,  60.35, 17.25, 77.60, 43.10,
  3500,  3600,  62.15, 17.75, 79.90, 44.40,
  3600,  3700,  63.85, 18.25, 82.10, 45.60,
  3700,  3800,  65.65, 18.75, 84.40, 46.90,
  3800,  3900,  67.35, 19.25, 86.60, 48.10,
  3900,  4000,  69.15, 19.75, 88.90, 49.40,

  4000,  4100,  70.85, 20.25, 91.10, 50.60,
  4100,  4200,  72.65, 20.75, 93.40, 51.90,
  4200,  4300,  74.35, 21.25, 95.60, 53.10,
  4300,  4400,  76.15, 21.75, 97.90, 54.40,
  4400,  4500,  77.85, 22.25,100.10, 55.60,
  4500,  4600,  79.65, 22.75,102.40, 56.90,
  4600,  4700,  81.35, 23.25,104.60, 58.10,
  4700,  4800,  83.15, 23.75,106.90, 59.40,
  4800,  4900,  84.85, 24.25,109.10, 60.60,
  4900,  5000,  86.65, 24.75,111.40, 61.90,

  5000,  5100,  88.35, 25.25,113.60, 63.10,
  5100,  5200,  90.15, 25.75,115.90, 64.40,
  5200,  5300,  91.85, 26.25,118.10, 65.60,
  5300,  5400,  93.65, 26.75,120.40, 66.90,
  5400,  5500,  95.35, 27.25,122.60, 68.10,
  5500,  5600,  97.15, 27.75,124.90, 69.40,
  5600,  5700,  98.85, 28.25,127.10, 70.60,
  5700,  5800, 100.65, 28.75,129.40, 71.90,
  5800,  5900, 102.35, 29.25,131.60, 73.10,
  5900,  6000, 104.15, 29.75,133.90, 74.40,

  6000,    Inf, 104.15, 29.75,133.90, 74.40
)

calc_socso_fc <- function(wage) {
  socso_act4 %>%
    filter(wage > wage_low, wage <= wage_high) %>%
    slice(1) %>%
    select(employer_fc, employee_fc, total_fc)
}

# =====================================================
# EIS TABLE 2025
# =====================================================
eis_act800 <- tribble(
  ~wage_low, ~wage_high, ~employer_eis, ~employee_eis, ~total_eis,

  0,    30,  0.05,  0.05,  0.10,
  30,    50,  0.10,  0.10,  0.20,
  50,    70,  0.15,  0.15,  0.30,
  70,   100,  0.20,  0.20,  0.40,
  100,   140,  0.25,  0.25,  0.50,
  140,   200,  0.35,  0.35,  0.70,
  200,   300,  0.50,  0.50,  1.00,
  300,   400,  0.70,  0.70,  1.40,
  400,   500,  0.90,  0.90,  1.80,
  500,   600,  1.10,  1.10,  2.20,
  600,   700,  1.30,  1.30,  2.60,
  700,   800,  1.50,  1.50,  3.00,
  800,   900,  1.70,  1.70,  3.40,
  900,  1000,  1.90,  1.90,  3.80,

  1000,  1100,  2.10,  2.10,  4.20,
  1100,  1200,  2.30,  2.30,  4.60,
  1200,  1300,  2.50,  2.50,  5.00,
  1300,  1400,  2.70,  2.70,  5.40,
  1400,  1500,  2.90,  2.90,  5.80,
  1500,  1600,  3.10,  3.10,  6.20,
  1600,  1700,  3.30,  3.30,  6.60,
  1700,  1800,  3.50,  3.50,  7.00,
  1800,  1900,  3.70,  3.70,  7.40,
  1900,  2000,  3.90,  3.90,  7.80,

  2000,  2100,  4.10,  4.10,  8.20,
  2100,  2200,  4.30,  4.30,  8.60,
  2200,  2300,  4.50,  4.50,  9.00,
  2300,  2400,  4.70,  4.70,  9.40,
  2400,  2500,  4.90,  4.90,  9.80,
  2500,  2600,  5.10,  5.10, 10.20,
  2600,  2700,  5.30,  5.30, 10.60,
  2700,  2800,  5.50,  5.50, 11.00,
  2800,  2900,  5.70,  5.70, 11.40,
  2900,  3000,  5.90,  5.90, 11.80,

  3000,  3100,  6.10,  6.10, 12.20,
  3100,  3200,  6.30,  6.30, 12.60,
  3200,  3300,  6.50,  6.50, 13.00,
  3300,  3400,  6.70,  6.70, 13.40,
  3400,  3500,  6.90,  6.90, 13.80,
  3500,  3600,  7.10,  7.10, 14.20,
  3600,  3700,  7.30,  7.30, 14.60,
  3700,  3800,  7.50,  7.50, 15.00,
  3800,  3900,  7.70,  7.70,  15.40,
  3900,  4000,  7.90,  7.90,  15.80,

  4000,  4100,  8.10,  8.10,  16.20,
  4100,  4200,  8.30,  8.30,  16.60,
  4200,  4300,  8.50,  8.50,  17.00,
  4300,  4400,  8.70,  8.70,  17.40,
  4400,  4500,  8.90,  8.90,  17.80,
  4500,  4600,  9.10,  9.10,  18.20,
  4600,  4700,  9.30,  9.30,  18.60,
  4700,  4800,  9.50,  9.50,  19.00,
  4800,  4900,  9.70,  9.70,  19.40,
  4900,  5000,  9.90,  9.90,  19.80,

  5000,  5100, 10.10, 10.10, 20.20,
  5100,  5200, 10.30, 10.30, 20.60,
  5200,  5300, 10.50, 10.50, 21.00,
  5300,  5400, 10.70, 10.70, 21.40,
  5400,  5500, 10.90, 10.90, 21.80,
  5500,  5600, 11.10, 11.10, 22.20,
  5600,  5700, 11.30, 11.30, 22.60,
  5700,  5800, 11.50, 11.50, 23.00,
  5800,  5900, 11.70, 11.70, 23.40,
  5900,  6000, 11.90, 11.90, 23.80,

  6000,   Inf, 11.90, 11.90, 23.80
)

calc_eis <- function(wage) {
  eis_act800 %>%
    filter(wage > wage_low, wage <= wage_high) %>%
    slice(1) %>%
    select(employer_eis, employee_eis, total_eis)
}

# =====================================================
# INCOME TAX (YOUR BRACKETS + YOUR RELIEFS APPROX)
# =====================================================
tax_annual_from_brackets <- function(chargeable_income) {
  x <- max(chargeable_income, 0)
  dplyr::case_when(
    x <= 5000 ~ 0,
    x <= 20000 ~ (x - 5000) * 0.01,
    x <= 35000 ~ 150 + (x - 20000) * 0.03,
    x <= 50000 ~ 600 + (x - 35000) * 0.06,
    x <= 70000 ~ 1500 + (x - 50000) * 0.11,
    x <= 100000 ~ 3700 + (x - 70000) * 0.19,
    x <= 400000 ~ 9400 + (x - 100000) * 0.25,
    x <= 600000 ~ 84400 + (x - 400000) * 0.26,
    x <= 2000000 ~ 136400 + (x - 600000) * 0.28,
    TRUE ~ 528400 + (x - 2000000) * 0.30
  )
}

calc_pcb <- function(monthly_salary, status, children, epf_employee_rate) {
  relief_self <- 9000
  relief_spouse <- ifelse(status == "married_spouse_not_working", 4000, 0)
  relief_children <- children * 2000

  epf_annual <- monthly_salary * epf_employee_rate * 12
  epf_relief <- min(epf_annual, 4000)

  chargeable_annual <-
    monthly_salary * 12 -
    relief_self - relief_spouse - relief_children - epf_relief

  annual_tax <- tax_annual_from_brackets(chargeable_annual)
  rebate_self <- ifelse(chargeable_annual <= 35000, 400, 0)

  round(max(annual_tax - rebate_self, 0) / 12, 2)
}

make_unique_sheet_name <- function(name, existing) {
  base <- substr(name, 1, 28)
  i <- 1
  new_name <- base
  while (new_name %in% existing) {
    new_name <- paste0(base, "_", i)
    i <- i + 1
  }
  new_name
}

find_qpdf <- function() {
  paths <- c(
    "/opt/homebrew/bin/qpdf",
    "/usr/local/bin/qpdf",
    "/usr/bin/qpdf",
    "C:/Program Files/qpdf/bin/qpdf.exe"
  )
  p <- paths[file.exists(paths)]
  if (length(p) == 0) stop("qpdf not found")
  p[1]
}

get_ic_password <- function(ic) {
  digits <- gsub("\\D", "", ic)
  substr(digits, nchar(digits) - 3, nchar(digits))
}

encrypt_pdf <- function(input, output, password) {
  processx::run(
    find_qpdf(),
    c("--encrypt", password, password, "256", "--", input, output),
    error_on_status = TRUE
  )
}


excel_to_pdf <- function(xlsx, out_dir) {
  soffice <- if (.Platform$OS.type == "windows") {
    "C:/Program Files/LibreOffice/program/soffice.exe"
  } else {
    "/Applications/LibreOffice.app/Contents/MacOS/soffice"
  }

  processx::run(
    soffice,
    c("--headless", "--convert-to", "pdf", "--outdir", out_dir, xlsx),
    error_on_status = TRUE
  )
}

generate_payslip_files <- function(emp, pay_month, pay_date, tmp_dir) {

  safe <- gsub("[^[:alnum:] _-]", "", emp$`STAFF NAME`)

  xlsx <- file.path(tmp_dir, paste0(safe, "_", pay_month, ".xlsx"))
  pdf  <- file.path(tmp_dir, paste0(safe, "_", pay_month, ".pdf"))
  enc  <- file.path(tmp_dir, paste0(safe, "_", pay_month, "_protected.pdf"))

  wb <- openxlsx::loadWorkbook("employeeslip.xlsx")

  write_one_payslip(
    wb, "Template", emp,
    pay_month = pay_month,
    pay_date  = pay_date
  )

  openxlsx::saveWorkbook(wb, xlsx, overwrite = TRUE)
  excel_to_pdf(xlsx, tmp_dir)

  pwd <- get_ic_password(emp$`IDENTIFICATION CARD`)
  encrypt_pdf(pdf, enc, pwd)

  list(pdf = enc, password = pwd)
}

write_one_payslip <- function(wb, sheet, emp, pay_month, pay_date) {
  # Header
  writeData(wb, sheet, emp$`STAFF NAME`,          startCol = 2, startRow = 2) # B2
  writeData(wb, sheet, emp$`IDENTIFICATION CARD`, startCol = 2, startRow = 3) # B3
  writeData(wb, sheet, pay_month,                 startCol = 4, startRow = 2) # D2
  writeData(wb, sheet, as.character(pay_date),    startCol = 4, startRow = 3) # D3

  # Earnings
  writeData(wb, sheet, emp$BASIC,                  startCol = 2, startRow = 6)  # B6
  writeData(wb, sheet, emp$OVERTIME,               startCol = 2, startRow = 7)  # B7
  writeData(wb, sheet, emp$`ATTENDANCE ALLOWANCE`, startCol = 2, startRow = 8)  # B8
  writeData(wb, sheet, emp$`SUNDAY PAY`,           startCol = 2, startRow = 9)  # B9
  writeData(wb, sheet, emp$`ALLOWANCE`,            startCol = 2, startRow = 10) # B10
  writeData(wb, sheet, emp$`COMMISSION`,           startCol = 2, startRow = 11) # B11
  writeData(wb, sheet, emp$`GROSS SALARY`,         startCol = 2, startRow = 12) # B12


  # Deductions
  writeData(wb, sheet, emp$`EMPLOYEE EPF`,         startCol = 4, startRow = 6)  # D6
  writeData(wb, sheet, emp$`EMPLOYEE SOCSO`,       startCol = 4, startRow = 7)  # D7
  writeData(wb, sheet, emp$`NO PAY LEAVE`,         startCol = 4, startRow = 8)  # D8 (unpaid basic due to absence)
  writeData(wb, sheet, emp$`EMPLOYEE EIS`,         startCol = 4, startRow = 9)  # D9
  writeData(wb, sheet, emp$`INCOME TAX`,           startCol = 4, startRow = 10) # D10
  writeData(wb, sheet, emp$`TOTAL DEDUCTIONS`,     startCol = 4, startRow = 12) # D12

  cash_adv <- emp$`CASH ADVANCE COMPANY` + emp$`CASH ADVANCE MANAGER`
  writeData(wb, sheet, cash_adv,                   startCol = 4, startRow = 11) # D11

  # Employer + bank + net pay
  writeData(wb, sheet, emp$`EMPLOYER EPF`,   startCol = 1, startRow = 16) # A16
  writeData(wb, sheet, emp$`EMPLOYER SOCSO`, startCol = 2, startRow = 16) # B16
  writeData(wb, sheet, emp$`EMPLOYER EIS`,   startCol = 2, startRow = 17) # B17
  writeData(wb, sheet, emp$`BANK NAME`,      startCol = 3, startRow = 16) # C16
  writeData(wb, sheet, emp$`ACCOUNT NUMBER`, startCol = 3, startRow = 17) # C17
  writeData(wb, sheet, emp$`FINAL PAY`,      startCol = 4, startRow = 18) # D18
}


# =====================================================
# UI
# =====================================================
ui <- fluidPage(
  titlePanel("Malaysia Payroll + SOCSO/EIS + Tax (Per Employee)"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Excel (optional)", accept = ".xlsx"),

      numericInput("days", "Working Days", 26, min = 20, max = 31),
      numericInput("hours", "Hours per Day", 8, min = 4, max = 12),

      numericInput("attendance_allowance", "Attendance Allowance (RM)", 300, min = 0, max = 10000),
      numericInput("attendance_threshold", "Attendance Allowance Threshold (RM)", 6000, min = 0),

      numericInput("ot_multiplier", "OT Multiplier", 1.5, step = 0.1),
      selectInput("ot_base", "OT Base",
                  c("Basic only" = "basic",
                    "Basic + Attendance Allowance" = "basic_att")),

      numericInput("sunday_multiplier", "Sunday Multiplier", 2.0, step = 0.1),
      selectInput("sunday_base", "Sunday Base",
                  c("Basic only" = "basic",
                    "Basic + Attendance Allowance" = "basic_att")),

      numericInput("epf_employee_rate", "EPF Employee Rate", 0.11, step = 0.01),
      textInput("pay_month", "Pay Month (e.g. August 2025)", "August 2025"),
      dateInput("pay_date", "Pay Date", value = Sys.Date()),

      hr(),
      downloadButton("download_sample", "Download Sample Excel")
    ),
    mainPanel(
      h4("Final Payroll Table"),
      tableOutput("final_tbl"),
      br(),
      downloadButton("download_csv", "Download CSV"),
      downloadButton("download_excel", "Download Excel"),
      downloadButton("download_payslips", "Download Payslips (Excel)"),
      actionButton("email_all", "📧 Email Payslips (PDF)"),
      hr(),
      h4("Bank Transfer Summary"),
      tableOutput("bank_tbl"),
      br(),
      downloadButton("download_bank_csv", "Download Bank Summary (CSV)"),
      downloadButton("download_bank_excel", "Download Bank Summary (Excel)")
    )
  )
)

generate_single_payslip <- function(
    emp,
    pay_month,
    pay_date,
    template_path,
    output_file
) {

  wb <- loadWorkbook(template_path)
  sheet <- names(wb)[1]  # template sheet

  # ------------------
  # HEADER
  # ------------------

  writeData(wb, sheet, emp$`STAFF NAME`, "B2", overwrite = TRUE)
  writeData(wb, sheet, emp$`IDENTIFICATION CARD`, "B3", overwrite = TRUE)
  writeData(wb, sheet, pay_month, "D2", overwrite = TRUE)
  writeData(wb, sheet, pay_date, "D3", overwrite = TRUE)

  writeData(wb, sheet, emp$BASIC, "B6", overwrite = TRUE)
  writeData(wb, sheet, emp$OVERTIME, "B7", overwrite = TRUE)
  writeData(wb, sheet, emp$`ATTENDANCE ALLOWANCE`, "B8", overwrite = TRUE)
  writeData(wb, sheet, emp$`SUNDAY PAY`, "B9", overwrite = TRUE)
  writeData(wb, sheet, emp$`ALLOWANCE`, "B10", overwrite = TRUE)
  writeData(wb, sheet, emp$`COMMISSION`, "B11", overwrite = TRUE)
  writeData(wb, sheet, emp$`GROSS SALARY`, "B12", overwrite = TRUE)

  writeData(wb, sheet, emp$`EMPLOYEE EPF`, "D6", overwrite = TRUE)
  writeData(wb, sheet, emp$`EMPLOYEE SOCSO`, "D7", overwrite = TRUE)
  writeData(wb, sheet, emp$`NO PAY LEAVE`, "D8", overwrite = TRUE)
  writeData(wb, sheet, emp$`EMPLOYEE EIS`, "D9", overwrite = TRUE)
  writeData(wb, sheet, emp$`INCOME TAX`, "D10", overwrite = TRUE)
  writeData(wb, sheet, emp$`TOTAL DEDUCTIONS`, "D12", overwrite = TRUE) # D12

  advance_total <- emp$`CASH ADVANCE COMPANY` + emp$`CASH ADVANCE MANAGER`
  writeData(wb, sheet, advance_total, "D11", overwrite = TRUE)

  writeData(wb, sheet, emp$`EMPLOYER EPF`,   "A16", overwrite = TRUE)
  writeData(wb, sheet, emp$`EMPLOYER SOCSO`, "B16", overwrite = TRUE)
  writeData(wb, sheet, emp$`EMPLOYER EIS`,   "B17", overwrite = TRUE)  # ✅
  writeData(wb, sheet, emp$`BANK NAME`,      "C16", overwrite = TRUE)  # ✅
  writeData(wb, sheet, emp$`ACCOUNT NUMBER`, "C17", overwrite = TRUE)  # ✅
  writeData(wb, sheet, emp$`FINAL PAY`,      "D18", overwrite = TRUE)

  saveWorkbook(wb, output_file, overwrite = TRUE)

}




# =====================================================
# SERVER
# =====================================================
server <- function(input, output) {

  raw_data <- reactive({
    df <- if (is.null(input$file)) {
      sample_data
    } else {
      read_excel(input$file$datapath)
    }

    # Ensure required columns exist (for uploaded Excel)
    if (!"IDENTIFICATION CARD" %in% names(df)) df$`IDENTIFICATION CARD` <- ""
    if (!"ACCOUNT NUMBER" %in% names(df)) df$`ACCOUNT NUMBER` <- ""
    if (!"BANK NAME" %in% names(df)) df$`BANK NAME` <- ""
    if (!"MEDICAL LEAVE" %in% names(df)) df$`MEDICAL LEAVE` <- 0
    if (!"ALLOWANCE" %in% names(df)) df$`ALLOWANCE` <- 0
    if (!"EMAIL" %in% names(df)) df$EMAIL <- NA_character_
    if (!"COMMISSION" %in% names(df)) df$`COMMISSION` <- 0

    # Employer EPF rate column (default 13%)
    if (!"EMPLOYER EPF RATE" %in% names(df)) {
      df$`EMPLOYER EPF RATE` <- 0.13
    } else {
      df$`EMPLOYER EPF RATE` <- ifelse(
        is.na(df$`EMPLOYER EPF RATE`) | df$`EMPLOYER EPF RATE` == "",
        0.13,
        as.numeric(df$`EMPLOYER EPF RATE`)
      )
      df$`EMPLOYER EPF RATE`[is.na(df$`EMPLOYER EPF RATE`)] <- 0.13
    }

    df
  })


  payroll <- reactive({
    df <- raw_data()

    df %>%
      mutate(
        # ---- daily / leave
        `DAILY PAY` = BASIC / input$days,
        `NO PAY LEAVE` = ifelse(ABSENCE > 0, `DAILY PAY` * ABSENCE, 0),
        `BASIC - NO PAY LEAVE` = BASIC - `NO PAY LEAVE`,

        # ---- attendance allowance (conditional; paid only if no absence and under threshold)
        `ATTENDANCE ALLOWANCE` =
          ifelse(ABSENCE > 0 | `MEDICAL LEAVE` > 2 | BASIC > input$attendance_threshold, 0,
          input$attendance_allowance
          ),


        # ---- OT base: include attendance ONLY if selected AND no absence
        ot_base_salary =
          ifelse(
            input$ot_base == "basic_att" & ABSENCE == 0 & `MEDICAL LEAVE` < 3,
            BASIC + `ATTENDANCE ALLOWANCE`,
            BASIC
          ),

        `OT/HOUR` = ot_base_salary / input$days / input$hours * input$ot_multiplier,
        OVERTIME  = `OT/HOUR` * OT,

        # ---- Sunday base (daily count, optional attendance allowance)
        sunday_base_salary =
          ifelse(
            input$sunday_base == "basic_att" & ABSENCE == 0 & `MEDICAL LEAVE` < 3,
            BASIC + `ATTENDANCE ALLOWANCE`,
            BASIC
          ),

        `SUNDAY PAY` =
          ifelse(
            SUNDAY > 0,
            sunday_base_salary / input$days * input$sunday_multiplier * SUNDAY,
            0
          ),

        # ---- statutory wage EXCLUDES attendance allowance
        `STATUTORY WAGE` =
          `BASIC - NO PAY LEAVE` + OVERTIME + `SUNDAY PAY`,


        # ---- gross display (includes attendance allowance; transport excluded)
        `GROSS SALARY` =
          `STATUTORY WAGE` + `ATTENDANCE ALLOWANCE` + `ALLOWANCE` + `COMMISSION`,

        # ---- EPF (basic - no pay leave)
        `EMPLOYEE EPF` = ceiling(`BASIC - NO PAY LEAVE` * input$epf_employee_rate),
        `EMPLOYER EPF` = ceiling(`BASIC - NO PAY LEAVE` * `EMPLOYER EPF RATE`)
      ) %>%
      rowwise() %>%
      mutate(
        # ---- SOCSO & EIS on STATUTORY WAGE (NOT gross, so attendance excluded)
        socso = list(calc_socso_fc(`STATUTORY WAGE`)),
        eis   = list(calc_eis(`STATUTORY WAGE`)),

        `EMPLOYER SOCSO` = socso$employer_fc,
        `EMPLOYEE SOCSO` = socso$employee_fc,

        `EMPLOYER EIS` = eis$employer_eis,
        `EMPLOYEE EIS` = eis$employee_eis,

        # ---- Income tax (PCB) on STATUTORY WAGE (attendance excluded)
        `INCOME TAX` = calc_pcb(
          monthly_salary     = `STATUTORY WAGE`,
          status             = `MARITAL STATUS`,
          children           = CHILDREN,
          epf_employee_rate  = input$epf_employee_rate
        ),

        # ---- TOTAL EMPLOYER COST (you pay attendance allowance, so include it)
        `TOTAL EMPLOYER COST` =
          `STATUTORY WAGE` +
          `ATTENDANCE ALLOWANCE` +
          `ALLOWANCE` +
          `COMMISSION` +
          `EMPLOYER EPF` +
          `EMPLOYER SOCSO` +
          `EMPLOYER EIS`
      ) %>%
      ungroup() %>%
      mutate(
        # ---- Nett pay after statutory deductions + tax (based on STATUTORY WAGE)
        `NETT PAY` =
          `STATUTORY WAGE` -
          `EMPLOYEE EPF` -
          `EMPLOYEE SOCSO` -
          `EMPLOYEE EIS` -
          `INCOME TAX`,

        # ---- Final pay: add attendance allowance at the end + transport - advances
        `FINAL PAY` =
          `NETT PAY` +
          `ATTENDANCE ALLOWANCE` +
          `ALLOWANCE` +
          `COMMISSION` -
          `CASH ADVANCE COMPANY` -
          `CASH ADVANCE MANAGER`,

        ## adding total deductions for payslip
        `TOTAL DEDUCTIONS` =
          `EMPLOYEE EPF` +
          `EMPLOYEE SOCSO` +
          `EMPLOYEE EIS` +
          `INCOME TAX` +
          `CASH ADVANCE COMPANY` +
          `CASH ADVANCE MANAGER`
      ) %>%
      mutate(
        across(where(is.numeric), ~ round(.x, 2))
      ) %>%
      select(
        `STAFF NAME`,
        EMAIL,                     # ✅ REQUIRED
        `BANK NAME`,
        `ACCOUNT NUMBER`,
        `IDENTIFICATION CARD`,
        `MARITAL STATUS`,
        CHILDREN,
        `EMPLOYER EPF RATE`,
        `ANNUAL LEAVE`,
        ABSENCE,
        `MEDICAL LEAVE`,
        BASIC, `DAILY PAY`,
        `OT/HOUR`,
        SUNDAY,
        OT,
        OVERTIME,
        `NO PAY LEAVE`,
        `SUNDAY PAY`,
        `BASIC - NO PAY LEAVE`,
        `ATTENDANCE ALLOWANCE`,
        `ALLOWANCE`,
        `COMMISSION`,
        `GROSS SALARY`,
        `EMPLOYER EPF`,
        `EMPLOYEE EPF`,
        `EMPLOYER SOCSO`,
        `EMPLOYEE SOCSO`,
        `EMPLOYER EIS`,
        `EMPLOYEE EIS`,
        `INCOME TAX`,
        `NETT PAY`,
        `CASH ADVANCE COMPANY`,
        `CASH ADVANCE MANAGER`,
        `FINAL PAY`,
        `TOTAL EMPLOYER COST`,
        `TOTAL DEDUCTIONS`
      )
  })

  output$final_tbl <- renderTable(payroll())

  output$download_sample <- downloadHandler(
    filename = "sample_payroll.xlsx",
    content = function(file) {
      openxlsx::write.xlsx(sample_data, file, overwrite = TRUE)
    },
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )

  output$download_csv <- downloadHandler(
    filename = "final_payroll.csv",
    content = function(file) write.csv(payroll(), file, row.names = FALSE)
  )

  output$download_excel <- downloadHandler(
    filename = "final_payroll.xlsx",
    content = function(file) {
      openxlsx::write.xlsx(payroll(), file, overwrite = TRUE)
    },
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )

  output$download_payslips <- downloadHandler(
    filename = function() paste0("Payslips_", input$pay_month, ".zip"),
    content = function(file) {

      tmp_dir <- tempfile("payslips_")
      dir.create(tmp_dir)

      df <- payroll()

      for (i in seq_len(nrow(df))) {
        emp <- df[i, ]

        wb <- openxlsx::loadWorkbook("employeeslip.xlsx")
        sheet <- "Template"  # must match your template sheet name

        write_one_payslip(
          wb        = wb,
          sheet     = sheet,
          emp       = emp,
          pay_month = input$pay_month,
          pay_date  = input$pay_date
        )

        safe_name <- gsub("[^[:alnum:] _-]", "", emp$`STAFF NAME`)
        out_xlsx  <- file.path(
          tmp_dir,
          paste0(safe_name, "_", input$pay_month, ".xlsx")
        )

        openxlsx::saveWorkbook(wb, out_xlsx, overwrite = TRUE)
      }

      zip::zipr(zipfile = file, files = list.files(tmp_dir, full.names = TRUE))
    },
    contentType = "application/zip"
  )

  bank_summary <- reactive({
    payroll() %>%
      transmute(
        `NAME` = `STAFF NAME`,
        `Bank` = `BANK NAME`,
        `Account Number` = `ACCOUNT NUMBER`,
        `Total (RM)` = round(`FINAL PAY`, 2)
      )
  })

  observeEvent(input$email_all, {

    df <- payroll()
    tmp_dir <- tempfile("email_")
    dir.create(tmp_dir)

    for (i in seq_len(nrow(df))) {

      emp <- df[i, ]

      if (is.na(emp$EMAIL) || emp$EMAIL == "") next

      files <- generate_payslip_files(
        emp       = emp,
        pay_month = input$pay_month,
        pay_date  = input$pay_date,
        tmp_dir   = tmp_dir
      )

      email <- compose_email(
        body = md(paste0(
          "Dear ", emp$`STAFF NAME`, ",\n\n",
          "Attached is your **password-protected payslip** for **",
          input$pay_month, "**.\n\n",
          "🔐 Password: **last 4 digits of your IC number**\n\n",
          "Regards,\nPayroll Team"
        ))
      ) %>%
        add_attachment(files$pdf)

      smtp_send(
        email,
        to = emp$EMAIL,
        from = "thevaasiinen@gmail.com",
        subject = paste0("Payslip – ", input$pay_month),
        credentials = creds_file("~/.gmail_payroll_creds")
      )
    }
  })



  output$bank_tbl <- renderTable({
    bank_summary()
  }, striped = TRUE, bordered = TRUE, spacing = "s")


  output$download_bank_csv <- downloadHandler(
    filename = function() {
      paste0("bank_transfer_summary_", input$pay_month, ".csv")
    },
    content = function(file) {
      write.csv(bank_summary(), file, row.names = FALSE)
    }
  )

  output$download_bank_excel <- downloadHandler(
    filename = function() {
      paste0("bank_transfer_summary_", input$pay_month, ".xlsx")
    },
    content = function(file) {
      openxlsx::write.xlsx(
        bank_summary(),
        file,
        overwrite = TRUE
      )
    },
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )

}

shinyApp(ui, server)
