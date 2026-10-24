# ==============================================================================
# crsp_stock_download.R
# Download daily CRSP stock data from WRDS for one or more tickers.
#
# You provide: ticker(s), start date, end date, and output format (csv / xlsx).
# The script writes a file to your working directory and tells you where it is.
#
# Data source : crsp.dsf_v2  (CRSP Stock Version 2, "CIZ" format, daily)
# Connection  : WRDS PostgreSQL server via DBI + RPostgres
#
# ONE-TIME SETUP
#   1. Install packages:
#        install.packages(c("DBI", "RPostgres", "writexl", "rstudioapi"))
#   2. (Optional but convenient) Store your WRDS username so you aren't asked
#      each time. Run usethis::edit_r_environ() (or edit ~/.Renviron) and add:
#        WRDS_USER=your_wrds_username
#      Do NOT store your password in a script. You'll be prompted for it, or
#      you can use a ~/.pgpass file (see WRDS documentation).
#   3. Your WRDS account will need Duo two-factor approval on first connection.
#
# USAGE
#   Interactive (prompts you for everything):
#        source("crsp_stock_download.R")
#        download_crsp()
#
#   Non-interactive (good for reproducible Quarto documents):
#        download_crsp(tickers = c("AAPL", "MSFT"),
#                      start_date = "2020-01-01",
#                      end_date   = "2024-12-31",
#                      format     = "csv")
# ==============================================================================

library(DBI)
library(RPostgres)

# ------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------

# Ask a question at the console and return the trimmed answer.
ask <- function(prompt) {
  trimws(readline(prompt))
}

# Parse a user-entered date; stop with a clear message if it isn't valid.
parse_date <- function(x, label) {
  d <- suppressWarnings(as.Date(x, format = "%Y-%m-%d"))
  if (is.na(d)) stop(label, " must be in YYYY-MM-DD format (e.g., 2020-01-31). You entered: ", x)
  d
}

# Connect to WRDS. Username from WRDS_USER env var (or prompt); password prompted.
connect_wrds <- function() {
  user <- Sys.getenv("WRDS_USER")
  if (user == "") user <- ask("WRDS username: ")

  pw <- Sys.getenv("WRDS_PASSWORD")   # only used if you set it yourself; not recommended
  if (pw == "") {
    if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
      pw <- rstudioapi::askForPassword("WRDS password")
    } else {
      pw <- ask("WRDS password: ")
    }
  }

  message("Connecting to WRDS (approve the Duo prompt if it appears)...")
  dbConnect(
    Postgres(),
    host     = "wrds-pgdata.wharton.upenn.edu",
    port     = 9737,
    dbname   = "wrds",
    sslmode  = "require",
    user     = user,
    password = pw
  )
}

# ------------------------------------------------------------------------------
# Main function
# ------------------------------------------------------------------------------

download_crsp <- function(tickers    = NULL,
                          start_date = NULL,
                          end_date   = NULL,
                          format     = NULL,
                          out_dir    = getwd()) {

  # --- 1. Collect inputs (prompt for anything not supplied) -------------------
  if (is.null(tickers)) {
    tickers <- ask("Ticker(s), separated by spaces or commas (e.g., AAPL MSFT): ")
  }
  tickers <- toupper(unlist(strsplit(tickers, "[ ,]+")))
  tickers <- tickers[tickers != ""]
  # Tickers are letters/digits (CRSP occasionally uses a period or hyphen).
  if (length(tickers) == 0 || any(!grepl("^[A-Z0-9.\\-]+$", tickers))) {
    stop("Please supply valid ticker symbols (letters, digits, '.' or '-').")
  }

  if (is.null(start_date)) start_date <- ask("Start date (YYYY-MM-DD): ")
  if (is.null(end_date))   end_date   <- ask("End date   (YYYY-MM-DD): ")
  start_date <- parse_date(start_date, "Start date")
  end_date   <- parse_date(end_date,   "End date")
  if (start_date > end_date) stop("Start date must be on or before the end date.")

  if (is.null(format)) format <- ask("Output format (csv or xlsx): ")
  format <- tolower(format)
  if (!format %in% c("csv", "xlsx")) stop("Format must be 'csv' or 'xlsx'.")

  # --- 2. Connect -------------------------------------------------------------
  con <- connect_wrds()
  on.exit(dbDisconnect(con), add = TRUE)

  # --- 3. Work out which columns exist in crsp.dsf_v2 -------------------------
  # Protects against variable-name changes by CRSP: we request only the
  # columns from our wish list that actually exist in the table.
  wanted <- c("permno", "permco", "ticker", "tradingsymbol", "securitynm",
              "dlycaldt", "dlyprc", "dlyret", "dlyretx", "dlyreti",
              "dlyvol", "dlycap", "shrout",
              "dlyopen", "dlyhigh", "dlylow", "dlyclose",
              "dlybid", "dlyask", "dlyfacprc")

  avail <- dbGetQuery(
    con,
    "SELECT column_name FROM information_schema.columns
     WHERE table_schema = 'crsp' AND table_name = 'dsf_v2'"
  )$column_name

  cols <- intersect(wanted, avail)
  if (!all(c("permno", "dlycaldt") %in% cols)) {
    stop("Could not find expected columns in crsp.dsf_v2. Check your CRSP access.")
  }
  missing_cols <- setdiff(wanted, avail)
  if (length(missing_cols) > 0) {
    message("Note: these columns were not found and were skipped: ",
            paste(missing_cols, collapse = ", "))
  }

  # --- 4. Build and run the query ---------------------------------------------
  ticker_list <- paste(dbQuoteString(con, tickers), collapse = ", ")

  sql <- sprintf(
    "SELECT %s
       FROM crsp.dsf_v2
      WHERE ticker IN (%s)
        AND dlycaldt BETWEEN %s AND %s
      ORDER BY ticker, permno, dlycaldt",
    paste(cols, collapse = ", "),
    ticker_list,
    dbQuoteString(con, as.character(start_date)),
    dbQuoteString(con, as.character(end_date))
  )

  message("Running query...")
  dat <- dbGetQuery(con, sql)

  if (nrow(dat) == 0) {
    stop("No rows returned. Check the ticker(s) and date range. ",
         "Note: `ticker` is the ticker as of each date, so a company that ",
         "changed its ticker will only appear under the symbol it used then.")
  }

  # Warn if one ticker maps to more than one PERMNO (e.g., ticker reused by a
  # different company over time, or multiple share classes).
  n_permno <- tapply(dat$permno, dat$ticker, function(x) length(unique(x)))
  multi <- names(n_permno)[n_permno > 1]
  if (length(multi) > 0) {
    warning("These tickers map to more than one PERMNO in your date range: ",
            paste(multi, collapse = ", "),
            ". Check the permno / securitynm columns to make sure you have the ",
            "company you intend.", call. = FALSE)
  }

  # Tickers requested but not found
  not_found <- setdiff(tickers, unique(dat$ticker))
  if (length(not_found) > 0) {
    warning("No data found for: ", paste(not_found, collapse = ", "), call. = FALSE)
  }

  # --- 5. Write the file ------------------------------------------------------
  stub <- paste0("crsp_daily_", paste(tickers, collapse = "-"), "_",
                 format(start_date, "%Y%m%d"), "_", format(end_date, "%Y%m%d"))
  # Keep file names reasonable if many tickers are requested
  if (nchar(stub) > 120) {
    stub <- paste0("crsp_daily_", length(tickers), "tickers_",
                   format(start_date, "%Y%m%d"), "_", format(end_date, "%Y%m%d"))
  }
  path <- file.path(out_dir, paste0(stub, ".", format))

  if (format == "csv") {
    write.csv(dat, path, row.names = FALSE)
  } else {
    if (!requireNamespace("writexl", quietly = TRUE)) {
      stop("Package 'writexl' is needed for xlsx output. Run install.packages('writexl').")
    }
    # Excel has a limit of 1,048,576 rows per sheet
    if (nrow(dat) > 1048575) {
      stop("Result has ", nrow(dat), " rows, which exceeds Excel's row limit. Use csv instead.")
    }
    writexl::write_xlsx(dat, path)
  }

  message("Done: ", nrow(dat), " rows written to\n  ", normalizePath(path))
  invisible(dat)   # also returns the data frame so you can keep working with it
}

# ------------------------------------------------------------------------------
# If this file is run directly (not sourced), start the interactive prompts.
# ------------------------------------------------------------------------------
if (sys.nframe() == 0 && interactive()) download_crsp()
