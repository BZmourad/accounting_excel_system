"""Column widths, row heights, freeze panes helpers."""

STANDARD_ROW_HEIGHT = 22
HEADER_ROW_HEIGHT   = 28
TITLE_ROW_HEIGHT    = 38
SUBTITLE_ROW_HEIGHT = 22
KPI_ROW_HEIGHT      = 60

# Column widths per sheet type
WIDTH_TX = {          # transaction sheets (revenue/expense)
    "A": 6,   # #
    "B": 18,  # institution
    "C": 22,  # category
    "D": 22,  # customer/desc
    "E": 30,  # description
    "F": 14,  # amount
    "G": 12,  # due date
    "H": 12,  # pay/collect date
    "I": 16,  # status
    "J": 16,  # bank
    "K": 24,  # notes
}

WIDTH_VEHICLES = {
    "A": 6,
    "B": 22,  # customer
    "C": 22,  # organization
    "D": 14,  # type
    "E": 10,  # count
    "F": 14,  # vehicle number
    "G": 16,  # plate
    "H": 14,  # delivery date
    "I": 16,  # monthly rental
    "J": 16,  # maintenance
    "K": 16,  # net rental
    "L": 18,  # contract status
    "M": 24,  # notes
}