INVOICE'S STATUS UPDATE PLAN
|
|-- SOURCE
|   |
|   |-- Read every input worksheet
|   |
|   `-- Filter: [Ref Type] = "IN"
|
|-- PREPARE
|   |
|   |-- [Acct]
|   |-- [Date] -> YYYY        -> dynamic year headers
|   `-- [Total]               -> amount to sum
|
|-- GROUP
|   |
|   |-- Acct
|   `-- YEAR(Date)
|
|-- CALCULATE
|   |
|   `-- SUM([Total])
|       for each:
|       Acct  + Year
|
|-- CREATE "Invoice's Status"
|   |
|   |-- Amazon title / invoice summary format
|   |
|   |-- Dynamic years
|   |   2019 | 2020 | ... | 2026
|   |
|   |-- Per account section
|   |   |
|   |   |-- AM9285 - AMAZON (Open IN's)
|   |  
|   |   |-- distinct Order# rows
|   |   |-- year amounts
|   |   `-- ACCOUNT TOTAL
|   |
|   `-- Total Amount
|
`-- CLEANUP RULES
    |
    |-- Skip an account section if its final Total Amount = 0.00
    |-- Remove unused formatting after the final visible account
    `-- Do not create empty account Invoice sheets