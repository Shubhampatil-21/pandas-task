# Pandas Excel Task

This script performs data processing on an Excel file using Python and pandas.

## Features

- Removes duplicate rows based on `Feed Name`
- Merges sheets on `Client` and `Feed Name`
- Calculates:
  - Addition of Import and Export product counts
  - Subtraction of Import and Export product counts
  - Multiplication of Import and Export product counts
  - Division of Import by Export product count (avoids divide-by-zero)
- Sorts data by `Export Product Count Today` in descending order
- Saves the result to a new sheet named `Sheet3`