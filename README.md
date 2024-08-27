# Project description
This project aims to make routine Excel operations (data cleaning,
data conversion, data formatting, data joins) quicker and standarized.

## Main goals
- [ ] Be able to read and write Excel files 
  - [ ] list sheets, 
  - [ ] read data, 
  - [ ] write data
- [ ] Be able to clean data
  - [ ] remove empty rows
  - [ ] remove empty columns
  - [ ] ungroup rows
  - [ ] detect attribute columns and their names
- [ ] Be able to convert data
  - [ ] extract numeric data from text, e.g. '<10' -> 10 
  - [ ] give several choices of sample label extraction, e.g. 'E2142 22 Pole2' -> 'E2142' or '22' or 'Pole2', etc.
  - [ ] convert quality data to numeric data, e.g. 'bardzo_wysoka' -> 5
- [ ] Have an usable GUI
  - [ ] list sheets
  - [ ] list columns
  - [ ] list data before and after operations
  - [ ] list operations and their examples
  - [ ] list results
  - [ ] provide a way to select operations
  - [ ] provide a way to select columns
- [ ] Have a complete test suite
---
### Project side goals
- Focus on building by utilizing the SOLID principles
- Focus on building a complete test suite
- Focus on building a complete documentation
- Focus on building an easy Qt GUI
