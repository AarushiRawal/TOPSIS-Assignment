# TOPSIS Assignment – 102303601

This repository contains the complete implementation of the TOPSIS (Technique for Order Preference by Similarity to Ideal Solution) method as part of the Predictive Statistics (UCS654) course assignment.

---

## Part 1: CLI Program

- Implemented TOPSIS using Python
- Accepts Excel file as input
- Takes weights and impacts via command-line arguments
- Generates ranked output Excel file

**Folder:** `Part-1_CLI`

---

## Part 2: PyPI Package

- Converted TOPSIS program into a Python package
- Published on PyPI
- Supports CLI execution via installed command

**PyPI Link:**  
https://pypi.org/project/Topsis-Aarushi-102303601/

**Folder:** `Part-2_PyPI`

---

## Part 3: Web Service

- Flask-based web application
- Upload Excel file through web interface
- Validates inputs:
  - correct number of weights
  - correct number of impacts
  - comma-separated format
  - valid email format
- Executes TOPSIS algorithm
- Sends result file to user via email

**Folder:** `Part-3_Web_Service`

---

## Technologies Used

- Python 3
- Pandas
- NumPy
- Flask


---

## Author

**Aarushi Rawal**  
Course: **UCS654 – Predictive Statistics**
