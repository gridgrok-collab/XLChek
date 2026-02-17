XLChek v0.1.0
Offline Excel Structural Risk Diagnostics

XLChek is an open-source, offline Excel formula analyzer and spreadsheet risk analysis tool that identifies structural risks in Excel workbooks.

It performs read-only Excel audits and generates a detailed HTML diagnostic report.

Designed for:

Financial model auditing

Spreadsheet risk diagnostics

Circular reference detection

Excel formula dependency analysis

Structural model validation

What XLChek Does

XLChek analyzes Excel workbooks for structural and dependency risks.

It detects:

Circular references

Hard-coded constants in formulas

Volatile Excel functions

Orphan calculations

Formula drift across ranges

Whole column / whole row references

External workbook references

XLChek does not validate business logic correctness.
It evaluates structural and dependency risk only.

Key Characteristics

100% offline

No internet required

No telemetry

No data upload

Read-only analysis

Generates HTML diagnostic report

Generates structured JSON output

Output

XLChek generates:

HTML risk report

JSON structured output

The report explains risk levels at:

Workbook level

Sheet level

Formula level

Requirements

Python 3.10 – 3.12

Windows / macOS / Linux

Tested on:

Python 3.12.7

Not yet tested on:

Python 3.13+

See requirements.txt for dependencies.

Installation

Install dependencies:

pip install -r requirements.txt


Run:

python -m sheetguard

Quick Start (Non-Technical Users)

Install Python 3.12 (Supported: 3.10–3.12)

Install dependencies

Run:

python -m sheetguard


Browse to your Excel file

Click Run Analysis

HTML report opens automatically

License

This project is licensed under the MIT License.
See LICENSE file for details.

Disclaimer

XLChek is provided “as is” without warranty of any kind.

Users are responsible for validating their own financial and operational models.

Version

v0.1.0 — Initial Public Release

Contact: gridgrok.contact@gmail.com

