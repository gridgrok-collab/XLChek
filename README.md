XLChek v0.1.0

Offline Excel Structural Risk Diagnostics



What XLChek Does

XLChek is an offline, read-only Excel analysis tool that identifies structural risks in spreadsheets.



It detects:

 	Circular references

 	Hard-coded constants

 	Volatile functions

 	Orphan calculations

 	Formula drift

 	Whole column/row references

 	External references



XLChek does NOT validate business logic correctness.

Tested: Python 3.12.7

Not yet tested: Python 3.13+



It evaluates structural and dependency risk.



Key Characteristics:

 	100% offline

 	No internet required

 	No telemetry

 	No data upload

 	Read-only analysis

 	Generates HTML diagnostic report



Requirements:

 	Supported Python 3.10 – 3.12

 	Windows / macOS / Linux



See requirements.txt

 	Install dependencies:

 	pip install -r requirements.txt

 	Run:

 	python -m sheetguard



Output:

XLChek generates:

 	HTML risk report

 	JSON structured output

 	The report explains risk levels at:

 	Workbook level

 	Sheet level

 	Formula level



Quick Start (For Non-Technical Users):



 	1. Install Python 3.12 (Supported version: Python 3.10 – 3.12)

 	2. Install dependencies:



 	pip install -r requirements.txt



 	3. Run:

 	python -m sheetguard



 	4. Browse to Excel file.

 	5. Click “Run Analysis”.

 	6. HTML report opens automatically.



License:

 	This project is licensed under the MIT License.

 	See LICENSE file for details.



Disclaimer:

XLChek is provided "as is" without warranty of any kind.

Users are responsible for validating their own financial and operational models.



Version:



 	v0.1.0 — Initial Public Release



Contact: XLChek.contact@gmail.com

