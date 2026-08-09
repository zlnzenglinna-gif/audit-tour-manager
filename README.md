# Audit Tour Manager

A Streamlit-based application for planning audit tours, managing hotel and travel information, calculating travel costs, and automatically generating Excel reports.

> **Demo Version:** All company names, persons, addresses, hotel information and costs used in this project are fictionalized for demonstration purposes.

## Overview

Audit Tour Manager was developed to improve an Excel-based workflow for organizing multi-day audit tours.

The original process required information to be maintained manually across several Excel files, including tour schedules, hotel records and travel cost calculations.

The application combines these steps into one workflow and reduces repetitive manual work.

## Key Features

### Tour Planning
- Upload an existing Excel Tourplan
- Read and display tour locations automatically
- Edit factories, hotels and travel information
- Add additional tour rows
- Record driving distance and travel time
- Export the updated Tourplan back to Excel

### Hotel Management
- Upload an existing hotel database
- Match factories with previously used hotels
- Use fuzzy name matching to handle different company-name formats
- Display historical hotel information and comments
- Record hotel prices, payment status and breakfast information

### Interpreter & Audit Period Management
- Create multiple interpreter periods
- Assign different interpreters to different date ranges
- Define the number of auditors for each period
- Handle additional company auditor participation
- Generate a separate cost report for each interpreter

### Automated Cost Calculation
The application automatically calculates:

- Hotel costs
- On-site hotel payments
- Driving distance
- Fuel costs
- Meal allowances
- Other costs / rounding
- Required travel budget

Fuel costs are calculated based on total driving distance, fuel consumption and fuel price.

Meal costs are calculated according to the number of travel days and participating persons.

The final budget can automatically be rounded using the "Sonstiges" position.

### Excel Automation
The application generates formatted Excel cost reports automatically.

The exported reports include:

- Interpreter name
- Travel period
- Hotel stays
- Hotel costs
- Fuel costs
- Meal costs
- Other costs
- Final total

Calculation formulas remain visible inside the generated Excel file for transparency.

### Save & Restore
Tour progress can be saved locally and restored when continuing work on the same tour.

## Workflow

1. Upload Tourplan
2. Upload hotel database
3. Edit and complete tour information
4. Review historical hotel recommendations
5. Enter distances and hotel costs
6. Define interpreter and auditor periods
7. Save the current work status
8. Export the updated Tourplan
9. Upload the cost-report template
10. Generate individual cost reports for each interpreter

## Tech Stack

- Python
- Streamlit
- pandas
- openpyxl
- RapidFuzz
- Excel

## Project Structure

```text
tour-planning-system/
│
├── app.py
├── services/
│   ├── cost_service.py
│   ├── excel_service.py
│   ├── hotel_service.py
│   ├── kostenerstellung_service.py
│   └── tourplan_service.py
│
├── utils/
├── sample_data/
├── screenshots/
├── requirements.txt
└── README.md

## Screenshots

### Application Overview
![Application Overview](screenshots/01_overview.png)

### Tour Planning
![Tour Planning](screenshots/02_tour_planning.png)

### Interpreter Period Management
![Interpreter Periods](screenshots/03_interpreter_periods.png)

### Automated Cost Report
![Cost Report](screenshots/04_cost_report.png)