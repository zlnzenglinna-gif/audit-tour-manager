# Audit Tour Manager

A business process automation tool for planning multi-day audit tours and generating travel cost reports.

The application transforms a previously Excel-based manual workflow into a structured process for tour planning, hotel management, interpreter assignment, travel distance tracking and automated cost reporting.

> **Demo Version:** All company names, persons, addresses, hotel information and costs used in this repository are fictionalized for demonstration purposes.

## Business Problem

Audit tour planning required information to be maintained manually across multiple Excel files. Tour schedules, historical hotel information, interpreter assignments, travel distances and cost calculations had to be checked and updated separately.

This created repetitive manual work and increased the risk of inconsistent information and calculation errors.

## Solution

Audit Tour Manager combines these activities into one Streamlit-based workflow.

The application allows users to:

- Upload and process existing Excel tour plans
- Match factories with previously used hotels
- Add and edit hotel information directly in the planning workflow
- Manage multiple interpreter assignment periods
- Record travel distances and hotel payment information
- Automatically calculate fuel and meal costs based on business rules
- Generate updated Tourplan Excel files
- Generate standardized Kostenerstellung Excel reports

## Business Value

The project demonstrates how an existing operational Excel workflow can be analyzed, structured and partially automated.

The solution reduces repetitive data entry, centralizes information from multiple Excel sources and applies predefined business rules consistently when generating operational reports.

## Tech Stack

- Python
- Streamlit
- pandas
- openpyxl
- RapidFuzz
- Excel
- Git / GitHub

## Project Structure

```text
audit-tour-manager/
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
```

## Screenshots

### Application Overview
![Application Overview](screenshots/01_overview.png)

### Tour Planning
![Tour Planning](screenshots/02_tour_planning.png)

### Interpreter Period Management
![Interpreter Periods](screenshots/03_interpreter_periods.png)

### Automated Cost Report
![Cost Report](screenshots/04_cost_report.png)