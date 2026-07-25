# Smart Attendance Reader

> AI-powered desktop application that converts handwritten attendance sheets into payroll-ready Excel files using Google Gemini Vision.

---

# Project Overview

Smart Attendance Reader is a desktop application developed in Python to automate monthly attendance processing.

Instead of manually reading attendance sheets and entering employee hours into Excel, the application sends the attendance sheet image directly to Google Gemini Vision, which extracts structured attendance data as JSON.

The extracted data is validated, processed, and exported into an Excel payroll template.

The project prioritizes simplicity, modularity, maintainability, and AI-first processing.

---

# Why This Project Exists

Current workflow:

Paper Attendance Sheet

↓

Manual Reading

↓

Manual Excel Entry

↓

Payroll Calculation

Problems:

- Time-consuming
- Human errors
- Repetitive work every month
- Difficult to maintain

The goal of this project is to automate the entire workflow.

---

# Workflow

Attendance Sheet Image

↓

Google Gemini Vision

↓

Structured JSON

↓

Data Processing

↓

Excel Writer

↓

Payroll Excel File

---

# Core Technologies

Language

- Python 3.12+

Desktop UI

- pywebview
- HTML
- CSS
- JavaScript

AI

- Google Gemini API
- gemini-flash-latest

Excel

- openpyxl

Version Control

- Git
- GitHub

---

# Project Structure

```
smart_attendance_reader/

├── main.py
│
├── gui/
│   ├── app_window.py
│   └── web/
│       ├── index.html
│       ├── style.css
│       └── app.js
│
├── ocr/
│   └── ocr_engine.py
│
├── processing/
│   └── data_extractor.py
│
├── excel/
│   ├── excel_writer.py
│   └── template_generator.py
│
├── templates/
│   └── attendance_template.xlsx
│
├── assets/
│
├── requirements.txt
│
└── .gitignore
```

---

# Module Responsibilities

## main.py

Application entry point.

Creates the application window and starts pywebview.

---

## gui/

Responsible for the desktop interface.

### app_window.py

Acts as the bridge between JavaScript and Python.

Responsibilities:

- receive requests from the frontend
- send images to Gemini
- save/load API key
- communicate with backend modules

---

### gui/web/

Contains the frontend.

- HTML
- CSS
- JavaScript

The frontend never communicates directly with Gemini.

All requests go through app_window.py.

---

## ocr/

Contains the AI extraction engine.

### ocr_engine.py

Responsibilities:

- load Gemini client
- send attendance image
- build prompt
- receive structured JSON
- validate response

No OCR libraries are used.

OpenCV and PaddleOCR are intentionally not part of this project.

---

## processing/

Converts Gemini JSON into structured Python objects.

Responsibilities:

- validate data
- normalize values
- prepare records for Excel

---

## excel/

Responsible for Excel generation.

### excel_writer.py

Writes employee data into Excel.

### template_generator.py

Creates payroll templates containing formulas.

---

## templates/

Contains reusable Excel templates.

---

## assets/

Application icons and static resources.

---

# Configuration

Application settings are stored locally.

```
~/.smart_attendance_reader/config.json
```

Currently stores:

- Gemini API Key

No user data is uploaded except the attendance image sent to Gemini.

---

# Current AI Model

Google Gemini

Model:

```
gemini-flash-latest
```

The model is responsible for:

- understanding attendance sheets
- extracting names
- extracting working hours
- returning structured JSON

---

# Data Flow

User

↓

Choose Image

↓

Python

↓

Gemini Vision

↓

JSON

↓

Data Processing

↓

Excel Writer

↓

Output Excel File

---

# Design Principles

The project follows these principles:

- modular architecture
- single responsibility
- maintainable code
- readable code
- scalable structure
- AI-first processing

---

# Development Rules

Every contributor (human or AI) should follow these rules.

1. Never break the existing architecture.

2. Keep each module responsible for one task only.

3. Do not duplicate code.

4. Explain major code changes.

5. Keep Python compatible with 3.12+.

6. Avoid unnecessary dependencies.

7. Prefer simple solutions.

8. Maintain clear folder organization.

9. Preserve API compatibility whenever possible.

10. Document important changes.

---

# Dependencies

Current dependencies:

- pywebview
- google-genai
- openpyxl

Removed dependencies:

- PaddleOCR
- PaddlePaddle
- OpenCV

---

# Current Status

Project Status:

🟡 Active Development

Current milestone:

Building the first working version that converts attendance images into Excel automatically using Gemini Vision.

---

# Author

Developer:

Yakoub

Country:

Algeria

Project:

Smart Attendance Reader
