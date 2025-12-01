# GLPI Software Audit Automation Script

![Python](https://img.shields.io/badge/Python-3.x-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Issues](https://img.shields.io/github/issues/<your-username>/<repo-name>)

---

## Overview

This Python script automates the process of performing a **software audit** using your GLPI server. It generates a structured, auditor-friendly **Excel report** that helps IT teams quickly identify:

- ✔️ Installed **approved software**  
- ❌ Any **unauthorized or suspicious software**  
- 🧑‍💼 Associated **user and department** for each system  

The goal is to make IT audits, compliance checks, and software lifecycle management **faster, easier, and fully automated**.

---

## Key Features

### 🔍 Automatic Software Inventory
- Connects to GLPI via API  
- Fetches all active computers and their installed software  
- No manual extraction required  

### 🧠 Built-in Allowed/Excluded Software Lists
- Approved software is maintained inside the script  
- Excluded keywords filter irrelevant components (drivers, KB updates, etc.)  
- You can easily **add or remove allowed software** in the script to match your IT policy  

### 🧑‍💼 User–Department Mapping
- Configuration file: `config.xlsx`  
- Single sheet `UserDeptMap` with mapping of systems to users and departments  

**Example `config.xlsx`:**

| System Name  | User Name        | Department      |
|--------------|-----------------|----------------|
| RC-ADM-123   | Hafiz Sheheryar | Administration |
| RC-ADM-02    | Hafiz Sheheryar | Administration |
| RC-HR-01     | Ayesha Khan     | HR             |
| RC-IT-05     | Ali Raza        | IT             |

> Only the `System Name` must match GLPI computer names. `User Name` and `Department` are used in the Excel report.

### ⚠️ Unauthorized Software Detection
- Flags any installed software **not allowed** and **not excluded**  
- Generates a dedicated sheet for easy review by IT or audit teams

---

## 📘 Excel Report

The script generates a professional Excel file with two sheets:

### 1️⃣ Software Audit (Main Sheet)
- Columns: System Name, User Name, Department, Allowed Software  
- Shows ✓ for installed allowed software  
- Clean layout suitable for auditing

### 2️⃣ Unauthorized Software (Second Sheet)
- Columns: System Name, User Name, Department, Unauthorized Software  
- Each unauthorized software is listed per system  

---

## 🧩 Flow Diagram

```mermaid
flowchart TD
    A[Start Script] --> B[Load config.xlsx (UserDeptMap)]
    B --> C[Connect to GLPI API]
    C --> D[Fetch all active computers]
    D --> E[Fetch installed software per computer]
    E --> F[Clean and filter software]
    F --> G[Check allowed vs unauthorized software]
    G --> H[Populate Excel: Main Sheet & Unauthorized Sheet]
    H --> I[Format & style Excel]
    I --> J[Save Excel Report]
    J --> K[End GLPI Session]
    K --> L[Script Completed ✅]



