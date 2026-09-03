# Employee Onboarding Automation (RPA)

> **7th Semester Robotic Process Automation (RPA) Project**  
> **Platform:** UiPath Studio (Modern Windows Target Framework)  
> **Integrations:** Microsoft Excel, Google Workspace (Gmail API via OAuth 2.0), Headless Chromium (Edge/Chrome)

---

## 1. Project Abstract & Motivation

Employee onboarding in enterprise environments is often a repetitive, time-consuming, and error-prone administrative task. Human Resources personnel must manually verify candidate records from spreadsheets, compile personalized appointment letters, export documents to PDF, create directory folders, and send individual welcome emails with attachments.

This project delivers an **unattended, fault-tolerant RPA software robot** designed to fully automate the end-to-end onboarding pipeline:
- Ingests structured candidate data from an Excel workbook (`Data/Employees.xlsx`).
- Validates data integrity (mandatory fields and RFC-compliant email syntax).
- Creates an organized employee directory structure (`Onboarding_Output/{EmpID}_{Name}/`).
- Generates a branded, print-ready, **single-page A4 PDF** welcome letter dynamically using an HTML/CSS template—with zero dependency on Microsoft Word or Office Interop.
- Dispatches personalized welcome emails with the PDF attached directly via the **Google Workspace (Gmail) API**.
- Implements **closed-loop state tracking & idempotency**: stamps `Status` and timestamps back into the Excel spreadsheet, automatically skipping already-processed records on subsequent runs.

---

## 2. Key Architecture & Features

### A. Modular Orchestrator-Worker Architecture
Separates orchestration concerns (`Main.xaml`) from individual business processing (`ProcessEmployee.xaml`), following standard enterprise RPA design principles for modularity, maintainability, and fault isolation.

### B. Idempotency & Closed-Loop Write-Back
To prevent duplicate processing or duplicate email deliveries if execution is interrupted, the robot:
1. Checks the `Status` column in Excel. If `Status == "Processed"`, the row is skipped.
2. Updates `Status` (`Processed` or `Failed: <reason>`) and `ProcessedDate` in memory.
3. Automatically persists the updated DataTable back to `Data/Employees.xlsx` upon completion using `excel:WriteRange`.

### C. Zero-Dependency, 1-Page PDF Generation Engine
Rather than relying on resource-intensive Microsoft Office COM Interop libraries, the bot uses native headless Chromium (`msedge.exe` / `chrome.exe`) CLI printing with `--no-pdf-header-footer` and `@page { size: A4 portrait; margin: 10mm 15mm; }`. This produces a crisp, strictly single-page corporate PDF without headers, footers, or third-party license costs.

### D. Non-Halting Validation & Exception Handling
- **Business Rule Failures**: Malformed records (missing fields, invalid email format) are classified as business exceptions. They log a `[SKIP]` warning, update Excel with the failure reason, and gracefully continue to subsequent records without halting the robot.
- **System Exception Boundary**: The worker sequence is enclosed in an isolated `TryCatch` block to prevent file system or network faults from crashing the outer loop.

### E. Modern Google Workspace Integration
Uses `UiPath.GSuite.Activities` (`SendEmailConnections`) powered by OAuth 2.0 Integration Service, ensuring secure, token-based cloud delivery without hardcoded SMTP passwords or insecure app passwords.

---

## 3. System Architecture & Diagrams

### System Architecture Diagram
![System Architecture](docs/architecture-diagram.svg)

---

### Process Control Flowchart
![Process Flowchart](docs/process-flowchart.svg)

---

## 4. Workflow Components & Arguments

### [`Main.xaml`](Main.xaml) (Orchestrator)
- **Role**: Master workflow that controls initialization, file I/O, looping, row state evaluation, and summary logging.
- **Variables**:
  - `dt_Employees`: System.Data.DataTable storing workbook records.
  - `templatePath`: Path to the HTML letter template (`Data\Template.html`).
  - `outputDir`: Directory for employee deliverables (`Onboarding_Output`).
  - `successCount`, `failedCount`, `alreadyProcessedCount`: Real-time telemetry counters.

### [`ProcessEmployee.xaml`](ProcessEmployee.xaml) (Worker)
- **Role**: Atomic workflow responsible for one employee's onboarding lifecycle.
- **Arguments**:

| Argument | Direction | Type | Description |
| :--- | :--- | :--- | :--- |
| `in_EmpID` | In | `String` | Employee identifier (e.g., `EMP001`) |
| `in_Name` | In | `String` | Candidate full name |
| `in_Email` | In | `String` | Candidate destination email address |
| `in_Department` | In | `String` | Assigned organizational department |
| `in_TemplatePath` | In | `String` | File path to `Template.html` |
| `in_OutputDir` | In | `String` | Base directory for output |
| `out_IsSuccess` | Out | `Boolean` | `True` if completely onboarded; `False` on skip/fail |
| `out_ErrorMessage`| Out | `String` | Validation or system error message |

---

## 5. Dataset Schema & Test Cases

The workbook [`Data/Employees.xlsx`](Data/Employees.xlsx) (Sheet: `Employees`) contains the following test suite designed to demonstrate positive runs and edge-case handling:

| EmpID | Name | Email | Department | Status (Initial) | Expected Result | Reason |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `EMP001` | John Doe | `achintydummy@gmail.com` | Engineering | *(Empty)* | **SUCCESS** | Valid record; PDF created and email sent |
| `EMP002` | Jane Smith | `achintydummy@gmail.com` | Human Resources | *(Empty)* | **SUCCESS** | Valid record; PDF created and email sent |
| `EMP003` | Alex Kumar | `achintydummy@gmail.com` | Marketing | *(Empty)* | **SUCCESS** | Valid record; PDF created and email sent |
| `EMP004` | Bad Data | `invalid-email-address` | Finance | *(Empty)* | **SKIP** | Missing `@` or `.` in email address |
| `EMP005` | *(empty)* | *(empty)* | Operations | *(Empty)* | **SKIP** | Missing mandatory ID, Name, and Email |
| `EMP006` | Sarah Connor | `achintydummy@gmail.com` | Cybersecurity | *(Empty)* | **SUCCESS** | Valid record; proves loop continues after skips |
| `EMP007` | Achintya | `achintydummy@gmail.com` | AI Engineering | *(Empty)* | **SUCCESS** | Valid record; customized verification |

---

## 6. Prerequisites & Execution Guide

### Prerequisites
1. **UiPath Studio** (v2024.10+ / v2026.x Community or Enterprise).
2. **Runtime Dependencies** (auto-restored via `project.json`):
   - `UiPath.Excel.Activities` (`3.6.1+`)
   - `UiPath.GSuite.Activities` (`3.11.10+`)
   - `UiPath.Mail.Activities` (`2.11.10+`)
   - `UiPath.System.Activities` (`26.6.3+`)
3. **Browser**: Microsoft Edge or Google Chrome installed in default Program Files location.
4. **Google Account**: Connected via UiPath Integration Service (OAuth 2.0).

### Execution Steps
1. Open **UiPath Studio**.
2. Click **Open a Local Project** and navigate to `RPA_Project/project.json`.
3. Open `Main.xaml` in the Designer.
4. Press **`F5`** (Run File) or click **Run**.
5. Observe execution in the **Output Panel**:
   ```text
   [Information] Starting Employee Onboarding Automation Process...
   [Information] Reading employee records from Data\Employees.xlsx (Sheet: Employees)...
   [Information] Total employee records to process: 7
   [Information] Processing onboarding for: EMP001 - John Doe
   [Information] [SUCCESS] Successfully onboarded John Doe (EMP001) - Welcome PDF and email dispatched.
   ...
   [Warning]     [SKIP] EMP004: Invalid or missing email format ('invalid-email-address')
   [Warning]     [SKIP] UNKNOWN: Mandatory fields missing (EmpID/Name/Department)
   ...
   [Information] [SUCCESS] Successfully onboarded Sarah Connor (EMP006) - Welcome PDF and email dispatched.
   [Information] ONBOARDING COMPLETE | Total: 7 | Newly Processed: 5 | Previously Processed (Skipped): 0 | Failed/Invalid: 2
   ```
6. Verify output files in `Onboarding_Output/` and delivered emails in your Gmail inbox.
