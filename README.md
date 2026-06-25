# Email Order Processing Automation RPA

A UiPath robotic process automation project that processes online orders received by email and updates a local inventory restock spreadsheet. Built with UiPath Studio Community, integrated with Outlook Classic and Microsoft Excel on Windows 11.

**Systems Engineering Project — FILS UPB, Year IV Semester I**

**Team:** SAKKA Mohamad Mario · MAHMOUD MIRGHANI Abdelrahman · ZAFAR Azzam · AL KHALIDY Essam

---

## Overview

Many companies receive a mix of B2B and retail orders by email. Even when the order content is simple, the processing steps are repetitive: a staff member must identify order emails, copy order lines into an internal file, download and store attachments, calculate totals, decide a delivery estimate, reply to the customer, and keep the inbox organized.

This project automates that end-to-end workflow. The robot reads a configured Outlook inbox, detects order emails by subject keywords, parses order lines from three supported body formats, computes totals and delivery estimates, saves attachments into structured folders, updates the restock spreadsheet, sends a confirmation email, moves the email to an outcome folder, and appends an audit record to a CSV log.

---

## What Is Automated

**In scope:**
- Reading a configured Outlook inbox
- Detecting candidate order emails using subject keywords and an unread filter
- Extracting order lines from three supported email body formats
- Computing order total, item count, and delivery estimate
- Saving attachments into a local folder structure
- Updating the restock spreadsheet on a local PC
- Sending a confirmation email with the estimated delivery date and item summary
- Moving emails into one of four outcome folders (Processed, HighPriority, NeedInfo, Failed)
- Logging processing outcomes into a CSV file

**Out of scope:**
- Payment processing and invoicing systems
- Warehouse dispatch and shipping carrier integration
- Customer account management and CRM integration
- Automatic inventory decrement and real-time stock validation

---

## Technologies

| Component | Detail |
|---|---|
| Automation platform | UiPath Studio Community 2026.0.182 STS |
| Email integration | Microsoft Outlook Classic |
| Spreadsheet integration | Microsoft Excel (XLSX) |
| Operating system | Windows 11 |
| Mailbox account | orderemailsautomationrpa@outlook.com |

---

## How It Works

### 1. Initialization and settings

At startup the robot loads settings from a settings file editable via a settings form. Settings include `AccountEmail`, subject keywords, `OnlyProcessUnreadEmails`, `MaxEmailsToRead`, `RunTime`, delivery thresholds, delivery days, and local file paths. It validates that configured paths exist and ensures that the Outlook folders Processed, HighPriority, NeedInfo, and Failed are present.

### 2. Email retrieval and detection

The robot retrieves unread emails from the configured inbox (up to `MaxEmailsToRead` per run, default 50). For each email it checks whether the subject contains any keyword from `SubjectKeywords` (defaults: `order`, `purchase`, `invoice`, `PO`). Non-matching emails are skipped.

### 3. Order parsing — three supported formats

| Format | Description |
|---|---|
| **Labeled** | Each line contains explicit `Name`, `Qty`, `Price` labels with semicolon separators |
| **CSV style** | Three comma-separated values per line: `Product, Quantity, UnitPrice` |
| **Pipe separated** | Three pipe-separated values per line: `Product \| Quantity \| UnitPrice` |

For all formats the robot validates that quantity is an integer and unit price is a decimal. If any required value is missing or invalid the email is routed to NeedInfo.

### 4. Business rules and delivery estimate

After parsing:
- `LineTotal` = Quantity × UnitPrice per line
- `TotalPrice` = sum of all LineTotals
- `ItemCount` = number of extracted lines
- `OrderID` = generated from the processing timestamp

**Priority classification:** `HighPriority` is true when `TotalPrice > FastDeliveryMinPrice` AND total quantity ≤ `FastDeliveryMaxQuantity`.

**Estimated delivery date:**
```
if HighPriority         → FastDaysDeliveryTime
else if Qty > SlowDeliveryMinQuantity → SlowDaysDeliveryTime
else                    → StandardDaysDelivery
ETA = today + delivery days
```

### 5. Attachments

When the email has attachments the robot creates a dedicated folder under `AttachmentsFolderPath` named with the OrderID and saves all files into it. If the save fails (permissions, invalid path) the email is routed to Failed.

### 6. Restock spreadsheet update

The robot opens the XLSX file at `RestockExcelPath`. For each parsed order line it searches the Product column: if a match exists it adds the quantity to `TotalQty`; if not it appends a new row. If the file is locked or unwritable the email is routed to Failed.

### 7. Confirmation email

For successfully processed orders the robot sends a reply to the sender with: estimated delivery date, delivery days used, item count, order total, and a full line summary including line totals. Subject format: `Order received - estimated delivery YYYY-MM-DD`.

For NeedInfo emails it sends a reply requesting the missing information and suggesting one of the supported formats.

### 8. Outcome routing and audit log

Each email is moved to one of four Outlook folders:

| Folder | Condition |
|---|---|
| **HighPriority** | HighPriority rule is true and processing succeeded |
| **Processed** | Processing succeeded and HighPriority is false |
| **NeedInfo** | Parsing or validation failed |
| **Failed** | Unexpected processing error |

One CSV row is appended to `OrdersLogPath` per email: `Timestamp, OrderID, SenderEmail, EmailSubject, Status, AttachmentCount, AttachmentFolderPath`.

---

## Configuration

All operator-adjustable settings are separated from workflow logic and editable through the settings form without modifying any workflow files:

| Setting | Description |
|---|---|
| `AccountEmail` | Outlook inbox account to monitor |
| `SubjectKeywords` | Comma-separated keywords for order detection |
| `OnlyProcessUnreadEmails` | Process unread emails only (default: true) |
| `MaxEmailsToRead` | Maximum emails per run (default: 50) |
| `RunTime` | Scheduled daily run time (default: 23:45) |
| `FastDeliveryMinPrice` | Minimum order value for fast delivery |
| `FastDeliveryMaxQuantity` | Maximum quantity for fast delivery |
| `FastDaysDeliveryTime` | Days for fast delivery |
| `SlowDeliveryMinQuantity` | Minimum quantity threshold for slow delivery |
| `SlowDaysDeliveryTime` | Days for slow delivery |
| `StandardDaysDelivery` | Days for standard delivery |
| `AttachmentsFolderPath` | Root folder for saved attachments |
| `RestockExcelPath` | Path to the restock XLSX spreadsheet |
| `OrdersLogPath` | Path to the CSV audit log |

---

## Requirements

### Operational
- **OR1** Windows 11 workstation with Outlook Classic and Microsoft Excel installed
- **OR2** Configured Outlook mailbox; replies sent from the same account
- **OR3** Daily schedule configured by the operator (default 23:45)
- **OR4** Process at most 50 emails per run by default (configurable)
- **OR5** Process only unread emails by default (configurable)

### Functional (key requirements)
- **FR1–FR2** Read inbox and identify order emails by subject keywords
- **FR4–FR5** Extract order lines in three formats; extract ProductName, Quantity, UnitPrice per line
- **FR6–FR8** Compute LineTotal, TotalPrice, ItemCount
- **FR9** Generate unique OrderID from processing timestamp
- **FR10–FR11** Classify HighPriority; compute delivery estimate
- **FR12–FR14** Save attachments; update or create rows in restock spreadsheet
- **FR15** Send confirmation email with delivery date, item count, order total, line summary
- **FR16** Move email to one of four outcome folders
- **FR17** Append one CSV log record per processed email
- **FR18** Route to NeedInfo and send clarification request when parsing fails
- **FR19** Route to Failed on unexpected error

### Non-functional
- **NFR1** Process each email in 30 seconds or less under normal conditions
- **NFR2** At least 95% successful runs (workflow completes without crashing)
- **NFR3** CSV log allows tracing each email to a timestamp, order ID, folder status, and attachment path
- **NFR4** No full email body content stored in any persistent file
- **NFR5** Attachments folder, restock spreadsheet, and orders log stored in a restricted local directory
- **NFR6** All keywords, thresholds, delivery days, file paths, and schedule time changeable via settings form without workflow edits

---

## Project Structure

```
Order-Email-Processing-Automation-RPA/
├── Order-Email-Processing-Automation-RPA/   # UiPath project files
└── Docs/
    ├── Email_Order_Processing_Automation_Report.pdf   # Full technical report
    ├── Email_Order_Processing_Automation_Report.docx
    ├── Email Order Processing Automation RPA.pptx
    ├── Email Order Processing Automation RPA - Risks Log.xlsx
    ├── Diagrams/                                       # SysML diagrams
    └── SysEngProjDemo.zip
```

---

## System Design (SysML)

The system was modeled using SysML with two structure diagrams (Block Definition Diagram and Internal Block Diagram), two behavior diagrams (Use Case and Sequence), and a Requirements Diagram tracing key requirements to main blocks.

**Main blocks:** `OutlookMailbox`, `OrderEmail`, `Order`, `OrderLine`, `DeliveryPolicy`, `RestockSpreadsheet`, `AttachmentStore`, `AuditLog`

**Core modules:** Settings, Parser, Rules Engine, Outlook Adapter, Attachment Handler, Restock Updater, Audit Logger

---

## Quality Assurance

Two quality metrics were defined and measured over a two-week observation window (50 emails read, ~40 true order emails):

**Metric 1 — Order Extraction Accuracy**
> Number of order emails where all extracted fields are correct for every line, divided by total true order emails processed. An email is correct only when ProductName, Quantity, and UnitPrice match ground truth for every line.

**Metric 2 — Restock Update Fail Rate**
> Number of restock file update attempts that fail, divided by total update attempts. A failure includes: cannot open file, cannot write changes, file locked, path invalid, or any exception during the update workflow.

---

## Risk Management

Eleven risks were identified and managed. Major risks:

| ID | Risk | Strategy |
|---|---|---|
| OA-R001 | False positives (non-order emails match keywords) | Mitigation |
| OA-R002 | False negatives (real orders missed by limited keywords) | Mitigation |
| OA-R003 | Email body format differs from supported patterns → NeedInfo spike | Mitigation |
| OA-R004 | Incorrect parsing silently produces wrong totals and restock updates | Avoidance/Mitigation |
| OA-R005 | Duplicate processing causes double restock increments | Avoidance/Mitigation |
| OA-R006 | Attachment filename collisions cause overwrite | Mitigation |
| OA-R007 | Restock spreadsheet locked or schema changed | Mitigation |
| OA-R008 | Misconfigured delivery-days rule generates incorrect ETAs | Mitigation |
| OA-R009 | Mailbox foldering fails, inbox cluttered | Mitigation |
| OA-R010 | Log record missing or does not reflect final status (audit gap) | Avoidance/Mitigation |
| OA-R011 | Attachments or logs stored in insecure location | Avoidance/Mitigation |

Mitigations include: configuration-driven keyword lists, `OnlyProcessUnreadEmails` by default, numeric validation of extracted values, per-email isolation so one failure does not break the full run, NeedInfo routing for uncertain cases, deterministic OrderID-based folder naming, and consistent log records.

---

## Team Roles

| Member | Role |
|---|---|
| MAHMOUD MIRGHANI Abdelrahman | Project manager and coordinator; UiPath development |
| ESSAM AL KHALIDY | Business analyst; requirements and stakeholder analysis; report writing |
| ZAFAR Azzam | SysML modeling and diagram production |
| SAKKA Mohamad Mario | UiPath development and quality assurance; presentation |

Project timeline: 1 November 2025 – 24 December 2025
