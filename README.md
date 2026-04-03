# PRINT & S.P.A.R.K
### IT Deployment and Setup Automation Toolkit

A pair of PowerShell scripts designed to simplify **workstation setup** by automating **printer installation** and **core software deployment**.

---

## 📦 Included Tools

| Tool | Description |
|----|-----------|
| **PRINT** | Printer Registration & Installation Network Tool |
| **S.P.A.R.K** | Software Package & Resource Kit |

Each tool can be used independently or together during workstation provisioning.

---

## 🖨️ PRINT  
### Printer Registration & Installation Network Tool

PRINT automates printer driver installation and network printer creation using IP addresses.

### ✨ Features
- Supports **ZIP, EXE, and MSI** printer driver packages
- Automatically detects installer type
- Silent installs for EXE and MSI packages
- INF driver installation via `pnputil`
- Automatic TCP/IP printer port creation
- Interactive, technician‑guided workflow
- Standardized, log‑friendly console output
- Loops until valid driver files are provided

### 📁 Supported Driver Formats

| Format | Behavior |
|------|---------|
| ZIP | Extracts and installs INF drivers |
| EXE | Runs vendor installer silently |
| MSI | Installs using `msiexec` silently |

### ✅ Typical Usage
1. Download printer drivers from the manufacturer
2. Place the driver file in the same folder as `PRINT.ps1`
3. Run PRINT as **Administrator**
4. Install drivers and add network printers

---

## ⚡ S.P.A.R.K  
### Software Package & Resource Kit

S.P.A.R.K automates package manager setup and standard software deployment.

### ✨ Features
- Installs and initializes **Winget** and **Chocolatey**
- Automatically updates package managers
- Installs predefined core applications
- Optional software selected interactively
- Silent installs where supported
- Tracks installation results with timestamps
- Displays a clear installation summary

### 📦 Core Software Installed
(Default list – easily customizable)
- Microsoft Teams
- Microsoft Office
- 7‑Zip
- Google Chrome
- Adobe Acrobat Reader
- Zoom

### ➕ Optional Software Examples
- Zoom Outlook Plugin
- DisplayLink Graphics Driver
- Dell Command Update

---

## 🔁 How They Work Together

PRINT and S.P.A.R.K are designed to complement each other:

1. **S.P.A.R.K** prepares the workstation with core applications
2. **PRINT** installs printer drivers and configures printers
3. Both scripts:
   - Require Administrator privileges
   - Are interactive and technician‑friendly
   - Provide clear success and failure feedback

Scripts may be run independently or in sequence.

---

## ✅ Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Administrator privileges
- Internet access (required for S.P.A.R.K)

---

## 📂 Recommended Repository Structure

/PRINT
├─ PRINT.ps1
└─ ExtractedDrivers/

/SPARK
└─ SPARK.ps1
