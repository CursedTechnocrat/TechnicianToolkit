# M.A.G.I.C., S.P.A.R.K. & U.P.K.E.E.P.
### IT Deployment and Setup Automation Toolkit

A set of PowerShell scripts designed to simplify **workstation provisioning and maintenance** by automating:

- 🖨️ Printer installation
- ⚡ Core software deployment
- 🔄 Windows Update and system maintenance

---

## 📦 Included Tools

| Tool | Description |
|------|-------------|
| **M.A.G.I.C.** | Machine Automated Graphical Ink Configurator |
| **S.P.A.R.K.** | Software Package & Resource Kit |
| **U.P.K.E.E.P.** | Update Package Keeping Everything Efficiently Prepared |

Each tool can be used independently or together during workstation setup and lifecycle management.

---

## 🖨️ M.A.G.I.C.
### Machine Automated Graphical Ink Configurator

M.A.G.I.C. automates printer driver installation and network printer creation using IP addresses.

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
|--------|----------|
| ZIP | Extracts and installs INF drivers |
| EXE | Runs vendor installer silently |
| MSI | Installs using `msiexec` silently |

### ✅ Typical Usage

1. Download printer drivers from the manufacturer  
2. Place the driver file in the same folder as `MAGIC.ps1`  
3. Run **M.A.G.I.C.** as **Administrator**  
4. Install drivers and add network printers  

---

## ⚡ S.P.A.R.K.
### Software Package & Resource Kit

S.P.A.R.K automates package manager setup and standard software deployment.

### ✨ Features

- Installs and initializes **Winget** and **Chocolatey**
- Automatically updates package managers
- Installs predefined core applications
- Optional software selected via parameters or prompts
- Silent installs where supported
- Tracks installation results with timestamps
- Exports results to CSV for auditing
- Writes events to the Windows Event Log

### 📦 Core Software Installed

*(Default list – easily customizable)*

- Microsoft Office
- Microsoft Edge
- 7‑Zip
- Adobe Acrobat Reader
- Zoom

### ➕ Optional Software Examples

- Zoom Outlook Plugin
- Dell Command Update
- Dell Command Suite

---

## 🔄 U.P.K.E.E.P.
### Update Package Keeping Everything Efficiently Prepared

U.P.K.E.E.P. automates **Windows Update detection, installation, and reboot handling** while ensuring the system remains awake and stable during the update process.

### ✨ Features

- Disables system sleep and monitor timeouts during execution
- Automatically installs and imports the **PSWindowsUpdate** module
- Scans for available Windows updates *(drivers excluded)*
- Displays a clear summary of pending updates
- Installs updates **without forcing an automatic reboot**
- Detects whether a reboot is required
- Prompts the technician before rebooting
- Provides a **30‑second reboot countdown with cancel option**
- Standardized output aligned with M.A.G.I.C. and S.P.A.R.K.

### ✅ Typical Usage

1. Run `UPKEEP.ps1` as **Administrator**
2. Allow the script to install and apply Windows updates
3. Review update results
4. Reboot immediately or defer when prompted

### ⚠️ Notes

- Internet access is required to download updates and modules
- Unsaved work should be closed before execution
- Designed for technician‑initiated maintenance or post‑deployment cleanup

---

## 🔁 How They Work Together

M.A.G.I.C., S.P.A.R.K., and U.P.K.E.E.P. are designed to complement each other across the workstation lifecycle:

1. **S.P.A.R.K.** installs and upgrades core software  
2. **M.A.G.I.C.** installs printer drivers and configures printers  
3. **U.P.K.E.E.P.** ensures Windows is fully updated and reboot‑clean  

All scripts:

- Require **Administrator privileges**
- Use a **consistent color and banner format**
- Provide clear success, warning, and error feedback
- Can be run independently or in sequence

---

## ✅ Requirements

- Windows 10 or Windows 11
- PowerShell **5.1 or later**
- Administrator privileges
- Internet access *(required for S.P.A.R.K. and U.P.K.E.E.P.)*

---

## 📂 Recommended Repository Structure

```text
/MAGIC
├─ MAGIC.ps1
└─ ExtractedDrivers/

/SPARK
└─ SPARK.ps1

/UPKEEP
└─ UPKEEP.ps1
