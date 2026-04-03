PRINT & S.P.A.R.K
IT Deployment and Setup Automation Toolkit

A pair of PowerShell scripts designed to simplify workstation setup by automating printer installation and core software deployment in a consistent, technician‑friendly way.

Overview
This repository contains two complementary PowerShell tools:



Script	Purpose
PRINT	Installs printer drivers and configures network printers
S.P.A.R.K	Installs Winget and Chocolatey, then deploys core and optional software
They can be used independently or together during new workstation builds, rebuilds, or user onboarding.

PRINT
Printer Registration & Installation Network Tool

PRINT automates the installation of printer drivers and the creation of network printers using IP addresses.

Key Features
Supports ZIP, EXE, and MSI printer driver packages
Automatically detects installer type and installs accordingly
Silent installation for EXE and MSI drivers
Installs INF‑based drivers using pnputil
Creates TCP/IP printer ports automatically
Interactive, technician‑guided workflow
Standardized, log‑friendly console output
Loops until valid driver files are provided
Supported Driver Formats


Format	Behavior
ZIP	Extracts and installs INF drivers
EXE	Runs installer silently
MSI	Installs using msiexec silently
Typical Use Case
Download printer drivers from manufacturer
Place driver file in the same folder as the script
Run PRINT as Administrator
Install drivers and add printers in one workflow
S.P.A.R.K
Software Package & Resource Kit

S.P.A.R.K automates the installation of package managers and standard workstation software.

Key Features
Installs and initializes Winget and Chocolatey
Automatically updates package managers
Installs a predefined set of core applications
Optional software is selected interactively
Uses silent installs where supported
Tracks installation results with timestamps
Displays a clear installation summary at completion
Core Software Installed
(Default list – easily customizable)

Microsoft Teams
Microsoft Office
7‑Zip
Google Chrome
Adobe Acrobat Reader
Zoom
Optional Software Examples
Zoom Outlook Plugin
DisplayLink Graphics Driver
Dell Command Update
How the Scripts Work Together
PRINT and S.P.A.R.K are designed to complement each other:

S.P.A.R.K prepares the workstation by installing core applications
PRINT configures printers and drivers afterward
Both scripts:
Are interactive
Require Administrator privileges
Provide clear success and failure feedback
They can be run in any order depending on deployment needs.

Requirements
Windows 10 or Windows 11
PowerShell 5.1 or later
Administrator privileges
Internet access (for S.P.A.R.K)
Repository Structure (Recommended)


/PRINT
  PRINT.ps1
  ExtractedDrivers/
 /SPARK
  SPARK.ps1
 README.md
Driver files for PRINT should be placed in the same folder as PRINT.ps1.

Updating and Maintenance
Updating PRINT
Add new driver handling logic inside the driver installation function
Vendor‑specific silent switches can be expanded as needed
Logging or unattended mode can be added without breaking workflow
Updating S.P.A.R.K
Core software is defined in the $coreSoftware array
Optional software is defined in the $optionalSoftware array
Additional package managers can be added modularly
Silent install flags can be tuned per application
Both scripts are modular and intentionally structured to make future changes straightforward.

Known Limitations
Some EXE installers may not respect standard silent switches
Some software installations may require a reboot
Printer driver auto‑matching depends on vendor naming conventions
