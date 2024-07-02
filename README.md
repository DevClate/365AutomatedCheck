# 365AutomatedCheck

365 Automated Checking tool to find non company standards

## Key Features

- Find and fix non compliant fields in Microsoft 365
- Easy to view HTML reports
- Add in your own Pester Tests
- And more to come..

The purpose of this module is two-fold, one, it is to make sure all users have company compliant values in their Microsoft 365 tenant. Two, find out if anyone within the company is not following company standards or even worse if a bad actor creates an account for bad intentions.

This is a community open source project and welcome PRs and feedback.

## Getting Started

### Installation

```powershell
Install-Module -Name 365AutomatedCheck -Scope CurrentUser
```

### Customize Validation Parameters

>Copy or update Excel workbook located at Assets/365ValidationParameters.xlsx to your company standards

Example: Runs everything with Excel Validation
Invoke-365AutomatedCheck

Example: Check to see if all fields are filled out without using company standard Excel:
Invoke-365AutomatedCheck -ExcludeTag "CompanyStandard"

Notes:

- When validating with Excel it with have an option for NA, that is for right now as I couldn't get it to remove empty options without doing that.
