<div align="center">
  <img src="assets/images/365AutomatedCheckCircle-200x200.png" alt="365AutomatedCheck Logo" width="300" height="300">
</div>

# 365AutomatedCheck

365AutomatedCheck is a tool to find non company standards using Pester tests and regular functions depending on your needs.

- [365AutomatedCheck](#365automatedcheck)
  - [Key Features](#key-features)
  - [Getting Started](#getting-started)
    - [Requirements](#requirements)
    - [Installation](#installation)
    - [Customize Validation Parameters](#customize-validation-parameters)
    - [Running Tests](#running-tests)
  - [Examples](#examples)
  - [Regular Functions](#regular-functions)
  - [Screenshots](#screenshots)

To view entire changelog [click here](changelog.md)

## Key Features

- Find and fix non compliant fields in Microsoft 365
- Easy to view HTML reports
- Add in your own Pester Tests
- And more to come..

The purpose of this module is two-fold, one, it is to make sure all users have company compliant values in their Microsoft 365 tenant. Two, find out if anyone within the company is not following company standards or even worse if a bad actor creates an account for bad intentions.

This is a community open source project and welcome PRs and feedback.

## Getting Started

### Requirements

- Pester 5
- ImportExcel
- ExchangeOnlineManagement
- Microsoft.Graph.Users
- Microsoft.Graph.Groups
- Microsoft.Graph.Identity.DirectoryManagement
- Microsoft.Graph.Users.Actions
- PSFramework


### Installation

```powershell
Install-Module -Name 365AutomatedCheck -Scope CurrentUser
```

### Customize Validation Parameters

Copy or update Excel workbook located at Assets/365ValidationParameters.xlsx to your company standards.

> Note: If you move the location of the file or rename it, you'll use that path when running Invoke-365AutomatedCheck (Invoke-365AutomatedCheck -ExcelFilePath "/Users/demo/Desktop/365ValidationParameters.xlsx")

> Note: If you have any empty values in a column, you will see "NA" as an option when updating for now. Working on a way so that isn't needed

### Running Tests

If you have configured your Excel workbook run:

```powershell
# This will export report to current directory /365ACReports/currentdate-currenttime
# If you haven't connected to graph do so now: Connect-MgGraph
Invoke-365AutomatedCheck
```

If you haven't configured your Excel workbook run:

```powershell
# This will export report to current directory /365ACReports/currentdate-currenttime
# If you haven't connected to graph do so now: Connect-MgGraph
Invoke-365AutomatedCheck -Tag Basic,SignIns -NoExcel $true
```

## Examples

Example 1: Check to see if all fields are filled out without using company standard Excel:

```powershell
Invoke-365AutomatedCheck -ExcludeTag "CompanyStandard" -NoExcel $true
```

Example 2: Change the Export path of html report

```powershell
invoke-365automatedcheck -OutputHtmlPath "/Users/Demo/Desktop/365Reports/testreport.html"
```

Example 3: Check Users last login

```powershell
Invoke-365AutomatedCheck -Tag SignIns -NoExcel $true
```

Example 4: Run tests to see in terminal with Excel validation in default path

```powershell
Invoke-365AutomatedCheck -Verbosity "normal"
```

Example 5: Run “Communication” tests to test if identities have a mobile phone, business phone, and fax number

```powershell
Invoke-365AutomatedCheck -tag "communication" -NoExcel $true
```

## Regular Functions

The regular functions are there for another way to test. They also will export to Terminal, HTML, or Excel. There will be more added in the future. Please let us know if you feel they are a benefit.

## Screenshots

View of successful tests:
![View of how successful test looks](assets/images/ghsuccessview.png)

View of failed tests:
![View of how failed test looks](assets/images/ghfailedview.png)

View of skipping test:
![View of how skipping a record looks](assets/images/ghskipview.png)

View of updating test:
![View of how updated record looks](assets/images/ghupdateview.png)