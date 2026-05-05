# Subaward Processing Application

## Overview

This .NET console application reads Excel spreadsheets, extracts subrecipient names from the **"G. Other Direct Costs"** section, and calculates total subaward amounts across multiple files.

The output is formatted to be easily readable by non-technical users.

---

## Features

* Reads multiple Excel files from a folder
* Extracts subrecipient names in the format: `Subaward: Name`
* Displays subrecipients for each file
* Aggregates total subaward amounts across all files
* Supports variable number of rows
* Includes unit testing

---

## Technologies Used

* .NET 9 / .NET 10
* C#
* EPPlus (Excel processing)
* xUnit (unit testing)

---

## How to Run

### Run Application

cd SubawardApp
dotnet run

### Run Tests

cd SubawardTests
dotnet test

---
## Project Structure

```
SubawardExercise/
├── SubawardApp/
│   ├── Data/
│   │   ├── SubawardBudgetExample1.xlsx
│   │   ├── SubawardBudgetExample2.xlsx
│   │   └── SubawardBudgetExample3.xlsx
│   ├── Program.cs
│   └── SubawardApp.csproj
│
├── SubawardTests/
│   ├── UnitTest1.cs
│   └── SubawardTests.csproj
│
└── SubawardExercise.sln
```

## Assumptions

* Subrecipient names follow: "Subaward: Name"
* Data exists in the first worksheet
* Amount values are numeric
* Excel structure is consistent

---

## Questions

* Can file formats vary?
* Can multiple sheets exist?
* Should name matching be case-sensitive?

---

## Testing

A unit test verifies presence of:

* Indiana
* Mayo
* Purdue
* Florida
