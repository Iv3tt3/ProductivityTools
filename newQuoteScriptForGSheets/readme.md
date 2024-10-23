# Script for GSheets: New Spreadsheet Generator with Selected Sheets

This Google Sheets script automates the process of creating new spreadsheets with specific selected sheets, tailored for a company's budgeting workflow.

The script enables sales team to generate budget templates dynamically based on different products and their specific conditions. By utilizing this system, sales teams can work efficiently without the need to manage multiple static templates, saving time and reducing the chance of errors.

## Problem Statement

The purpose of this script is to streamline and simplify the budget creation process for the sales team by automating it. This reduces the likelihood of human errors and enhances both efficiency and productivity.

### Problem

The company offers a variety of products, each with its own unique budget conditions. The traditional method of managing multiple templates for each product becomes cumbersome and prone to errors. The sales team requires a streamlined process that:

- Ensures different budget conditions are applied for each product.
- Simplifies the work by using a single template.
- Avoids the risk of human error.

### Solution

This script addresses the problem by providing the following workflow:

The salesperson fills out a main "Budget" sheet where they enter customer information. They select the type of budget they require via checkboxes.
By pressing a button in the sheet, a new Google Sheets file is automatically generated with the appropriate pre-selected sheets that match the type of budget.

- The new spreadsheet is linked back to the main sheet "Budget" with its URL for easy access.
- The salesperson can open and edit the new budget sheet without altering the original template.
- Only the necessary sheets are copied into the new file, facilitating ease of use and streamlined printing.
- Fields that should not be modified by the sales representative are protected, avoiding accidental changes.
- Google Sheets' built-in validation ensures proper functionality and reduces the risk of errors.

## How it works

## Usage
