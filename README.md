# Asset Request Tool

This script generates personalized email address to aid in the collection of hardware information

# Features

- Takes in an Excel spreadsheet and parses it using pandas and openpyxl
- Generates email templates based on login data and machine data from the Excel data
  - Standard email which includes a link to a form
  - Alternate link if a duplciate user is found, which manually asks for information
- Writes back the personalized email to the Excel sheet, preserving formatting thanks to openpyxl 
