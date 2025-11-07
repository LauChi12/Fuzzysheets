# Fuzzysheets
A Python data utility for data reconciliation/cleanup using algorithmic fuzzy matching (Levenshtein) and a dual architecture for both local CSV I/O and secure Google Sheets API integration via OAuth 2.0.


## Overview
This tool connects to a Google Sheet, applies fuzzy matching (Levenshstein) to a target column, and outputs the matching rows. It features: 
  - Graphical User Interface (Tkinter)
  - Real-time status updates and log messages
  - Fuzzy matching with an adjustable similarity threshold
  - Automatic creation of a results tab in Google Sheets


## Technical Notes
This project was developed to explore and demonstrate: 
  - Fuzzy string matching algorithms and threshold tuning
  - Google Sheets API integration and handling OAuth authentication securely
  - GUI design with TKinter, including responsive threading for operations
  - Software packaging and distribution to ensure that non-technical users can run Python tools without installing dependencies

This project is a response to a common issue in data collection at my workplace: text input boxes make it difficult to group identical entries due to differences in nomenclature, spelling, and punctuation. 


## Fuzzy Matching Logic
Algorithm used: Levenshtein Distance (via Python/Rapidfuzz)
Purpose/Reasoning: Many issues with data collected from text input sources (ex: addresses) stem from misplaced punctuation (Rd. vs Rd) , slight spelling errors (Cypriss vs. Cypress) , and nomenclature differences (Drive vs Dr). I decided a fuzzy matching logic would be a good way to identify similar terms that mean the same thing. 

## Process: 
- Normalize input strings (lowercase, remove punctuation). 
- Compute Levenshtein distance between input and all entries in the identified column.
- Convert the distance to a similarity score (0-100).
- Return all rows with scores above the threshold. 

### Example:
  - Search String:
  - In-text String:
  - Threshold: 
  - Levenshtein Distance:
  - Score Output:
  - Result:

Possible Improvements: Using the Damerau-Levenshtein distance algorithm would be more sensitive to common mispellings: ie.,adjacent letter swaps. The matching logic would be more accurate using that algorithm. 

## Features
- Input Spreadsheet URL or name
- Specify the target column, search term, and similarity threshold (0-100)
- Automatic deletion of previous results sheet, if any.
- Runs in a separate thread to keep the UI responsive. 
- Color-coded log for progress and errors

## How to Run
### Windows
- Download ‘FuzzySheets_Windows.zip’ from Google Drive. 
- Unzip the folder.
- Double click ‘main.exe’ to launch the program. 

### Mac
- Download ‘FuzzySheets_Mac.zip’ from Google Drive.
- Unzip the folder.
- Double click ‘main’ to launch the program. 
You may see a security prompt. Go to System Preferences ->  Security & Privacy and allow the app. 

## First-Time Setup
- The program will ask you to log into your Google Account to access the sheet. 
- The authentication token is saved locally, so you only need to log in once. 

## Using the Program
- Enter your Google Sheet URL or name.
- Enter the sheet tab name (example: Sheet 1, Sheet 2)
- Enter the name of the Column to Filter (ex: Column A)
- Enter the search term.
- Enter the similarity threshold (0-100)
	Note: Threshold = 0 means everything matches. Threshold = 100 means exact matches only.
- Click ‘Run Filter’

The results will appear in a new sheet called ‘Fuzzy_Match_Results_OUTPUT’ .
Progress and messages are shown in the log box. 


IMPORTANT: Existing output tabs in the sheet (ie: from previous times you run the program) will be deleted when you click ‘Run Filter’. Make sure to CHANGE THE NAME of any output data you want to save. 

—
License
MIT License - see ‘LICENSE.txt’. 
