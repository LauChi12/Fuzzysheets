
# Fuzzysheets
A Python data utility for data reconciliation/cleanup using algorithmic fuzzy matching (Levenshtein) and a dual architecture for both local CSV I/O and secure Google Sheets API integration via OAuth 2.0.

## **Overview**

This tool connects to a Google Sheet or CSV/Excel File (in local I/O version), applies fuzzy matching (Levenshstein) to a target column, and outputs the matching rows. It features: 
  - Graphical User Interface (Tkinter)
  - Real-time status updates and log messages
  - Fuzzy matching with an adjustable similarity threshold
  - Automatic creation of a results tab in Google Sheets, or a new output CSV file. 


## Technical Notes
This project was developed to explore and demonstrate: 
  - Fuzzy string matching algorithms and threshold tuning
  - Google Sheets API integration and handling OAuth authentication securely
  - GUI design with TKinter, including responsive threading for operations
  - Software packaging and distribution to ensure that non-technical users can run Python tools without installing dependencies

This project is a response to a common issue in data collection at my workplace: text input boxes make it difficult to group identical entries due to differences in nomenclature, spelling, and punctuation. 


## Fuzzy Matching Logic
Algorithm used: Levenshtein Distance (via Python/Rapidfuzz)
Purpose/Reasoning: Many issues with data collected from text input sources (ex: addresses) stem from misplaced punctuation (Rd. vs Rd) , slight spelling errors (Cypriss vs. Cypress) , capitalization differences (RD vs. rd vs. Rd), and nomenclature differences (Drive vs Dr). I decided a fuzzy matching logic would be a good way to identify similar terms that mean the same thing. I chose to implement tools from the library 'rapidfuzz' for ease and speed of processing and automatic data cleanup(removing capitalization and punctuation).  

### Process: 
- Normalize input strings (lowercase, remove punctuation). 
- Compute Levenshtein distance between input and all entries in the identified column.
- Convert the distance to a similarity score (0-100).
- Return all rows with scores above the threshold. 

#### Example:
	Search String: "Main Ave"
	In-text String: "Main Av"
	Threshold: 80
	Levenshtein Distance:85.71
	Score Output:86
	Result: PASS The in-text string would be output to the new sheet.

Possible Improvements: Using the Damerau-Levenshtein distance algorithm would be more sensitive to common mispellings: ie.,adjacent letter swaps. The matching logic would be more accurate using that algorithm. However, I'd have to either use the slower ``fuzzywuzzy`` instead of ``rapidfuzz`` or encode the algorithm in the program itself, which is resourse intensive. 


## Features
- Select input file with a file search button
- Specify the target column, search term, and similarity threshold (0-100)
- Automatic deletion of previous results sheet, if any.
- Runs in a separate thread to keep the UI responsive. 
- Color-coded log for progress and errors

## How to Run
### Windows(not out yet)
- Download ``FuzzySheets_Windows.zip``. 
- Unzip the folder.
- Double click ``main.exe`` to launch the program. 

## #Mac
- Download ``FuzzySheets_Mac.zip``.
- Unzip the folder.
- Double click ``main`` to launch the program. 
You may see a security prompt. Go to System Preferences ->  Security & Privacy and allow the app. 

## First-Time Setup
-
-

## Using the Program
- Click Search button to search for the input file.
- Enter the name of the Column to Filter (ex: Column A)
- Enter the search term.
- Enter the similarity threshold (0-100)
	Note: Threshold = 0 means everything matches. Threshold = 100 means exact matches only.
- Click ``Run Filter``

The results will appear in a new sheet on your desktop called ``Fuzzy_Match_Results_OUTPUT``.
Progress and messages are shown in the log box. 


![Warning Sign](https://www.shareicon.net/data/16x16/2016/05/28/572061_warning_32x32.png)
***IMPORTANT:*** Existing output tabs in the sheet (ie: from previous times you run the program) will be deleted when you click ``Run Filter``. Make sure to **CHANGE THE NAME/FILE LOCATION** of any output data you want to save. ![Warning Sign](https://www.shareicon.net/data/16x16/2016/05/28/572061_warning_32x32.png)


## License
MIT License - see ``LICENSE.txt``. 

