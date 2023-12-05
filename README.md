# Windows Browser Data to CSV
# Author: Tach Lehman
# Last Modified: December 5, 2023

 This python script was primarily thought up as a solution to missing functionality present in Windows File Analyzer
 which contained a non-functional button for grabbing the Index.dat of Internet Explorer (which no longer is
 supported and may as well not exist anymore). In this script, I have created a spiritual successor to that
 functionality which aims to pull the same kind of data from modern Windows web browsers. For this purpose,
 I have targeted the new default Windows browser, Microsoft Edge, and the most popular browser, Google Chrome.
 Accommodating both is fairly trivial as both are based on Chromium, and as such their file structures are much the
 same.
 
 This tool pulls browser data including History, Bookmarks, and Cookies from both Chrome and Edge and outputs the
 data to CSVs for each individual browser as well as combined data of both. This automates the collection of these
 browser artifacts and formats them in a way that can be easily searched via spreadsheet software like Microsoft
 Excel or Google Sheets.

# Usage:
python WBD.py
(CSV files are created in the same folder as the script)

# Dependencies:
Pandas (used for working with the data contained in the various files) - pip install pandas
