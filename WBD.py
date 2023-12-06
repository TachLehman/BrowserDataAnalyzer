#
# Windows Browser Data to CSV
# Author: Tach Lehman
# Last Modified: December 5, 2023
#
# This python script was primarily thought up as a solution to missing functionality present in Windows File Analyzer
# which contained a non-functional button for grabbing the Index.dat of Internet Explorer (which no longer is
# supported and may as well not exist anymore). In this script, I have created a spiritual successor to that
# functionality which aims to pull the same kind of data from modern Windows web browsers. For this purpose,
# I have targeted the new default Windows browser, Microsoft Edge, and the most popular browser, Google Chrome.
# Accommodating both is fairly trivial as both are based on Chromium, and as such their file structures are much the
# same.
# 
# This tool pulls browser data including History, Bookmarks, and Cookies from both Chrome and Edge and outputs the
# data to CSVs for each individual browser as well as combined data of both. This automates the collection of these
# browser artifacts and formats them in a way that can be easily searched via spreadsheet software like Microsoft
# Excel or Google Sheets.
#

import os
import sqlite3
import pandas
import shutil
import datetime


# convert_timestamp - converts timestamps from chromium data format to a readable time format
def convert_timestamp(timestamp):
    try:
        # Convert the timestamp string to an integer
        timestamp_int = int(timestamp)
        # Chromium timestamp conversion
        epoch_start = datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=timestamp_int)
        # Return readable date format
        return epoch_start.strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        # Return the original timestamp if conversion is not possible
        return timestamp


# extract_history_cookies - extracts the history and cookies. Both are stored in sqlite databases, so can be used for
# both.
def extract_history_cookies(browser_path, query, columns_to_convert=None):
    # Check if the file even exists
    if not os.path.exists(browser_path):
        print(f"File not found: {browser_path}")
        return pandas.DataFrame()

    # try-catch used to handle potential errors during data ingest
    try:
        # Backup file used to avoid issues with the database being open
        backup_path = browser_path + '.backup'
        shutil.copyfile(browser_path, backup_path)

        # Connect to the backed up database
        with sqlite3.connect(backup_path) as conn:
            # feed SQL query results into a pandas data frame
            df = pandas.read_sql_query(query, conn)

            # Check for timestamp columns that need converted from the chromium timestamp format, convert them
            if columns_to_convert:
                for column in columns_to_convert:
                    df[column] = df[column].apply(convert_timestamp)

            # Return the data frame
            return df

    except sqlite3.Error as e:
        print(f"Error reading from {browser_path}: {e}")
        return pandas.DataFrame()


# extract_bookmarks - extracts bookmarks from the relevant JSON file
def extract_bookmarks(bookmarks_path):
    # Check if the file even exists
    if not os.path.exists(bookmarks_path):
        print(f"Bookmarks file not found: {bookmarks_path}")
        return pandas.DataFrame()

    # Try-catch to catch potential errors
    try:
        # Open the bookmarks file
        with open(bookmarks_path, 'r', encoding='utf-8') as file:
            # Read contents with pandas
            bookmarks_json = pandas.read_json(file)

            # Check if the necessary data exists in the JSON structure (the file can exist, but be empty)
            if 'roots' in bookmarks_json and 'bookmark_bar' in bookmarks_json['roots'] and 'children' in bookmarks_json['roots']['bookmark_bar']:
                bookmarks_data = pandas.json_normalize(bookmarks_json['roots']['bookmark_bar']['children'])

                # Check if 'date_added' column is available for conversion
                if 'date_added' in bookmarks_data.columns:
                    bookmarks_data['date_added'] = bookmarks_data['date_added'].apply(convert_timestamp)

                # Return processed bookmark data
                return bookmarks_data[['name', 'type', 'url', 'date_added']]
            else:
                print("Bookmark data is missing or in an unexpected format.")
                return pandas.DataFrame()

    except Exception as e:
        print(f"Error reading from {bookmarks_path}: {e}")
        return pandas.DataFrame()


# Path to current user, used in locating the relevant AppData folder
user_path = os.path.expanduser('~')

# Chrome paths - History and Bookmarks are together, but cookies are in a different folder
chrome_path = os.path.join(user_path, r'AppData\Local\Google\Chrome\User Data\Default')
chrome_cookies_path = os.path.join(chrome_path, 'Network', 'Cookies')

# Edge paths - History and Bookmarks are together, but cookies are in a different folder
edge_path = os.path.join(user_path, r'AppData\Local\Microsoft\Edge\User Data\Default')
edge_cookies_path = os.path.join(edge_path, 'Network', 'Cookies')

# Timestamp columns present in each file (these are the same across Edge and Chrome (Chromium based))
timestamp_cookies = ['creation_utc', 'expires_utc', 'last_access_utc']
timestamp_history = ['last_visit_time']
timestamp_bookmarks = ['date_added']

# History - executes extract_history_cookies for both Chrome and Edge's histories
chrome_history = extract_history_cookies(
    os.path.join(chrome_path, 'History'),
    "SELECT url, title, visit_count, last_visit_time FROM urls",
    columns_to_convert=timestamp_history
)

edge_history = extract_history_cookies(
    os.path.join(edge_path, 'History'),
    "SELECT url, title, visit_count, last_visit_time FROM urls",
    columns_to_convert=timestamp_history
)

# Cookies - executes extract_history_cookies for both Chrome and Edge's cookies
chrome_cookies = extract_history_cookies(chrome_cookies_path,
    "SELECT host_key, name, value, creation_utc, expires_utc, last_access_utc FROM cookies",
    columns_to_convert=timestamp_cookies)

edge_cookies = extract_history_cookies(edge_cookies_path,
    "SELECT host_key, name, value, creation_utc, expires_utc, last_access_utc FROM cookies",
    columns_to_convert=timestamp_cookies)

# Bookmarks - executes extract_bookmarks to obtain both Chrome and Edge's bookmarks, timestamp conversion done as well
chrome_bookmarks = extract_bookmarks(os.path.join(chrome_path, 'Bookmarks'))

edge_bookmarks = extract_bookmarks(os.path.join(edge_path, 'Bookmarks'))


# Save to CSVs
chrome_history.to_csv("chrome_history.csv", index=False)
edge_history.to_csv("edge_history.csv", index=False)
chrome_cookies.to_csv("chrome_cookies.csv", index=False)
edge_cookies.to_csv("edge_cookies.csv", index=False)
chrome_bookmarks.to_csv("chrome_bookmarks.csv", index=False)
edge_bookmarks.to_csv("edge_bookmarks.csv", index=False)

# Create combined CSVs using pandas
combined_history = pandas.concat(
    [chrome_history.assign(browser='Chrome'), edge_history.assign(browser='Edge')])
combined_cookies = pandas.concat(
    [chrome_cookies.assign(browser='Chrome'), edge_cookies.assign(browser='Edge')])
combined_bookmarks = pandas.concat(
    [chrome_bookmarks.assign(browser='Chrome'), edge_bookmarks.assign(browser='Edge')])

combined_history.to_csv("combined_history.csv", index=False)
combined_cookies.to_csv("combined_cookies.csv", index=False)
combined_bookmarks.to_csv("combined_bookmarks.csv", index=False)

# Print confirmation that the program ran successfully
print("Data exported to CSV files")
