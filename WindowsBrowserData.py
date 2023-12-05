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
        epoch_start = datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=timestamp_int)
        return epoch_start.strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        # Return the original timestamp if conversion is not possible
        return timestamp


# extract_history_cookies - extracts the history and cookies. Both are stored in sqlite databases.
def extract_history_cookies(browser_data_path, query, columns_to_convert=None):
    if not os.path.exists(browser_data_path):
        print(f"File not found: {browser_data_path}")
        return pandas.DataFrame()

    try:
        backup_path = browser_data_path + '.backup'
        shutil.copyfile(browser_data_path, backup_path)
        with sqlite3.connect(backup_path) as conn:
            df = pandas.read_sql_query(query, conn)
            if columns_to_convert:
                for column in columns_to_convert:
                    df[column] = df[column].apply(convert_timestamp)
            return df
    except sqlite3.Error as e:
        print(f"Error reading from {browser_data_path}: {e}")
        return pandas.DataFrame()


# extract_bookmarks - extracts the bookmarks. Bookmarks are stored in json files.
def extract_bookmarks(bookmarks_path):
    if not os.path.exists(bookmarks_path):
        print(f"Bookmarks file not found: {bookmarks_path}")

    try:
        with open(bookmarks_path, 'r', encoding='utf-8') as file:
            bookmarks_json = pandas.read_json(file)
            bookmarks_data = pandas.json_normalize(bookmarks_json['roots']['bookmark_bar']['children'])
            return bookmarks_data[['name', 'type', 'url', 'date_added']]
    except Exception as e:
        print(f"Error reading from {bookmarks_path}: {e}")
        return pandas.DataFrame()


# Path to current user, used in locating the relevant AppData folder
user_profile_path = os.path.expanduser('~')

# Chrome paths
chrome_data_path = os.path.join(user_profile_path, r'AppData\Local\Google\Chrome\User Data\Default')
chrome_cookies_path = os.path.join(chrome_data_path, 'Network', 'Cookies')

# Edge paths
edge_data_path = os.path.join(user_profile_path, r'AppData\Local\Microsoft\Edge\User Data\Default')
edge_cookies_path = os.path.join(edge_data_path, 'Network', 'Cookies')

# Timestamp columns present in each file (these are the same across Edge and Chrome (Chromium based))
timestamp_cookies = ['creation_utc', 'expires_utc', 'last_access_utc']
timestamp_history = ['last_visit_time']
timestamp_bookmarks = ['date_added']

# History
chrome_history = extract_history_cookies(
    os.path.join(chrome_data_path, 'History'),
    "SELECT url, title, visit_count, last_visit_time FROM urls",
    columns_to_convert=timestamp_history
)

edge_history = extract_history_cookies(
    os.path.join(edge_data_path, 'History'),
    "SELECT url, title, visit_count, last_visit_time FROM urls",
    columns_to_convert=timestamp_history
)

# Cookies
chrome_cookies = extract_history_cookies(chrome_cookies_path,
    "SELECT host_key, name, value, creation_utc, expires_utc, last_access_utc FROM cookies",
    columns_to_convert=timestamp_cookies)

edge_cookies = extract_history_cookies(edge_cookies_path,
    "SELECT host_key, name, value, creation_utc, expires_utc, last_access_utc FROM cookies",
    columns_to_convert=timestamp_cookies)

# Bookmarks
chrome_bookmarks = extract_bookmarks(os.path.join(chrome_data_path, 'Bookmarks'))
chrome_bookmarks['date_added'] = chrome_bookmarks['date_added'].apply(convert_timestamp)
edge_bookmarks = extract_bookmarks(os.path.join(edge_data_path, 'Bookmarks'))
edge_bookmarks['date_added'] = edge_bookmarks['date_added'].apply(convert_timestamp)

# Save to CSVs
chrome_history.to_csv("chrome_history.csv", index=False)
edge_history.to_csv("edge_history.csv", index=False)
chrome_cookies.to_csv("chrome_cookies.csv", index=False)
edge_cookies.to_csv("edge_cookies.csv", index=False)
chrome_bookmarks.to_csv("chrome_bookmarks.csv", index=False)
edge_bookmarks.to_csv("edge_bookmarks.csv", index=False)

# Create combined CSVS
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
