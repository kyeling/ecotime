import pyautogui as pg
import pandas as pd
import datetime
from datetime import *
import time

# call function: python ecotime.py
# have ecotime open in another window, then after running the script, go to the ecotime window
# CHANGE INPUT FILE BELOW:
FILENAME = 'G:/My Drive/CSE/CSE95 FA22 Hours Distribution.xlsx'

# may need to change these constants depending on your device/browser speed
SAVE_DELAY = 3.5  # delay time for page refresh on save
DAY_DELAY  = 1    # delay time for pressing enter to select day
TAB_DELAY  = 0.01 # delay time between key presses otherwise too fast for ecotime

# experimental times:
# SAVE_DELAY = 5
# DAY_DELAY  = 1
# TAB_DELAY  = .5

TIME_IN_STEPS      = 19 # number of tabs from Timesheet Employee Information to Time In
NEXT_TIME_IN_STEPS = 12 # number of tabs to next Time In
EOP_STEPS          = 12 # number of tabs from Total to end of page
SAVE_STEPS         = 16 # number of tabs from top of page to Save button #originally 13
RESET_STEPS        = 23 # number of tabs from top of page (after page refresh) to Time In

# pre-process weekly hours dataframe by dropping empty rows and converting times to H, M, AM/PM
# modifies original df so don't need to return df
# x = the number of the dataframe, if df2, add 7 to weekday integers to account for number of tabs
def preprocess(df, x):
    df.dropna(thresh=4, inplace=True)
    df['Weekday'] = (pd.to_datetime(df['Day']).dt.weekday + 1) % 7
    if x == 2: df['Weekday'] = df['Weekday'] + 7

    format_time = lambda s : datetime.strptime(str(s), '%H:%M:%S').time().strftime('%I, %M, %p')
    df['Start Time'] = df['Start Time'].apply(format_time)
    df['End Time'] = df['End Time'].apply(format_time)
 
# press tab n times
tab = lambda n : pg.press('tab', presses=n)

# press tab n times with interval t in between presses
tab_delay = lambda n, t : pg.press('tab', presses=n, interval=t)

# press downkey n times
down = lambda n : pg.press('down', presses=n)

# takes in a list of start and end hours, minutes, periods and 
# fills in one row of timesheet
def fill_time(times):
    for t in times:
        if t == 'AM': pass # down(0) bc automatically defaults to AM
        elif t == 'PM': down(1)
        elif t == '00' or int(t) > 12: down(int(t) // 15) # minutes
        else: down(int(t)) # hours
        tab(1) 
    # currently at Overnight, fill in Position ID and Pay Code   
    tab(2)
    for i in range(2):
        tab(1)
        down(1)
        time.sleep(0.25)
    tab(2) # tab to next Time In row

# go to end of page based on number of unfilled rows and
# save all time entries for a given day
def save_progress(row_count):
    tab(row_count * NEXT_TIME_IN_STEPS + EOP_STEPS)
    time.sleep(0.5)
    tab_delay(SAVE_STEPS, TAB_DELAY) 
    pg.press('enter')               
    time.sleep(SAVE_DELAY) # allow page to refresh
    tab(RESET_STEPS)            

# select day to enter time for based on number corresponding to weekday
def select_day(day):    
    tab_delay(int(day), TAB_DELAY)
    pg.press('enter')           
    time.sleep(DAY_DELAY) # allow page to refresh
    tab(TIME_IN_STEPS + 1)

# select day for first day with time entered
def select_first_day(day):    
    tab(2 + int(day)) # tab to first Sunday, then first day with time entries
    pg.press('enter')
    time.sleep(DAY_DELAY) # allow page to refresh
    tab(TIME_IN_STEPS)


# sheets listed in reverse chronological order
# get most recent 2 weeks (previous week as df1 then current week as df2)
df1 = pd.read_excel(FILENAME, sheet_name=1)
df2 = pd.read_excel(FILENAME, sheet_name=0)
preprocess(df1, 1)
preprocess(df2, 2)
df = pd.concat([df1, df2]).reset_index(drop=True)

time.sleep(5) # delay to click on window with ecotime        

num_entries = len(df)  
row_count = 3 # number of remaining unfilled rows on a given day
for i in range(num_entries):
    row = df.iloc[i]
    day = row['Weekday']

    # first day
    if i == 0:
        select_first_day(day)
    # new day   
    elif not pd.isna(day): 
        # go to end of current day and save
        save_progress(row_count)
        # reset row_count and go to next day
        row_count = 3
        select_day(day)
    # same day
    else:
        row_count -= 1

    times = row['Start Time'].split(', ') + row['End Time'].split(', ')
    fill_time(times)
    
save_progress(row_count)

    
    


