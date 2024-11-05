import pandas as pd
import numpy as np
import datetime
import math
import plotly.express as px
import tkinter as tk
from tkinter import filedialog
import openpyxl
import locale
import string
import re

root = tk.Tk()
root.withdraw()

default_directory = " D:/GitHub/planningManipulation"

file_path = filedialog.askopenfilename(
    title="Selectionner un planning Excel",
    initialdir=default_directory,
    filetypes=[("Excel files", "*.xlsm"),("Excel files", "*.xlsx"), ("Excel files", "*.xls")]
)
# Print selected file path
if file_path:
    print(f"Selected file: {file_path}")
else:
    print("No file selected.")

xls = pd.ExcelFile(file_path)
name = 'lorine'

#fetch names and content of sheets

sheet_names = xls.sheet_names
# print(sheet_names)
sheets_content = []
sheet_head = []


for sheet_name in sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    if re.sub(r"\s+", "", df.iloc[0, 0]) != re.sub(r"\s+", "", sheet_name):
        df = df.iloc[1:]
        sheet_head.append(True)
    else :
        sheet_head.append(False)
    sheets_content.append(df)

workbook = openpyxl.load_workbook(file_path)

for sheet_number in range(0,len(sheets_content)) :

    #convert sheet content to dict
    sheet_dict_lists = sheets_content[sheet_number].to_dict(orient='list')
    # print(sheet_dict_lists)

    #search rows with specified name
    UC_name = name.upper()
    rows_with_name = []

    for index, value in enumerate(sheet_dict_lists[0]):
        if value == UC_name:
            rows_with_name.append(index)

    # print(rows_with_name)

    #search rows with days of the week dates
    row_with_week = []
    for index, value in enumerate(sheet_dict_lists[0]):
        if re.sub(r"\s+", "", str(value)) == re.sub(r"\s+", "", sheet_names[sheet_number]):
            row_with_week.append(index)

    # print(row_with_week)

    #get all dates in the week
    locale.setlocale(locale.LC_TIME, 'fr_FR')
    week_dates = []

    for week in row_with_week :
        week_dates.append(sheet_dict_lists[1][week].strftime("%d %B %Y"))

    # print(week_dates)

    #fetch half-hours in each day
    planning = np.full((6, 27), '', dtype=str)
    i = 0
    j = 0

    for row_name in rows_with_name :
        for half_hour in range (1, 25) :
            planning[i, half_hour-1] = sheet_dict_lists[list(sheet_dict_lists.keys())[half_hour]][row_name]
            j+=1
        i+=1

    # print(planning)
    
    #fetch name cell color (garde)
    sheet = workbook[sheet_names[sheet_number]]

    sheet_offset = 0
    if(sheet_head[sheet_number]):
        sheet_offset = 2
    else:
        sheet_offset = 1

    for i in range(0, len(rows_with_name)):
        cell = sheet[f"A{rows_with_name[i]+sheet_offset}"]
        if isinstance(cell.fill.start_color.rgb, str):
            bg_color = cell.fill.start_color.rgb
            if bg_color == "FFFF0000":
                planning[i, 23] = '1'
                planning[i, 24] = '1'
                planning[i, 25] = '1'

    def workFrame(planning) :
        frames = []
        half = np.arange(16, 44, 1)
        isBegin = False
        row = [0,0]
        for i in range(0, half.size-1) :
            if(planning[i] == '1' and not isBegin) :
                isBegin = True
                row[0] = half[i]/2
            if(planning[i] != '1' and isBegin) :
                isBegin = False
                row[1] = half[i]/2
                frames.append([row[0], row[1]])

        return frames
    # print(sheet_names[sheet_number])
    print("------------------------")
    # for day in planning :
    #     print(workFrame(day))
    
    data = []
    date = 0
    start_time = datetime.datetime(2024, 1, 1, 0, 0)
    days_of_week = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']

    alphabet_enum = list(string.ascii_uppercase)

    sheet_offset = 0
    if(sheet_head[sheet_number]):
        sheet_offset = 2
    else:
        sheet_offset = 1

    for day in planning :
        frame = workFrame(day)
        if frame == []:
            start = start_time + datetime.timedelta(hours=8)
            finish = start_time + datetime.timedelta(hours=20)
            activity_name = f"{days_of_week[date]} {week_dates[date]}"
            tache = "off"
            for i in range (0, day.size-4):
                cell = sheet[f"{alphabet_enum[i+1]}{rows_with_name[date]+sheet_offset}"]
                if isinstance(cell.fill.start_color.rgb, str):
                    bg_color = cell.fill.start_color.rgb
                    if bg_color == "FFFFFF00":
                        tache = "vacances"
                    elif bg_color == "FF92D050":
                        tache = "repos"
                    elif bg_color == "FFFF99CC":
                        tache = "arrêt"
            data.append(dict(Work=activity_name, Start=start, Finish=finish, Tâche=tache))
        for interval in frame :
            start_minute = 0
            if interval[0]%1 != 0:
                start_minute = 30
            start = start_time + datetime.timedelta(hours=math.floor(interval[0]), minutes=start_minute)
            finish_minute  = 0
            if interval[1]%1 != 0:
                finish_minute = 30
            finish = start_time + datetime.timedelta(hours=math.floor(interval[1]), minutes=finish_minute)
            activity_name = f"{days_of_week[date]} {week_dates[date]}"
            data.append(dict(Work=activity_name, Start=start, Finish=finish, Tâche="travail"))
        date +=1

    # Data with hour-based tasks (use a dummy date like '2024-01-01')
    df = pd.DataFrame(data)

    # Convert Start and Finish to datetime format
    df['Start'] = pd.to_datetime(df['Start'])
    df['Finish'] = pd.to_datetime(df['Finish'])

    # Calculate the duration of each task and total time elapsed
    df['Duration'] = df['Finish'] - df['Start']
    task_work_df = df[df['Tâche'] == "travail"]

    # Sum the total duration for "Task A"
    total_duration_task_work = task_work_df['Duration'].sum()

    # Retrieve total time in hours for "Task A"
    total_hours_task_work = total_duration_task_work.total_seconds() / 3600

    # Define the exact order of tasks as they appear in the data
    task_order = df["Work"].tolist()

    color_map = {
        "travail": "rgb(141,237,217)",
        "off": "rgb(253,88,110)",
        "repos": "rgb(238,9,121)",
        "vacances": "rgb(192,15,191)",
        "arrêt": "rgb(174,225,242)"
    }

    fig = px.timeline(
        df,
        x_start="Start",
        x_end="Finish",
        y="Work",
        color="Tâche",
        category_orders={"Work": task_order},
        color_discrete_map=color_map
    )

    # Adjust x-axis to show tick marks at 30-minute intervals
    fig.update_layout(
        xaxis=dict(
            tickformat="%H:%M",  # Format to show only hours and minutes
            title="Temps",
            side="top",  # Move the x-axis to the top
            dtick=1800000  # 30 minutes in milliseconds (30 * 60 * 1000)
        ),
        title=f"Total heures : {total_hours_task_work} h"
    )

    fig.show()