import pandas as pd
import numpy as np
import datetime
import math
import plotly.express as px

xls = pd.ExcelFile('./_09_2024.xlsm')
name = 'lorine'

#fetch names and content of sheets

sheet_names = xls.sheet_names
# print(sheet_names)
sheets_content = []

for sheet_name in sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    if df.iloc[0, 0] != sheet_name:
        df = df.iloc[1:]
    sheets_content.append(df)

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
        if value == sheet_names[sheet_number]:
            row_with_week.append(index)

    # print(row_with_week)

    #get all dates in the week
    week_dates = []

    for week in row_with_week :
        week_dates.append(sheet_dict_lists[1][week].strftime("%d %B %Y"))

    # print(week_dates)

    #fetch half-hours in each day
    planning = np.full((6, 24), '', dtype=str)
    i = 0
    j = 0

    for row_name in rows_with_name :
        for half_hour in range (1, 25) :
            planning[i, half_hour-1] = sheet_dict_lists[list(sheet_dict_lists.keys())[half_hour]][row_name]
            j+=1
        i+=1

    # print(planning)

    def workFrame(planning) :
        frames = []
        half = np.arange(16, 41, 1)
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
    print(sheet_names[sheet_number])
    for day in planning :
        print(workFrame(day))
    
    data = []
    date = 0
    start_time = datetime.datetime(2024, 1, 1, 0, 0)
    days_of_week = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']

    for day in planning :
        frame = workFrame(day)
        if frame == []:
            start = start_time + datetime.timedelta(hours=8)
            finish = start_time + datetime.timedelta(hours=20)
            activity_name = f"{days_of_week[date]} {week_dates[date]}"
            data.append(dict(Work=activity_name, Start=start, Finish=finish, Tâche="off"))
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
        date = date+1


    # Data with hour-based tasks (use a dummy date like '2024-01-01')
    df = pd.DataFrame(data)

    # Define the exact order of tasks as they appear in the data
    task_order = df["Work"].tolist()

    color_map = {
        "travail": "rgb(141,237,217)",
        "off": "rgb(238,9,121)"
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
        )
    )

    fig.show()