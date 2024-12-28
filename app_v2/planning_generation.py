import pandas as pd
import numpy as np
import datetime
import math
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import locale
import string
import re
import datetime

def log(content):
    try:
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Format the log entry
        log_entry = f"--> [{current_time}]: {content}\n"

        with open('./planning_generation.log', 'a') as file:
            file.write(log_entry)
    except FileNotFoundError:
        print("log file not found")
        pass
    except Exception as e:
        print(f"Cannot log : {e}")
        pass

def retrieve_sheets(file_path):
    try:
        xls = pd.ExcelFile(file_path)

        sheet_names = xls.sheet_names
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

        return sheet_names, sheets_content, sheet_head

    except Exception as e:
        print(e)
        log(f"An error occurred generating planning : {e}")

def retrieve_week_rows(sheet_dict, sheet_name):
    row_with_week = []
    for index, value in enumerate(sheet_dict[0]):
        if re.sub(r"\s+", "", str(value)) == re.sub(r"\s+", "", sheet_name):
            row_with_week.append(index)
    return row_with_week

def retrieve_name_rows(sheet_dict, name):
    rows_with_name = []
    for index, value in enumerate(sheet_dict[0]):
        if value == name:
            rows_with_name.append(index)
    return rows_with_name

def get_week_dates(sheet_dict, row_with_week):
    locale.setlocale(locale.LC_TIME, 'fr_FR')
    week_dates = []
    for week in row_with_week :
        week_dates.append(sheet_dict[1][week].strftime("%d %B %Y"))
    return week_dates

def get_half_hour(sheet_dict, rows_with_name):
    #fetch half-hours in each day
    planning = np.full((6, 27), '', dtype=str)
    i = 0
    j = 0
    for row_name in rows_with_name :
        for half_hour in range (1, 25) :
            planning[i, half_hour-1] = sheet_dict[list(sheet_dict.keys())[half_hour]][row_name]
            j+=1
        i+=1
    return planning

def get_color_hours(planning, color_sheet, sheet_head, row_with_name):
    sheet_offset = 0
    if(sheet_head): #sheet_head[sheet_number]
        sheet_offset = 2
    else:
        sheet_offset = 1

    for i, row_with_n in enumerate(row_with_name):
        cell = color_sheet[f"A{row_with_n+sheet_offset}"]
        if isinstance(cell.fill.start_color.rgb, str):
            bg_color = cell.fill.start_color.rgb
            if bg_color == "FFFF0000":
                planning[i, 23] = '1'
                planning[i, 24] = '1'
                planning[i, 25] = '1'
    return planning

def build_work_frame(planning) :
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

def build_planning_frame(planning, color_sheet, sheet_head, week_dates, rows_with_name):
    data = []
    start_time = datetime.datetime(2024, 1, 1, 0, 0)
    days_of_week = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']

    alphabet_enum = list(string.ascii_uppercase)

    sheet_offset = 0
    if(sheet_head):
        sheet_offset = 2
    else:
        sheet_offset = 1

    for index, day in enumerate(planning) :
        frame = build_work_frame(day)
        print("--")
        if frame == []:
            start = start_time + datetime.timedelta(hours=8)
            finish = start_time + datetime.timedelta(hours=20)
            activity_name = f"{days_of_week[index]} {week_dates[index]}"
            tache = "inconnu"
            for i in range (0, day.size-4):
                cell = color_sheet[f"{alphabet_enum[i+1]}{rows_with_name[index]+sheet_offset}"]
                print(cell)
                if isinstance(cell.fill.start_color.rgb, str):
                    bg_color = cell.fill.start_color.rgb
                    if bg_color == "FFFFFF00":
                        tache = "vacances"
                    elif bg_color == "FF92D050":
                        tache = "repos"
                    elif bg_color == "FFFF99CC":
                        tache = "arrêt"
            data.append(dict(Work=activity_name, Start=start, Finish=finish, Tâche=tache))
        else:
            print("frame not empty")
        for interval in frame :
            start_minute = 0
            if interval[0]%1 != 0:
                start_minute = 30
            start = start_time + datetime.timedelta(hours=math.floor(interval[0]), minutes=start_minute)
            finish_minute  = 0
            if interval[1]%1 != 0:
                finish_minute = 30
            finish = start_time + datetime.timedelta(hours=math.floor(interval[1]), minutes=finish_minute)
            activity_name = f"{days_of_week[index]} {week_dates[index]}"
            data.append(dict(Work=activity_name, Start=start, Finish=finish, Tâche="travail"))

    return data

def define_week_work_duration(data_frame):
    task_work_df = data_frame[data_frame['Tâche'] == "travail"]

    # Sum the total duration for "Task A"
    total_duration_task_work = task_work_df['Duration'].sum()

    # Retrieve total time in hours for "Task A"
    total_hours_task_work = total_duration_task_work.total_seconds() / 3600

    return total_hours_task_work

def generate_planning(file_path, config):
    try:
        sheet_names, sheet_content, sheet_head = retrieve_sheets(file_path)
        workbook = openpyxl.load_workbook(file_path)

        figures = []

        if 'name' in config and 'colors' in config:
            user_name = str(config['name']).upper()

            config_colors = config['colors']
            color_map = {
                "travail": config_colors['work'],
                "inconnu": config_colors['undefined'],
                "repos": config_colors['off'],
                "vacances": config_colors['vacation'],
                "arrêt": config_colors['sick']
            }

            for sheet_number, sheet in enumerate(sheet_content):
                try:
                    sheet_dict_list = sheet.to_dict(orient='list')

                    row_with_week = retrieve_week_rows(sheet_dict_list, sheet_names[sheet_number])
                    row_with_name = retrieve_name_rows(sheet_dict_list, user_name)

                    if len(row_with_week) != 6 or len(row_with_name) != 6:
                        if len(row_with_name) == 0:
                            raise Exception("The name you provided does not exist in the file")
                        elif len(row_with_week) < 6:
                            raise Exception(f"The file contains a week number error at {sheet_names[sheet_number]}, only {len(row_with_week)} instances")
                        else:
                            raise Exception("Error about user name or week number in the file")

                    week_dates = get_week_dates(sheet_dict_list, row_with_week)

                    half_hour_planning = get_half_hour(sheet_dict_list, row_with_name)

                    color_sheet =  workbook[sheet_names[sheet_number]]
                    # duty day
                    half_hour_planning = get_color_hours(half_hour_planning, color_sheet, sheet_head[sheet_number], row_with_name)

                    data = build_planning_frame(half_hour_planning, color_sheet, sheet_head[sheet_number], week_dates, row_with_name)

                    # Data with hour-based tasks (use a dummy date like '2024-01-01')
                    df = pd.DataFrame(data)
                    # Convert Start and Finish to datetime format
                    df['Start'] = pd.to_datetime(df['Start'])
                    df['Finish'] = pd.to_datetime(df['Finish'])

                    # Calculate the duration of each task and total time elapsed
                    df['Duration'] = df['Finish'] - df['Start']
                
                    total_hours_task_work = define_week_work_duration(df)

                    # Define the exact order of tasks as they appear in the data
                    task_order = df["Work"].tolist()

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

                    figures.append(fig)

                    print("------------------------")

                except Exception as e:
                    log(f"An error occurred generating planning : {e}")
                    fig = go.Figure()
                    fig.add_annotation(
                        text="Une erreur est survenue lors de la génération du planning (vérifier le format du fichier .xls*)",
                        xref="paper", yref="paper",
                        x=0.5, y=0.5,
                        showarrow=False,
                        font=dict(size=20, color="red")
                    )
                    fig.update_layout(
                        title="Erreur",
                        xaxis=dict(visible=False),
                        yaxis=dict(visible=False)
                    )
                    figures.append(fig)
                    print(e)
                    pass

        else:
            print("NO NAME OR NO COLORS IN CONFIG !!!")
            log(f"An error occurred generating planning : No name or no colors in config file")

        return figures
    except Exception as e:
        log(f"An error occurred generating planning : {e}")
        print(e)



