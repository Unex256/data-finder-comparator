import re
import sys

import pandas as pd

import os
import PySimpleGUI as sg
import configparser
import numpy as np
from openpyxl import load_workbook

from threading import Thread, Event, Lock
from queue import Queue, Empty
import time

from images import cogwheel_image_data


def read_file_to_df(filepath, sheet=0):
    d = pd.read_excel(filepath, sheet)
    return d


def visual_levenshtein_distance(s: str, t: str) -> tuple:
    # Convert strings to lowercase
    s_clean = re.sub(r'[^a-zA-Z0-9]', '', s.lower())
    t_clean = re.sub(r'[^a-zA-Z0-9]', '', t.lower())

    # Create a matrix of size (len(s) + 1) x (len(t) + 1)
    d = [[0] * (len(t_clean) + 1) for i in range(len(s_clean) + 1)]

    # Initialize the first row and column
    for i in range(len(s_clean) + 1):
        d[i][0] = i
    for j in range(len(t_clean) + 1):
        d[0][j] = j

    # Calculate the minimum edit distance
    for j in range(1, len(t_clean) + 1):
        for i in range(1, len(s_clean) + 1):
            if s_clean[i - 1] == t_clean[j - 1]:
                d[i][j] = d[i - 1][j - 1]
            else:
                d[i][j] = min(d[i - 1][j] + 1,  # Deletion
                              d[i][j - 1] + 1,  # Insertion
                              d[i - 1][j - 1] + 1)  # Substitution

    # Generate the list of symbols representing differences
    symbol_list = []
    s_index = 0
    t_index = 0
    while s_index < len(s_clean) and t_index < len(t_clean):
        if s_clean[s_index] == t_clean[t_index]:
            symbol_list.append(1)
            s_index += 1
            t_index += 1
        else:
            if d[s_index + 1][t_index + 1] == d[s_index][t_index] + 1:  # Substitution
                symbol_list.append(0)
                s_index += 1
                t_index += 1
            elif d[s_index + 1][t_index + 1] == d[s_index + 1][t_index] + 1:  # Deletion
                symbol_list.append(0)
                s_index += 1
            elif d[s_index + 1][t_index + 1] == d[s_index][t_index + 1] + 1:  # Insertion
                symbol_list.append(0)
                t_index += 1

    # Return the list of symbols representing differences
    return d[len(s_clean)][len(t_clean)], symbol_list


def old_visual_levenshtein_distance(s: str, t: str) -> tuple:
    # Convert strings to lowercase
    s_clean = re.sub(r'[^a-zA-Z0-9]', '', s.lower())
    t_clean = re.sub(r'[^a-zA-Z0-9]', '', t.lower())

    # Create a matrix of size (len(s) + 1) x (len(t) + 1)
    d = [[0] * (len(t_clean) + 1) for i in range(len(s_clean) + 1)]

    # Initialize the first row and column
    for i in range(len(s_clean) + 1):
        d[i][0] = i
    for j in range(len(t_clean) + 1):
        d[0][j] = j

    # Calculate the minimum edit distance
    for j in range(1, len(t_clean) + 1):
        for i in range(1, len(s_clean) + 1):
            if s_clean[i-1] == t_clean[j-1]:
                d[i][j] = d[i-1][j-1]
            else:
                d[i][j] = min(d[i-1][j] + 1, # Deletion
                              d[i][j-1] + 1, # Insertion
                              d[i-1][j-1] + 1) # Substitution

    # Generate the list of symbols representing differences
    symbol_list = []
    s_index = 0
    t_index = 0
    while s_index < len(s_clean) and t_index < len(t_clean):
        if s_clean[s_index] == t_clean[t_index]:
            symbol_list.append("*")
            s_index += 1
            t_index += 1
        else:
            if d[s_index + 1][t_index + 1] == d[s_index][t_index] + 1:  # Substitution
                symbol_list.append(t[t_index])
                s_index += 1
                t_index += 1
            elif d[s_index + 1][t_index + 1] == d[s_index + 1][t_index] + 1:  # Deletion
                symbol_list.append("*")
                s_index += 1
            elif d[s_index + 1][t_index + 1] == d[s_index][t_index + 1] + 1:  # Insertion
                symbol_list.append(t[t_index])
                t_index += 1

    # Add remaining symbols to the end of the list
    while s_index < len(s):
        symbol_list.append("*")
        s_index += 1
    while t_index < len(t):
        symbol_list.append(t[t_index])
        t_index += 1

    # Return the minimum edit distance and the list of symbols representing differences
    return d[len(s_clean)][len(t_clean)], symbol_list


def visualise_differences(s1, s2):
    # In case inputs are not strings
    s1 = str(s1)
    s2 = str(s2)
    # for testing: CH-S09FTXD-BL/SC CH-S09FTXAL-SC
    min_index = 0
    match_list = []
    match_index_list = []
    skip_need = 0
    for index, char in enumerate(s1):
        if skip_need > 0:
            skip_need -= 1
            continue
        step = 1
        last_match = None
        if index + step == len(s1):
            break
        char += s1[index + step]
        while True:
            match = re.search(char, s2[min_index:])
            if match is not None:
                skip_need += 1
                last_match = match
                step += 1
                if index + step == len(s1):
                    if last_match is not None:
                        start_index, end_index = last_match.span()
                        match_list.append(s2[min_index + start_index:min_index + end_index])
                        match_index_list.append([min_index + start_index, min_index + end_index])
                        min_index += end_index
                    break
                char += s1[index + step]
            else:
                if last_match is not None:
                    start_index, end_index = last_match.span()
                    match_list.append(s2[min_index + start_index:min_index + end_index])
                    match_index_list.append([min_index + start_index, min_index + end_index])
                    min_index += end_index
                break
    return match_list, match_index_list


def levenshtein_distance(s: str, t: str) -> int:
    # In case that values are not strings
    s = str(s)
    t = str(t)
    # Convert strings to lowercase
    s = re.sub(r'[^a-zA-Z0-9]', '', s.lower())
    t = re.sub(r'[^a-zA-Z0-9]', '', t.lower())

    # Create a matrix of size (len(s) + 1) x (len(t) + 1)
    d = [[0] * (len(t) + 1) for i in range(len(s) + 1)]

    # Initialize the first row and column
    for i in range(len(s) + 1):
        d[i][0] = i
    for j in range(len(t) + 1):
        d[0][j] = j

    # Calculate the minimum edit distance
    for j in range(1, len(t) + 1):
        for i in range(1, len(s) + 1):
            if s[i-1] == t[j-1]:
                d[i][j] = d[i-1][j-1]
            else:
                d[i][j] = min(d[i-1][j] + 1, # Deletion
                              d[i][j-1] + 1, # Insertion
                              d[i-1][j-1] + 1) # Substitution

    # Return the minimum edit distance between the two strings
    return d[len(s)][len(t)]


def find_matches(search_val, match_list, threshold=3):
    exact_matches = []
    best_list = []
    best_distance = float('inf')
    potential_matches_list = []
    poor_best_list = []

    for i, match_val in enumerate(match_list):
        if pd.isna(match_val):
            continue
        distance = levenshtein_distance(search_val, match_val)

        if distance == 0:
            exact_matches.append([match_val, distance, i])

        elif distance <= threshold:
            potential_matches_list.append([match_val, distance, i])
            if distance <= best_distance:
                if distance < best_distance:
                    best_list = []
                    best_distance = distance
                best_list.append([match_val, distance, i])

    if len(best_list) > 0:
        potential_matches_list = [value for value in potential_matches_list if value not in best_list]
    else:
        # find the best match outside the threshold distance
        for i, match_val in enumerate(match_list):
            if pd.isna(match_val):
                continue
            distance = levenshtein_distance(search_val, match_val)
            if distance <= best_distance:
                if distance < best_distance:
                    poor_best_list = []
                    best_distance = distance
                poor_best_list.append([match_val, distance, i])

    return exact_matches, best_list, potential_matches_list, poor_best_list


def visual_find_matches(search_val, match_list):
    exact_matches = []
    best_list = []
    best_distance = float('inf')
    potential_matches_list = []
    poor_best_list = []

    for match_val in match_list:
        distance, difference = visual_levenshtein_distance(search_val, match_val)

        if distance == 0:
            exact_matches.append([match_val, distance, difference])

        elif distance <= threshold:
            potential_matches_list.append([match_val, distance, difference])
            if distance <= best_distance:
                if distance < best_distance:
                    best_list = []
                    best_distance = distance
                best_list.append([match_val, distance, difference])

    if len(best_list) > 0:
        potential_matches_list = [value for value in potential_matches_list if value not in best_list]
    else:
        # find the best match outside the threshold distance
        for match_val in match_list:
            distance, difference = visual_levenshtein_distance(search_val, match_val)
            if distance <= best_distance:
                if distance < best_distance:
                    poor_best_list = []
                    best_distance = distance
                poor_best_list.append([match_val, distance, difference])

    return exact_matches, best_list, potential_matches_list, poor_best_list


def colored_text(string, red_indices):
    # Initialize the list of Text elements
    text_elements = []
    start_index = 0

    # Iterate through the red indices and add colored and uncolored Text elements to the list
    for i in red_indices:
        # Add the uncolored Text element for the substring between the previous index and the current one
        if i > start_index:
            text_elements.append(sg.Text(string[start_index:i], pad=(0, 0)))
        # Add the colored Text element for the current substring
        text_elements.append(sg.Text(string[i], text_color='red', pad=(0, 0)))
        start_index = i + 1

    # Add the final uncolored Text element for the substring after the last index
    if start_index < len(string):
        text_elements.append(sg.Text(string[start_index:], pad=(0, 0)))

    return text_elements


def get_splices_with_indices(match_splices, match_indices, s2):
    s2 = str(s2)
    result = []
    if len(match_indices) == 0:
        result.append((0, s2))
        return result
    if match_indices[0][0] != 0:
        result.append((0, s2[:match_indices[0][0]]))
    for i in range(len(match_splices)):
        result.append((1, match_splices[i]))
        if i != len(match_splices) - 1:
            result.append((0, s2[match_indices[i][1]:match_indices[i+1][0]]))
    if match_indices[-1][1] != len(s2):
        result.append((0, s2[match_indices[-1][1]:]))
    return result


def settings_page():
    global out_1
    global out_2
    global column_index1
    global column_index2
    global threshold
    global settings_changed
    # Get the column names from the DataFrames
    columns1 = df_search.columns.tolist()
    columns2 = df_data.columns.tolist()

    # Create the table data
    table_data1 = [[col] for col in columns1]
    table_data2 = [[col] for col in columns2]

    # Create the layout
    layout = [[sg.Text('Select a column from DataFrame 1:')],
              [sg.Input(default_text=columns1[0], key='text1', size=(30, 0), disabled=True)],
              [sg.Push(), sg.Table(values=table_data1, headings=['DataFrame 1'], key='table1', enable_events=True)],
              [sg.Text('Select a column from DataFrame 2:')],
              [sg.Input(default_text=columns2[0], key='text2', size=(30, 0), disabled=True)],
              [sg.Push(), sg.Table(values=table_data2, headings=['DataFrame 2'], key='table2', enable_events=True)],
              [sg.Button('Apply')]]

    layout_right = [[sg.Text("Threshold for good matches")],
                    [sg.Spin([i for i in range(1, 11)],
                             initial_value=threshold,
                             key='-THRESHOLD-',
                             size=(20, 0))],
                    [sg.Checkbox('Output DataFrame1',
                     key='-OUT_1-',
                     default=out_1)],
                    [sg.Checkbox('Output DataFrame2',
                     key='-OUT_2-',
                     default=out_2)]]

    layout = [[sg.Column(layout), sg.VerticalSeparator(), sg.Column(layout_right)]]

    window = sg.Window('Settings', layout, finalize=True)

    window["table1"].update(select_rows=[column_index1])
    window["table2"].update(select_rows=[column_index2])

    window['-THRESHOLD-'].update(threshold)

    while True:
        event, values = window.read()
        print(event)
        if event == 'table1':
            row_index = values['table1'][0]
            window['text1'].update(columns1[row_index])
        elif event == 'table2':
            row_index = values['table2'][0]
            window['text2'].update(columns2[row_index])
        elif event == 'Apply':
            settings_changed.set()

            row_index = values['table1'][0]
            column_index1 = row_index
            config.set("settings", "column_index1", str(row_index))

            row_index = values['table2'][0]
            column_index2 = row_index
            config.set("settings", "column_index2", str(row_index))

            out_1 = values['-OUT_1-']
            config.set("settings", "out_1", str(out_1))

            out_2 = values['-OUT_2-']
            config.set("settings", "out_2", str(out_2))

            threshold = values['-THRESHOLD-']
            config.set("settings", "threshold", str(threshold))

            with open('config.ini', 'w') as configfile:
                config.write(configfile)

            window.close()
            return True

        elif event == sg.WIN_CLOSED:
            break

    window.close()


def display_matches(dis_value, dis_matches):
    if not any(len(lst) > 0 for lst in dis_matches):
        print("Column does not contain anything")

    max_match_size = (30, 1)
    max_match_size_col = (253, 20)
    distance_size = (5, 1)

    layout = [[sg.Text("Search element")],
              [sg.Text(dis_value, size=max_match_size),
               sg.Button("Confirm"),
               sg.Button("Skip")]]

    for i, match_type in enumerate(["Exact", "Best", "Potential", "Poor"]):
        layout.append([sg.HorizontalSeparator()])
        layout.append([sg.Text(match_type, size=max_match_size), sg.Text("Dist.")])
        for e, match in enumerate(dis_matches[i]):
            match_splices, match_indices = visualise_differences(dis_value, match[0])
            splices_with_indices = get_splices_with_indices(match_splices, match_indices, match[0])
            column_element = [[]]
            for j, splice_with_index in enumerate(splices_with_indices):
                if j == 0:
                    if splice_with_index[0] == 1:
                        column_element[0].append(sg.Text(splice_with_index[1], pad=((3, 0), 0)))
                    else:
                        column_element[0].append(sg.Text(splice_with_index[1], text_color="red", pad=((3, 0), 0)))
                else:
                    if splice_with_index[0] == 1:
                        column_element[0].append(sg.Text(splice_with_index[1], pad=(0, 0)))
                    else:
                        column_element[0].append(sg.Text(splice_with_index[1], text_color="red", pad=(0, 0)))

            row = [sg.Column(column_element, size=max_match_size_col, pad=(0, 0)), sg.Text(match[1], pad=distance_size)]

            row.extend([sg.Button("Replace", key=("Replace", i, e, match[0], match[2])),
                        sg.Button("Keep", key=("Keep", i, e, match[2]))])

            layout.append(row)

    layout.append([sg.HorizontalSeparator()])

    layout.append([sg.Button("Confirm"),
                   sg.Button("Skip")])

    local_table = create_table()

    settings_button = sg.Button(image_data=cogwheel_image_data,
                                 button_color=(sg.theme_background_color(),
                                               sg.theme_background_color()),
                                 border_width=0,
                                 key='-SETTINGS-',
                                 size=(25, 25))

    layout = [[sg.Column(layout), sg.Column([[sg.Push(), settings_button],
                                             [sg.VerticalSeparator(), local_table]],
                                            vertical_alignment="top"
                                            )]]

    window = sg.Window("Matches", layout, finalize=True)
    local_table.update(select_rows=[global_table_row])

    while True:
        event, values = window.read()
        print(f"event: {event}, {type(event)}")
        if event == sg.WIN_CLOSED:
            sys.exit()

        elif type(event) == tuple:
            if event[0] == "Replace":
                window.close()
                return event
            elif event[0] == "Keep":
                window.close()
                return event

        elif type(event) == str:
            if re.match("Confirm", event):
                window.close()
                return "Confirm"
            elif re.match("Skip", event):
                window.close()
                return "Skip"
            elif event == '-GLOBAL_TABLE-':
                row_index = values['-GLOBAL_TABLE-'][0]
                print(f'Row {row_index} was clicked')
            elif event == '-SETTINGS-':
                settings_page()


def process_value(queue):
    #has a bug that gui thread is faster and grabs value that is still old after setting change
    global global_table_row

    global column_index1
    global column_index2

    global settings_changed

    cfg_changed = False

    if column_index1 >= df_search.shape[1]:
        cfg_changed = True
        column_index1 = 0
        config.set("settings", "column_index1", str(column_index1))

    if column_index2 >= df_data.shape[1]:
        cfg_changed = True
        column_index2 = 0
        config.set("settings", "column_index2", str(column_index2))

    if cfg_changed:
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

    # Get the column of data to process
    data_column = df_search.iloc[:, column_index1]
    # Initialize an index variable to keep track of the current position in the column
    index = 0
    # Use a while loop to process the column
    while index < len(data_column):
        # Get the current value from the column
        value = data_column.iloc[index]
        # Check if the settings have changed
        if settings_changed.is_set():
            with queue_lock:
                # If the settings have changed, clear the queue
                while not queue.empty():
                    queue.get()
                # Reset the index variable to start processing from the beginning of the column
                index = 0
                value = data_column.iloc[index]
                # Reset the settings_changed event
                settings_changed.clear()
                # Add at least one item into the queue under the lock
                matches = find_matches(value, df_data.iloc[:, column_index2].values.tolist(), 3)
                matches[2].sort(key=lambda x: x[1])
                queue.put((value, matches))
                print(f"Matches are ready under lock for Value: {value}")
                # Increment the index variable to move to the next value in the column
                index += 1
        else:
            # If the settings have not changed, continue processing as normal
            matches = find_matches(value, df_data.iloc[:, column_index2].values.tolist(), 3)
            matches[2].sort(key=lambda x: x[1])
            queue.put((value, matches))
            print(f"Matches are ready for Value: {value}")
            # Increment the index variable to move to the next value in the column
            index += 1


def process_value_legacy(queue):
    for value in df_search.iloc[:, column_index1]:
        matches = find_matches(value, df_data.iloc[:, column_index2].values.tolist(), 3)
        matches[2].sort(key=lambda x: x[1])

        queue.put((value, matches))
        print(f"Matches are ready for Value: {value}")


def gui_process(queue):
    folder_selection_screen()
    while True:
        with queue_lock:
            try:
                item = queue.get()
                print("item got")
            except Exception as e:
                print(e)
                continue

        print(f'Processing Value: {item[0]}')
        action = display_matches(*item)
        queue.task_done()
        if action == "Confirm":
            # do something
            pass
        elif action == "Skip":
            # do something else
            pass


queue_lock = Lock()


def main_with_threading():

    queue = Queue(maxsize=3)

    gui_thread = Thread(
        target=gui_process,
        args=(queue,)
    )

    gui_thread.start()

    matching_thread = Thread(
        target=process_value,
        args=(queue,)
    )
    matching_thread.start()

    queue.join()


def keep(df_index):
    if not os.path.isfile("output.xlsx"):
        df_search.iloc[[df_index]].to_excel("output.xlsx", index=False)
    else:
        with pd.ExcelWriter(
                "output.xlsx",
                engine='openpyxl',
                mode='a',
                if_sheet_exists='overlay') as writer:
            reader = pd.read_excel("output.xlsx")
            df_search.iloc[[df_index]].to_excel(
                writer,
                startrow=reader.shape[0] + 1,
                index=False,
                header=False)


def replace(df_index, new_value):
    df_search.at[df_index, column_index1] = new_value
    if not os.path.isfile("output.xlsx"):
        df_search.iloc[[df_index]].to_excel("output.xlsx", index=False)
    else:
        with pd.ExcelWriter(
                "output.xlsx",
                engine='openpyxl',
                mode='a',
                if_sheet_exists='overlay') as writer:
            reader = pd.read_excel("output.xlsx")
            df_search.iloc[[df_index]].to_excel(
                writer,
                startrow=reader.shape[0] + 1,
                index=False,
                header=False)


def main():
    global global_table_row

    global column_index1
    global column_index2

    cfg_changed = False

    if column_index1 >= df_search.shape[1]:
        cfg_changed = True
        column_index1 = 0
        config.set("settings", "column_index1", str(column_index1))

    if column_index2 >= df_data.shape[1]:
        cfg_changed = True
        column_index2 = 0
        config.set("settings", "column_index2", str(column_index2))

    if cfg_changed:
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

    for df_index, value in enumerate(df_search.iloc[:, column_index1]):
        matches = find_matches(value, df_data.iloc[:, column_index2].values.tolist(), 3)
        matches[2].sort(key=lambda x: x[1])

        action = display_matches(value, matches)
        print(action)
        if type(action) == str:
            if action == "Confirm":
                global_table_row += 1
                print("Confirm")

            elif action == "Skip":
                global_table_row += 1
                print("Skip")

        elif type(action) == tuple:
            if action[0] == "Replace":
                global_table_row += 1
                replace(df_index, action[3])

            elif action[0] == "Keep":
                global_table_row += 1
                keep(df_index)


def create_table():
    df = df_search.iloc[:, [column_index1]].copy()
    df.reset_index(inplace=True)
    df["index"] = df["index"] + 1
    return sg.Table(df.values.tolist(),
                    headings=["index", "search_val"],
                    enable_events=True,
                    key='-GLOBAL_TABLE-',
                    size=(0, 40))


def folder_selection_screen():
    global search_file
    global data_folder
    global multithreading

    layout = [
        [sg.Text('Select xlsx file:'), sg.Input(default_text=search_file, key='-FILE-'),
         sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
        [sg.Text('Select folder with xlsx files:'),
         sg.Input(default_text=data_folder, key='-FOLDER-'),
         sg.FolderBrowse()],
        [sg.Checkbox('Enable multithreading',
                     key='-MULTITHREADING-',
                     default=multithreading),
         sg.Button('Submit')]]

    window = sg.Window('Browse Documents', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            sys.exit()

        if event == 'Submit':
            print(f"Selected file: {values['-FILE-']}")
            search_file = values['-FILE-']
            config.set("pre_browse", "search_file", search_file)
            print(f"Selected folder: {values['-FOLDER-']}")
            data_folder = values['-FOLDER-']
            config.set("pre_browse", "data_folder", data_folder)
            print(f"Multithreading enabled: {values['-MULTITHREADING-']}")
            change_check = multithreading
            multithreading = values['-MULTITHREADING-']
            config.set("multithreading", "multithreading", str(multithreading))
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
            window.close()

            if change_check != multithreading:
                sys.exit()
            break


if __name__ == '__main__':

    script_path = sys.argv[0]
    script_dir = os.path.dirname(script_path)

    config_file = 'config.ini'
    config_path = os.path.join(script_dir, config_file)

    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')

    if config.get("multithreading", "multithreading") == "True":
        multithreading = True
    else:
        multithreading = False

    if config.get("pre_browse", "search_file") == "":
        search_file = None
    else:
        search_file = config.get("pre_browse", "search_file")
    if config.get("pre_browse", "data_folder") == "":
        search_file = None
    else:
        data_folder = config.get("pre_browse", "data_folder")

    if config.get("settings", "out_1") == "True":
        out_1 = True
    else:
        out_1 = False
    if config.get("settings", "out_2") == "True":
        out_2 = True
    else:
        out_2 = False

    if int(config.get("settings", "column_index1")):
        column_index1 = int(config.get("settings", "column_index1"))
    else:
        column_index1 = 0
    if int(config.get("settings", "column_index2")):
        column_index2 = int(config.get("settings", "column_index2"))
    else:
        column_index2 = 0

    if int(config.get("settings", "threshold")):
        threshold = int(config.get("settings", "threshold"))
    else:
        threshold = 0

    data_folder = config.get("pre_browse", "data_folder")

    df_search = read_file_to_df(search_file)

    df_data = pd.DataFrame()

    data_files = os.listdir(data_folder)
    for file in data_files:
        path = f"{data_folder}/{file}"
        temp = pd.read_excel(path)
        df_data = pd.concat([df_data, temp], ignore_index=True)

    global_table_row = 0

    if multithreading:
        settings_changed = Event()
        main_with_threading()
    else:
        folder_selection_screen()
        main()




