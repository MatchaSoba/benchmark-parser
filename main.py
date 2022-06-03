import PySimpleGUI as sg
import os.path
from functions import *
from xmlfunctions import *

# -----------------------------------ALL GUI HERE-------------------------------------------
# First the window layout in 2 columns

file_list_column = [
    [
        sg.Text('Folder with reports:', text_color='#44d62c', background_color='black'),
        sg.In(size=(25, 1), enable_events=True, key='-FOLDER-'),
        sg.FolderBrowse(button_color=('black', '#44d62c')),
    ],
    [
        sg.Listbox(
            values=[],
            enable_events=True,
            size=(50, 20),
            key='-FILE LIST-',
            sbar_background_color='#44d62c',
            sbar_arrow_color='black',
            sbar_frame_color='black',
        )
    ]
]

functions_column = [
    [sg.Text('NOTE: Select the FOLDER with the reports, then click PARSE NOW', text_color='#44d62c',
             background_color='black')],
    [sg.Button('Parse Now', button_color=('black', '#44d62c'), size=(10, 2))],
    [sg.Text('Status: Not parsed', key='-STATUS1-', text_color='#44d62c', background_color='black')],
    [sg.Text('Where would you like your output file to go?', text_color='#44d62c', background_color='black')],
    [
        sg.In(size=(25, 1), enable_events=True, key='-OUTPUTPATH-'),
        sg.FolderBrowse(button_color=('black', '#44d62c'))
    ],
    [sg.Button('Create', button_color=('black', '#44d62c'), size=(10, 2))],
    [sg.Text('Status: Not created, not parsed', key='-STATUS2-', text_color='red', background_color='black')]
]

# Layout assembly
layout = [
    [
        sg.Column(file_list_column, background_color='black'),
        sg.VSeperator(color='#44d62c'),
        sg.Column(functions_column, background_color='black'),
    ]
]

# Window assembly
window = sg.Window("Benchmark Data Parser v1.0.0", layout, background_color='black', icon='razer.ico')

# Event Loop
while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break

    # Folder name was filled in, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []
        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
               and f.lower().endswith('.txt')
               or f.lower().endswith('.xml')
        ]
        fpaths = [
            folder + '/' + f for f in fnames
        ]

        window["-FILE LIST-"].update(fnames)

    if event == 'Parse Now':
        parsed_numbers_602 = []
        parsed_numbers_700 = []
        parsed_numbers_804 = []
        parsed_numbers_cinebench = []
        parsed_numbers_3dmark = []
        parsed_numbers_3dmark_raw = []
        try:
            for file_path in fpaths:
                while True:
                    try:
                        # Getting user input for file paths:
                        parsed = parse_text_cdm(file_path)
                        # Checking app version for CDM Reports and appending to respective lists:
                        if parsed[2] == ['6.0.2']:
                            parsed_numbers_602.append(parsed)
                        elif parsed[2] == ["7.0.0"]:
                            parsed_numbers_700.append(parsed)
                        break
                    # For 8.0.4, CDM uses UTF-8 encoding, thus a specific encoding hint for parse_text function.
                    except UnicodeError:
                        try:
                            parsed_numbers_804.append(parse_text_cdm(file_path, 'utf8'))
                            break
                        except IndexError:
                            try:
                                parsed_numbers_cinebench.append(parse_text_cinebench(file_path, 'utf8'))
                                break
                            except IndexError:
                                parsed_numbers_3dmark_raw.append(parse_text_3d(file_path))
                                for item in parsed_numbers_3dmark_raw:
                                    print(item)
                                    for data in item:
                                        print(data)
                                        if 'Score' in data[0]:
                                            parsed_numbers_3dmark.append(data)
                                        elif 'score' in data[0]:
                                            parsed_numbers_3dmark.append(data)
                                        elif 'Dlss' in data[0]:
                                            parsed_numbers_3dmark.append(data)
                                        elif 'Raytracing' in data[0]:
                                            parsed_numbers_3dmark.append(data)
                                    for reading in parsed_numbers_3dmark:
                                        if 'ForPass' in reading[0]:
                                            parsed_numbers_3dmark.remove(reading)
                                        elif 'forpass' in reading[0]:
                                            parsed_numbers_3dmark.remove(reading)
                                    new_parsed_numbers_3dmark = []
                                    for reading in parsed_numbers_3dmark:
                                        if reading not in new_parsed_numbers_3dmark:
                                            new_parsed_numbers_3dmark.append(reading)
                                    parsed_numbers_3dmark = new_parsed_numbers_3dmark
                            break

            window['-STATUS1-'].update('Files parsed!')
            window['-STATUS2-'].update('Not created, ready', text_color='#44d62c')
        except NameError:
            continue

    if event == '-OUTPUTPATH-':
        outputpath = values["-OUTPUTPATH-"]

    if event == 'Create':
        try:
            new_xlsx(parsed_numbers_602,
                     parsed_numbers_700,
                     parsed_numbers_804,
                     parsed_numbers_cinebench,
                     parsed_numbers_3dmark,
                     output_file_path=outputpath)
            window['-STATUS2-'].update('File created!', text_color='#44d62c')
        # except IndexError:
        #     print('Index error')
        #     continue
        except NameError:
            print('name error')
            window['-STATUS2-'].update('No files parsed yet!', text_color='red')
            continue
        except FileNotFoundError:
            window['-STATUS2-'].update('Invalid Destination, please use the browse function.\n'
                                       'Please press Parse Now again before creating file.', text_color='red')
            continue
        except PermissionError:
            window['-STATUS2-'].update(''
                                       'Permission denied. Please close the Excel file.\n'
                                       'Please press Parse Now again before creating file.', text_color='red')

