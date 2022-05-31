import PySimpleGUI as sg
import os.path
from re import compile, findall
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


# ---------------------------------------All functions here-------------------------------------------------------
def parse_text_cinebench(input_file_path, encoding='utf8'):
    # Regex expressions for different patterns in cinebench:
    cinebench_type_pattern = compile(r'Single|Multiple')
    cinebench_score_pattern = compile(r'CB (\d*.\d*)')
    cinebench_version_pattern = compile(r'R20|R23')

    with open(input_file_path, "r", encoding=encoding) as opened_txt:
        read_txt: str = opened_txt.read()
        cinebench_version_raw = findall(cinebench_version_pattern, read_txt)
        cinebench_version = cinebench_version_raw[0]
        score_raw = findall(cinebench_score_pattern, read_txt)
        score = float(score_raw[0])
        test_type_raw = findall(cinebench_type_pattern, read_txt)
        test_type = test_type_raw[0]
        output_list = [cinebench_version, test_type, score]
        print(output_list)
        return output_list


def parse_text_cdm(input_file_path, encoding='utf16'):
    # Regex expressions for different patterns
    mb_pattern = compile(r'(\d*).(\d*) (?=MB)')
    iops_pattern = compile(r'(\d*).(\d*) (?=IOPS)')
    sequential_pattern = compile(r'(seq+\w*)')
    random_pattern = compile(r'(R\w*nd)')
    all_headers_pattern = compile(r'(R\w*nd)|(seq+\w*)')
    read_pattern = compile(r'(read)')
    write_pattern = compile(r'(write)')
    cdm_version = compile(r'\d\.\d\.\d')
    test_pattern = compile(r' \d\d?\d?\d?\d? M?G?iB ')

    # opening CDM report, with read permission and corresponding encoding
    with open(input_file_path, "r", encoding=encoding) as opened_txt:
        read_txt: str = opened_txt.read()

    # Finding all patterns with regex
    app_version = findall(cdm_version, read_txt)
    mb_readings = findall(mb_pattern, read_txt)[1:]
    test_type = findall(test_pattern, read_txt)
    mb_readings_formatted = []
    print(test_type[0] + 'read!')
    mb_readings_formatted.append(test_type[0])
    for pair in mb_readings:
        mb_readings_formatted.append(pair[0] + '.' + pair[1])
    iops_readings = findall(iops_pattern, read_txt)
    iops_readings_formatted = []
    for pair in iops_readings:
        iops_readings_formatted.append(pair[0] + '.' + pair[1])

    output_list = [mb_readings_formatted, iops_readings_formatted, app_version]
    return output_list


def new_xlsx(pn_602=[], pn_700=[], pn_804=[], pn_cinebench=[], output_file_path=''):
    dest_filename = 'output.xlsx'
    dest_filepath = output_file_path
    wb = Workbook()
    ws = wb.active
    header_cdm = ['Crystal Disk Mark Results']
    header_cinebench = ['Cinebench Results']
    header_3dmark = ['3DMark Results']
    header_pcmark = ['PCMark10 Results']
    headers_602 = [
        'CDM Version',
        'Test Type',
        'Sequential Read Q32T1 (MB/s)',
        'Sequential Write Q32T1 (MB/s)',
        'Random Read Q8T8 (MB/s)',
        'Random Write Q8T8 (MB/s)',
        'Random Read Q32T4 (IOP/s)',
        'Random Write Q32T4 (IOP/s)',
        'Random Read Q1T1 (IOP/s)',
        'Random Write Q1T1 (IOP/s)'
    ]
    headers_700 = [
        'CDM Version',
        'Test Type',
        'Sequential Read Q8T1 (MB/s)',
        'Sequential Write Q8T1 (MB/s)',
        'Sequential Read Q1T1 (MB/s)',
        'Sequential Write Q1T1 (MB/s)',
        'Random Read Q32T16 (IOP/s)',
        'Random Write Q32T16 (IOP/s)',
        'Random Read Q1T1 (IOP/s)',
        'Random Write Q1T1 (IOP/s)'
    ]
    headers_804 = [
        'CDM Version',
        'Test Type',
        'Sequential Read Q8T1 (MB/s)',
        'Sequential Write Q8T1 (MB/s)',
        'Sequential Read Q32T1 (MB/s)',
        'Sequential Write Q32T1 (MB/s)',
        'Random Read Q32T16 (IOP/s)',
        'Random Write Q32T16 (IOP/s)',
        'Random Read Q1T1 (IOP/s)',
        'Random Write Q1T1 (IOP/s)'
    ]
    headers_cinebench = [
        'Cinebench Version',
        'Test Type',
        'Score'
    ]
    ws.append(header_cdm)
    ws.append(headers_602)
    for test in pn_602:
        results_602 = test[2] + test[0][0:5] + test[1][2:6]
        results_602[2:] = [float(num) for num in results_602[2:]]
        ws.append(results_602)
    ws.append(headers_700)
    for test in pn_700:
        order_mb = [0, 1, 5, 2, 6]
        order_iops = [2, 6, 3, -1]
        test[0] = [test[0][i] for i in order_mb]
        test[0][1:] = [float(num) for num in test[0][1:]]
        test[1] = [test[1][i] for i in order_iops]
        test[1] = [float(num) for num in test[1]]
        ws.append(test[2] + test[0] + test[1])
    ws.append(headers_804)
    for test in pn_804:
        order_mb = [0, 1, 5, 2, 6]
        order_iops = [2, 6, 3, -1]
        test[0] = [test[0][i] for i in order_mb]
        test[0][1:] = [float(num) for num in test[0][1:]]
        test[1] = [test[1][i] for i in order_iops]
        test[1] = [float(num) for num in test[1]]
        ws.append(test[2] + test[0] + test[1])
    ws.append(header_cinebench)
    ws.append(headers_cinebench)
    for test in pn_cinebench:
        ws.append(test)

    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value
        wb.save(dest_filepath + '/' + dest_filename)


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
            size=(40, 20),
            key='-FILE LIST-',
            sbar_background_color='#44d62c',
            sbar_arrow_color='black',
            sbar_frame_color='black',
        )
    ]
]

functions_column = [
    [sg.Text('NOTE: Select the FOLDER with the reports, then click PARSE NOW', text_color='#44d62c', background_color='black')],
    [sg.Button('Parse Now', button_color=('black', '#44d62c'))],
    [sg.Text('Status: Not parsed', key='-STATUS1-', text_color='#44d62c', background_color='black')],
    [sg.Text('Where would you like your output file to go?', text_color='#44d62c', background_color='black')],
    [
        sg.In(size=(25, 1), enable_events=True, key='-OUTPUTPATH-'),
        sg.FolderBrowse(button_color=('black', '#44d62c'))
    ],
    [sg.Button('Create', button_color=('black', '#44d62c'))],
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
window = sg.Window("Benchmark Data Parser v1.0.0", layout, background_color='black')

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
        ]
        fpaths = [
            folder + '/' + f for f in fnames
        ]

        window["-FILE LIST-"].update(fpaths)

    if event == 'Parse Now':
        try:
            parsed_numbers_602 = []
            parsed_numbers_700 = []
            parsed_numbers_804 = []
            parsed_numbers_cinebench = []
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
                            parsed_numbers_cinebench.append(parse_text_cinebench(file_path, 'utf8'))
                            break

            window['-STATUS1-'].update('Files parsed!')
            window['-STATUS2-'].update('Not created, ready', text_color='#44d62c')
        except NameError:
            continue

    if event == '-OUTPUTPATH-':
        outputpath = values["-OUTPUTPATH-"]

    if event == 'Create':
        try:
            new_xlsx(parsed_numbers_602, parsed_numbers_700, parsed_numbers_804, parsed_numbers_cinebench, outputpath)
            window['-STATUS2-'].update('File created!', text_color='#44d62c')
        except IndexError:
            continue
        except NameError:
            continue
