# functions.py>
from re import compile, findall
from openpyxl import Workbook


# functions
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


def new_xlsx(pn_602=None, pn_700=None, pn_804=None, pn_cinebench=None, pn_3dmark=None, pn_pcmark=None, output_file_path=''):
    if pn_602 is None:
        pn_602 = []
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
    ws.append([])
    ws.append(header_cinebench)
    ws.append(headers_cinebench)
    for test in pn_cinebench:
        ws.append(test)
    ws.append(header_3dmark)
    for item in pn_3dmark:
        ws.append(item)
    ws.append([])
    # ws.append(pn_pcmark)
    # ws.append(header_pcmark)

    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value
        wb.save(dest_filepath + '/' + dest_filename)
        
