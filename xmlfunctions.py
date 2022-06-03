import xml.etree.ElementTree as ET


def parse_text_3d(input_file_path):
    mytree = ET.parse(input_file_path)
    myroot = mytree.getroot()
    data = []
    for result in myroot.iter('result'):
        for child in result:
            data.append([child.tag, child.text])
    return data


def parse_text_pc(input_file_path):
    mytree = ET.parse(input_file_path)
    myroot = mytree.getroot()
