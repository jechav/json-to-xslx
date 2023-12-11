import json
from xlsxwriter import Workbook
from dictor import dictor
from pprint import pprint


def create_xlsx_file(file_path: str, headers: dict, items: list):
    with Workbook(file_path) as workbook:
        worksheet = workbook.add_worksheet()
        worksheet.write_row(row=0, col=0, data=headers.values())
        header_keys = list(headers.keys())
        for index, item in enumerate(items):
            row = map(lambda field_id: item.get(field_id, ''), header_keys)
            worksheet.write_row(row=index + 1, col=0, data=row)

def get_json_content(filename):
    with open(filename) as f:
        lines = f.readlines()
        return json.loads(''.join(lines))

def dictionary_check(keys, input, parent_key=None):
    if isinstance(input, dict):
        for key,value in input.items():
            new_key = key
            if parent_key is not None:
                new_key = parent_key + "." + key
            dictionary_check(keys, value, new_key)
    elif isinstance(input, list):
        for index, item in enumerate(input):
            new_key = str(index)
            if parent_key is not None:
                new_key = parent_key + "." + str(index)
            dictionary_check(keys, item, new_key)
    else:
        keys.append(parent_key)

    return keys

def get_json_headers(created_dealers):
    element = created_dealers[0]
    keys = []
    dictionary_check(keys, element)
    return keys

def convert_header_list_to_dict(headers):
    headers_result = {}
    for header in headers:
        headers_result[header] = header

    return headers_result


def get_json_items(created_dealers, headers):
    items = []
    for element in created_dealers:
        item = {}
        for header in headers:
            item[header] = dictor(element, header)

        items.append(item)

    return items

def main():
    json_content = get_json_content('response.json')
    created_dealers = dictor(json_content, 'newDealersCreated')

    headers = get_json_headers(created_dealers)
    headers.sort();
    headers_to_xlsx = convert_header_list_to_dict(headers);

    items = get_json_items(created_dealers, headers)
    create_xlsx_file("output.xlsx", headers_to_xlsx, items)

if __name__ == '__main__':
    main()
