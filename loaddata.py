import argparse
import json
import re

import requests
from openpyxl import load_workbook


def login(host, email, password):
    response = requests.post(f'{host}/api/auth/login', data={
        'email': email, 'password': password
    }, headers={
        'X-Requested-With': 'XMLHttpRequest'
    })

    body = response.json()

    return body.get('data').get('token')


def entering_warehouse(host, token, product_name, product_batch, spec, weight, entered_at, made_at):
    recyclable_type = ''
    if product_name.find('SP') >= 0 and product_name.find('内袋') >= 0:
        recyclable_type = 'bucket'
    elif product_name.find('A-9060') >= 0 and product_name.find('内袋') >= 0:
        recyclable_type = 'bucket'
    elif product_name.find('内袋') >= 0:
        recyclable_type = 'box'

    amount = 0
    match = re.match(r'^\d+\.?\d+', spec)
    if match:
        per_weight = float(match.group())
        if recyclable_type == 'box' and per_weight < 10:
            per_weight = 10
        amount = int(weight/per_weight)

    response = requests.post(f'{host}/api/entering-warehouses', data={
        'product_name': product_name,
        'product_batch': product_batch,
        'spec': spec,
        'weight': weight,
        'entered_at': entered_at,
        'made_at': made_at,
        'recyclable_type': recyclable_type,
        'amount': amount,
    }, headers={
        'X-Requested-With': 'XMLHttpRequest',
        'Authorization': f'Bearer {token}'
    })


def read_file(filename):
    wb = load_workbook(filename)
    ws = wb['成品流水账']

    for row in ws[f'A3:J{ws.max_row}']:
        _type, stored_at, NO, custmor, code, name, spec, batch, weight, made_at = row
        if _type.value == "产品进仓":
            print(_type, stored_at, NO, custmor, code, name, spec, batch, weight, made_at)
            r = requests.post('http://httpbin.org/post', data={'key': 'value'})
        elif _type.value == "往来销售":
            pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='导入流水数据')
    parser.add_argument("excel_file", help="文件名")
    args = parser.parse_args()
    # args.excel_file
