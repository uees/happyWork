# -*- coding: utf-8 -*-
'''
Created on 2015年7月10日

@author: Wan
'''
from _functools import reduce
from openpyxl import load_workbook
from sqlalchemy import and_
from database import Product, db_session, Base, get_table_class


def search_product(keywords):
    criteria = []
    for keyword in keywords.split():
        criteria.append(Product.internal_name.ilike('%{}%'.format(keyword)))
    q = reduce(and_, criteria)
    return Product.query.filter(q)


def get_product_by_id(pid):
    return Product.query.filter_by(id=pid).first()


def insert_product(product):
    db_session.add(product)
    db_session.commit()


def insert_product_to_xlsx(product, file, sheet):
    wb = load_workbook(file)
    ws = wb.get_sheet_by_name(sheet)
    index = len(ws.rows)+1
    ws.cell("A{}".format(index), value=product.internal_name)
    ws.cell("B{}".format(index), value=product.template)
    ws.cell("C{}".format(index), value=product.viscosity)
    ws.cell("D{}".format(index), value=product.viscosity_width)
    ws.cell("E{}".format(index), value=product.market_name)

   
def fetch_product_data(file, sheet):
    wb = load_workbook(filename=file)
    ws = wb.get_sheet_by_name(sheet)
    for row in ws.iter_rows('A2:I{}'.format(ws.max_row)):
        internal_name,template,viscosity,viscosity_width,market_name,category,part_a,part_b,ratio = row
        product = Product(internal_name=internal_name.value,
                          template=template.value,
                          viscosity=viscosity.value,
                          viscosity_width=viscosity_width.value,
                          market_name=market_name.value,
                          category=category.value,
                          part_a=part_a.value,
                          part_b=part_b.value,
                          ratio=ratio.value)
        db_session.add(product)
    db_session.commit()
    print("插入了 %s行数据到data/info.db." % ws.max_row)


def init_database():
    Base.metadata.drop_all() 
    Base.metadata.create_all()


def reset_table(tableClass):
    if isinstance(tableClass, str):
        tableClass = get_table_class(tableClass)
    tableClass.__table__.drop(checkfirst=True)
    tableClass.__table__.create()