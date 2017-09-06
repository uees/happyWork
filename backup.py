# -*- coding: utf-8 -*-
import argparse
import logging
import os
import shutil
from datetime import datetime
import time

import win32com.client
from pywintypes import com_error
from sqlalchemy import Column, Integer, String, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import scoped_session, sessionmaker

# inits
logging.basicConfig(
    format='[%(asctime)s][%(name)s][%(levelname)s] - %(message)s',
    datefmt='%Y-%d-%m %H:%M:%S',
    level=logging.DEBUG,
    filename='backup.log',
    filemode='w'
)
# 定义一个StreamHandler，将INFO级别或更高的日志信息打印到标准错误，并将其添加到当前的日志处理对象
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
console.setFormatter(formatter)
log = logging.getLogger('BACKUP.lib')
log.addHandler(console)

backup_dir = 'E:/备份'
if not os.path.exists(backup_dir):
    os.mkdir(backup_dir)

engine = create_engine("sqlite:///%s/backup.db" % backup_dir, convert_unicode=True)
db_session = scoped_session(sessionmaker(autocommit=False,
                                         autoflush=False,
                                         bind=engine))
Base = declarative_base(bind=engine)
Base.query = db_session.query_property()


def get_table_class(name):
    models = [obj for obj in globals().values() if hasattr(obj, '__tablename__')]
    for model in models:
        if model.__tablename__ == name:
            return model


def reset_table(tableClass):
    if isinstance(tableClass, str):
        tableClass = get_table_class(tableClass)
    tableClass.__table__.drop(checkfirst=True)
    tableClass.__table__.create()


def init_database():
    Base.metadata.drop_all()
    Base.metadata.create_all()

    # 初始化一些数据
    db_session.bulk_insert_mappings(
        Option,
        [dict(name='sheet%s_start' % i, value='5')
            for i in range(1, 4)]
    )
    db_session.commit()


class Manager(object):

    def __init__(self):
        self.MSOffice = win32com.client.Dispatch('Excel.Application')
        self.filefrom = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'FQC检测记录表格.xlsx')

        try:
            self.wb = self.MSOffice.Workbooks(self.filefrom)
        except com_error:
            self.wb = self.MSOffice.Workbooks.Open(self.filefrom)

        self.config = [
            {
                'name': 'FQC阻焊表格',
                'start': int(Option.at('sheet1_start')),
                'done_column': 18,
                'num_columns': 28,
                'model': Sheet1,
                'fileds': [
                    'date_at',
                    'time_at',
                    'name',
                    'lot_NO',
                    'nian_du6',
                    'nian_du60',
                    'color',
                    'xi_du',
                    'liu_you',
                    'liu_you2',
                    'hong_wai',
                    'ban_mian',
                    'gu_hua',
                    'phototonus',
                    'show_shadow40',
                    'show_shadow60',
                    'show_shadow70',
                    'show_shadow80',
                    'show_shadow12h',
                    'show_shadow24h',
                    'bridge_line',
                    'tester',
                    'hardness',
                    'fu_zhuo_li',
                    'nai_han_xing',
                    'nai_hua_xing',
                    'gui_ying',
                    'show_shadow_sk',
                ]
            },
            {
                'name': 'FQC湿膜表格',
                'start': int(Option.at('sheet2_start')),
                'done_column': 16,
                'num_columns': 22,
                'model': Sheet2,
                'fileds': [
                    'date_at',
                    'time_at',
                    'name',
                    'lot_NO',
                    'nian_du',
                    'color',
                    'xi_du',
                    'hong_wai',
                    'ban_mian',
                    'gu_hua',
                    'hardness',
                    'fu_zhuo_li',
                    'zhen_kong',
                    'dieban',
                    'phototonus',
                    'show_shadow',
                    'phototonus24h',
                    'show_shadow24h',
                    'jie_xiang',
                    'qu_mo',
                    'nai_hua_xing',
                    'tester',
                ],
            },
            {
                'name': 'FQC其他油墨表格',
                'start': int(Option.at('sheet3_start')),
                'done_column': 17,
                'num_columns': 17,
                'model': Sheet3,
                'fileds': [
                    'date_at',
                    'time_at',
                    'name',
                    'lot_NO',
                    'nian_du',
                    'color',
                    'xi_du',
                    'hong_wai',
                    'ban_mian',
                    'gu_hua',
                    'hardness',
                    'fu_zhuo_li',
                    'show_shadow',
                    'qu_mo',
                    'nai_han_xing',
                    'nai_hua_xing',
                    'tester',
                ],
            },
        ]

    def backup2file(self):
        fileto = os.path.join(
            backup_dir,
            '%s fqc_backup.xlsx' % datetime.strftime(datetime.now(), '%Y-%m-%d %H-%M-%S')
        )
        shutil.copy(self.filefrom, fileto)

    def backup2db(self):
        for idx, conf in enumerate(self.config):
            ws = self.wb.Worksheets(conf['name'])
            start = conf['start']
            num_columns = conf['num_columns']
            done_row = self.get_done_row(ws, conf['done_column'], start)

            data = ws.Range(ws.Cells(start, 1), ws.Cells(done_row, num_columns)).Value

            objects = []
            for row in data:
                row = list(row)
                # 修复日期
                if isinstance(row[0], datetime):
                    row[0] = datetime.strftime(row[0], '%Y-%m-%d')
                # fix 批号, 有些色浆的批号是日期
                if isinstance(row[3], datetime):
                    row[3] = datetime.strftime(row[3], '%Y-%m-%d')
                objects.append(conf['model'](**dict(zip(conf['fileds'], row))))

            db_session.bulk_save_objects(objects)

            Option.query.filter_by(name='sheet%s_start' % (idx + 1,)).update({'value': done_row + 1})

        db_session.commit()

    def get_done_row(self, ws, done_col, start):
        max_row = ws.UsedRange.Rows.Count
        for row in range(max_row, start - 1, -1):
            done_cell = ws.Cells(row, done_col)
            if done_cell.Value:
                return row

    def fix_data(self, row, idx):
        if isinstance(row[idx], datetime):
            row = list(row)
            row[idx] = datetime.strftime(row[idx], '%Y-%m-%d')
        return row

    def run(self):
        self.backup2file()
        try:
            self.backup2db()
        except:
            import traceback

            log.error(traceback.format_exc())


class Option(Base):
    __tablename__ = 'options'
    id = Column(Integer, primary_key=True)
    name = Column(String(64), default='', index=True)
    value = Column(String(255), default='')

    @classmethod
    def at(cls, name):
        op = cls.query.filter_by(name=name).first()
        if op:
            return op.value


class Sheet1(Base):
    __tablename__ = 'sheet1'

    id = Column(Integer, primary_key=True)
    date_at = Column(String(64), default='')
    time_at = Column(String(64), default='')
    name = Column(String(64), default='')
    lot_NO = Column(String(64), default='')
    nian_du6 = Column(String(64), default='')
    nian_du60 = Column(String(64), default='')
    color = Column(String(64), default='')
    xi_du = Column(String(64), default='')
    liu_you = Column(String(64), default='')
    liu_you2 = Column(String(64), default='')
    hong_wai = Column(String(64), default='')
    ban_mian = Column(String(64), default='')
    gu_hua = Column(String(64), default='')
    phototonus = Column(String(64), default='')
    show_shadow40 = Column(String(64), default='')
    show_shadow60 = Column(String(64), default='')
    show_shadow70 = Column(String(64), default='')
    show_shadow80 = Column(String(64), default='')
    show_shadow12h = Column(String(64), default='')
    show_shadow24h = Column(String(64), default='')
    bridge_line = Column(String(64), default='')
    tester = Column(String(64), default='')
    hardness = Column(String(64), default='')
    fu_zhuo_li = Column(String(64), default='')
    nai_han_xing = Column(String(64), default='')
    nai_hua_xing = Column(String(64), default='')
    gui_ying = Column(String(64), default='')
    show_shadow_sk = Column(String(64), default='')


class Sheet2(Base):
    __tablename__ = 'sheet2'

    id = Column(Integer, primary_key=True)
    date_at = Column(String(64), default='')
    time_at = Column(String(64), default='')
    name = Column(String(64), default='')
    lot_NO = Column(String(64), default='')
    nian_du = Column(String(64), default='')
    color = Column(String(64), default='')
    xi_du = Column(String(64), default='')
    hong_wai = Column(String(64), default='')
    ban_mian = Column(String(64), default='')
    gu_hua = Column(String(64), default='')
    hardness = Column(String(64), default='')
    fu_zhuo_li = Column(String(64), default='')
    zhen_kong = Column(String(64), default='')
    dieban = Column(String(64), default='')
    phototonus = Column(String(64), default='')
    show_shadow = Column(String(64), default='')
    phototonus24h = Column(String(64), default='')
    show_shadow24h = Column(String(64), default='')
    jie_xiang = Column(String(64), default='')
    qu_mo = Column(String(64), default='')
    nai_hua_xing = Column(String(64), default='')
    tester = Column(String(64), default='')


class Sheet3(Base):
    __tablename__ = 'sheet3'

    id = Column(Integer, primary_key=True)
    date_at = Column(String(64), default='')
    time_at = Column(String(64), default='')
    name = Column(String(64), default='')
    lot_NO = Column(String(64), default='')
    nian_du = Column(String(64), default='')
    color = Column(String(64), default='')
    xi_du = Column(String(64), default='')
    hong_wai = Column(String(64), default='')
    ban_mian = Column(String(64), default='')
    gu_hua = Column(String(64), default='')
    hardness = Column(String(64), default='')
    fu_zhuo_li = Column(String(64), default='')
    show_shadow = Column(String(64), default='')
    qu_mo = Column(String(64), default='')
    nai_han_xing = Column(String(64), default='')
    nai_hua_xing = Column(String(64), default='')
    tester = Column(String(64), default='')


# fix first run error
if not os.path.exists(os.path.join(backup_dir, "backup.db")):
    init_database()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--initdb", action="store_true", default=False, help="初始化数据库")

    args = parser.parse_args()
    if args.initdb:
        init_database()
    else:
        print("正在备份检测记录...\n请不要关闭此窗口\n\n")
        Manager().run()
        time.sleep(2)
        print("备份成功，O(∩_∩)O")
        time.sleep(2)
