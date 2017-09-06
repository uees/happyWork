#!/usr/bin/env python
# -- coding: utf-8 -*-
'''
Created on 2016年6月4日

@author: Wan
'''
from datetime import datetime

from sqlalchemy import (Column, DateTime, Float, ForeignKey, Integer, String,
                        create_engine)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship, scoped_session, sessionmaker

engine = create_engine("sqlite:///data/database.sdb3", convert_unicode=True)
db_session = scoped_session(sessionmaker(autocommit=False,
                                         autoflush=False,
                                         bind=engine))
Base = declarative_base(bind=engine)
Base.query = db_session.query_property()


class Hebing(Base):
    __tablename__ = 'hebing'
    id = Column(Integer, primary_key=True)
    code = Column(String)
    customer = Column(String)

    cg_2012 = Column(Float)
    yf_2012 = Column(Float)
    ye_2012 = Column(Float)

    cg_2013 = Column(Float)
    yf_2013 = Column(Float)
    ye_2013 = Column(Float)

    cg_2014 = Column(Float)
    yf_2014 = Column(Float)
    ye_2014 = Column(Float)

    cg_2015 = Column(Float)
    yf_2015 = Column(Float)
    ye_2015 = Column(Float)

    cg_2016 = Column(Float)
    yf_2016 = Column(Float)
    ye_2016 = Column(Float)


class IQCMaterial(Base):
    __tablename__ = 'iqc_materials'
    id = Column(Integer, primary_key=True)
    name = Column(String(64))
    qc_items = Column(String(250))
    spec = Column(String(250))
    qc_method = Column(String(250))


class Product(Base):
    '''产品信息'''
    __tablename__ = 'products'
    id = Column(Integer, primary_key=True)
    internal_name = Column(String(64), default='', index=True)  # 内部品名（生产单品名）
    market_name = Column(String(64), default='')  # 销售品名
    category = Column(String(64), default='')  # 类别
    template = Column(String(64), default='')  # 检验报告模板
    viscosity = Column(String(64), default='0')  # 检验报告粘度
    viscosity_width = Column(String(64), default='0')  # 粘度幅度
    part_a = Column(String(64), default='')
    part_b = Column(String(64), default='')
    ratio = Column(Float, default=0)  # ratio = part_b/part_a


class BuliangFangan(Base):
    '''不良品处理方案'''
    __tablename__ = 'buliang_fangan'
    id = Column(Integer, primary_key=True)
    product_name = Column(String(64), default='')
    buliang_name = Column(String(64), default='')
    chuliliang = Column(Float, default=0)


class CangkuLiushui(Base):
    '''仓库流水表
    ALTER TABLE liushui RENAME TO product_liushui
    '''
    __tablename__ = 'cangku_liushui'
    id = Column(Integer, primary_key=True)
    yewu_type = Column(String(64), default='')
    jilu_date = Column(DateTime)
    danhao = Column(String(64), default='')
    kehu = Column(String(64), default='')
    chanpin_bianma = Column(String(64), default='')
    product_name = Column(String(64), default='')
    spec = Column(String(64), default='')
    batch = Column(String(64), default='')
    amount = Column(String(64), default='')
    product_date = Column(DateTime)
    peifang_version = Column(String(64), default='')
    note = Column(String(64), default='')


class ProductLiushui(Base):
    '''生产流水表
    '''
    __tablename__ = 'product_liushui'
    id = Column(Integer, primary_key=True)
    kind = Column(String(64), default='')  # 类别
    product_name = Column(String(64), default='')  # 品名
    product_date = Column(DateTime)
    batch = Column(String(64), default='')  # 批号
    ji_hua_zhong = Column(Float, default=0)  # 计划重
    pei_liao_liang = Column(Float, default=0)  # 配料量
    he_zhong_liang = Column(Float, default=0)  # 核重量
    yan_mo_hou = Column(Float, default=0)  # 研磨后
    yan_mo_sun_hao = Column(Float, default=0)  # 研磨损耗
    jia_liao_liang = Column(Float, default=0)  # 加料量
    san_lei = Column(Float, default=0)  # 三类
    fan_hui_you = Column(Float, default=0)  # 返回油
    jia_liao_hou = Column(Float, default=0)  # 加料后
    sheng_yu_you = Column(Float, default=0)  # 剩余油
    guan_shu = Column(String(32), default='')
    gui_ge = Column(String(32), default='')
    ru_ku_liang = Column(Float, default=0)  # 入库量
    gui_ge_2 = Column(String(32), default='')
    gu_hua_ji = Column(Float, default=0)  # 固化剂量
    bao_zhuang_sun_hao = Column(Float, default=0)  # 包装损耗
    zong_sun_hao = Column(Float, default=0)  # 总损耗
    sun_hao_lv = Column(Float, default=0)  # 损耗率
    wan_cheng_ri = Column(DateTime)


class BuliangKuchun(Base):
    '''不良库存量'''
    __tablename__ = 'buliang_kuchun'
    id = Column(Integer, primary_key=True)
    product_name = Column(String(64), default='')
    amount = Column(Float, default=0)


class ProductClassification(Base):
    """ 财务对账使用的产品分类表 """
    __tablename__ = 'product_classification'
    id = Column(Integer, primary_key=True)
    part_id = Column(Integer, default=0)
    accounting_classification = Column(String(128), default="")  # 记账分类
    costing_classification = Column(String(128), default="")  # 成本核算分类
    new_costing_classification = Column(String(128), default="")  # 新成本核算分类
    model_name = Column(String(64), default="", index=True)  # 型号
    slug = Column(String(128), default="", index=True)  # 简称
    unit = Column(String(16), default="kg")  # 结算单位
    product_code = Column(String(32), default="00000000**")
    create_time = Column(DateTime)  # 新增时间
    delete_time = Column(DateTime)  # 取消时间
    note = Column(String(250), default="")   # 备注
    people_name = Column(String(16), default="")  # 工程师


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
