#!/usr/bin/env python
#-- coding: utf-8 -*-
'''
Created on 2016年6月4日

@author: Wan
'''
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, scoped_session, relationship
from datetime import datetime

engine = create_engine("sqlite:///data/database.sdb3", convert_unicode=True)
db_session = scoped_session(sessionmaker(autocommit=False,
                                         autoflush=False,
                                         bind=engine))
Base = declarative_base(bind=engine)
Base.query = db_session.query_property()

class Product(Base):
    '''产品信息'''
    __tablename__ = 'products'
    id = Column(Integer, primary_key=True)
    internal_name = Column(String(64), default='', index=True) #内部品名（生产单品名）
    market_name = Column(String(64), default='') #销售品名
    category = Column(String(64), default='')    #类别
    template = Column(String(64), default='')    #检验报告模板
    viscosity = Column(String(64), default='0')  #检验报告粘度
    viscosity_width = Column(String(64), default='0')  #粘度幅度
    part_a = Column(String(64), default='')
    part_b = Column(String(64), default='')
    ratio = Column(Float, default=0)   #ratio = part_b/part_a
    
    
class Batch(Base):
    '''批次信息'''
    __tablename__ = 'batchs'
    id = Column(Integer, primary_key=True)
    name = Column(String(64), default='', index=True) #生产品名
    spec = Column(String(64), default='') #规格
    batch = Column(String(64), default='', index=True) #批号
    product_amount = Column(Float, default=0) #生产数量
    warehouse_amount = Column(Float, default=0) #入库数量
    product_date = Column(DateTime) #生产日期
    warehouse_date = Column(DateTime) #入库日期


class Formula(Base):
    '''配方列表'''
    __tablename__ = 'formulas'
    id = Column(Integer, primary_key=True)
    name = Column(String(64), default='')
    create_time = Column(DateTime)
    update_time = Column(DateTime, default=datetime.utcnow())
    status = Column(String(64), default='')     #状态：正式、待确认、失效
    viscosity = Column(String(64), default='0') #粘度要求
    note = Column(String(64), default='')       #注意事项
    
    materials = relationship("FormulaInfo")
    
    
class FormulaInfo(Base):
    '''配方信息'''
    __tablename__ = 'formula_info'
    id = Column(Integer, primary_key=True)
    material_name = Column(String(64), default='') #材料名
    material_weight = Column(Float, default=0)     #材料重量 kg
    material_ratio = Column(Float, default=0)      #材料重量比例 %
    material_volume = Column(Float, default=0)     #材料体积 ml
    material_area = Column(String(64), default='') #材料区域
    material_note = Column(String(64), default='') #注意事项
    formula_id = Column(Integer, ForeignKey('formulas.id'))
    
    
class Material(Base):
    '''材料'''
    __tablename__ = 'materials'
    id = Column(Integer, primary_key=True)
    name = Column(String(64), default='') #材料名
    note = Column(String(64), default='') #备注
    
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
    product_name = Column(String(64), default='') #品名
    batch = Column(String(64), default='') #批号
    ji_hua_zhong = Column(Float, default=0) #计划重
    pei_liao_liang = Column(Float, default=0) #配料量
    he_zhong_liang = Column(Float, default=0) #核重量
    yan_mo_hou = Column(Float, default=0) #研磨后
    yan_mo_sun_hao = Column(Float, default=0) #研磨损耗
    jia_liao_liang = Column(Float, default=0) #加料量
    san_lei = Column(Float, default=0) #三类
    fan_hui_you = Column(Float, default=0) #返回油
    jia_liao_hou = Column(Float, default=0) #加料后
    sheng_yu_you = Column(Float, default=0) #剩余油
    ru_ku_liang = Column(Float, default=0) #入库量
    gu_hua_ji = Column(Float, default=0) #固化剂量
    bao_zhuang_sun_hao = Column(Float, default=0) #包装损耗
    zong_sun_hao = Column(Float, default=0) #总损耗
    sun_hao_lv = Column(Float, default=0) #损耗率
    wan_cheng_ri = Column(DateTime)
    
class BuliangKuchun(Base):
    '''不良库存量'''
    __tablename__ = 'buliang_kuchun'
    id = Column(Integer, primary_key=True)
    product_name = Column(String(64), default='')
    amount = Column(Float, default=0)
    
def get_table_class(name):
    table_class = dict(materials=Material,
                       formula_info=FormulaInfo,
                       formulas=Formula,
                       batchs=Batch,
                       products=Product,
                       buliang_fangan=BuliangFangan,
                       cangku_liushui=CangkuLiushui,
                       product_liushui=ProductLiushui,
                       buliang_kuchun=BuliangKuchun)
    return table_class[name]