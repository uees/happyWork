#!/usr/bin/env python
#-- coding: utf-8 -*-
'''
Created on 2016年7月4日

@author: Wan
'''
import os
from openpyxl import load_workbook, Workbook
from datetime import datetime
from sqlalchemy.sql.expression import and_

from common import module_path
from database import BuliangFangan, db_session, CangkuLiushui, ProductLiushui,\
    BuliangKuchun
from win32office import Excel
from data_warp import reset_table

data_dir = os.path.join(module_path(), 'data')

def init_buliang_fangan():
    reset_table('buliang_fangan')
    wb_file = os.path.join(data_dir, '不良品处理对照表.xlsx')
    wb = load_workbook(wb_file)
    ws = wb.get_sheet_by_name('对照表')
    for row in ws.iter_rows('A2:C{}'.format(ws.max_row)):
        product_name, buliang_name, bili = row
        fangan = BuliangFangan(product_name=product_name.value,
                               buliang_name=buliang_name.value,
                               chuliliang=bili.value)
        db_session.add(fangan)
    db_session.commit()
    
def load_buliang_kuchun(file_path, start_row, stop_row):
    reset_table('buliang_kuchun')
    excel = Excel(file_path)
    excel.select()
    for row in range(start_row, stop_row+1):
        product_name = excel.get_cell_value(row=row, col=1)
        amount = excel.get_cell_value(row=row, col=2)
        kuchun = BuliangKuchun(product_name=product_name, amount=amount)
        db_session.add(kuchun)
        print('%s \t %s' % (product_name, amount))
    db_session.commit()
    excel.quit()
    
def load_cangku_liushui(file_path, start_row, stop_row):
    #reset_table('cangku_liushui')
    excel = Excel(file_path)
    excel.select('流水表')
    for row in range(start_row, stop_row+1):
        yewu_type, jilu_date, danhao, kehu, chanpin_bianma,\
        product_name, spec, batch, amount, product_date,\
        peifang_version, note = [excel.get_cell_value(row=row, col=col) for col in range(1, 13)]
        if not isinstance(jilu_date, datetime):
            jilu_date = None
        if not isinstance(product_date, datetime):
            product_date = None
        liushui = CangkuLiushui(yewu_type=yewu_type,
                          jilu_date=jilu_date,
                          danhao=danhao,
                          kehu=kehu,
                          chanpin_bianma=chanpin_bianma,
                          product_name=product_name,
                          spec=spec,
                          batch=batch,
                          amount=amount,
                          product_date=product_date,
                          peifang_version=peifang_version,
                          note=note)
        db_session.add(liushui)
        print('add %s \t %s \t %s' % (jilu_date, product_name, batch))
    db_session.commit()
    excel.quit()
    
def load_product_liushui(file_path, start_row, stop_row):
    excel = Excel(file_path)
    excel.select()
    for row in range(start_row, stop_row+1):
        (kind, product_name, batch, 
        ji_hua_zhong, pei_liao_liang, he_zhong_liang, 
        yan_mo_hou, yan_mo_sun_hao, jia_liao_liang, 
        san_lei, fan_hui_you, jia_liao_hou, 
        sheng_yu_you, ru_ku_liang, gu_hua_ji, 
        bao_zhuang_sun_hao, zong_sun_hao, sun_hao_lv, 
        wan_cheng_ri) = [excel.get_cell_value(row=row, col=col) for col in range(1, 20)]
        if not isinstance(wan_cheng_ri, datetime):
            wan_cheng_ri = None
        liushui = ProductLiushui(kind=kind, 
                                 product_name=product_name, 
                                 batch=batch, 
                                 ji_hua_zhong=ji_hua_zhong, 
                                 pei_liao_liang=pei_liao_liang, 
                                 he_zhong_liang=he_zhong_liang, 
                                 yan_mo_hou=yan_mo_hou, 
                                 yan_mo_sun_hao=yan_mo_sun_hao, 
                                 jia_liao_liang=jia_liao_liang, 
                                 san_lei=san_lei,
                                 fan_hui_you=fan_hui_you, 
                                 jia_liao_hou=jia_liao_hou, 
                                 sheng_yu_you=sheng_yu_you, 
                                 ru_ku_liang=ru_ku_liang, 
                                 gu_hua_ji=gu_hua_ji, 
                                 bao_zhuang_sun_hao=bao_zhuang_sun_hao, 
                                 zong_sun_hao=zong_sun_hao, 
                                 sun_hao_lv=sun_hao_lv, 
                                 wan_cheng_ri=wan_cheng_ri)
        db_session.add(liushui)
        print('%s \t %s \t %s' % (wan_cheng_ri, product_name, batch))
    db_session.commit()
    excel.quit()
    
def jisuan_ruku_liushui(out_file):
    excel = Excel(out_file)
    excel.select('成品入库流水')
    liushui = CangkuLiushui.query.filter(and_(CangkuLiushui.yewu_type=="产品进仓",
                                        CangkuLiushui.jilu_date>=datetime(2016, 5, 31),
                                        CangkuLiushui.jilu_date<=datetime(2016, 6, 30))).all()
    buliangfang = BuliangFangan.query.all()
    data = []
    for batch in liushui:
        flag_find = False
        jilu_date = datetime.strftime(batch.jilu_date,'%Y-%m-%d')
        for fangan in buliangfang:
            if batch.product_name.find(fangan.product_name) >= 0:
                flag_find = True
                data.append([jilu_date, batch.product_name, batch.batch, batch.amount, 
                             fangan.buliang_name, fangan.chuliliang])
                break
        if not flag_find:
            # 兑入同型号10%
            data.append([jilu_date, batch.product_name, batch.batch, batch.amount, 
                         '理论可处理同型号10%', 0.1])
    excel.set_range_value(1, 2, data)
    excel.save()
    
def jisuan_product_liushui(out_file):
    excel = Excel(out_file)
    excel.select('生产流水')
    liushui = ProductLiushui.query.filter(and_(ProductLiushui.wan_cheng_ri>=datetime(2016, 5, 31),
                                               ProductLiushui.wan_cheng_ri<=datetime(2016, 6, 30))).all()
    buliangfang = BuliangFangan.query.all()
    data = []
    for batch in liushui:
        flag_find = False
        wan_cheng_ri = datetime.strftime(batch.wan_cheng_ri,'%Y-%m-%d')
        for fangan in buliangfang:
            if batch.product_name.find(fangan.product_name) >= 0:
                flag_find = True
                data.append([wan_cheng_ri, batch.product_name, batch.batch, batch.yan_mo_hou, batch.jia_liao_hou,
                             batch.san_lei, batch.fan_hui_you, fangan.buliang_name, fangan.chuliliang])
                break
        if not flag_find:
            flag_find = False
            # 兑入同型号10%
            data.append([wan_cheng_ri, batch.product_name, batch.batch, batch.yan_mo_hou, batch.jia_liao_hou,
                         batch.san_lei, batch.fan_hui_you, '理论可处理同型号10%', 0.1])
    excel.set_range_value(1, 2, data)
    excel.save()
        
def jisuan_kunchun(out_file):
    excel = Excel(out_file)
    excel.select('不良库存')
    liushui = ProductLiushui.query.filter(and_(ProductLiushui.wan_cheng_ri>=datetime(2016, 5, 31),
                                               ProductLiushui.wan_cheng_ri<=datetime(2016, 6, 30))).all()
    buliangfang = BuliangFangan.query.all()
    kunchun = BuliangKuchun.query.all()
    data = []
    
    
if __name__ == "__main__":
    #init_buliang_fangan()
    #load_buliang_kuchun(os.path.join(data_dir, '惠东不良品仓6月份流水帐2.xlsx'), start_row=4, stop_row=141)
    #load_cangku_liushui(os.path.join(data_dir, '6月份仓库流水表.xlsx'), start_row=4, stop_row=4420)
    #load_product_liushui(os.path.join(data_dir, '6月份生产流水表.xlsx'), start_row=4, stop_row=1364)
    #jisuan_ruku_liushui(os.path.join(data_dir, '6月理论处理量.xlsx'))
    #jisuan_product_liushui(os.path.join(data_dir, '6月理论处理量.xlsx'))
    pass