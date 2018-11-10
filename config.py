# -*- coding: utf-8 -*-

APP_NAME = "QcReport"

APP_VERSION = "v0.1"

ALL_FQC_ITEMS = ['外观颜色', '细度', '反白条', '粘度', '板面效果', '固化性', '硬度',
                 '附着力', '感光性', '显影性', '解像性', '去膜性', '耐焊性', '耐化学性', '红外图谱']

FQC_ITEMS = dict(h9100=['外观颜色', '细度', '反白条', '粘度', '板面效果', '固化性', '硬度', '附着力', '感光性',
                        '显影性', '耐焊性', '耐化学性', '红外图谱'],
                 h8100=['外观颜色', '细度', '反白条', '粘度', '板面效果', '固化性', '硬度', '附着力', '感光性',
                        '显影性', '耐焊性', '耐化学性', '红外图谱'],
                 a9=['外观颜色', '细度', '反白条', '粘度', '板面效果', '固化性', '硬度',
                     '附着力', '感光性', '显影性', '解像性', '去膜性', '红外图谱'],
                 a2=['外观颜色', '细度', '反白条', '粘度', '板面效果',
                     '固化性', '附着力', '去膜性', '红外图谱'],
                 k2=['外观颜色', '细度', '反白条', '粘度', '板面效果',
                     '固化性', '附着力', '去膜性', '红外图谱'],
                 tm3100=['外观颜色', '细度', '反白条', '粘度', '板面效果',
                         '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'],
                 ts3000=['外观颜色', '细度', '反白条', '粘度', '板面效果',
                         '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'],
                 uvs1000=['外观颜色', '细度', '反白条', '粘度', '板面效果',
                          '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'],
                 uvm1800=['外观颜色', '细度', '反白条', '粘度', '板面效果',
                          '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'],
                 uvw5d10=['外观颜色', '细度', '反白条', '粘度',
                          '板面效果', '固化性', '硬度', '附着力', '红外图谱'],
                 uvw5d65=['外观颜色', '细度', '反白条', '粘度',
                          '板面效果', '固化性', '附着力', '去膜性', '红外图谱'],
                 uvw5d35=['外观颜色', '细度', '反白条', '粘度', '板面效果', '固化性', '附着力', '去膜性', '红外图谱'])

CONF = [
    {
        'slug': 'h9100',
        'name': '感光阻焊油墨',
        'template': 'H-9100_感光阻焊油墨',
        'alias': ['H-9100'],
    },
    {
        'slug': 'h8100',
        'name': '感光阻焊油墨',
        'template': 'H-8100_感光阻焊油墨',
        'alias': ['H-8100'],
    },
    {
        'slug': 'h9100d',
        'name': '感光阻焊油墨',
        'template': 'H-9100D',
        'alias': ['H-9100D'],
    },
    {
        'slug': 'h9100_jx',
        'name': '感光阻焊油墨',
        'template': 'H-9100_金像',
        'customer': '金像',
    },
    {
        'slug': 'h8100_jx',
        'template': 'H-8100_金像',
        'name': '感光阻焊油墨',
        'customer': '金像',
    },
    {
        'slug': 'h9100_jd',
        'template': 'H-9100_健鼎',
        'name': '感光阻焊油墨',
        'customer': '健鼎',
    },
    {
        'slug': 'h8100_jd',
        'template': 'H-8100_健鼎',
        'name': '感光阻焊油墨',
        'customer': '健鼎',
    },
    {
        'slug': 'h9100_fsk',
        'template': 'H-9100_烟台富士康',
        'name': '感光阻焊油墨',
    },
    {
        'slug': 'h9100_cd',
        'template': 'H-9100_崇达',
        'name': '感光阻焊油墨',
        'customer': '崇达',
        'ext_info': '(要求打发货数量)',
    },
    {
        'slug': 'h8100_cd',
        'template': 'H-8100_崇达',
        'name': '感光阻焊油墨',
        'customer': '崇达',
        'ext_info': '(要求打发货数量)',
    },
    {
        'slug': 'h9100_jy_sk',
        'template': 'H-9100_SK_建业',
        'name': '感光阻焊油墨',
        'customer': '建业',
    },
    {
        'slug': 'h8100_jy_sk',
        'template': 'H-8100_SK_建业',
        'name': '感光阻焊油墨',
        'customer': '建业',
    },
    {
        'slug': 'h9100_dj',
        'template': 'H-9100_达进',
        'name': '感光阻焊油墨',
        'customer': '达进',
    },
    {
        'slug': 'h8100_dj',
        'template': 'H-8100_达进',
        'name': '感光阻焊油墨',
        'customer': '达进',
    },
    {
        'slug': 'h9100_jw',
        'template': 'H-9100_景旺',
        'name': '感光阻焊油墨',
        'customer': '景旺',
    },
    {
        'slug': 'h9100_ntsn',
        'template': 'H-9100_南通深南',
        'name': '感光阻焊油墨',
        'customer': '南通深南',
    },
    {
        'slug': 'h8100_ntsn',
        'template': 'H-8100_南通深南',
        'name': '感光阻焊油墨',
        'customer': '南通深南',
    },
    {
        'slug': 'a9060a',
        'template': 'A-9060A_内层湿膜',
        'alias': ['A-9060A', 'A-9060B', 'A-9060C', 'A-9060D'],
        'name': '感光线路油墨内层',
    },
    {
        'slug': 'a9060a_sn',
        'template': 'A-9060A_内层湿膜_深南',
        'name': '感光线路油墨内层',
    },
    {
        'slug': 'a9060a_jd',
        'template': 'A-9060B_内层湿膜_健鼎',
        'name': '感光线路油墨内层',
        'customer': '健鼎',
    },
    {
        'slug': 'a9060a_my',
        'template': 'A-9060C_有色_明阳',
        'name': '感光线路油墨内层',
        'customer': '明阳',
    },
    {
        'slug': 'a9060a_jw',
        'template': 'A-9060C01_景旺',
        'name': '感光线路油墨内层',
        'customer': '景旺',
    },
    {
        'slug': 'a9000',
        'template': 'A-9000_外层湿磨',
        'alias': ['A-9000', 'A-9060', 'A-9070', 'A-9080'],
        'name': '感光线路油墨外层',
    },
    {
        'slug': 'a2000',
        'template': 'A-2000_耐酸抗蚀油墨',
        'alias': ['A-2000', 'A-2000BK', 'A-2000BL'],
        'name': '耐酸抗蚀油墨',
    },
    {
        'slug': 'a2100',
        'template': 'A-2100_紫外光固化线路油墨',
        'alias': ['A-2100', 'A-2100F', 'A-2100G'],
        'name': '紫外光固化线路油墨',
    },
    {
        'slug': 'k2500',
        'template': 'K-2500_耐碱抗蚀油墨',
        'alias': ['K-2500', 'K-2500BL', 'K-2500BK'],
        'name': '耐碱抗蚀油墨',
    },
    {
        'slug': 'uvs1000',
        'template': 'UVS-1000_紫外光固化阻焊油墨',
        'alias': ['UVS-1000'],
        'name': '紫外光固化阻焊油墨',
    },
    {
        'slug': 'uvm1800',
        'template': 'UVM-1800_紫外光固化字符油墨',
        'alias': ['UVM-1800'],
        'name': '紫外光固化字符油墨',
    },
    {
        'slug': 'ts3000',
        'template': 'TS-3000_热固化保护油',
        'alias': ['TS-3000'],
        'name': '热固化保护油',
    },
    {
        'slug': 'tm3100',
        'template': 'TM-3100_热固化文字油墨',
        'alias': ['TM-3100'],
        'name': '热固化文字油墨',
    },
    {
        'slug': 'uvw5d10',
        'template': 'UVW-5D10',
    },
    {
        'slug': 'uvw5d65',
        'template': 'UVW-5D65',
    },
    {
        'slug': 'uvw5d65d28',
        'template': 'UVW-5D65D28',
    },
    {
        'slug': 'uvw5d85',
        'template': 'UVW-5D85',
    },
    {
        'slug': 'thw5d35',
        'template': 'THW-5D35',
    },
    {
        'slug': 'di7j84',
        'template': 'DI-7J84',
    },
    {
        'slug': 'xsj',
        'template': '稀释剂',
        'alias': ['XSJ'],
    },
    {
        'slug': 'xd',
        'template': 'XD704',
    },
    {
        'slug': 'rdj',
        'template': 'COA_RDJ',
        'alias': ['RDJ'],
    },
    {
        'slug': 'rdi6000',
        'template': 'RDI-6000',
    },
    {
        'slug': 'rdi600003',
        'template': 'RDI-600003',
    },
    {
        'slug': 'other',
        'template': '未定义',
    },
    {
        'slug': 'wujin',
        'template': '未定义',
    },
]

SPECS = {"1": "1kg±5g",
         "4": "4kg±20g",
         "5": "5kg±20g",
         "15": "15kg±50g",
         "20": "20kg±50g"}

formula = {
    'name': '6GHB',
    'version': 'B-01',
    'category': 'H-9100',
    'common_name': '湿绿油',
    'description': '无卤素',

    'mixing_note': '配料要求',

    'grind_times': '3',  # 研磨次数
    'grind_temperature': '<=45',  # 出料温度要求
    'grind_machine': '三辊机',  # 研磨设备
    'grind_granule': '<=20um',  # 研磨细度
    'grind_speed': '120kg/h',  # 研磨速度
    'grind_note': '研磨要求',  # 研磨其他要求

    'viscosity': "260~270dpas/25℃",
    'after_adding_note': '加料要求',

    'package_machine': '升降机',  # 包装过滤设备方式
    'package_bag': '100T',  # 过滤袋规格
    'package_specification': '5kg',
    'package_ratio': '3:1',
    'package_part_b': 'HD21',
    'package_label': '160dpas,两头贴',  # 标签要求
    'package_note': '包装其他要求',  # 包装要求

    'materials': dict(name='', amount='', workshop='', note=''),
}
