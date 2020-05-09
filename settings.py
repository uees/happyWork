# -*- coding: utf-8 -*-
import os

from dotenv import load_dotenv


BASE_DIR = os.path.abspath(os.path.dirname(__file__))

load_dotenv(dotenv_path=os.path.join(BASE_DIR, '.env'))

APP_NAME = "QcReport"

APP_VERSION = "v0.1"

ALL_FQC_ITEMS = [
    '外观颜色', '细度', '反白条', '粘度', '板面效果',
    '固化性', '硬度', '附着力', '感光性', '显影性',
    '解像性', '去膜性', '耐焊性', '耐化学性', '红外图谱',
]

FQC_ITEMS = dict(
    h9100=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '感光性',
        '显影性', '耐焊性', '耐化学性', '红外图谱'
    ],
    h8100=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '感光性',
        '显影性', '耐焊性', '耐化学性', '红外图谱'
    ],
    a9=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '感光性', '显影性', '解像性',
        '去膜性', '红外图谱'
    ],
    a2=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '附着力', '去膜性', '红外图谱'
    ],
    k2=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '附着力', '去膜性', '红外图谱'
    ],
    tm3100=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'
    ],
    ts3000=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'
    ],
    uvs1000=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'
    ],
    uvm1800=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果',
        '固化性', '硬度', '附着力', '耐焊性', '耐化学性', '红外图谱'
    ],
    uvw5d10=[
        '外观颜色', '细度', '反白条', '粘度',
        '板面效果', '固化性', '硬度', '附着力', '红外图谱'
    ],
    uvw5d65=[
        '外观颜色', '细度', '反白条', '粘度',
        '板面效果', '固化性', '附着力', '去膜性', '红外图谱'
    ],
    uvw5d35=[
        '外观颜色', '细度', '反白条', '粘度', '板面效果', '固化性',
        '附着力', '去膜性', '红外图谱'
    ],
)

CONF = [
    {
        "slug": "lpi_360gs",
        "name": "360树脂",
        "template": "LPI-360GS",
    },
    {
        'slug': 'h9100',
        'name': '感光阻焊油墨',
        'template': 'H-9100_感光阻焊油墨',
        'alias': ['H-9100'],
    },
    {
        'slug': 'h9100_22c',  # 22 摄氏度粘度模版
        'name': '感光阻焊油墨',
        'template': 'H-9100_感光阻焊油墨_22C',
    },
    {
        'slug': 'h8100',
        'name': '感光阻焊油墨',
        'template': 'H-8100_感光阻焊油墨',
        'alias': ['H-8100'],
    },
    {
        'slug': 'h8100_22c',  # 22 摄氏度粘度模版
        'name': '感光阻焊油墨',
        'template': 'H-8100_感光阻焊油墨_22C',
    },
    {
        "slug": "h8100_ldi",
        "name": "LDI 感光阻焊油墨",
        "template": "LDIBL01",
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
        'slug': 'h8100_bomin',
        'template': 'H-8100_博敏',
        'name': '感光阻焊油墨',
        'customer': '博敏',
    },
    {
        'slug': 'h9100_bomin',
        'template': 'H-9100_博敏',
        'name': '感光阻焊油墨',
        'customer': '博敏',
    },
    {
        'slug': 'h9100_weg',
        'template': 'H-9100_威尔高',
        'name': '感光阻焊油墨',
        'customer': '威尔高',
    },
    {
        'slug': 'h8100_weg',
        'template': 'H-8100_威尔高',
        'name': '感光阻焊油墨',
        'customer': '威尔高',
    },
    {
        'slug': 'h9100_jw',
        'template': 'H-9100_景旺',
        'name': '感光阻焊油墨',
        'customer': '景旺',
    },
    {
        'slug': 'h9100_jw22',
        'template': 'H-9100_景旺_22C',
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
        'slug': 'a9060c_hbjd',
        'template': 'A-9060C_内层湿膜_湖北健鼎',
        'name': '感光线路油墨内层',
        'customer': '湖北健鼎',
    },
    {
        'slug': 'a9060c01',
        'template': 'A-9060C01_有色',
        'name': '感光线路油墨内层',
    },
    {
        'slug': 'a9060c0101',
        'template': 'A-9060C0101_有色',
        'name': '感光线路油墨内层',
        'customer': '江西景旺',
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
        'slug': 'tm3100_mls',
        'template': 'TM-3100_木林森',
        'name': '热固化文字油墨',
    },
    {
        'slug': 'uvw5d10',
        'template': 'UVW-5D10',
    },
    {
        'slug': 'uvw5d110d1',
        'template': 'UVW-5D110D1',
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
        'slug': 'uvw5d65d28a',
        'template': 'UVW-5D65D28A',
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
        'slug': 'thw5d37',
        'template': 'THW-5D37',
    },
    {
        'slug': 'thw4d46',
        'template': 'THW-4D46',
    },
    {
        'slug': 'thw6102m17',
        'template': 'THW-6102M17',
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
        'slug': 'xsj_amount',
        'template': '稀释剂_数量',
        'alias': ['XSJ_AMOUNT'],
    },
    {
        'slug': 'xd',
        'template': 'XD704',
    },
    {
        'slug': 'rd391',
        'template': 'RD-391',
    },
    {
        'slug': 'rdi6000',
        'template': 'RDI-6000',
    },
    {
        'slug': 'rdi6030',
        'template': 'RDI-6030',
    },
    {
        'slug': 'rdi6040gl',
        'template': 'RDI-6040GL01',
    },
    {
        'slug': 'rdi600001',
        'template': 'RDI-600001',
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

SPECS = {
    "1": "1kg±5g",
    "4": "4kg±20g",
    "5": "5kg±20g",
    "15": "15kg±50g",
    "20": "20kg±50g",
}
