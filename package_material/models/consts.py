# kind -> package category
KIND_PACKAGES = {
    'undefined': {},
    'H-8100': {
        '5kg': '5kg双组分包装',
        '10kg': '10kg双组分包装',
        '20kg': '20kg双组分包装',
        '10kg内袋': '10kg双组分内袋包装',
        '20kg内袋': '20kg双组分内袋包装',
        '20kg固内': '20kg双组分固化剂内袋包装',
        '25kg内袋': '25kg双组分内袋包装',
    },
    'H-9100': {
        '5kg': '5kg双组分包装',
        '10kg': '10kg双组分包装',
        '20kg': '20kg双组分包装',
        '10kg内袋': '10kg双组分内袋包装',
        '20kg内袋': '20kg双组分内袋包装',
        '20kg固内': '20kg双组分固化剂内袋包装',
        '25kg内袋': '25kg双组分内袋包装',
    },
    'H-8100B/H-9100B': {
        "10kg": "10kg单组份包装",
        "20kg": "SP固化剂包装",
        '20kg内袋': 'SP固化剂内袋包装',
        "20kg固内": 'SP固化剂内袋包装',
    },
    'H-8100 SP': {
        '10kg': '10kg双组分包装',  # 这是有些丝印绿油选错类别
        '10kg内袋': '10kg双组分内袋包装',  # 这是有些丝印绿油选错类别
        '20kg': "20L桶包装",
        '20kg内袋': "20L桶内袋包装",
    },
    'H-9100 SP': {
        '10kg': '10kg双组分包装',  # 这是有些丝印绿油选错类别
        '10kg内袋': '10kg双组分内袋包装',  # 这是有些丝印绿油选错类别
        '20kg': "20L桶包装",
        '20kg内袋': "20L桶内袋包装",
    },
    'A-9060A': {
        '10kg': "10kg单组份包装",
        '20kg': "20L桶包装",
        '20kg内袋': "20L桶内袋包装",
    },
    'A-9000': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'A-2100': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'A-2000': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'K-2500': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'TS-3000': {
        '10kg': "热固化油墨包装",
    },
    'TM-3100': {
        '10kg': "热固化油墨包装",
    },
    'UVM-1800': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'UVS-1000': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'TH-XX': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'P-2700': {
        '10kg': "10kg单组份包装",
        '20kg': "20kg单组份包装",
        '20kg内袋': "20kg单组份内袋包装",
    },
    'SJ': {},
}

# package category info
PACKAGE_CATEGORIES = {
    "10kg单组份包装": {
        "box_type": "3#箱",         # 箱子类型
        "box_amount": 1,            # 箱子数量
        "part_a_jar_type": "1L罐",  # A 组分罐子类别
        "part_a_jar_amount": 10,    # A 组分罐子数量
        "part_b_jar_type": None,    # B 组分罐子类别
        "part_b_jar_amount": 0,     # B 组分罐子数量
        "weight": 10,               # 每箱重量
        "label_amount": 11,         # 每箱用标签总数
    },
    "20kg单组份包装": {
        "box_type": "1#箱",
        "box_amount": 1,
        "part_a_jar_type": "4L罐",
        "part_a_jar_amount": 4,
        "part_b_jar_type": None,
        "part_b_jar_amount": 0,
        "weight": 20,
        "label_amount": 5,
    },
    "5kg双组分包装": {
        "box_type": "3#箱",
        "box_amount": 1,
        "part_a_jar_type": "1L罐",
        "part_a_jar_amount": 5,
        "part_b_jar_type": "1L罐",
        "part_b_jar_amount": 5,
        "weight": 5,
        "label_amount": 11,
    },
    "10kg双组分包装": {
        "box_type": "4#箱",
        "box_amount": 1,
        "part_a_jar_type": "1L罐",
        "part_a_jar_amount": 10,
        "part_b_jar_type": "0.3L罐",
        "part_b_jar_amount": 10,
        "weight": 10,
        "label_amount": 21,
    },
    "20kg双组分包装": {
        "box_type": "5#箱",
        "box_amount": 1,
        "part_a_jar_type": "5L罐",
        "part_a_jar_amount": 4,
        "part_b_jar_type": "1L罐",
        "part_b_jar_amount": 4,
        "weight": 20,
        "label_amount": 9,
    },
    "热固化油墨包装": {
        "box_type": "7#箱",
        "box_amount": 1,
        "part_a_jar_type": "1L罐",
        "part_a_jar_amount": 10,
        "part_b_jar_type": "200ml罐",
        "part_b_jar_amount": 10,
        "weight": 10,
        "label_amount": 21,
    },
    "10kg双组分内袋包装": {
        "box_type": "18#箱",
        "box_amount": 1,
        "part_a_jar_type": "1L袋",
        "part_a_jar_amount": 10,
        "part_b_jar_type": "0.3L袋",
        "part_b_jar_amount": 10,
        "weight": 10,
        "label_amount": 21,
    },
    "20kg双组分内袋包装": {
        "box_type": "18#箱",
        "box_amount": 1,
        "part_a_jar_type": "4L袋",
        "part_a_jar_amount": 4,
        "part_b_jar_type": "1L袋",
        "part_b_jar_amount": 4,
        "weight": 20,
        "label_amount": 9,
    },
    "25kg双组分内袋包装": {
        "box_type": "17#箱",
        "box_amount": 1,
        "part_a_jar_type": "4L袋",
        "part_a_jar_amount": 5,
        "part_b_jar_type": "1L袋",
        "part_b_jar_amount": 5,
        "weight": 25,
        "label_amount": 11,
    },
    "20kg双组分固化剂内袋包装": {
        "box_type": "5#箱",
        "box_amount": 1,
        "part_a_jar_type": "5L罐",
        "part_a_jar_amount": 4,
        "part_b_jar_type": "1L袋",
        "part_b_jar_amount": 4,
        "weight": 20,
        "label_amount": 9,
    },
    "20kg单组份内袋包装": {
        "box_type": "18#箱",
        "box_amount": 1,
        "part_a_jar_type": "5L袋",
        "part_a_jar_amount": 4,
        "part_b_jar_type": None,
        "part_b_jar_amount": 0,
        "weight": 20,
        "label_amount": 5,
    },
    "20L桶包装": {
        "box_type": None,
        "box_amount": 0,
        "part_a_jar_type": "20L罐",
        "part_a_jar_amount": 1,
        "part_b_jar_type": None,
        "part_b_jar_amount": 0,
        "weight": 20,
        "label_amount": 1,
    },
    "20L桶内袋包装": {
        "box_type": None,
        "box_amount": 0,
        "part_a_jar_type": "20L罐",
        "part_a_jar_amount": 1,
        "part_b_jar_type": "PE袋",
        "part_b_jar_amount": 1,
        "weight": 20,
        "label_amount": 1,
    },
    "SP固化剂包装": {
        "box_type": "1#箱",
        "box_amount": 1,
        "part_a_jar_type": "5L罐",
        "part_a_jar_amount": 4,
        "part_b_jar_type": None,
        "part_b_jar_amount": 0,
        "weight": 0,  # 重量不确定
        "label_amount": 5,
    },
    "SP固化剂内袋包装": {
        "box_type": "18#箱",
        "box_amount": 1,
        "part_a_jar_type": "4L袋",
        "part_a_jar_amount": 4,
        "part_b_jar_type": None,
        "part_b_jar_amount": 0,
        "weight": 0,
        "label_amount": 5,
    },
}

COL_INDEXES = {
    '1#箱': 'K',
    '3#箱': 'L',
    '4#箱': 'M',
    '5#箱': 'N',
    '7#箱': 'O',
    '9#箱': 'P',
    '18#箱': 'Q',
    '0.3L罐': 'R',
    '1L罐': 'S',
    '4L罐': 'T',
    '5L罐': 'U',
    '20L罐': 'V',
    '200ml罐': 'W',
    'PE袋': 'X',
    '标签': 'Y',
    '0.3L袋': 'Z',
    '1L袋': 'AA',
    '4L袋': 'AB',
    '5L袋': 'AC',
    '17#箱': 'AD'
}
