# -*- coding: utf-8 -*-

"""全局常量配置"""

import re

# 默认属性列表
DEFAULT_PROPERTIES = [
    "Assembly/Cast unit Mark",
    "Assembly/Cast unit position code",
    "Assembly/Cast unit top elevation"
]

# 需要跳过的实体类型
SKIP_ENTITY_TYPES = {
    "IfcProject", "IfcSite", "IfcBuilding", "IfcBuildingStorey",
    "IfcSpace", "IfcAnnotation", "IfcGrid", "IfcStructuralItem"
}

# 日期格式
DATE_FORMAT = "%m-%d"

# 处理块大小
CHUNK_SIZE = 50

# 文件名非法字符正则
INVALID_FILENAME_CHARS = re.compile(r'[\\/*?:"<>|]')