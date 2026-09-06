from typing import List, Any, Sequence, Dict, Optional

from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def lookup(ws:Worksheet, lookup_value, lookup_col: int, return_col: int, min_row=1):
    for row in ws.iter_rows(min_row=min_row):
        c = row[lookup_col]
        if c.value == lookup_value:
            return row[return_col].value
    return None

def lookup_fuzzy(ws, lookup_value, lookup_col: int, return_col: int, min_row=1):
    if lookup_value is None:
        return None
    target = str(lookup_value).strip()
    for row in ws.iter_rows(min_row=min_row):
        cell_val = row[lookup_col -1 ].value
        if cell_val is None:
            continue
        s = str(cell_val)
        # 包含匹配：源单元格包含搜索关键词
        if target in s:
            return row[return_col -1 ].value
    return None

def find_cell_index(row: List[Cell], target_value: Any) -> int | None:
    """
    在一行单元格中查找等于 target_value 的单元格，返回索引(0开始)
    :param row: openpyxl iter_rows 得到的一行单元格列表
    :param target_value: 需要匹配的值（数字/字符串/None）
    :return: 索引int，找不到返回None
    """
    for idx, cell in enumerate(row):
        # cell.value 取出单元格实际内容
        val = cell.value
        # 全等匹配；如果需要忽略空格可以改成 str(val).strip() == str(target_value).strip()
        if val == target_value:
            return idx
    return None

#解析表头
def parse_master_header(header_row: List[Cell]) -> Dict[str, Optional[int]]:
    """
    解析master表头行，返回各字段列索引字典
    key: 内存字段名, value:列索引(0开始) / None
    """
    # 内存key : excel表头文本
    mapping = [
        ("desk_id", "桌面ID"),
        ("src_ip", "IP"),
        ("ser_id", "使用人工号"),
        ("user_name", "使用人姓名"),
        ("migrate_date", "迁移日"),
        ("dest_name", "新桌面名称"),
        ("dest_ip", "新桌面IP"),
        ("dest_mac", "新桌面MAC"),
        ("dest_hci", "目标HCI"),
        ("dest_vdc_ip", "目标VDC"),
        ("dest_storage", "存储位置"),
        ("dest_vlan", "目标网络交换机"),
        ("dest_vdc_role", "vdc角色"),
        ("dest_vdc_resource", "vdc资源名称"),
        ("scmt_ip","迁移服务器")
    ]

    col_cfg: Dict[str, Optional[int]] = {}
    for key, excel_title in mapping:
        col_cfg[excel_title] = find_cell_index(header_row, excel_title)
    return col_cfg

def check_required_cols(col_cfg: Dict[str, Optional[int]]):
    """校验必填列，任意一个None打印缺失字段并sys.exit(1)退出"""
    required_keys = [
        "桌面ID",
        "IP",
        "使用人工号",
        "使用人姓名",
        "迁移日",
        "新桌面名称",
        "新桌面IP",
        "新桌面MAC",
        "目标HCI",
        "目标VDC",
        "存储位置",
        "目标网络交换机",
        "vdc角色",
        "vdc资源名称",
        "迁移服务器",
    ]
    missing = []
    for k in required_keys:
        if col_cfg[k] is None:
            missing.append(k)
    if missing:
        print(f"[ERROR] 缺失必填表头字段,请检查：{missing}")
        exit(1)
    return col_cfg
