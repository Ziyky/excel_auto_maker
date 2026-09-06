import sys
from pathlib import Path
from typing import Dict, Optional
from openpyxl.worksheet.worksheet import Worksheet

from auto_gen_package.common import parse_master_header, check_required_cols


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent

# 全局预先定义好各个输出路径
APP_ROOT = get_app_dir()
DIR_INPUT = APP_ROOT / "data_input"
DIR_TEMPLATE = DIR_INPUT / "template"
DIR_OUTPUT = APP_ROOT / "data_output"
DIR_LOG = APP_ROOT / "logs"    # 如果以后要加日志目录，在这里加

def init_work_dir():
    """程序初始化：一次性创建全部需要的文件夹"""
    DIR_INPUT.mkdir(parents=True,exist_ok=True)
    DIR_TEMPLATE.mkdir(parents=True,exist_ok=True)
    DIR_OUTPUT.mkdir(parents=True, exist_ok=True)
    DIR_LOG.mkdir(parents=True, exist_ok=True)

MASTER_HEADER: Dict[str, Optional[int]] | None = None

def init_filed_index(master:Worksheet):
    global MASTER_HEADER
    #获取表头数据
    header_row = list(next(master.iter_rows(min_row=1, max_row=1)))
    col_cfg = parse_master_header(header_row)
    #执行检查
    check_required_cols(col_cfg)
    MASTER_HEADER = col_cfg

