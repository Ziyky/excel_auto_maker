import time
from openpyxl.reader.excel import load_workbook

from auto_gen_package.app_init import init_work_dir, DIR_INPUT, DIR_TEMPLATE, DIR_OUTPUT, init_filed_index
from auto_gen_package.template import make_template, make_dest_monitor_info, make_script_ips, make_html_info, \
    make_not_online, lst_online

import warnings
warnings.filterwarnings("ignore", message=r"Print area cannot be set to Defined name.*")
def main():
    # 打印欢迎界面
    print("==============================")
    print("欢迎使用本程序")
    print("版本：1.1")
    print("系统会自动识别./date_input/template下面带\"任务\"关键字的表格")
    print("template文件夹下命名规范-->\"10 任务 08-21 迁移任务\"<--")
    print("绝大部分数据依赖于表格 \"master.xlsx\"")
    print("==============================")
    input("按下回车键继续...")

    #--------------初始化--------------
    print("正在进行初始化...")
    #文件夹初始化
    init_work_dir()
    master_file = DIR_INPUT / "master.xlsx"
    if not master_file.is_file():
        print(f"错误：找不到主表格文件 {master_file}")
        input("按回车键退出...")
        exit(1)
    wb_master = load_workbook(master_file)
    ws_master = wb_master.active
    init_filed_index(ws_master)
    #--------------初始化--------------

    sync_date = input(" 请输入数据同步时间[xxxx-xx-xx xx:xx:xx]: ")
    #--------------模板制作--------------

    print("正在进行模板制作")
    #获取template文件夹所有任务文件并制作
    for f in DIR_TEMPLATE.glob("*迁移任务*.xlsx"):
        if f.is_file():
            wb_tp = load_workbook(f)
            scmt = f.name[:2]
            res = make_template(wb_tp,ws_master,scmt,sync_date)
            res.save(DIR_OUTPUT / f"output-{f.name}")

    time.sleep(0.5)

    #--------------三个文件制作--------------
    print("正在制作三个文件")
    wb_dest_monitor_info = make_dest_monitor_info(ws_master)
    wb_dest_monitor_info.save(DIR_OUTPUT / "output-目标端监控ips.xlsx")
    make_html_info(ws_master)
    make_script_ips(ws_master)

    #统计不在线情况
    make_not_online(lst_online,ws_master)

    print("制作成功,已全部存放在data_output目录!")


if __name__ == '__main__':
    main()

