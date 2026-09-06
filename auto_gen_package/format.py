#excel的xlookup功能
from openpyxl.worksheet.worksheet import Worksheet

from auto_gen_package import app_init
from auto_gen_package.common import lookup, lookup_fuzzy

#格式化磁盘列文本
def disk_text_format(disk_info_text:str):
    disk_info_text = disk_info_text.replace("\n磁盘1(16MB)","")
    disk_info_text = disk_info_text.replace("磁盘1(16MB)\n","")
    return disk_info_text

#一列分组
###问题函数 已解决
def group_name_transform(date):
    if date is not None:
        date = str(date)
        parts = date.split(".")
        if len(parts) < 2:
            return None
        # 月、日补零到2位
        mm = parts[0].zfill(2)
        dd = parts[1].zfill(2)
        return f"{mm}-{dd}"
    return None

def get_net_out_config(master:Worksheet,tp_hci_info:Worksheet,str_net_out_config:str,desk_id):
    config = ""
    #获取网络出口配置的网卡
    str_net_out_config_parts = str_net_out_config.split(",")
    eth=str_net_out_config_parts[0]

    #拼接网卡
    if eth is not None:config += eth+","
    #拼接DHCP
    config +="DHCP,"
    #获取原始VLAN
    vlan_text = ""
    vlan_id = lookup(master,desk_id,app_init.MASTER_HEADER["桌面ID"],app_init.MASTER_HEADER["目标网络交换机"])
    if vlan_id is not None:

        #模糊查询,获取到hci_info里的text
        hci_vlan = lookup_fuzzy(tp_hci_info, lookup(master,desk_id,app_init.MASTER_HEADER["桌面ID"],app_init.MASTER_HEADER["目标HCI"]),2,8)
        hci_vlan_parts = str(hci_vlan).split("\n")
        for part in hci_vlan_parts:
            if vlan_id[-4:] == part[-4:]:
                vlan_text = part

    #拼接vlan
    if vlan_text is not None:config += vlan_text+","

    #获取mac
    mac = lookup(master,desk_id,app_init.MASTER_HEADER["桌面ID"],app_init.MASTER_HEADER["新桌面MAC"])
    if mac is not None :config += mac

    return config
