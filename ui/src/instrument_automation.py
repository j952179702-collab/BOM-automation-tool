# 111.py - 仪表自动化处理器（修改版）

import os
import sys
import re
import logging
from pathlib import Path
from typing import Optional, Tuple
import numpy as np
import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side
from base_processor import BaseProcessor
import datetime
from openpyxl import load_workbook

# ====================== QTextEditLogger ======================
class QTextEditLogger(logging.Handler):
    def __init__(self):
        super().__init__()
        self.widget = None

    def set_widget(self, widget):
        self.widget = widget

    def emit(self, record):
        msg = self.format(record)
        if self.widget:
            self.widget.append(msg)

# ====================== 主功能类 ======================
class InstrumentAutomationProcessor(BaseProcessor):
    def __init__(self, input_file):
        super().__init__(input_file)
        self.input_file = input_file
        self.df = None
        self.df_sort = None
        self.processed_workbook = None
        self.logger = None
        self.log_handler = QTextEditLogger()
        # 注意：不要在这里调用 load_csv，因为它需要 logger 已经初始化

    @staticmethod
    def resource_path(relative_path):
        """兼容源码运行与 PyInstaller 单文件运行的资源路径。"""
        base_path = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parents[2]))
        return base_path / relative_path
    
    def save_processed_file(self, save_path: str) -> bool:
        """实现抽象方法：保存处理结果"""
        if self.df_sort is None:
            self.logger.error("❌ 没有可保存的数据（df_sort 为 None）")
            return False

        try:
            # 使用 pandas 保存为 Excel
            self.df_sort.to_excel(save_path, index=False, engine='openpyxl')
            self.logger.info(f"✅ 结果已保存到: {save_path}")
            return True
        except Exception as e:
            self.logger.error(f"❌ 保存文件失败: {str(e)}")
            return False

    def setup_logging(self, text_edit_widget):
        """由 UI 传入 QTextEdit，初始化日志"""
        self.logger = logging.getLogger(f'Processor_{id(self)}')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers.clear()

        # 设置 handler
        self.log_handler.set_widget(text_edit_widget)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S')
        self.log_handler.setFormatter(formatter)
        self.logger.addHandler(self.log_handler)

        # 控制台输出
        console = logging.StreamHandler()
        console.setFormatter(formatter)
        self.logger.addHandler(console)

        # ✅ 现在才能用 logger
        self.logger.info("日志系统已启动")

    def load_file(self) -> pd.DataFrame:
        """读取 CSV/XLSX；CSV 自动识别常见的中文文件编码。"""
        file_path = Path(self.input_file)
        suffix = file_path.suffix.lower()

        try:
            if suffix == '.xlsx':
                df = pd.read_excel(file_path, engine='openpyxl')
                self.logger.info(f"✅ 成功加载 Excel 文件: {file_path}")
                return df

            if suffix != '.csv':
                raise ValueError(f"不支持的文件格式: {suffix}，请选择 CSV 或 XLSX 文件")

            decode_errors = []
            for encoding in ('utf-8-sig', 'utf-8', 'gb18030', 'gbk'):
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    self.logger.info(
                        f"✅ 成功加载 CSV 文件: {file_path}，编码: {encoding}"
                    )
                    return df
                except UnicodeDecodeError as error:
                    decode_errors.append(f"{encoding}: {error}")

            raise UnicodeError(
                "无法识别 CSV 编码，已尝试 utf-8-sig、utf-8、gb18030、gbk。"
                + " | ".join(decode_errors)
            )
        except Exception as error:
            self.logger.error(f"❌ 文件加载失败: {error}")
            raise

    def load_csv(self) -> pd.DataFrame:
        """保留旧调用入口，统一交给自动编码读取。"""
        return self.load_file()
        
    def extract_chinese(self,text) -> str:
        # 提取所有中文字符
        if pd.isna(text) or not isinstance(text, str):
            text = str(text) if not pd.isna(text) else ""
        chinese = re.findall(r"[\u4e00-\u9fa5]+", text)
        return "".join(chinese) if chinese else "未知"  

    def generate_code(self):
        try:
            if self.df is None:
                self.df = self.load_csv()
            self.df_sort = self.df.sort_values(by=self.df.columns[1], ascending=False).copy()
            self.logger.info(f"已排序，共 {len(self.df_sort)} 行数据")

            # 初始化列（防止 KeyError）
            column_defaults = {
                "名称": "编码错误",
                "量程": "编码错误",
                "表壳材质": "编码错误",
                "仪表材质": "编码错误",
                "探杆长度": 0,
                "安装方式": "编码错误",
                "仪表类型": "编码错误",
                "线缆长度": "编码错误",
            }
            for col, default in column_defaults.items():
                if col not in self.df_sort.columns:
                    self.df_sort[col] = default

            # 量程映射
            range_mapping_rules = {
                "双金属温度计": {"-40~0": "L", "0~100": "M", "0~200": "H"},
                "压力变送器": {"-0.1~0.1": "A", "0~0.25": "B", "0~0.5": "C", "0~1.0": "D"},
                "温度变送器": {"-40~0": "L", "0~100": "M", "0~200": "H"},
                "压力表": {"-0.1~0.1": "L", "0~1.0": "M", "0~0.5": "N", "0~0.25": "O", "0~1.6": "G"}
            }

            flat_map = {}
            for inst_type, ranges in range_mapping_rules.items():
                for range_val, code in ranges.items():
                    flat_map[(inst_type, range_val)] = code

            range_mapped = self.df_sort["量程"].astype(str).map(flat_map)
            range_val_series = pd.Series(
                np.where(
                    range_mapped.isna(),
                    self.df_sort["量程"].astype(str),  # 回退到原始量程(Y)
                    range_mapped                     # 使用映射编码
                ),
                index=self.df_sort.index
            ).replace("nan", "ERR")  # 处理可能的 'nan' 字符串
            range_val = range_val_series
            
            # 材质映射
            material_mapping = {
                "304": "SS", "316L": "SS1", "TA2": "Ti", "Ta": "Ta",
                "HC": "HC", "2205": "SS2", "压铸铝": "C", "编码错误": "ERR"
            }
            self.shell_material = self.df_sort["表壳材质"].fillna("编码错误").map(material_mapping).fillna("ERR")
            self.material_value = self.df_sort["仪表材质"].fillna("编码错误").map(material_mapping).fillna("ERR")

            # 其他字段
            probe = (
                pd.to_numeric(self.df_sort["探杆长度"], errors="coerce")
                .fillna(0)
                .astype(int)
            )
            install_type = self.df_sort["安装方式"].fillna("编码错误").astype(str) if "安装方式" in self.df_sort.columns else " "
            long = self.df_sort["线缆长度"].fillna("编码错误").astype(str) if "线缆长度" in self.df_sort.columns else " "

            def get_param_text(medium):
                if medium in medium_map:
                    # 按中文分号拆分，去掉空，再用 \n 连接
                    parts = [p.strip() for p in medium_map[medium].split("；") if p.strip()]
                    return "\n".join(parts)  # ✅ 用换行符连接，存入一个格子
                else:
                    return "无"

            parameter_ins = {
                "压力表": "（*）. 精度：1.0%FS；\n（*）. 防护等级：IP65；\n（*）. 安装方式：径向直接式",
                "压力变送器": "（*）.精度：0.5%FS；\n（*）.输出信号：4-20mA；\n（*）.电压：两线制；\n（*）.通讯协议：无；\n（*）.现场显示：一体式多功能LCD显示表；\n（*）.防护等级：IP65；\n（*）.防爆等级：无；\n（*）.电气密封接口：M20*1.5；\n（*）.安装方式：螺纹；",
                "法兰液位仪表": "（*）. 精度：0.5%FS；\n（*）. 输出信号：4-20mA，二线制；\n（*）. 通讯协议：无；\n（*）. 防护等级：IP65；\n（*）. 电气密封接口：M20*1.5；\n（*）. 防爆等级：无",
                "投入液位仪表": "（*）. 精度：0.5%FS；\n（*）. 输出信号：4-20mA，二线制；\n（*）. 通讯协议：无；\n（*）. 现场显示：LCD显示，带调零功能；\n（*）. 防护等级：IP65；\n（*）. 电气密封接口：M20*1.5；\n（*）. 防爆等级：无",
                "液位计开关": "（*）.输出信号：开关量；\n（*）.防护等级：IP65；\n（*）.防爆等级：/；\n（*）.电气密封接口：不带接线盒；\n（*）.安装/接管方式：浸没",
                "电磁流量计":"（*）.精度：0.5%FS；\n（*）.输出信号：DC24V，4-20mA；\n（*）.电压：四线制；\n（*）.通讯协议：modbus485；\n（*）.现场显示：一体式多功能LCD显示表；\n（*）.防护等级：IP65；\n（*）.防爆等级：无；\n（*）.电气密封接口：M20*1.5；",
                "涡街流量计":"（*）.精度：1.0级；\n（*）.输出信号：DC24V，4-20mA；\n（*）.电压：四线制；\n（*）.通讯协议：modbus485；\n（*）.现场显示：一体式多功能LCD显示表；\n（*）.防护等级：IP65；\n（*）.防爆等级：无；\n（*）.电气密封接口：M20*1.5；",
                "热式流量计":"（*）.精度：1.0级；\n（*）.输出信号：DC24V，4-20mA；\n（*）.电压：四线制；\n（*）.通讯协议：modbus485；\n（*）.现场显示：一体式多功能LCD显示表；\n（*）.防护等级：IP65；\n（*）.防爆等级：无；\n（*）.电气密封接口：M20*1.5；",
                "温度变送器":"（*）.精度：0.5%FS；\n（*）.输出信号：4-20mA；\n（*）.电压：DC24V,两线制；\n（*）.通讯协议：/；\n（*）.现场显示：无；\n（*）.防护等级：IP65；\n（*）.防爆等级：/；\n（*）.电气密封接口：M20*1.5；\n（*）.保护管类型：螺纹式直型保护管；\n（*）.保护管直径：φ12；\n（*）.备注：配聚四氟乙烯垫片1个；\n（*）.安装/接管方式：固定G1/2外螺纹,PN10；",
                "双金属温度计":"（*）.精度：1.0%FS；\n（*）.现场显示：表盘直径φ100，表头填充硅油；\n（*）.防护等级：IP65；\n（*）.形式：耐震，万向型\n（*）.安装方式：轴向安装\n（*）.鞘直径：6mm；\n（*）.鞘安装接头：G1/2外螺纹，PN10；\n（*）.保护管类型：螺纹式直型保护管；\n（*）.保护管直径：φ12；\n（*）.保护管与鞘连接：G1/2内螺纹，PN10；\n（*）.安装/接管方式：固定G1/2外螺纹；",
                "分析仪表":"（*）.输出信号：4-20mA；\n（*）.电压：DC24V，两线制；\n（*）.通讯协议：RS485；\n（*）.线长：20m；\n（*）.现场显示：一体式多功能LCD显示表；\n（*）.防护等级：IP65；\n（*）.防爆等级：无；\n（*）.电气密封接口：2-M20*1.5；",
                }

            self.df_sort['参数2'] = self.df_sort['名称'].apply(lambda x: self.extract_chinese(x)).map(parameter_ins)
            unmatched_mask = self.df_sort['参数2'].isna()
            if unmatched_mask.any():
                unmatched = self.df_sort.loc[unmatched_mask, '名称'].unique()
                self.logger.error(f"⚠️ 以下仪表名称未在标准参数库中找到：{unmatched.tolist()}")
                self.df_sort['参数2'] = self.df_sort['参数2'].fillna("")
            else:
                self.logger.info("✅ 所有仪表名称均已匹配标准参数")

            # 生成 SKU
            sku_list = []
            medium_df = pd.read_excel(
                self.resource_path("dataset/medium.xlsx"),
                engine='openpyxl',
            )
            medium_df.dropna(subset=['介质', '参数'], inplace=True)
            medium_map = dict(zip(medium_df["介质"], medium_df["参数"]))
            self.df_sort["参数1"] = self.df_sort["介质"].apply(get_param_text)
            self.df_sort['参数'] = self.df_sort['参数1'].astype(str) + "\n" + self.df_sort['参数2'].astype(str)
            for idx in self.df_sort.index:
                self.ins_name =  self.extract_chinese(str(self.df_sort.loc[idx, "名称"])).strip()
                range_val_str = str(self.df_sort.loc[idx, "量程"])
                key = (self.ins_name,range_val_str)
                if key in flat_map:
                    self.r_val = flat_map[key]
                    self.logger.info(f"✅ 匹配成功: {key} → {self.r_val}")
                else:
                    self.r_val = range_val_str  # 回退到原始量程
                    self.logger.error(f"❌ 未匹配: {key} → 使用原始值")
                mat_val = self.material_value.loc[idx]
                self.shell_mat = self.shell_material.loc[idx]
                self.prob = probe.loc[idx]
                self.install_val = install_type.loc[idx]
                self.long_val = long.loc[idx]

                shell_value = self.df_sort.loc[idx,'表壳材质']
                material_value = self.df_sort.loc[idx,'仪表材质']
                
                try:
                    appen_text = ""  # 初始化 appen_text 为空字符串
                    if self.ins_name == "双金属温度计":
                        sku = f"WSS-AOA{self.r_val}{self.shell_mat}-12{mat_val}{self.prob}-LWG1/2"
                        appen_text = f"\n（*）.保护管长度：{self.prob}mm；\n（*）.量程：{range_val_str}℃；"
                        self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：{shell_value}；\n（*）.保护管材质：{material_value}（含接头）"
                    elif self.ins_name == "温度变送器":
                        sku = f"SWBZ-A0-24D2NNA0P{self.r_val}0-12{mat_val}{self.prob}-LWG1/2"
                        appen_text = f"\n（*）.保护管长度：{self.prob}mm；\n（*）.量程：{range_val_str}℃；"
                        self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.保护管材质：{material_value}（含接头）"
                    elif self.ins_name == "压力表":
                        inst_type = self.df_sort.loc[idx, "仪表类型"]
                        if str(inst_type) == "耐震":
                            sku = f"YTHN-A1A{self.r_val}{self.shell_mat}-{mat_val}-LWG1/2W"
                            appen_text = f"\n（*）.形式：普通型耐震；\n（*）.现场显示：表盘直径φ100，表头填充硅油；\n（*）.量程：{range_val_str}MPa（G）；\n（*）.安装/接管方式：固定G1/2外螺纹,PN10；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：{shell_value}；\n（*）.接液材质：{material_value}"
                        elif str(inst_type) == "耐震隔膜":
                            if mat_val == "PP":
                                sku = f"YMN-A1A{self.r_val}PP-PP-LWG1/2N"
                                appen_text = f"\n（*）.现场显示：表盘直径φ60，表头填充 silicone；\n（*）.量程：{range_val_str}MPa（G）；\n（*）.形式：耐震隔膜型；\n（*）.安装/接管方式：固定G1/2内螺纹,PN10；"
                                self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：{shell_value}；\n（*）.膜片材质：{material_value}"
                            else:
                                    sku = f"YMN-A1A{self.r_val}{self.shell_mat}-{mat_val}-FL120RF1.0"
                                    appen_text = f"\n（*）.现场显示：表盘直径φ60，表头填充 silicone；\n（*）.量程：{range_val_str}MPa（G）；\n（*）.形式：耐震隔膜型；\n（*）.安装/接管方式：法兰连接，DN20,PN10"
                                    self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：{shell_value}；\n（*）.膜片材质：{material_value}"
                    elif self.ins_name == "压力变送器":
                        if self.r_val in ["0~0.25", "B"]:
                            sku = f"PT-H-B-1-{mat_val}-50-NSR"
                            appen_text = f"\n（*）.量程：{range_val_str}MPa（G）；\n（*）.安装/接管方式：法兰DN50，PN10 HG/T-20592-2009，RF，B型；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}"
                        else:
                            sku = f"PT-H-{self.r_val}-1-{mat_val}-G1/2-NSR"
                            appen_text = f"\n（*）.量程：{range_val_str}MPa（G）；\n（*）.安装/接管方式：固定G1/2外螺纹,PN10；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}"
                    elif self.ins_name == "电磁流量计":
                        sku = f"EMF-1C-0-A{mat_val}CS-FL{self.install_val.split('DN')[-1]}RF1.0"
                        appen_text = f"\n（*）.介质流量：{range_val_str}m³/h；\n（*）.安装/接管方式：法兰DN{self.install_val.split('DN')[-1]}，PN10 HG/T-20592-2009，RF，B型；"
                        self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.电极材质：{material_value}；\n（*）.接液材质：PTFE；\n（*）.法兰材质：碳钢"
                    elif self.ins_name == "热式流量计":
                        dn_match = re.search(r"DN\s*(\d+)", self.install_val, re.IGNORECASE)
                        dn_number = int(dn_match.group(1)) if dn_match else None
                        if dn_number is not None and dn_number <= 80:
                            sku = f"TFC-A{self.r_val.replace('-','~').split('~')[-1]}-FL{self.install_val.split('DN')[-1]}"
                            appen_text = f"\n（*）.介质流量：{range_val_str}m³/h；\n（*）.安装/接管方式：法兰DN{self.install_val.split('DN')[-1]}，PN10 HG/T-20592-2009，RF，B型；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.法兰材质：304"
                        else:
                            sku = f"TFC-A{self.r_val.replace('-','~').split('~')[-1]}-LWG1/2-{self.install_val.split('DN')[-1]}"
                            appen_text = f"\n（*）.介质流量：{range_val_str}m³/h；\n（*）.安装/接管方式：固定G1/2外螺纹；管道规格：{self.install_val}；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.螺纹材质：304"
                    elif self.ins_name == "浮子流量计":
                        if any("DN" in item for item in self.install_val):
                            sku = f"FF-{mat_val}-{self.r_val.replace('-','~').split('~')[-1]}-FL{self.install_val.split('DN')[-1]}"
                            appen_text = f"\n（*）.介质流量：{range_val_str}m³/h；\n（*）.法兰DN{self.install_val.split('DN')[-1]}，PN10 HG/T-20592-2009，RF，B型；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.材质：UPVC"
                        else:
                            sku = f"FF-{mat_val}-{self.r_val.replace('-','~').split('~')[-1]}-CC{self.install_val.split('DN')[-1]}"
                            appen_text = f"\n（*）.介质流量：{range_val_str}m³/h；\n（*）.安装/接管方式：承插{self.install_val}；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.材质：UPVC"
                    elif self.ins_name =="涡街流量计":
                        sku =  f"VF-1C-{mat_val}{mat_val}-FL{self.install_val.split('DN')[-1]}RF1.6"
                        appen_text = f"\n（*）.介质流量：{range_val_str}m³/h；\n（*）.安装/接管方式：法兰DN{self.install_val.split('DN')[-1]}，PN10 HG/T-20592-2009，RF，B型；"
                        self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.法兰材质：{material_value}"
                    elif self.ins_name == "法兰液位仪表":
                        if self.df_sort.loc[idx,"仪表类型"] == "磁翻板":
                            if "KP" in install_type:
                                sku = f"LG-FQC-{self.r_val.replace('-','~').split('~')[-1]}-{mat_val}{mat_val}-KP50"
                                appen_text = f"\n（*）.量程：{range_val_str}mm；\n（*）.变送器位置:液位计上方；\n（*）.安装/接管方式：卡盘直径50.5；"
                                self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.法兰材质：{material_value}"
                            else:
                                sku = f"LG-FQC-{self.r_val.replace('-','~').split('~')[-1]}-{mat_val}{mat_val}-FL25RF1.6"
                                appen_text = f"\n（*）.量程：{range_val_str}mm；\n（*）.变送器位置:液位计上方；\n（*）.安装/接管方式：法兰DN25，PN16 HG/T-20592-2009，RF，B型；"
                                self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.法兰材质：{material_value}"
                        elif self.df_sort.loc[idx,"仪表类型"] == "单法兰":
                            sku = f"LG-DFC-{self.r_val.replace('-','~').split('~')[-1]}-SS{mat_val}-FL50RF1.6"
                            appen_text = f"\n（*）.量程：{range_val_str}mm；\n（*）.安装/接管方式：法兰DN50，PN10 HG/T-20592-2009，RF，B型；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.法兰材质：304"
                        elif self.df_sort.loc[idx,"仪表类型"] == "双法兰":
                            sku = f"LG-SFC-{self.r_val.replace('-','~').split('~')[-1]}-SS{mat_val}-FL50RF1.6"
                            appen_text = f"\n（*）.现场显示：LCD显示，带调零功能；\n（*）.量程：{range_val_str}mm；\n（*）.安装/接管方式：法兰DN50，PN10 HG/T-20592-2009，RF，B型；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}；\n（*）.法兰材质：304"
                    elif self.ins_name == "投入液位仪表":
                        sku = f"LG-TRC-{self.r_val.replace('-','~').split('~')[-1]}-{mat_val}-{mat_val}-TR"
                        appen_text = f"\n（*）.量程：{range_val_str}mm；\n（*）.安装/接管方式：投入式；"
                        self.df_sort.loc[idx,'材质'] = f"（*）.表壳材质：压铸铝；\n（*）.接液材质：{material_value}"
                    elif self.ins_name == "液位计开关":
                        sku = f"CFBS-{self.long_val}-PP-0-0"
                        appen_text = f"\n（*）.线缆长度{self.long_val}m；"
                        self.df_sort.loc[idx,'材质'] = f"（*）.接液材质：{material_value}"
                    elif self.ins_name == "分析仪表":
                        if self.df_sort.loc[idx,"仪表类型"] == "Ω":
                            range_val_str = str(range_val.loc[idx],"量程")
                            if "2000" in range_val_str:
                                sku = f"EC-W1-A1/2000-SS1-LWG1/2"
                                appen_text = f"\n（*）.精度：0.5%FS；\n（*）.量程：1-2000us/cm，k=1；\n（*）.线长：20m；\n（*）.安装/接管方式：固定G1/2外螺纹；"
                                self.df_sort.loc[idx,'材质'] = f"（*）.接液材质：{material_value}"
                            elif "20000" in range_val_str:
                                sku = f"EC-W1-A10/20000-SS1-LWG1/2"
                                appen_text = f"\n（*）.精度：0.5%FS；\n（*）.量程：10-20000us/cm，k=10；\n（*）.线长：20m；\n（*）.安装/接管方式：固定G1/2外螺纹；"
                                self.df_sort.loc[idx,'材质'] = f"（*）.接液材质：{material_value}"
                            else:
                                sku = "编码失败"
                        elif self.df_sort.loc[idx,"仪表类型"] == "PH":
                            sku = f"PH-0-DC24VA/14-GL-LWG3/4"
                            appen_text = f"\n（*）.量程：0-14；\n（*）.精度：±0.01；\n（*）.安装/接管方式：带支架，G3/4螺纹；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.接液材质：玻璃"
                        elif self.df_sort.loc[idx,"仪表类型"] == "DO":
                            sku = f"DO-0-A/020-SS1-JM"
                            appen_text = f"\n（*）.精度：0.5%FS；\n（*）.电缆长度：10m；\n（*）.安装/接管方式：投入式；"
                            self.df_sort.loc[idx,'材质'] = f"（*）.保护管材质：PP；\n（*）.接液材质：316L"
                    else:
                        sku = f"无编码"
                        self.df_sort.loc[idx,"参数"] = f"\n无编码；"
                        self.df_sort.loc[idx,'材质'] = f"无编码"
                    sku_list.append(sku)
                    self.df_sort.loc[idx, '参数'] += appen_text
                except Exception as e:
                    sku = f"ERROR_{idx}"
                    sku_list.append(sku)
                    self.logger.error(f"行 {idx} 生成失败: {e}")
            self.df_sort['SKU编码'] = sku_list
            self.logger.info("✅ SKU 生成完成！")
            return self.df_sort
        except Exception as e:
            self.logger.error(f"❌ 生成 SKU 时出错: {str(e)}")
            raise
    
    def set_metadata(self, project_number):
        """设置项目号元数据。"""
        if self.df_sort is not None:
            self.df_sort['项目号'] = project_number
            self.df_sort['所属项目'] = project_number
        else:
            raise ValueError("df_sort 未初始化，请先加载数据")
    
    def get_note(self, df):
        """
        为 DataFrame 生成备注列：项目号 + 仪表名称
        """      
        required_columns = ['项目号', '仪表名称']  
        # ✅ 详细的列检查
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            error_msg = f"❌ 生成备注失败：缺少必要列 {missing_cols}"
            self.logger.error(error_msg)
            self.logger.info(f"📋 当前DataFrame的列名: {list(df.columns)}")
            raise ValueError(error_msg)

        try:
            # ✅ 安全转换为字符串，处理 NaN
            project_num = df['项目号'].fillna('').astype(str)
            ins_name = df['仪表名称'].fillna('').astype(str)           
            # 确保字符串连接时有空格分隔
            df['备注'] = project_num.str.strip()+ ins_name.str.strip()
            # ✅ 检查备注列是否成功创建
            if '备注' not in df.columns:
                raise ValueError("备注列未成功创建")
            
            self.logger.info(f"✅ 备注列生成成功，前3个值: {df['备注'].head(3).tolist()}")
            self.logger.info(f"✅ 备注列数据类型: {df['备注'].dtype}")
            return df
        except Exception as e:
            error_msg = f"❌ 生成备注列时发生未知错误: {e}"
            self.logger.error(error_msg)
            raise
    def merge_by_SKU(self):
        self.df_sort['备注'] = self.df_sort['备注'].fillna('')
        aggregation = {
            '仪表类型': 'first',
            '参数': 'first',
            '材质': 'first',
            '备注': lambda x:'\n'.join(item for item in x if item.strip() != ''),
        }
        self.df_group = self.df_sort.groupby('SKU编码').agg(aggregation).reset_index()
        self.df_group['*申请数量'] = self.df_sort.groupby('SKU编码').size().values
        self.df_group = self.df_group.rename(columns={
            'SKU编码': '*SKU编号',
            '仪表类型': '产品名称',
        })
        output_columns = [
            '*SKU编号',
            '产品名称',
            '参数',
            '材质',
            '*申请数量',
            '备注',
        ]
        self.df_output = self.df_group[output_columns].reset_index(drop=True)

        book = Workbook()
        sheet = book.active
        sheet.title = "仪表"

        for column_index, column_name in enumerate(self.df_output.columns, start=1):
            cell = sheet.cell(row=1, column=column_index, value=column_name)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for row_index, row in enumerate(
            self.df_output.itertuples(index=False, name=None),
            start=2,
        ):
            for column_index, value in enumerate(row, start=1):
                cell = sheet.cell(
                    row=row_index,
                    column=column_index,
                    value='' if pd.isna(value) else value,
                )
                cell.alignment = Alignment(vertical='top', wrap_text=True)

        self.book = book
        self.logger.info(f"✅ 按SKU合并完成，数据形状: {self.df_group.shape}")
        return book



    def process(self, project_number):
        try:
            # 1. 生成SKU编码
            self.df_sort = self.generate_code()
            if self.df_sort is None:
                raise ValueError("SKU编码生成失败，df_sort 为空")
            self.logger.info(f"✅ 生成SKU编码完成，数据形状: {self.df_sort.shape}")
            
            # 2. 设置元数据
            # 确保项目号列存在并设置值
            if '项目号' not in self.df_sort.columns:
                self.df_sort['项目号'] = ''
            self.df_sort['项目号'] = project_number
            self.logger.info(f"✅ 设置元数据完成，项目号: {project_number}")
            
            # 3. 生成备注列
            if '仪表名称' not in self.df_sort.columns:
                # 假设仪表名称就是从“名称”列里提取的中文部分
                self.df_sort['仪表名称'] = self.df_sort['名称'].apply(lambda x: self.extract_chinese(str(x)))
            self.logger.info(f"开始生成备注列，当前列: {self.df_sort.columns.tolist()}")
            self.df_sort = self.get_note(df=self.df_sort)
            if '备注' not in self.df_sort.columns:
                raise ValueError("备注列生成失败")
            self.df_output = self.merge_by_SKU()
            
            self.logger.info("✅ 文件处理完成")
            return True
        except Exception as e:
            self.logger.error(f"❌ 处理失败: {str(e)}")
            return False
