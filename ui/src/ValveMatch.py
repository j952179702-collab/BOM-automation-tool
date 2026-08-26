import pandas as pd
import os
import logging

class ValveProtocolMatcher:
    def __init__(self, xlsx_path, logger=None):
        self.xlsx_path = xlsx_path
        self.master_selection = None # 存放合并后的选型手册
        self.master_params = None    # 存放合并后的参数表
        self.logger = logger or logging.getLogger(__name__)
        
        # === 1. 配置：材质映射 (物理材质 -> 材质分组) ===
        self.MATERIAL_MAP = {
            # 通用组
            "304": "General", "316L": "General", "C-F": "General", 
            "WCB": "General", "碳钢": "General", "1.4529": "General","石墨铸铁":"General",
            # 钛材组
            "TA2": "Titanium", "TC4": "Titanium", "钛材": "Titanium", "纯钛": "Titanium"
        }

    def load_and_preprocess(self):
        """一次性加载所有8个Sheet，构建两张超级总表"""
        try:
            # 读取所有 Sheet
            all_sheets = pd.read_excel(self.xlsx_path, sheet_name=None)
            self.logger.info(f"✅ 已加载“阀门选型手册”，包含 Sheet: {list(all_sheets.keys())}")
        except Exception as e:
            self.logger.error(f"❌ 读取失败: {e}")
            return

        # ==========================================
        # A. 构建【选型手册总表】 (处理 2 个 Sheet)
        # ==========================================
        selection_configs = [
            {"sheet": "选型手册", "group": "General"},
            {"sheet": "选型手册 (钛)", "group": "Titanium"} # 注意文件名匹配
        ]
        
        sel_dfs = []
        for config in selection_configs:
            sheet_name = config["sheet"]
            group_name = config["group"]
            if sheet_name in all_sheets:
                df = all_sheets[sheet_name].copy()
            # 关键步骤：逆透视 (Melt)
            df_hand = df[['手阀协议号', '手阀编码']].copy().rename(
                columns={'手阀协议号': '协议号', '手阀编码': 'Match_Code'}
            )
            df_hand['驱动类型'] = '手动'
            # --- 处理电动 ---
            df_elec = df[['电动阀协议号', '电动阀编码']].copy().rename(
                columns={'电动阀协议号': '协议号', '电动阀编码': 'Match_Code'}
            )
            df_elec['驱动类型'] = '电动'
            
            # --- 处理气动 ---
            df_pneu = df[['气动阀协议号', '气动阀编码']].copy().rename(
                columns={'气动阀协议号': '协议号', '气动阀编码': 'Match_Code'}
            )
            df_pneu['驱动类型'] = '气动'
            df_sheet_combine = pd.concat([df_hand, df_elec, df_pneu], ignore_index=True)
            df_sheet_combine['材质组'] = group_name
            sel_dfs.append(df_sheet_combine)
        if sel_dfs:
            self.master_selection = pd.concat(sel_dfs, ignore_index=True).dropna(subset=['Match_Code'])
            self.master_selection['Match_Code'] = self.master_selection['Match_Code'].astype(str).str.strip()
            self.logger.info("✅ 选型手册总表预处理完成")
        
        param_sheet_map = {
            "手阀参数":       ("手动", "General"),
            "电动阀参数":     ("电动", "General"),
            "气动阀参数":     ("气动", "General"),
            "手阀参数 (钛)":   ("手动", "Titanium"),
            "电动阀参数 (钛)": ("电动", "Titanium"),
            "气动阀参数 (钛)": ("气动", "Titanium")
        }
        param_sheet_df= []
        for sheet_name, (drive_type, material_group) in param_sheet_map.items():
            if sheet_name in all_sheets:
                df = all_sheets[sheet_name].copy()
                df['驱动类型'] = drive_type
                df['材质组'] = material_group
                param_sheet_df.append(df)
            if param_sheet_df:
                self.master_params = pd.concat(param_sheet_df,ignore_index =True)        
        return True

    def run_match(self,df_target):
        result = df_target.copy()
        result['材质组']  = result['阀门材质'].map(self.MATERIAL_MAP).fillna('General')
        result['*SKU编号'] = (
            result['*SKU编号']
            .fillna('')          # 把 NaN 变成空字符串
            .astype(str)         # 确保全是字符串
            .str.strip()         # 去除首尾空白
                )
        self.master_selection['Match_Code'] = (
                self.master_selection['Match_Code']
                .fillna('')
                .astype(str)
                .str.strip()
            )

        # --- 诊断开始 ---
        test_val = "Q00F-50-2E"

        # 检查左表（你生成的 71 行数据）
        left_row = result[result['*SKU编号'] == test_val]
        print(f"🔍 诊断：左表是否存在 '{test_val}'? {'是' if not left_row.empty else '否'}")

        # 检查右表（协议库总表）
        if self.master_selection is not None:
            right_row = self.master_selection[self.master_selection['Match_Code'] == test_val]
            print(f"🔍 诊断：协议库是否存在 '{test_val}'? {'是' if not right_row.empty else '否'}")
            
            if not right_row.empty:
                print(f"🔍 诊断：协议库中对应的协议号是: {right_row['协议号'].values}")
        # --- 诊断结束 ---

        def get_drive_type(row):
            vt = str(row.get('名称',''))
            if "电动" in vt:
                return "电动"
            if "气动" in vt:
                return "气动"
            return "手动"
        result['驱动类型'] = result.apply(get_drive_type, axis=1)
        result['*SKU编号'] = result['*SKU编号'].astype(str).str.strip()
        if self.master_selection is not None:
            result = pd.merge(
                result,
                self.master_selection[['Match_Code', '协议号']],
                left_on=['*SKU编号'],
                right_on=['Match_Code'],
                how='left'
            )
        
        if self.master_params is not None:
            result = pd.merge(
                result,
                self.master_params[['产品编码', '参数']],
                left_on = ['*SKU编号'],
                right_on = ['产品编码'],
                how = 'left'
            )
        self.logger.info(f"✅ 匹配完成，结果包含 {len(result)} 行")
        return result

class ParameterFiller:
    def __init__(self, param_path, logger=None):
        self.param_path = param_path
        self.logger = logger or logging.getLogger(__name__)
        
        # 1. 加载介质库
        try:
            medium_df = pd.read_excel(param_path, engine='openpyxl')
            # 建立映射字典：介质名称 -> 介质参数
            self.medium_dict = dict(zip(medium_df["介质"], medium_df["参数"]))
            self.logger.info("✅ 介质库加载成功")
        except Exception as e:
            self.logger.error(f"介质库读取失败: {e}")
            self.medium_dict = {}

        # 2. 结构规则库 (注意：这里只保留【物理结构】，去掉了传动和开度，因为那些是动态的)
        self.STRUCTURE_RULES = {
            "球阀": "（*）.阀门结构形式：浮动球式，直通；",
            "蝶阀": "（*）.阀门结构形式：单偏心式；\n（*）.阀体颜色：RAL9006白铝色或RAL9006灰铝色（金属色）；\n（*）.泄露等级：V级；",
            "止回阀": "（*）.阀门结构形式：旋启式单瓣；",
            "截止阀": "（*）.阀门结构形式：单座直通；",
            "针阀": "（*）.阀门结构形式：直通；",
            "闸阀": "（*）.阀门结构形式：单闸板；",
            "疏水阀": "（*）.阀门结构形式：杠杆浮球 ；\n（*）.最大压降：0.1-0.5MPa（出口直接外排）\n（*）.疏水量：",
            "上展式": "（*）.阀门结构形式：上展式；",
            "减压阀": "（*）.精度：±1%；\n（*）.流量特性：等百分比特性\n（*）.取压方式：阀后定压，阀外取压\n（*）.取压接头：G1/2外螺纹\n（*）.阀门结构形式：单座直通式；\n（*）.防护等级：IP65；\n（*）.防爆等级：/；\n（*）.泄露等级：IV级；",
            "安全阀": "（*）.整定压力：；\n（*）.起跳压力：；\n（*）.介质过气量：；\n（*）.分子量：水蒸气，18，\n备注：\n1、排量可按最大过气量计算\n2、安全阀口径可按计算的排气量更改。\n3、安装位置：蒸汽主管\n（*）.其余要求：带扳手弹簧微起式；"
        }

        self.drive_map = {     
            "电动": "\n（*）.精度：±1%；\n（*）.传动方式：电动；\n（*）.流量特性：等百分比特性；\n（*）.输入/输出信号：4-20mA；\n（*）.电源电压：AC220V，四线制；\n（*）.防护等级：IP65；\n（*）.防爆等级：/；\n（*）.泄露等级：V级；\n（*）.电气密封接口：共2个M20*1.5电气接口(1个供电用,1个输入/输出信号共用)；\n",
            "气动": "\n（*）.传动方式：气动；\n（*）.输入信号：数字量；\n（*）.输出信号：数字量（带限位开关）；\n（*）.阀门结构形式：浮动球式，直通；\n（*）.泄露等级：V级；\n（*）.气缸类型：双作用气缸；\n（*）.气缸进气接口：φ8内螺纹；\n（*）.备注：不带电磁阀、二联件",
            "手动": "\n（*）.传动方式：手动；\n"
        }

    def fill_dataframe(self,df):
        self.logger.info(f"🔍 已进入 fill_dataframe，共 {len(df)} 行")
        for index, row in df.iterrows():
            protocol_column = row['协议号']
            valve_name = row['名称']
            drive_type = row['驱动类型']
            medium = row['介质']
            diameter = row['阀门规格']
            temp_in = row.get("入口温度", "")
            temp_out = row.get("出口温度", "")

            self.connection_map = {
                "法兰": f"（*）.连接方式：法兰连接{diameter}，PN10,HG/T-20592-2009，RF，B型；",
                "减压": f"（*）.连接方式：法兰连接{diameter}，PN10,HG/T-20592-2009，RF，B型；",
                "卡箍": f"（*）.安装/接管方式：通径{diameter}，卡箍φ50.5",
                "焊接": f"（*）.安装/接管方式：{diameter}焊接",
                "对夹": f"（*）.连接方式：对夹安装{diameter}，PN10；配套 HG/T 20592-2009 B型 RF法兰；",
                "螺纹": f"（*）.安装/接管方式：内螺纹{diameter}",
                "上展": f"（*）.连接方式：法兰连接{diameter}，PN10,HG/T-20592-2009，RF，B型；"
            }

            if "气动" in valve_name:
                kaibi_fill = "\n（*）.阀门开度：0/100；\n"
            elif "止回阀" in valve_name:
                kaibi_fill = ""
            elif "截止阀" in valve_name:
                kaibi_fill = ""
            elif "疏水阀" in valve_name:
                kaibi_fill = ""
            else:
                kaibi_fill = "\n（*）.阀门开度：0~100；\n"

            structure_found = False
            for key, val in self.connection_map.items():
                if key in valve_name:
                    connection_fill = val
                    structure_found = True
                    self.logger.info(f"行 {index} 的连接方式参数填充为: {connection_fill}")
                    break  
            if not structure_found:
                connection_fill = "未知连接方式参数"
            found = False
            for key, val in self.STRUCTURE_RULES.items():
                if key in valve_name:
                    structure_fill = val
                    found = True
                    break  # 可选：找到就停（推荐）
            if not found:
                structure_fill = "未知结构参数"
            for key, val in self.drive_map.items():
                if key in drive_type:
                    drive_fill = val
                    break
            medium_fill = self.medium_dict.get(medium, "未知介质参数")
            self.logger.info(f"行 {index} 的介质参数填充为: {medium_fill}")
            if not protocol_column or pd.isna(protocol_column):
                try:
                    if "减压" in valve_name:
                        parameter_template = f"（*）介质：蒸汽；\n（*）介质温度：入口{temp_in}℃，出口{temp_out}℃；\n（*）.介质流量：；\n（*）.介质压力:\n{kaibi_fill}；\n{structure_fill}{drive_fill}{connection_fill}"
                        df.at[index, "参数"] = parameter_template
                        self.logger.info(f"✅ 行 {index} 减压阀参数填充完成")
                    else:
                        parameter_tamplate = f"{medium_fill}{kaibi_fill}{structure_fill}{drive_fill}{connection_fill}"
                        parameter_list = []
                        parameter_tamplate = f"{medium_fill}{kaibi_fill}{structure_fill}{drive_fill}{connection_fill}"
                        parameter_list.append(parameter_tamplate)
                        df.at[index, '参数'] = ''.join(parameter_list)
                        self.logger.info(f"✅ 行 {index} 参数填充完成")
                except Exception as e:
                    self.logger.error(f"❌ 行 {index} 参数填充失败: {e}")
        return df   

