import pandas as pd
import os
import logging

class ValveProtocolMatcher:
    def __init__(self, xlsx_path):
        self.xlsx_path = xlsx_path
        self.master_selection = None # 存放合并后的选型手册
        self.master_params = None    # 存放合并后的参数表
        self.logger = logging.getLogger(__name__)
        
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
        result['SKU'] = (
            result['SKU']
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
        left_row = result[result['SKU'] == test_val]
        print(f"🔍 诊断：左表是否存在 '{test_val}'? {'是' if not left_row.empty else '否'}")

        # 检查右表（协议库总表）
        if self.master_selection is not None:
            right_row = self.master_selection[self.master_selection['Match_Code'] == test_val]
            print(f"🔍 诊断：协议库是否存在 '{test_val}'? {'是' if not right_row.empty else '否'}")
            
            if not right_row.empty:
                print(f"🔍 诊断：协议库中对应的协议号是: {right_row['协议号'].values}")
        # --- 诊断结束 ---

        def get_drive_type(row):
            vt = str(row.get('阀门形式',''))
            if "电动" in vt:
                return "电动"
            if "气动" in vt:
                return "气动"
            return "手动"
        result['驱动类型'] = result.apply(get_drive_type, axis=1)
        result['SKU'] = result['SKU'].astype(str).str.strip()
        if self.master_selection is not None:
            result = pd.merge(
                result,
                self.master_selection[['Match_Code', '协议号']],
                left_on=['SKU'],
                right_on=['Match_Code'],
                how='left'
            )
        
        if self.master_params is not None:
            result = pd.merge(
                result,
                self.master_params[['产品编码', '参数']],
                left_on = ['SKU'],
                right_on = ['产品编码'],
                how = 'left'
            )
        self.logger.info(f"✅ 匹配完成，结果包含 {len(result)} 行")
        return result

class ParameterFiller:
    def __init__(self, param_path):
        self.param_path = param_path
        self.logger = logging.getLogger(__name__)
        
        # 1. 加载介质库
        try:
            medium_df = pd.read_excel(param_path, engine='openpyxl')
            # 建立映射字典：介质名称 -> 介质参数
            self.medium_dict = dict(zip(medium_df["介质"], medium_df["参数"]))
        except Exception as e:
            self.logger.error(f"介质库读取失败: {e}")
            self.medium_dict = {}

        # 2. 结构规则库 (注意：这里只保留【物理结构】，去掉了传动和开度，因为那些是动态的)
        self.STRUCTURE_RULES = {
            "球阀": "（*）.阀门结构形式：浮动球式，直通；",
            "蝶阀": "（*）.阀门结构形式：单偏心式；",
            "止回阀": "（*）.阀门结构形式：旋启式单瓣；",
            "截止阀": "（*）.阀门结构形式：单座直通；",
            "针阀": "（*）.阀门结构形式：直通；",
            "闸阀": "（*）.阀门结构形式：单闸板；",
            "疏水阀": "（*）.阀门结构形式：杠杆浮球 ；\n（*）.最大压降：0.1-0.5MPa（出口直接外排）\n（*）.疏水量：",
            "上展式": "（*）.阀门结构形式：上展式；",
            "减压阀": "（*）.精度：±1%；\n（*）.流量特性：等百分比特性\n（*）.取压方式：阀后定压，阀外取压\n（*）.取压接头：G1/2外螺纹\n（*）.阀门结构形式：单座直通式；\n（*）.防护等级：IP65；\n（*）.防爆等级：/；\n（*）.泄露等级：IV级；",
            "安全阀": "（*）.整定压力：；\n（*）.起跳压力：；\n（*）.介质过气量：；\n（*）.分子量：水蒸气，18，\n备注：\n1、排量可按最大过气量计算\n2、安全阀口径可按计算的排气量更改。\n3、安装位置：蒸汽主管\n（*）.其余要求：带扳手弹簧微起式；"
        }

    def _get_connection_param(self, info_string, diameter):
        """
        辅助方法：根据字符串（如阀门名称或连接方式）里的关键字，匹配连接参数
        """
        info_string = str(info_string) # 防止空值报错
        
        # 定义关键字映射规则
        rules = {
            "法兰": f"（*）.连接方式：法兰连接{diameter}，PN10,HG/T-20592-2009，RF，B型；",
            "减压": f"（*）.连接方式：法兰连接{diameter}，PN10,HG/T-20592-2009，RF，B型；",
            "卡箍": f"（*）.安装/接管方式：通径{diameter}，卡箍φ50.5",
            "焊接": f"（*）.安装/接管方式：{diameter}焊接",
            "对夹": f"（*）.连接方式：对夹安装{diameter}，PN10；配套 HG/T 20592-2009 B型 RF法兰；",
            "螺纹": f"（*）.安装/接管方式：内螺纹{diameter}",
            "上展": f"（*）.连接方式：法兰连接{diameter}，PN10,HG/T-20592-2009，RF，B型；"
        }
        
        # 循环判断关键字是否存在
        for key, val in rules.items():
            if key in info_string:
                return val
        
        return "" # 没匹配到返回空

    def process_single_row(self, row):
        """
        处理单行逻辑：返回组装好的参数字典
        """
        # 1. 提取基础信息
        media_key = row.get('介质', '')
        actuator_key = row.get('驱动类型', '') # 比如：气动/电动/手动
        valve_key = row.get('阀门形式', '')     # 比如：气动法兰球阀
        diameter = str(row.get("阀门规格", "")).strip()
        
        # 如果阀门形式为空，可能不需要填充，或者根据具体需求处理
        if not valve_key or pd.isna(valve_key):
            return {} 

        # --- 步骤1：匹配介质参数 ---
        p1_media = self.medium_dict.get(media_key, "")

        # --- 步骤2：生成动力源 & 开度参数 (动态生成！) ---
        # 默认值
        drive_method = "手动" 
        open_range = "0-100"
        
        if "气动" in str(actuator_key) or "气动" in str(valve_key):
            drive_method = "气动"
            open_range = "0或100" # 气动特有逻辑
        elif "电动" in str(actuator_key) or "电动" in str(valve_key):
            drive_method = "电动"
            open_range = "0-100"
        else:
            drive_method = "手动" # 或者是其他默认值
            open_range = "0-100"
            
        # 组装动力源部分的字符串
        p2_actuator_str = f"\n（*）.传动方式：{drive_method}；\n（*）.阀门开度：{open_range}；"

        # --- 步骤3：匹配纯物理结构参数 ---
        p3_structure_body = ""
        for key, val in self.STRUCTURE_RULES.items():
            if key in str(valve_key):
                p3_structure_body = val
                break # 找到匹配项就停止，防止比如“截止阀”匹配到“止回阀”出错（虽然这里key不同，但是个好习惯）
        
        # 将动力源参数 + 物理结构参数 合并
        # 结果示例：(*).传动方式：气动；(*).阀门开度：0或100；(*).阀门结构形式：浮动球式...
        final_structure_param = p2_actuator_str + "\n" + p3_structure_body

        # --- 步骤4：匹配连接参数 ---
        # 优先看有没有“连接方式”这列，如果没有就去“阀门形式”里找关键字
        conn_source = row.get('连接方式') if pd.notna(row.get('连接方式')) else valve_key
        p4_connection = self._get_connection_param(conn_source, diameter)

        # 返回结果
        return {
            "介质参数": p1_media,
            "结构参数": final_structure_param,
            "连接参数": p4_connection
        }

    def fill_dataframe(self, df):
        """
        主入口：循环每一行并填充
        """
        # 找到协议号为空的行（假设这是填充条件）
        mask = df['协议号'].isna()
        
        if mask.any():
            print(f"检测到 {mask.sum()} 行数据需要填充...")
            # 对符合条件的行应用 process_single_row
            # 结果会是一个包含字典的 Series
            results = df['参数'].apply(self.process_single_row, axis=1)
            
            # 将字典拆分回 DataFrame 的列
            # 注意：这里需要确保你的 DataFrame 里有这几列，或者新建这些列
            result_df = pd.DataFrame(results.tolist(), index=results.index)
        return result_df
