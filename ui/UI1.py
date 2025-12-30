import sys
import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QTextEdit, QListWidget, QGroupBox,
    QScrollArea, QMessageBox, QSplitter,QStackedWidget

)
from PyQt5.QtCore import Qt
# 导入您修改后的功能类
from instrument_automation import InstrumentAutomationProcessor
from valve_automation import ValveAutomationProcess  
import logging
from PyQt5.QtCore import QAbstractTableModel, Qt
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTableView, QHeaderView
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


class Automationtool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("自动化工具集")
        self.setGeometry(100, 100, 1200, 800)
        self.processor = None  # 存储功能类实例
        self.initUI()
        self.setup_logging()
    # 添加到 UI1.py 的 Automationtool 类中
    def show_data_preview(self):
        """弹出窗口显示当前处理器的 df_sort 数据"""
        if not hasattr(self, 'current_processor') or self.current_processor is None:
            QMessageBox.warning(self, "提示", "没有处理器实例")
            return
            
        if not hasattr(self.current_processor, 'df_sort') or self.current_processor.df_sort is None:
            QMessageBox.warning(self, "提示", "当前没有处理后的数据 (df_sort 为空)")
            return

        # 创建弹窗
        dialog = QDialog(self)
        dialog.setWindowTitle(f"数据预览 - 共 {len(self.current_processor.df_sort)} 行")
        dialog.resize(1000, 600) # 设置窗口大小

        # 创建表格视图
        layout = QVBoxLayout()
        table_view = QTableView()
        
        # 加载数据模型
        model = PandasModel(self.current_processor.df_sort)
        table_view.setModel(model)
        
        # 稍微美化一下表格
        header = table_view.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive) # 允许拖动列宽
        header.setStretchLastSection(True)
        
        layout.addWidget(table_view)
        dialog.setLayout(layout)
        
        dialog.exec_() # 显示窗口

    def setup_logging(self):
        logger = logging.getLogger('UI')
        logger.setLevel(logging.INFO)
        logger.handlers.clear()

        handler = QTextEditLogger()
        handler.set_widget(self.log_text_edit)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(handler)

        console = logging.StreamHandler()
        console.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
        logger.addHandler(console)
        self.logger = logger

    def initUI(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        # 使用 QSplitter 实现左右分隔
        h_splitter = QSplitter(Qt.Horizontal)
        main_layout = QHBoxLayout(main_widget)
        main_layout.addWidget(h_splitter)

        # 1. 左侧：工具列表 (Sidebar)
        self.sidebar = self.create_sidebar()
        h_splitter.addWidget(self.sidebar)
        self.sidebar.setFixedWidth(200)

        # 2. 右侧：工具内容区 (QStackedWidget) + 日志区
        # 为了让日志区独立显示，我们用一个 QVBoxLayout 来容纳 StackedWidget 和 Log
        
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        # 2a. 工具页面切换区
        self.stacked_widget = QStackedWidget()
        right_layout.addWidget(self.stacked_widget, 4) # 权重 4

        # 2b. 日志区
        log_group = QGroupBox("日志输出")
        log_layout = QVBoxLayout()
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        log_layout.addWidget(self.log_text_edit)
        log_group.setLayout(log_layout)
        right_layout.addWidget(log_group, 3) # 权重 1 (日志区占1/5的高度)

        h_splitter.addWidget(right_panel)
        h_splitter.setSizes([200, 1000])


        # 3. 创建并添加所有工具的页面 (这一步您做得很好)
        self.pages = {}
        # 注意：您需要实现这些创建方法，并在内部创建独立的输入框属性
        self.pages["阀门提单自动化工具"] = self.create_valve_page() 
        self.pages["仪表提单自动化"] = self.create_instrument_page()
        #self.pages["功率表自动化工具"] = self.create_power_meter_page()#待开发 

        for page in self.pages.values():
            self.stacked_widget.addWidget(page)

        # 4. 初始化日志系统 (注意在创建 self.log_text_edit 后立即调用)
        self.setup_logging() 
        
        # 5. 默认选中第一个工具并显示其页面
        self.tool_list.setCurrentRow(0)

    def create_sidebar(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        title_label = QLabel("工具列表")
        layout.addWidget(title_label)
        self.tool_list = QListWidget()
        self.tool_list.addItems(["阀门提单自动化工具", "仪表提单自动化", "功率表自动化工具"])
        self.tool_list.currentRowChanged.connect(self.on_tool_changed)
        layout.addWidget(self.tool_list)
        layout.addStretch()
        return widget

    # 界面整体窗口布局
    def create_instrument_page(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)

        # 文件选择
        file_group = QGroupBox("文件操作")
        file_layout = QVBoxLayout()
         # 第一行：输入文件
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("输入文件:"))
        self.instr_input_file_path = QLineEdit()
        self.instr_input_file_path.setReadOnly(True)
        btn_browse = QPushButton("浏览")
        btn_browse.clicked.connect(self.browse_instr_file)
        input_layout.addWidget(self.instr_input_file_path)
        input_layout.addWidget(btn_browse)
        file_layout.addLayout(input_layout)
        file_group.setLayout(file_layout)

        # 第二行：申购人
        applicant_layout = QHBoxLayout() # 创建一个从左到右的水平布局
        applicant_layout.addWidget(QLabel("申购人:")) #创建文字标签
        self.instr_input_applicant = QLineEdit() #创建输入框
        self.instr_input_applicant.setPlaceholderText("请输入申购人姓名") #提示底纹
        applicant_layout.addWidget(self.instr_input_applicant) #把上述的东西放入水平布局中
        file_layout.addLayout(applicant_layout)  # 添加到布局

        # 第三行：申购日期
        date_layout = QHBoxLayout() # 创建一个从左到右的水平布局
        date_layout.addWidget(QLabel("需求日期:")) #创建文字标签
        self.instr_input_date = QLineEdit() #创建输入框
        self.instr_input_date.setPlaceholderText("需求日期") #提示底纹
        date_layout.addWidget(self.instr_input_date) #把上述的东西放入水平布局中
        file_layout.addLayout(date_layout)  # 添加到布局

        # 第四行：项目号
        number_layout = QHBoxLayout() # 创建一个从左到右的水平布局
        number_layout.addWidget(QLabel("项目号:")) #创建文字标签
        self.instr_input_number = QLineEdit() #创建输入框
        self.instr_input_number.setPlaceholderText("项目号") #提示底纹
        number_layout.addWidget(self.instr_input_number) #把上述的东西放入水平布局中
        file_layout.addLayout(number_layout)  # 添加到布局


        # 执行按钮
        exec_group = QGroupBox("执行操作")
        exec_layout = QVBoxLayout()
        self.instr_btn_process = QPushButton("开始处理")
        self.instr_btn_process.clicked.connect(self.start_processing)
        exec_layout.addWidget(self.instr_btn_process)
        exec_group.setLayout(exec_layout)

        # 另存为按钮
        save_group = QGroupBox("保存结果")
        save_layout = QVBoxLayout()
        self.instr_btn_save_as = QPushButton("另存为 Excel")
        self.instr_btn_save_as.clicked.connect(self.save_as_excel)
        self.instr_btn_save_as.setEnabled(False)
        save_layout.addWidget(self.instr_btn_save_as)
        save_group.setLayout(save_layout)

        # 添加到布局
        content_layout.addWidget(file_group)
        content_layout.addWidget(exec_group)
        content_layout.addWidget(save_group)
        content_layout.addStretch()

        scroll_area.setWidget(content_widget)
        return scroll_area

    def create_valve_page(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)

        # 文件选择
        file_group = QGroupBox("文件操作")
        file_layout = QVBoxLayout()
         # 第一行：输入文件
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("输入文件:"))
        self.valve_input_file_path = QLineEdit()
        self.valve_input_file_path.setReadOnly(True)
        btn_valve_browse = QPushButton("浏览")
        btn_valve_browse.clicked.connect(self.browse_valve_file)
        input_layout.addWidget(self.valve_input_file_path)
        input_layout.addWidget(btn_valve_browse)
        file_layout.addLayout(input_layout)
        file_group.setLayout(file_layout)

        # 第二行：申购人
        applicant_layout = QHBoxLayout() # 创建一个从左到右的水平布局
        applicant_layout.addWidget(QLabel("申购人:")) #创建文字标签
        self.valve_input_applicant = QLineEdit() #创建输入框
        self.valve_input_applicant.setPlaceholderText("请输入申购人姓名") #提示底纹
        applicant_layout.addWidget(self.valve_input_applicant) #把上述的东西放入水平布局中
        file_layout.addLayout(applicant_layout)  # 添加到布局

        # 第三行：申购日期
        date_layout = QHBoxLayout() # 创建一个从左到右的水平布局
        date_layout.addWidget(QLabel("需求日期:")) #创建文字标签
        self.valve_input_date = QLineEdit() #创建输入框
        self.valve_input_date.setPlaceholderText("需求日期") #提示底纹
        date_layout.addWidget(self.valve_input_date) #把上述的东西放入水平布局中
        file_layout.addLayout(date_layout)  # 添加到布局

        # 第四行：项目号
        number_layout = QHBoxLayout() # 创建一个从左到右的水平布局
        number_layout.addWidget(QLabel("项目号:")) #创建文字标签
        self.valve_input_number = QLineEdit() #创建输入框
        self.valve_input_number.setPlaceholderText("项目号") #提示底纹
        number_layout.addWidget(self.valve_input_number) #把上述的东西放入水平布局中
        file_layout.addLayout(number_layout)  # 添加到布局


        # 执行按钮
        exec_group = QGroupBox("执行操作")
        exec_layout = QVBoxLayout()
        self.valve_btn_process = QPushButton("开始处理")
        self.valve_btn_process.clicked.connect(self.start_processing)
        exec_layout.addWidget(self.valve_btn_process)
        exec_group.setLayout(exec_layout)

        # 另存为按钮
        save_group = QGroupBox("保存结果")
        save_layout = QVBoxLayout()
        self.valve_btn_save_as = QPushButton("另存为 Excel")
        self.valve_btn_save_as.clicked.connect(self.save_as_excel)
        self.valve_btn_save_as.setEnabled(False)
        save_layout.addWidget(self.valve_btn_save_as)
        save_group.setLayout(save_layout)

        # 添加到布局
        content_layout.addWidget(file_group)
        content_layout.addWidget(exec_group)
        content_layout.addWidget(save_group)
        content_layout.addStretch()

        scroll_area.setWidget(content_widget)
        return scroll_area
    def browse_instr_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择仪表输入文件", "", "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.instr_input_file_path.setText(file_path)

    def browse_valve_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择阀门输入文件", "", "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.valve_input_file_path.setText(file_path)
    

    def start_processing(self):
        # 使用 self.current_tool_name 获取当前激活的工具，这个值在 on_tool_changed 中设置
        tool_name = getattr(self, 'current_tool_name', None)
        
        if not tool_name:
            QMessageBox.warning(self, "警告", "请先在左侧选择一个自动化工具！")
            return

        # --- 1. 定义工具配置映射 ---
        tool_map = {
            "仪表提单自动化": {
                "processor_class": InstrumentAutomationProcessor,
                "inputs": {
                    "file": self.instr_input_file_path,
                    "applicant": self.instr_input_applicant,
                    "date": self.instr_input_date,
                    "number": self.instr_input_number,
                },
                # 定义仪表工具的执行步骤
                "processing_steps": [
                    {"method": "load_csv","store_as": "df"}, 
                    # 执行 generate_code (仪表有此方法)
                    {"method": "generate_code","store_as": "df_sort"}, 
                    # 调用 get_note，需要传入 self.current_processor.df_sort 作为 df 参数
                    {"method": "get_note", "args": {"df": "df_sort"}, "store_as": "df"}, 
                    # 执行 merge_by_SKU
                    {"method": "merge_by_SKU","store_as": "book"}, 
                ],
                "save_button": self.instr_btn_save_as,
            },
            "阀门提单自动化工具": {
                "processor_class": ValveAutomationProcess,
                "inputs": {
                    "file": self.valve_input_file_path,
                    "applicant": self.valve_input_applicant,
                    "date": self.valve_input_date,
                    "number": self.valve_input_number,
                },
                # 定义仪表工具的执行步骤
                "processing_steps": [
                    {"method": "load_csv","store_as": "df"}, 
                    # 执行 generate_code (仪表有此方法)
                    {"method": "generate_code","store_as": "df_sort"},              
                    ],
                "save_button": self.valve_btn_save_as,
            }
            
        }
        
        if tool_name not in tool_map:
            QMessageBox.critical(self, "错误", f"未识别的工具名称: {tool_name}")
            return

        config = tool_map[tool_name]
        
        # --- 2. 校验输入数据 ---
        input_values = {key: widget.text().strip() for key, widget in config["inputs"].items()}
        
        if not input_values["file"]:
            QMessageBox.warning(self, "警告", "请先选择输入文件！")
            return
        if not input_values["applicant"]:
            QMessageBox.warning(self, "警告", "请输入申购人！")
            return
        if not input_values["date"]:
            QMessageBox.warning(self, "警告", "请输入需求日期！")
            return
        if not input_values["number"]:
            QMessageBox.warning(self, "警告", "请输入项目号！")
            return

        file_path = input_values["file"]
        # --- 3. 执行处理逻辑 ---
        try:
            self.logger.info(f"🚀 开始处理【{tool_name}】...")
            ProcessorClass = config["processor_class"]

            # ... (ProcessorClass None 检查)

            # 实例化处理器，并设置日志
            self.current_processor = ProcessorClass(file_path)
            # 假设 setup_logging 存在于处理器类中
            if hasattr(self.current_processor, 'setup_logging'):
                self.current_processor. setup_logging(self.log_text_edit)
            self.logger.info("✅ 处理器实例和日志系统已初始化。")

            # ----------------------------------------------------
            # 核心修改区域：用反射循环替换所有硬编码处理步骤
            # ----------------------------------------------------
            for step in config.get("processing_steps", []):
                method_name = step["method"]
                
                # 1. 检查方法是否存在（解决 generate_code 不存在时报错的问题）
                if not hasattr(self.current_processor, method_name):
                    self.logger.warning(f"跳过：处理器 {tool_name} 不包含方法 '{method_name}'。")
                    continue

                method = getattr(self.current_processor, method_name)
                
                # 2. 准备参数
                args = {}
                if "args" in step:
                    for arg_name, attr_name in step["args"].items():
                        if hasattr(self.current_processor, attr_name):
                            args[arg_name] = getattr(self.current_processor, attr_name)
                        else:
                            self.logger.error(f"执行 {method_name} 失败：参数属性 '{attr_name}' 在处理器中不存在。")
                            raise AttributeError(f"Missing required attribute for {method_name}: {attr_name}")

                self.logger.info(f"▶️ 正在执行方法: {method_name}")
                
                # 3. 执行方法
                result = method(**args)

                # 4. 存储结果
                if "store_as" in step and result is not None:
                    setattr(self.current_processor, step["store_as"], result)
                    self.logger.info(f"结果已存储到 self.{step['store_as']}")

            # ----------------------------------------------------
            # 统一写入申购人信息 (必须在处理循环后，且在保存前)
            # ----------------------------------------------------

            # ⚠️ 修正语法错误，并检查 df_sort 是否由循环创建成功
            if hasattr(self.current_processor, 'df_sort') and self.current_processor.df_sort is not None:
                self.current_processor.df_sort['申购人'] = input_values["applicant"]
                self.current_processor.df_sort['申购日期'] = input_values["date"]
                self.current_processor.df_sort['项目号'] = input_values["number"]
                
                # 假设 get_note 的功能是写入 Note，它应该在循环中被调用。
                # 如果您一定要在这里做更多操作，请确保不会重复。

                # ⚠️ 注意：移除了原来这里的 merge_by_SKU，因为它现在在配置循环中执行。
                
                self.logger.info("写入申购人信息完成。")
            else:
                # 如果 df_sort 仍未创建，则日志会输出此警告（说明配置或类方法有问题）
                self.logger.warning("处理器中没有找到 df_sort 属性，跳过写入申购信息。")

            # --- 4. 启用保存按钮 ---
            config["save_button"].setEnabled(True)
            self.logger.info(f"🎉 【{tool_name}】处理成功！可点击【另存为】保存结果。")

        except Exception as e:
            self.logger.error(f"❌ 处理失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"【{tool_name}】处理失败：{str(e)}")

        # 在 start_processing 函数的最后
        self.logger.info(f"🎉 【{tool_name}】处理成功！")
        # 自动弹出预览窗口
        self.show_data_preview()
    def save_as_excel(self):
        """
        弹出文件保存对话框，并根据当前激活的工具，调用其处理器保存结果。
        该方法优先尝试保存 book 对象，其次尝试保存 df_sort 对象。
        """
        # 1. 前置检查
        # 检查是否有正在处理的实例
        if not hasattr(self, 'current_processor') or self.current_processor is None:
            QMessageBox.warning(self, "警告", "请先执行【开始处理】以生成数据！")
            self.logger.warning("❌ 保存操作失败: 处理器实例不存在。")
            return
            
        # --- 2. 确定当前工具的保存文件名和标题 ---
        tool_name = getattr(self, 'current_tool_name', None)
        
        if tool_name == "仪表提单自动化":
            title = "仪表申购单"
            input_number = getattr(self.current_processor, 'instr_input_number', self.instr_input_number).text().strip()
        elif tool_name == "阀门提单自动化工具":
            title = "阀门申购单"
            input_number = getattr(self.current_processor, 'valve_input_number', self.valve_input_number).text().strip()
        elif tool_name == "功率表自动化工具":
            title = "功率表申购单"
            input_number = getattr(self.current_processor, 'power_input_number', self.power_input_number).text().strip()
        else:
            QMessageBox.critical(self, "错误", "无法识别当前工具类型，无法保存。")
            self.logger.error("❌ 保存操作失败: 无法识别当前工具类型。")
            return

        default_name = f"【{input_number}】{title}"
        
        # --- 3. 确定要保存的对象 (book 或 df_sort) ---
        save_object = None
        
        # 优先检查是否存在 self.book (openpyxl Workbook 对象)
        if hasattr(self.current_processor, 'book') and self.current_processor.book is not None:
            save_object = self.current_processor.book
            self.logger.info("✅ 找到最终保存对象: book (Workbook)。")
            
        # 其次检查是否存在 self.df_sort (Pandas DataFrame)
        elif hasattr(self.current_processor, 'df_sort') and self.current_processor.df_sort is not None:
            save_object = self.current_processor.df_sort
            self.logger.warning("⚠️ 找到中间结果: df_sort (DataFrame)。")
        
        if save_object is None:
            QMessageBox.warning(self, "警告", "当前处理器中没有可保存的数据 (book 和 df_sort 均为空)！")
            self.logger.error("❌ 处理器中没有找到可保存的数据。")
            return

        # --- 4. 弹出保存对话框 ---
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            f"另存为 {title} 文件", 
            f"{default_name}-{timestamp}.xlsx", 
            "Excel 文件 (*.xlsx)"
        )
        
        if not file_path:
            return # 用户取消保存

        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        # --- 5. 执行保存操作 ---
        try:
            if hasattr(save_object, 'save'):
                # 如果是 openpyxl.Workbook 对象，使用 save 方法
                save_object.save(file_path)
            elif hasattr(save_object, 'to_excel'):
                # 如果是 pandas DataFrame，使用 to_excel 方法
                save_object.to_excel(file_path, index=False)
            else:
                raise TypeError("处理器结果不包含有效的保存方法 (book.save 或 df.to_excel)。")
                
            self.logger.info(f"🎉 文件已成功保存：{file_path}")
            QMessageBox.information(self, "成功", f"文件已保存：\n{file_path}")
            
        except Exception as e:
            self.logger.error(f"❌ 文件保存失败: {str(e)}")
            QMessageBox.critical(self, "保存失败", f"无法保存文件：\n{str(e)}")

    def on_tool_changed(self, index):
        """
        处理左侧工具列表选择变化的槽函数。
        根据选中的工具，切换 QStackedWidget 中显示的页面。
        """
        # 确保索引有效
        if index < 0:
            return

        # 1. 获取当前选中工具的名称
        # self.tool_list 是 QListWidget
        tool_name = self.tool_list.item(index).text()
        
        # 2. 查找并切换 QStackedWidget 页面
        # self.pages 是一个字典，存储了工具名称到对应页面的映射
        if tool_name in self.pages:
            page_widget = self.pages[tool_name]
            # 找到该页面在 QStackedWidget 中的索引
            page_index = self.stacked_widget.indexOf(page_widget)
            
            # 切换到对应的页面
            self.stacked_widget.setCurrentIndex(page_index)
            
            # 3. 记录当前激活的工具名称
            # 这一步非常重要，用于 start_processing 方法中判断应该执行哪个工具的逻辑
            self.current_tool_name = tool_name
            
            self.logger.info(f"切换到工具: {tool_name}")
        else:
            self.logger.warning(f"未找到工具 '{tool_name}' 的 UI 页面。")
    
    # 放在 UI1.py 文件的最下面，或者类定义之外
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super(PandasModel, self).__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Automationtool()
    win.show()
    sys.exit(app.exec_())