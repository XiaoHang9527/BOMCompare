import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, font as tk_font
from tkinter.font import Font
import pandas as pd
import datetime
import re  # 添加re模块导入
import traceback
from datetime import datetime
import random
import time
import threading  # 添加threading模块导入
import difflib
# 添加更新功能所需的库
import requests
import tempfile
import shutil
import zipfile
import subprocess
import platform
from packaging import version as pkg_version

# 定义版本信息和更新相关常量
APP_VERSION = "1.4"
GITHUB_REPO = "XiaoHang9527/BOMCompare"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_CHECK_INTERVAL = 7  # 天

# 定义下载重试次数和超时时间
DOWNLOAD_MAX_RETRIES = 3  # 最大重试次数
DOWNLOAD_TIMEOUT = 30     # 下载超时时间(秒)
DOWNLOAD_CHUNK_SIZE = 8192  # 下载块大小

# 避免Windows上打包后的UTF-8编码问题
if sys.platform.startswith('win'):
    import locale
    locale.setlocale(locale.LC_ALL, 'C')

class BOMComparer:
    def __init__(self, parent_window=None):
        """初始化BOM比较器

        Args:
            parent_window: 父窗口，用于显示弹出对话框
        """
        # 设置父窗口引用
        self.parent_window = parent_window

        # 设置默认的字段映射
        self.field_mappings = {
            'Item': ['Item', 'item', '序号', 'Number'],
            'P/N': ['P/N', '料号', '物料编码', '物料编号', 'Part Number', '型号'],
            'Reference': ['Reference', 'Ref', 'ref', '位号'],
            'Description': ['Description', '描述', '物料描述'],
            'MPN': ['Manufacturer P/N', 'MPN', '制造商料号', '厂家料号', '生产商料号']
        }

        # 报告显示设置
        self.show_mpn_in_report = True  # 默认显示MPN信息

        # 替代料映射字典 - 可以由用户配置
        self.alternative_map = {}

        # 进度回调函数
        self.progress_callback = None

        # 存储单个BOM文件的全局变量
        self.bom_a = None
        self.bom_b = None

        # 用于记录处理时间
        self.start_time = None
        self.end_time = None  # 结束时间

        # 错误信息映射字典
        self.error_messages = {
            'FileNotFoundError': '文件不存在，请检查文件路径是否正确',
            'PermissionError': '无法访问文件，请检查文件是否被其他程序占用',
            'EmptyDataError': '文件内容为空，请检查文件是否有数据',
            'XLRDError': '不支持的文件格式，请使用Excel文件(.xlsx或.xls)',
            'HeaderError': '无法识别表头，请检查文件格式是否正确',
            'MissingFields': '缺少必要的字段，请检查文件是否包含所需的列',
            'InvalidFormat': '文件格式不正确，请使用标准的BOM表格式',
            'ProcessingError': '处理过程中出错，请检查文件内容是否符合要求',
            'EncodingError': '文件编码错误，请使用UTF-8编码保存文件',
            'MemoryError': '文件过大，内存不足，请尝试减小文件大小',
            'DependencyError': '缺少必要的依赖库，请安装所需的Python库',
            'default': '发生未知错误，请检查文件格式和内容是否正确'
        }

    def set_progress_callback(self, callback):
        """设置进度回调函数"""
        self.progress_callback = callback

    def set_field_mappings(self, field_mappings):
        """设置字段映射字典

        Args:
            field_mappings (dict): 字段映射字典，格式为 {标准字段名: [可能的别名列表]}
        """
        # 确保必要的字段存在
        required_fields = ['Item', 'P/N', 'Reference', 'Description', 'MPN']
        for field in required_fields:
            if field not in field_mappings:
                field_mappings[field] = self.field_mappings.get(field, [])

        self.field_mappings = field_mappings

    def update_progress(self, progress, message=""):
        """更新进度信息"""
        if self.progress_callback:
            self.progress_callback(progress, message)

    def optimize_column_widths(self, tree, data, columns, sample_size=50):
        """智能优化列宽度

        Args:
            tree: 表格树控件
            data: 数据DataFrame
            columns: 列名列表
            sample_size: 用于检查的数据行数
        """
        if len(data) == 0:
            return

        # 字体实例用于计算文本宽度
        font = tk_font.Font()

        # 为特定列设置最大宽度限制
        column_max_widths = {
            'Item': 80,       # Item列宽度限制为80像素
            'Quantity': 80    # Quantity列宽度限制为80像素
        }

        # 对每一列优化宽度
        for col in columns:
            # 初始宽度为列标题宽度
            header_width = font.measure(str(col)) + 20

            # 采样数据行以计算内容的最大宽度
            max_content_width = 0
            sample_indices = list(range(min(sample_size, len(data))))
            if len(data) > sample_size:
                # 随机选择一些行
                sample_indices = random.sample(range(len(data)), sample_size)

            for i in sample_indices:
                content = str(data.iloc[i].get(col, ''))
                content_width = font.measure(content) + 20
                max_content_width = max(max_content_width, content_width)

            # 获取该列的最大宽度限制（如果有）
            max_width_limit = column_max_widths.get(col, 300)

            # 设置最终宽度 - 标题宽度和内容宽度的最大值，但不超过列的最大宽度限制
            final_width = min(max_width_limit, max(header_width, max_content_width, 80))
            tree.column(col, width=final_width)

    def load_bom(self, file_path):
        """加载BOM文件并处理"""
        try:
            print(f"\n=== 加载文件: {os.path.basename(file_path)} ===")

            # 检查文件是否存在
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"文件不存在: {file_path}")

            # 检查文件扩展名
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext not in ['.xlsx', '.xls']:
                raise ValueError("不支持的文件格式，请使用Excel文件(.xlsx或.xls)")

            # 首先不指定header，读取整个文件作为原始数据
            try:
                df_raw = pd.read_excel(file_path, header=None)
            except Exception as e:
                error_msg = str(e)
                if "XLRDError" in error_msg:
                    raise ValueError("不支持的文件格式，请使用Excel文件(.xlsx或.xls)")
                elif "EmptyDataError" in error_msg:
                    raise ValueError("文件内容为空，请检查文件是否有数据")
                elif "PermissionError" in error_msg:
                    raise ValueError("无法访问文件，请检查文件是否被其他程序占用")
                elif "Missing optional dependency" in error_msg and "xlrd" in error_msg:
                    raise ValueError("缺少读取Excel所需的库，请安装xlrd库: pip install xlrd>=2.0.1")
                else:
                    raise ValueError(f"读取文件失败: {error_msg}")

            print(f"原始数据行数: {len(df_raw)}")

            # 使用实例的字段映射字典，而不是硬编码的
            field_mappings = self.field_mappings

            # 尝试识别表头行
            header_row = -1
            max_matches = 0

            # 检查前20行，寻找最可能的表头行
            for row_idx in range(min(20, len(df_raw))):
                row = df_raw.iloc[row_idx]
                match_count = 0

                # 计算该行与期望字段的匹配度
                for field_aliases in field_mappings.values():
                    for alias in field_aliases:
                        if any(str(cell).lower() == alias.lower() for cell in row):
                            match_count += 1
                            break

                # 记录匹配度最高的行
                if match_count > max_matches:
                    max_matches = match_count
                    header_row = row_idx

            # 如果找不到合适的表头行，尝试读取第一行作为表头
            if header_row < 0:
                header_row = 0

            # 重新读取文件，使用识别出的表头行
            df = pd.read_excel(file_path, header=header_row)

            # 用于存储实际列名到标准列名的映射
            column_map = {}

            # 检查并映射列名
            for std_field, possible_names in field_mappings.items():
                found = False
                for name in possible_names:
                    for col in df.columns:
                        # 精确匹配
                        if str(col).lower() == name.lower():
                            column_map[std_field] = col
                            found = True
                            break
                        # 包含匹配（例如"物料编码(P/N)"包含"物料编码"）
                        elif name.lower() in str(col).lower():
                            column_map[std_field] = col
                            found = True
                            break
                    if found:
                        break

            # 检查是否所有必要字段都找到了映射
            missing_fields = []
            required_fields = ['P/N', 'Reference']  # 只有这两个字段是绝对必要的
            optional_fields = ['Description', 'MPN']  # 这些字段是可选的

            for field in required_fields:
                if field not in column_map:
                    missing_fields.append(field)

            # 如果有必要字段缺失，尝试弹出对话框让用户选择
            if missing_fields and hasattr(self, 'parent_window'):
                for missing_field in missing_fields[:]:  # 使用切片创建副本，避免在循环中修改
                    # 构建提示信息
                    if missing_field == 'Reference':
                        title = "选择位号列"
                        message = "未能自动识别位号列，请手动选择:"
                    elif missing_field == 'P/N':
                        title = "选择物料编号列"
                        message = "未能自动识别物料编号列，请手动选择:"
                    else:
                        title = f"选择{missing_field}列"
                        message = f"未能自动识别{missing_field}列，请手动选择:"

                    # 弹出对话框让用户选择
                    dialog = FieldMappingDialog(
                        self.parent_window,
                        title,
                        message,
                        [str(col) for col in df.columns],
                        field_type=missing_field
                    )

                    # 如果用户选择了字段
                    if dialog.result:
                        selected_column = dialog.result['field']
                        remember = dialog.result['remember']

                        # 更新映射
                        column_map[missing_field] = selected_column
                        missing_fields.remove(missing_field)

                        # 如果用户选择记住选择，更新字段映射
                        if remember:
                            self.field_mappings[missing_field].insert(0, selected_column)
                    else:
                        # 用户取消了选择，中断处理
                        raise ValueError(f"用户取消了{missing_field}列的选择")

            # 如果仍有缺失字段，报错
            if missing_fields:
                missing_fields_str = '、'.join(missing_fields)
                raise ValueError(f"BOM文件缺少必要字段: {missing_fields_str}，请检查文件格式")

            # 添加可选字段的默认值
            for field in optional_fields:
                if field not in column_map:
                    print(f"警告: 未找到字段 '{field}'，将使用默认空值")
                    # 添加空列
                    df[field] = ""
                    column_map[field] = field

            # 重命名列以标准化
            df_renamed = df.rename(columns={v: k for k, v in column_map.items()})

            # 如果没有找到Item字段，创建一个
            if 'Item' not in column_map:
                df_renamed['Item'] = range(1, len(df_renamed) + 1)

            # 确保所有必需字段都存在
            for field in ['Item', 'P/N', 'Reference', 'Description', 'MPN']:
                if field not in df_renamed.columns:
                    df_renamed[field] = '' if field != 'Item' else range(1, len(df_renamed) + 1)

            # 查找实际数据开始的行（跳过项目信息行）
            data_start_row = 0

            # 尝试使用多种方法找到数据起始行

            # 方法1: 基于Item列中的数字
            if 'Item' in df_renamed.columns:
                try:
                    # 查找第一个数字开头的项
                    numeric_rows = df_renamed[df_renamed['Item'].astype(str).str.match(r'^\d+$')].index
                    if len(numeric_rows) > 0:
                        data_start_row = numeric_rows[0]
                except:
                    pass

            # 方法2: 基于Reference列的非空值
            if data_start_row == 0 and 'Reference' in df_renamed.columns:
                try:
                    # 查找第一个非空的Reference行
                    non_empty_refs = df_renamed[df_renamed['Reference'].astype(str).str.strip() != ''].index
                    if len(non_empty_refs) > 0:
                        data_start_row = non_empty_refs[0]
                except:
                    pass

            # 裁剪数据，保留实际的BOM行
            if data_start_row > 0:
                df_renamed = df_renamed.iloc[data_start_row:].reset_index(drop=True)

            # 确保Reference列是字符串类型
            df_renamed['Reference'] = df_renamed['Reference'].astype(str)

            # 移除空的Reference行
            df_renamed = df_renamed[df_renamed['Reference'].str.strip() != '']
            df_renamed = df_renamed[df_renamed['Reference'].str.strip().str.lower() != 'nan']

            # 移除完全相同的重复行
            df_renamed = df_renamed.drop_duplicates()

            # 重置索引
            df_renamed = df_renamed.reset_index(drop=True)

            # 数据统计和检查
            stats = {
                '总行数': len(df_renamed),
                '唯一物料数': len(df_renamed['P/N'].unique()),
                '唯一位号数': len(set(','.join(df_renamed['Reference'].astype(str)).split(','))),
                '空位号行数': len(df_renamed[df_renamed['Reference'].astype(str).str.strip() == '']),
                '重复物料行数': len(df_renamed) - len(df_renamed.drop_duplicates(['P/N']))
            }

            print("\n=== 数据统计 ===")
            for key, value in stats.items():
                print(f"{key}: {value}")

            # 检查潜在问题
            warnings = []
            if stats['空位号行数'] > 0:
                warnings.append(f"发现 {stats['空位号行数']} 行空位号数据")
            if stats['重复物料行数'] > 0:
                warnings.append(f"发现 {stats['重复物料行数']} 行重复物料")

            if warnings:
                print("\n=== 警告信息 ===")
                for warning in warnings:
                    print(f"警告: {warning}")

            # 检查是否有有效数据
            if len(df_renamed) == 0:
                raise ValueError("处理后的BOM数据为空，请检查文件格式和内容")

            # 在处理完数据后，根据Item列识别替代料关系
            def extract_main_item(item):
                """从Item值中提取主序号（例如：从'1.2'提取'1'）"""
                try:
                    # 尝试将item转换为字符串并分割
                    item_str = str(item)
                    if '.' in item_str:
                        return item_str.split('.')[0]
                    return item_str
                except:
                    return str(item)

            # 创建基于Item的替代料映射
            item_based_alt_map = {}

            # 检查是否有Item列
            if 'Item' in df_renamed.columns:
                # 按主序号分组
                item_groups = {}
                for _, row in df_renamed.iterrows():
                    item_val = row.get('Item')
                    if pd.notna(item_val):  # 确保不是空值
                        main_item = extract_main_item(item_val)
                        if main_item not in item_groups:
                            item_groups[main_item] = []
                        item_groups[main_item].append(row)

                # 对每个分组处理替代料关系
                for main_item, rows in item_groups.items():
                    if len(rows) > 1:  # 如果同一个主序号有多行
                        # 收集该组中的所有料号
                        all_pns = [row['P/N'] for row in rows]
                        # 将这些料号作为互为替代料关系
                        if len(all_pns) >= 2:  # 至少需要两个料号才构成替代料关系
                            # 直接将这些料号作为一个替代料组
                            # 对组中的每个料号，将其他所有料号作为其替代料
                            for pn in all_pns:
                                # 将其他所有料号作为当前料号的替代料
                                alt_pns = [p for p in all_pns if p != pn]
                                if alt_pns:  # 确保有替代料
                                    item_based_alt_map[pn] = alt_pns
                                    print(f"基于Item识别到替代料关系: {pn} -> {alt_pns}")

            # 将基于Item的替代料关系添加到替代料映射中
            for main_pn, alt_pns in item_based_alt_map.items():
                if main_pn not in self.alternative_map:
                    self.alternative_map[main_pn] = []
                for alt_pn in alt_pns:
                    if alt_pn not in self.alternative_map[main_pn]:
                        self.alternative_map[main_pn].append(alt_pn)

            # 返回处理后的DataFrame
            return df_renamed

        except Exception as e:
            error_type = type(e).__name__
            error_msg = str(e)

            # 获取友好的错误信息
            friendly_msg = self.error_messages.get(error_type, self.error_messages['default'])

            # 如果是已知的错误类型，使用预定义的友好消息
            if error_type in self.error_messages:
                error_msg = friendly_msg
            # 否则尝试将技术错误信息转换为友好消息
            else:
                if "XLRDError" in error_msg:
                    error_msg = self.error_messages['XLRDError']
                elif "EmptyDataError" in error_msg:
                    error_msg = self.error_messages['EmptyDataError']
                elif "PermissionError" in error_msg:
                    error_msg = self.error_messages['PermissionError']
                elif "encoding" in error_msg.lower():
                    error_msg = self.error_messages['EncodingError']
                elif "memory" in error_msg.lower():
                    error_msg = self.error_messages['MemoryError']
                elif "Missing optional dependency" in error_msg and "xlrd" in error_msg:
                    error_msg = "缺少读取Excel所需的库，请安装xlrd库: pip install xlrd>=2.0.1"
                else:
                    error_msg = f"{friendly_msg}"

            print(f"BOM文件处理出错: {error_msg}")
            if hasattr(self, 'update_progress'):
                self.update_progress(0, error_msg)
            raise ValueError(error_msg)

    def set_alternative_map(self, alt_map):
        """设置物料替代关系映射"""
        self.alternative_map = alt_map

    def get_material_key(self, pn):
        """获取物料主料号（处理替代料关系）"""
        for main_pn, alt_pns in self.alternative_map.items():
            if pn in alt_pns or pn == main_pn:
                return main_pn
        return pn

    def compare(self, bom_a, bom_b, is_dataframe=False):
        """比较两个BOM文件

        Args:
            bom_a: 基准BOM文件路径或DataFrame
            bom_b: 对比BOM文件路径或DataFrame
            is_dataframe: 如果为True，则bom_a和bom_b是DataFrame，否则是文件路径

        Returns:
            dict: 包含比较结果的字典
        """
        try:
            # 记录开始时间
            self.start_time = datetime.now()
            print(f"比较开始时间: {self.start_time}")

            # 加载或使用已有的DataFrame
            self.update_progress(5, "准备数据...")

            if not is_dataframe:
                # 从文件加载
                bom_a_df = self.load_bom(bom_a)
                self.update_progress(20, "基准BOM加载完成")

                bom_b_df = self.load_bom(bom_b)
                self.update_progress(40, "对比BOM加载完成")
            else:
                # 直接使用提供的DataFrame
                bom_a_df = bom_a
                bom_b_df = bom_b
                self.update_progress(40, "使用已加载的数据")

            print(f"BOM A 数据行数: {len(bom_a_df)}, 列: {list(bom_a_df.columns)}")
            print(f"BOM B 数据行数: {len(bom_b_df)}, 列: {list(bom_b_df.columns)}")

            # 提取A和B中的物料编号和位号信息
            self.update_progress(50, "分析BOM数据...")

            # 处理位号信息 - 分割位号字符串，创建位号到物料的映射
            ref_to_pn_a = {}
            pn_to_refs_a = {}
            mpn_map_a = {}  # 存储物料号到MPN的映射
            desc_map_a = {}  # 存储物料号到描述的映射

            for _, row in bom_a_df.iterrows():
                pn = str(row['P/N']).strip()
                refs = str(row['Reference']).strip()
                mpn = str(row.get('MPN', '')).strip()
                desc = str(row.get('Description', '')).strip()

                # 存储MPN和描述信息
                mpn_map_a[pn] = mpn
                desc_map_a[pn] = desc

                # 分割位号（C1,C2,C3 或 C1 C2 C3）
                if ',' in refs:
                    ref_list = [r.strip() for r in refs.split(',')]
                else:
                    ref_list = [r.strip() for r in refs.split()]

                # 构建映射
                for ref in ref_list:
                    if ref and ref.lower() != 'nan':
                        ref_to_pn_a[ref] = pn

                # 构建物料到位号的映射
                if pn not in pn_to_refs_a:
                    pn_to_refs_a[pn] = []
                pn_to_refs_a[pn].extend([r for r in ref_list if r and r.lower() != 'nan'])

            # 同样处理BOM B
            ref_to_pn_b = {}
            pn_to_refs_b = {}
            mpn_map_b = {}
            desc_map_b = {}

            for _, row in bom_b_df.iterrows():
                pn = str(row['P/N']).strip()
                refs = str(row['Reference']).strip()
                mpn = str(row.get('MPN', '')).strip()
                desc = str(row.get('Description', '')).strip()

                mpn_map_b[pn] = mpn
                desc_map_b[pn] = desc

                if ',' in refs:
                    ref_list = [r.strip() for r in refs.split(',')]
                else:
                    ref_list = [r.strip() for r in refs.split()]

                for ref in ref_list:
                    if ref and ref.lower() != 'nan':
                        ref_to_pn_b[ref] = pn

                if pn not in pn_to_refs_b:
                    pn_to_refs_b[pn] = []
                pn_to_refs_b[pn].extend([r for r in ref_list if r and r.lower() != 'nan'])

            print(f"BOM A 位号数: {len(ref_to_pn_a)}, 物料数: {len(pn_to_refs_a)}")
            print(f"BOM B 位号数: {len(ref_to_pn_b)}, 物料数: {len(pn_to_refs_b)}")

            # 分析结果
            self.update_progress(60, "分析差异...")

            # 1. 物料变更分析
            pn_added = []    # 在B中新增的物料
            pn_removed = []  # 从A中移除的物料
            pn_common = []   # A和B中共有的物料

            for pn in pn_to_refs_a:
                if pn not in pn_to_refs_b:
                    pn_removed.append(pn)
                else:
                    pn_common.append(pn)

            for pn in pn_to_refs_b:
                if pn not in pn_to_refs_a:
                    pn_added.append(pn)

            # 2. 位号变更分析
            ref_added = []   # 在B中新增的位号
            ref_removed = [] # 从A中移除的位号
            ref_changed = [] # 物料变更的位号

            for ref in ref_to_pn_a:
                if ref not in ref_to_pn_b:
                    ref_removed.append(ref)
                elif ref_to_pn_a[ref] != ref_to_pn_b[ref]:
                    # 检查是否是替代料
                    pn_a = ref_to_pn_a[ref]
                    pn_b = ref_to_pn_b[ref]

                    is_alternative = False
                    # 检查替代料关系
                    for main_pn, alt_pns in self.alternative_map.items():
                        if (pn_a == main_pn and pn_b in alt_pns) or (pn_b == main_pn and pn_a in alt_pns):
                            is_alternative = True
                            break
                        # 检查是否都是同一主料的替代料
                        if pn_a in alt_pns and pn_b in alt_pns:
                            is_alternative = True
                            break

                    ref_changed.append((ref, pn_a, pn_b, is_alternative))

            for ref in ref_to_pn_b:
                if ref not in ref_to_pn_a:
                    ref_added.append(ref)

            print(f"新增位号: {len(ref_added)}个, 示例: {ref_added[:5] if ref_added else '无'}")
            print(f"移除位号: {len(ref_removed)}个, 示例: {ref_removed[:5] if ref_removed else '无'}")
            print(f"变更位号: {len(ref_changed)}个, 示例: {ref_changed[:5] if ref_changed else '无'}")

            # 3. 共有物料位号数量变化分析
            pn_quantity_changes = []

            # 物料数量变更包含三种情况：
            # 1. 常规数量变更：物料在A和B中都存在，但数量不同
            # 2. 物料完全移除：物料在A中存在，但在B中不存在（数量从N变为0）
            # 3. 物料完全新增：物料在A中不存在，但在B中存在（数量从0变为N）

            # 1. 处理常规数量变更
            for pn in pn_common:
                count_a = len(pn_to_refs_a[pn])
                count_b = len(pn_to_refs_b[pn])

                if count_a != count_b:
                    pn_quantity_changes.append((pn, count_a, count_b))

            # 2. 处理物料完全移除的情况
            for pn in pn_removed:
                count_a = len(pn_to_refs_a[pn])
                count_b = 0  # 在B中不存在，数量为0
                pn_quantity_changes.append((pn, count_a, count_b))

            # 3. 处理物料完全新增的情况
            for pn in pn_added:
                count_a = 0  # 在A中不存在，数量为0
                count_b = len(pn_to_refs_b[pn])
                pn_quantity_changes.append((pn, count_a, count_b))

            # 生成报告
            self.update_progress(80, "生成报告...")

            result = []

            # 报告标题
            result.append("=== BOM对比报告 ===")
            result.append("")

            # 基本信息
            result.append("1. 基本信息")
            result.append(f"基准BOM(A)物料数: {len(pn_to_refs_a)}")
            result.append(f"基准BOM(A)位号数: {len(ref_to_pn_a)}")
            result.append(f"对比BOM(B)物料数: {len(pn_to_refs_b)}")
            result.append(f"对比BOM(B)位号数: {len(ref_to_pn_b)}")
            result.append("")

            # 物料变更汇总
            result.append("2. 物料变更汇总")
            result.append(f"新增物料: {len(pn_added)}个")
            result.append(f"移除物料: {len(pn_removed)}个")
            result.append(f"变更物料: {len(ref_changed)}个")
            result.append("")

            # 位号变更汇总
            result.append("3. 位号变更汇总")
            result.append(f"新增位号: {len(ref_added)}个")
            result.append(f"移除位号: {len(ref_removed)}个")
            result.append(f"变更位号: {len(ref_changed)}个")
            result.append("")

            # 查找同一位号物料替换的情况
            replaced_refs = set()

            # 查找同一位号上的替换情况：移除了一个物料并新增了另一个物料
            for ref, pn_a, pn_b, is_alt in ref_changed:
                replaced_refs.add(ref)

            # 位号变动详情
            result.append("4. 位号变动")

            # 位号变动计数器（统一所有类型的位号变动）
            position_change_counter = 1

            # 初始化变更类型字典，确保它在任何情况下都存在
            changes_by_type = {
                "[常规替换]": [],
                "[新物料引入]": [],
                "[物料整合]": []
            }

            # 1. 物料变更部分（原4.物料变更详情）
            if ref_changed:
                # 按原物料号分组
                changes_by_pn = {}
                for ref, pn_a, pn_b, is_alt in ref_changed:
                    key = (pn_a, pn_b, is_alt)
                    if key not in changes_by_pn:
                        changes_by_pn[key] = []
                    changes_by_pn[key].append(ref)

                # 分类变更类型
                changes_by_type = {
                    "[常规替换]": [],
                    "[新物料引入]": [],
                    "[物料整合]": []
                }

                # 输出分组信息
                for (pn_a, pn_b, is_alt), refs in sorted(changes_by_pn.items()):
                    mpn_a_info = f" (MPN: {mpn_map_a.get(pn_a, '')})" if self.show_mpn_in_report else ""
                    mpn_b_info = f" (MPN: {mpn_map_b.get(pn_b, '')})" if self.show_mpn_in_report else ""

                    alt_info = " [替代料]" if is_alt else ""

                    # 判断变更类型
                    change_type = ""

                    # 检查物料A是否在B中完全被移除（没有出现在任何其他位号）
                    pn_a_completely_removed = pn_a not in pn_to_refs_b

                    # 检查物料B是否是完全新增的（在A中不存在）
                    pn_b_completely_new = pn_b not in pn_to_refs_a

                    if not pn_a_completely_removed and not pn_b_completely_new:
                        # 情况1: 两个物料在BOM中都保留 - 只是位号上的调整
                        change_type = "[常规替换]"
                    elif pn_a_completely_removed and not pn_b_completely_new:
                        # 物料A被完全移除，物料B已存在于A中
                        change_type = "[物料整合]"
                    else:
                        # 情况2和3: 物料B是新增的(无论物料A是否完全移除)
                        change_type = "[新物料引入]"

                    # 存储变更信息，包含该变更涉及的位号、物料和变更类型
                    changes_by_type[change_type].append((refs, pn_a, pn_b, mpn_a_info, mpn_b_info, alt_info))

                # 按变更类型显示
                first_type = True
                for change_type, changes in changes_by_type.items():
                    if changes:
                        # 在不同类型之间添加空行（第一个类型前不添加）
                        if not first_type:
                            result.append("")
                        first_type = False

                        result.append(f"    位号变更{change_type}:")
                        for refs, pn_a, pn_b, mpn_a_info, mpn_b_info, alt_info in changes:
                            for ref in sorted(refs):
                                # 检查是否有替代料
                                alt_a_info = ""
                                alt_b_info = ""

                                # 找到A料号的替代料
                                if pn_a in self.alternative_map:
                                    alt_pns_a = self.alternative_map[pn_a]
                                    if alt_pns_a:
                                        alt_a_info = f" [替代料: {', '.join(alt_pns_a)}]"

                                # 找到B料号的替代料
                                if pn_b in self.alternative_map:
                                    alt_pns_b = self.alternative_map[pn_b]
                                    if alt_pns_b:
                                        alt_b_info = f" [替代料: {', '.join(alt_pns_b)}]"

                                result.append(f"    {position_change_counter}.{ref} : {pn_a}{mpn_a_info}{alt_a_info} → {pn_b}{mpn_b_info}{alt_b_info}")
                                position_change_counter += 1

            # 新增位号部分前添加空行（仅当有位号变更且有新增位号时）
            has_changes = any(changes for changes in changes_by_type.values())
            if has_changes and ref_added:
                result.append("")

            # 2. 新增位号部分（原6.位号变动的新增位号部分）
            if ref_added:
                # 首先按照物料的不同状态分组
                new_refs_with_new_material = []  # 新增位号对应新增物料
                new_refs_with_existing_material = []  # 新增位号对应原有物料

                for ref in sorted(ref_added):
                    pn = ref_to_pn_b.get(ref, "未知")
                    if pn not in pn_to_refs_a:
                        new_refs_with_new_material.append((ref, pn))
                    else:
                        new_refs_with_existing_material.append((ref, pn))

                # 先显示新增位号对应新增物料的情况
                if new_refs_with_new_material:
                    result.append("    新增位号[对应新增物料]:")
                    for ref, pn in sorted(new_refs_with_new_material):
                        mpn_info = f" (MPN: {mpn_map_b.get(pn, '')})" if self.show_mpn_in_report else ""

                        # 检查是否有替代料
                        alt_info = ""
                        alt_pns = []
                        for main_pn, alt_pn_list in self.alternative_map.items():
                            if pn == main_pn:
                                # 当前料号是主料号，查找其所有替代料
                                alt_pns = alt_pn_list
                                break
                            elif pn in alt_pn_list:
                                # 当前料号是替代料，找到主料号和其他替代料
                                alt_pns = [main_pn] + [p for p in alt_pn_list if p != pn]
                                break

                        # 如果有替代料，添加替代料信息
                        if alt_pns:
                            alt_info = " [替代料: " + ", ".join(alt_pns) + "]"

                        result.append(f"    {position_change_counter}.{ref} : {pn}{mpn_info}{alt_info}")
                        position_change_counter += 1

                # 增加一个空行，使子分类之间有间隔
                if new_refs_with_new_material and new_refs_with_existing_material:
                    result.append("")

                # 再显示新增位号对应原有物料的情况
                if new_refs_with_existing_material:
                    result.append("    新增位号[对应原有物料]:")
                    for ref, pn in sorted(new_refs_with_existing_material):
                        mpn_info = f" (MPN: {mpn_map_b.get(pn, '')})" if self.show_mpn_in_report else ""

                        # 检查是否有替代料
                        alt_info = ""
                        alt_pns = []
                        for main_pn, alt_pn_list in self.alternative_map.items():
                            if pn == main_pn:
                                # 当前料号是主料号，查找其所有替代料
                                alt_pns = alt_pn_list
                                break
                            elif pn in alt_pn_list:
                                # 当前料号是替代料，找到主料号和其他替代料
                                alt_pns = [main_pn] + [p for p in alt_pn_list if p != pn]
                                break

                        # 如果有替代料，添加替代料信息
                        if alt_pns:
                            alt_info = " [替代料: " + ", ".join(alt_pns) + "]"

                        result.append(f"    {position_change_counter}.{ref} : {pn}{mpn_info}{alt_info}")
                        position_change_counter += 1

            # 增加一个空行，使子分类之间有间隔
            if ref_added and ref_removed:
                result.append("")

            # 3. 移除位号部分（原6.位号变动的移除位号部分）
            if ref_removed:
                # 首先按照物料的不同状态分组
                removed_refs_with_removed_material = []  # 移除位号对应物料移除
                removed_refs_with_remaining_material = []  # 移除位号对应物料仍保留

                for ref in sorted(ref_removed):
                    pn = ref_to_pn_a.get(ref, "未知")
                    if pn not in pn_to_refs_b:
                        removed_refs_with_removed_material.append((ref, pn))
                    else:
                        removed_refs_with_remaining_material.append((ref, pn))

                # 先显示移除位号对应物料移除的情况
                if removed_refs_with_removed_material:
                    result.append("    移除位号[对应物料移除]:")
                    for ref, pn in sorted(removed_refs_with_removed_material):
                        mpn_info = f" (MPN: {mpn_map_a.get(pn, '')})" if self.show_mpn_in_report else ""

                        # 检查是否有替代料
                        alt_info = ""
                        alt_pns = []
                        for main_pn, alt_pn_list in self.alternative_map.items():
                            if pn == main_pn:
                                # 当前料号是主料号，查找其所有替代料
                                alt_pns = alt_pn_list
                                break
                            elif pn in alt_pn_list:
                                # 当前料号是替代料，找到主料号和其他替代料
                                alt_pns = [main_pn] + [p for p in alt_pn_list if p != pn]
                                break

                        # 如果有替代料，添加替代料信息
                        if alt_pns:
                            alt_info = " [替代料: " + ", ".join(alt_pns) + "]"

                        result.append(f"    {position_change_counter}.{ref} : {pn}{mpn_info}{alt_info}")
                        position_change_counter += 1

                # 增加一个空行，使子分类之间有间隔
                if removed_refs_with_removed_material and removed_refs_with_remaining_material:
                    result.append("")

                # 再显示移除位号对应物料仍保留的情况
                if removed_refs_with_remaining_material:
                    result.append("    移除位号[对应物料仍保留]:")
                    for ref, pn in sorted(removed_refs_with_remaining_material):
                        mpn_info = f" (MPN: {mpn_map_a.get(pn, '')})" if self.show_mpn_in_report else ""

                        # 检查是否有替代料
                        alt_info = ""
                        alt_pns = []
                        for main_pn, alt_pn_list in self.alternative_map.items():
                            if pn == main_pn:
                                # 当前料号是主料号，查找其所有替代料
                                alt_pns = alt_pn_list
                                break
                            elif pn in alt_pn_list:
                                # 当前料号是替代料，找到主料号和其他替代料
                                alt_pns = [main_pn] + [p for p in alt_pn_list if p != pn]
                                break

                        # 如果有替代料，添加替代料信息
                        if alt_pns:
                            alt_info = " [替代料: " + ", ".join(alt_pns) + "]"

                        result.append(f"    {position_change_counter}.{ref} : {pn}{mpn_info}{alt_info}")
                        position_change_counter += 1

            result.append("")

            # 过滤掉已经在变更中报告过的位号
            filtered_pn_added = []
            for pn in pn_added:
                # 检查该物料的所有位号是否都已经在变更中报告过
                refs = set(pn_to_refs_b[pn])
                if not all(ref in replaced_refs for ref in refs):
                    filtered_pn_added.append(pn)

            filtered_pn_removed = []
            for pn in pn_removed:
                # 检查该物料的所有位号是否都已经在变更中报告过
                refs = set(pn_to_refs_a[pn])
                if not all(ref in replaced_refs for ref in refs):
                    filtered_pn_removed.append(pn)

            # 详细变更信息
            # 物料变动（新增物料和移除物料）
            if filtered_pn_added or filtered_pn_removed:
                result.append("5. 物料变动")

                # 物料变动计数器
                material_change_counter = 1

                # 1. 新增物料部分
                if filtered_pn_added:
                    result.append("    新增物料:")
                    for pn in sorted(filtered_pn_added):
                        # 过滤掉已经报告的位号
                        all_refs = set(pn_to_refs_b[pn])
                        unreported_refs = [r for r in all_refs if r not in replaced_refs]

                        if not unreported_refs:
                            continue

                        mpn_info = f" (MPN: {mpn_map_b.get(pn, '')})" if self.show_mpn_in_report else ""

                        for ref in sorted(unreported_refs):
                            result.append(f"    {material_change_counter}.{pn}{mpn_info} : {ref}")
                            material_change_counter += 1

                # 增加一个空行，使分类之间有间隔
                if filtered_pn_added and filtered_pn_removed:
                    result.append("")

                # 2. 移除物料部分
                if filtered_pn_removed:
                    result.append("    移除物料:")
                    for pn in sorted(filtered_pn_removed):
                        # 过滤掉已经报告的位号
                        all_refs = set(pn_to_refs_a[pn])
                        unreported_refs = [r for r in all_refs if r not in replaced_refs]

                        if not unreported_refs:
                            continue

                        mpn_info = f" (MPN: {mpn_map_a.get(pn, '')})" if self.show_mpn_in_report else ""

                        for ref in sorted(unreported_refs):
                            result.append(f"    {material_change_counter}.{pn}{mpn_info} : {ref}")
                            material_change_counter += 1

                result.append("")

            # 数量变更
            if pn_quantity_changes:
                result.append("6. 物料数量变更")

                # 数量变更计数器
                quantity_change_counter = 1

                # 自定义排序函数，按变更类型分组排序
                # 排序优先级：
                # 1. 完全移除的物料（count_b=0）
                # 2. 完全新增的物料（count_a=0）
                # 3. 数量增加的物料（count_b>count_a）
                # 4. 数量减少的物料（count_b<count_a）
                # 每组内按物料号字典序排序
                def sort_key(item):
                    pn, count_a, count_b = item
                    if count_b == 0:  # 完全移除
                        return 0, pn  # 第一组：完全移除的物料
                    elif count_a == 0:  # 完全新增
                        return 1, pn  # 第二组：完全新增的物料
                    elif count_b > count_a:  # 数量增加
                        return 2, pn  # 第三组：数量增加的物料
                    else:  # 数量减少
                        return 3, pn  # 第四组：数量减少的物料

                # 按变更类型分组排序
                sorted_changes = sorted(pn_quantity_changes, key=sort_key)
                total_changes = len(sorted_changes)

                # 记录前一个项目的类型，用于在类型变化时添加空行和类型标识
                prev_type = None

                # 遍历排序后的物料数量变更
                for i, (pn, count_a, count_b) in enumerate(sorted_changes):
                    # 确定当前项目的类型和对应的标识文本
                    if count_b == 0:
                        current_type = "removed"
                        type_label = "    【物料完全移除】"
                    elif count_a == 0:
                        current_type = "added"
                        type_label = "    【物料完全新增】"
                    elif count_b > count_a:
                        current_type = "increased"
                        type_label = "    【物料数量增加】"
                    else:
                        current_type = "decreased"
                        type_label = "    【物料数量减少】"

                    # 处理类型变化和类型标识的显示
                    if prev_type is None:
                        # 第一个类型，添加类型标识
                        result.append(type_label)
                    elif prev_type != current_type:
                        # 类型发生变化，添加空行和新类型标识
                        result.append("")
                        result.append(type_label)

                    # 更新前一个类型
                    prev_type = current_type

                    # 生成差异描述文本
                    change = count_b - count_a
                    if count_b == 0:  # 物料完全移除的情况
                        direction = "移除"
                        difference_text = "完全移除"
                    elif count_a == 0:  # 物料完全新增的情况
                        direction = "新增"
                        difference_text = "完全新增"
                    else:
                        direction = "增加" if change > 0 else "减少"
                        difference_text = f"{direction}{abs(change)}"

                    # 根据情况选择正确的MPN信息源
                    # 对于新增物料，使用B中的MPN信息；对于其他情况，使用A中的MPN信息
                    if count_a == 0:  # 完全新增的物料
                        mpn_info = f" (MPN: {mpn_map_b.get(pn, '')})" if self.show_mpn_in_report else ""
                    else:
                        mpn_info = f" (MPN: {mpn_map_a.get(pn, '')})" if self.show_mpn_in_report else ""

                    # 添加主数量变更信息行
                    result.append(f"    {quantity_change_counter}.{pn}{mpn_info} : {count_a} → {count_b} (差异: {difference_text})")

                    # 获取位号差异：添加的位号和移除的位号
                    refs_a_set = set(pn_to_refs_a.get(pn, []))
                    refs_b_set = set(pn_to_refs_b.get(pn, []))

                    added_refs = refs_b_set - refs_a_set
                    removed_refs = refs_a_set - refs_b_set

                    # 显示移除的位号（所有类型都可能有）
                    if removed_refs:
                        for ref in sorted(removed_refs):
                            result.append(f"\t移除位号: {ref} → {pn}{mpn_info}")

                    # 显示新增的位号（只有当物料没有完全移除时才显示）
                    if added_refs and count_b > 0:
                        for ref in sorted(added_refs):
                            result.append(f"\t新增位号: {ref} → {pn}{mpn_info}")

                    # 添加项目间的空行分隔（只在同一类型内的项目之间添加）
                    if i < total_changes - 1:  # 不是最后一项
                        next_pn, next_count_a, next_count_b = sorted_changes[i + 1]

                        # 确定下一项的类型
                        if next_count_b == 0:
                            next_type = "removed"
                        elif next_count_a == 0:
                            next_type = "added"
                        elif next_count_b > next_count_a:
                            next_type = "increased"
                        else:
                            next_type = "decreased"

                        # 仅当下一项与当前项类型相同时，添加空行
                        if current_type == next_type:
                            result.append("")

                    # 递增计数器
                    quantity_change_counter += 1

            # 记录结束时间
            self.end_time = datetime.now()
            processing_time = self.end_time - self.start_time

            # 添加处理时间信息到报告开头
            time_info = [
                "=== 处理时间统计 ===",
                f"开始时间: {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}",
                f"结束时间: {self.end_time.strftime('%Y-%m-%d %H:%M:%S')}",
                f"总耗时: {processing_time.total_seconds():.2f}秒\n"
            ]

            result = time_info + result

            self.update_progress(100, "处理完成")
            return "\n".join(result)

        except Exception as e:
            trace = traceback.format_exc()
            return f"生成对比报告时出错:\n{str(e)}\n\n详细错误信息:\n{trace}"

    def create_bom_table(self, parent_frame, is_bom_a=True):
        """创建BOM数据显示表格"""
        # 创建表格框架
        table_frame = ttk.Frame(parent_frame)
        table_frame.pack(fill="both", expand=True, padx=3, pady=3)  # 从5减小到3

        # 创建表头 - 初始不指定列，将在加载数据时动态设置
        columns = []

        # 创建Treeview组件
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse", height=8)  # 设置为browse模式，确保单击选择功能正常

        # 确保树视图的样式设置正确，支持选择高亮
        style = ttk.Style()
        style.configure("Treeview",
                        background="#ffffff",  # 正常背景色
                        fieldbackground="#ffffff",  # 字段背景色
                        foreground="#000000")  # 前景色（文字颜色）

        # 配置选中项的样式
        style.map("Treeview",
                  background=[("selected", "#0078d7")],  # 选中项的背景色
                  foreground=[("selected", "#ffffff")])  # 选中项的文字颜色

        # 绑定单击事件
        tree.bind('<<TreeviewSelect>>', lambda event: self.comparer.on_tree_select(event, tree, is_bom_a))

        # 绑定双击事件 - 用于搜索结果
        tree.bind('<Double-1>', lambda event: self.comparer.on_tree_double_click(event, tree, is_bom_a))

        # 创建滚动条
        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # 放置表格和滚动条
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        # 保存引用
        if is_bom_a:
            self.bom_a_tree = tree
            self.bom_a_xscroll = scrollbar_x
            self.bom_a_yscroll = scrollbar_y
        else:
            self.bom_b_tree = tree
            self.bom_b_xscroll = scrollbar_x
            self.bom_b_yscroll = scrollbar_y

        return tree

    def on_tree_select(self, event, tree, is_bom_a):
        """处理树视图选择事件"""
        selected_items = tree.selection()
        if not selected_items:
            return

        # 获取选中的项
        item = selected_items[0]
        values = tree.item(item, 'values')

        # 获取P/N列索引
        pn_col_idx = 1  # 默认值
        column_headers = tree.cget('columns')
        for i, col in enumerate(column_headers):
            header_text = tree.heading(col, 'text')
            if header_text and ('型号' in header_text or 'P/N' in header_text):
                pn_col_idx = i
                break

        if len(values) > pn_col_idx:
            pn = values[pn_col_idx]
            if pn:
                # 在单击选择模式下，我们不触发自定义高亮，保留系统默认的蓝色高亮
                # 用户必须双击才会触发黄色高亮并清除选择
                pass
                # 原来的代码：self.highlight_material_in_both_trees(pn)

    def on_tree_double_click(self, event, tree, is_bom_a):
        """处理树视图双击事件"""
        item = tree.identify('item', event.x, event.y)
        if not item:
            return

        values = tree.item(item, 'values')

        # 获取P/N列索引
        pn_col_idx = 1  # 默认值
        column_headers = tree.cget('columns')
        for i, col in enumerate(column_headers):
            header_text = tree.heading(col, 'text')
            if header_text and ('型号' in header_text or 'P/N' in header_text):
                pn_col_idx = i
                break

        if len(values) > pn_col_idx:
            pn = values[pn_col_idx]
            if pn:
                # 先清除所有树视图的选择状态，避免同时存在蓝色选择和黄色高亮
                if hasattr(self, 'bom_a_tree'):
                    self.bom_a_tree.selection_remove(self.bom_a_tree.selection())
                if hasattr(self, 'bom_b_tree'):
                    self.bom_b_tree.selection_remove(self.bom_b_tree.selection())

                # 在结果文本区域中搜索该型号
                self.search_in_results(pn)

                # 使用自定义高亮（黄色背景）
                self.highlight_material_in_both_trees(pn)

    def search_in_results(self, search_text):
        """在结果文本区域中搜索并滚动到相关位置"""
        # 获取父窗口中的结果文本框
        if not self.parent_window or not hasattr(self.parent_window, 'result_text') or not search_text:
            return

        # 使用父窗口的结果文本框
        result_text = self.parent_window.result_text

        # 清除之前的所有标记
        result_text.tag_remove('found', '1.0', tk.END)

        # 重置搜索位置
        start_pos = '1.0'

        # 搜索文本
        count = 0
        while True:
            start_pos = result_text.search(search_text, start_pos, stopindex=tk.END, nocase=True)
            if not start_pos:
                break

            end_pos = f"{start_pos}+{len(search_text)}c"
            # 添加标记
            result_text.tag_add('found', start_pos, end_pos)
            # 更新下一次搜索的起始位置
            start_pos = end_pos
            count += 1

        # 配置标记样式 - 黄色背景突出显示
        result_text.tag_configure('found', background='yellow', foreground='black')

        # 滚动到第一个匹配处
        if count > 0:
            # 获取第一个带有'found'标记的位置
            first_match = result_text.tag_nextrange('found', '1.0')
            if first_match:
                result_text.see(first_match[0])
                # 将焦点放在文本区域
                result_text.focus_set()

    def highlight_material_in_both_trees(self, pn):
        """在两个BOM树中高亮显示指定物料编号的行"""
        print(f"在两个BOM树中高亮显示物料: {pn}")

        # 确保父窗口存在
        if not self.parent_window:
            return False

        # 让父窗口来执行高亮逻辑
        if hasattr(self.parent_window, 'highlight_material_in_both_trees'):
            return self.parent_window.highlight_material_in_both_trees(pn)

        return False

class BOMComparerGUI:
    def __init__(self, root):
        """初始化GUI"""
        self.root = root
        self.root.title("BOM比对工具 V1.4")

        # 设置窗口图标
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass  # 如果没有找到图标文件，则使用默认图标

        # 设置窗口大小 - 根据屏幕尺寸自适应
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 根据屏幕大小调整窗口尺寸
        # 对于小屏幕设备（如笔记本），使用较小的窗口尺寸
        if screen_width <= 1366 or screen_height <= 768:  # 常见的笔记本分辨率
            window_width = min(1000, int(screen_width * 0.85))  # 屏幕宽度的85%
            window_height = min(700, int(screen_height * 0.85))  # 屏幕高度的85%
            min_width = 800  # 较小的最小宽度
            min_height = 600  # 较小的最小高度
        else:  # 对于大屏幕设备，使用较大的窗口尺寸
            window_width = min(1200, int(screen_width * 0.75))  # 屏幕宽度的75%
            window_height = min(900, int(screen_height * 0.75))  # 屏幕高度的75%
            min_width = 1000  # 原来的最小宽度
            min_height = 750  # 原来的最小高度

        # 调整窗口大小，适应小屏幕
        if screen_height <= 900:  # 对于小屏幕，进一步减小窗口高度
            window_height = min(window_height, int(screen_height * 0.85))  # 最多使用屏幕高度的85%
            window_width = min(window_width, int(screen_width * 0.85))  # 最多使用屏幕宽度的85%

        # 计算窗口位置 - 使用固定偏移而不是百分比
        x = (screen_width - window_width) // 2  # 水平居中

        # 对于窗口高度，使用固定的上边距
        y = 20  # 固定距离屏幕顶部20像素

        # 确保坐标不为负
        x = max(0, x)
        y = max(0, y)

        # 设置窗口位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(min_width, min_height)  # 根据屏幕大小设置最小尺寸

        # 确保窗口居中并显示
        self.root.update_idletasks()  # 处理所有待处理的窗口事件，确保几何管理器已完成布局计算

        # 打印调试信息
        print(f"屏幕尺寸: {screen_width}x{screen_height}")
        print(f"窗口尺寸: {window_width}x{window_height}")
        print(f"窗口位置: +{x}+{y}")

        # 将窗口置于前台
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(lambda: self.root.attributes('-topmost', False))

        # 创建BOM比较器实例，传入root作为父窗口
        self.comparer = BOMComparer(parent_window=self.root)

        # 设置进度回调
        self.comparer.set_progress_callback(self.update_progress)

        # 进度变量
        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)

        # 文件选择的路径（两个文件选择共享一个路径）
        self.last_dir = os.path.expanduser("~")  # 默认为用户主目录

        # 清理配置文件中的无效字段
        self.clean_config_files()

        # 从配置文件加载设置
        self.load_config_from_file()

        # 检查更新相关方法
        self.check_updates_manually = self._check_updates_manually
        self.check_updates_on_startup = self._check_updates_on_startup

        # 设置界面样式
        self.setup_styles()

        # 设置UI组件
        self.setup_ui()

        # 窗口关闭前保存配置
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _check_updates_manually(self):
        """手动检查更新"""
        self.update_progress(0, "正在检查更新...")
        threading.Thread(target=self._check_updates, args=(True,), daemon=True).start()

    def _check_updates_on_startup(self):
        """启动时自动检查更新"""
        # 等待一些时间再检查，避免影响程序启动速度
        time.sleep(2)
        self._check_updates(False)

    def _check_updates(self, is_manual_check=False):
        """检查更新的实际实现"""
        try:
            # 检查更新
            has_update, latest_version, download_url, changelog, is_exe_update = check_for_updates(APP_VERSION)

            # 如果是手动检查，更新进度条
            if is_manual_check:
                self.update_progress(100, "检查更新完成")

            # 如果有更新，显示更新通知
            if has_update:
                # 在主线程中显示更新通知
                self.root.after(0, lambda: self._show_update_dialog(latest_version, download_url, changelog, is_exe_update))
            elif is_manual_check:
                # 如果是手动检查且没有更新，显示提示
                self.root.after(0, lambda: messagebox.showinfo("检查更新",
                                                            f"当前版本 {APP_VERSION} 已是最新版本。"))
        except Exception as e:
            print(f"检查更新时出错: {str(e)}")
            if is_manual_check:
                self.update_progress(100, "检查更新失败")
                self.root.after(0, lambda: messagebox.showerror("检查更新失败",
                                                             f"检查更新时出错: {str(e)}"))

    def _show_update_dialog(self, latest_version, download_url, changelog, is_exe_update):
        """显示更新对话框"""
        # 显示更新通知对话框
        if show_update_notification(self.root, APP_VERSION, latest_version, changelog, download_url, is_exe_update):
            # 用户选择更新，开始下载
            self._download_update(latest_version, download_url, is_exe_update)

    def _download_update(self, latest_version, download_url, is_exe_update):
        """下载更新"""
        try:
            # 获取屏幕尺寸
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()

            # 根据屏幕大小调整对话框尺寸
            if screen_width <= 1366 or screen_height <= 768:  # 小屏幕设备
                width = min(350, int(screen_width * 0.5))
                height = min(140, int(screen_height * 0.3))
            else:  # 大屏幕设备
                width = 400
                height = 150

            # 计算居中位置
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2

            # 创建下载进度对话框
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("下载更新")
            progress_dialog.transient(self.root)  # 设置为父窗口的临时窗口

            # 先设置位置和大小，再显示窗口
            progress_dialog.geometry(f"{width}x{height}+{x}+{y}")
            progress_dialog.resizable(False, False)
            progress_dialog.withdraw()  # 先隐藏窗口

            # 设置为模态对话框
            progress_dialog.grab_set()  # 模态对话框

            # 创建内容框架
            frame = ttk.Frame(progress_dialog, padding=10)
            frame.pack(fill="both", expand=True)

            # 标题
            title_label = ttk.Label(frame, text=f"正在下载版本 {latest_version}", font=("微软雅黑", 10, "bold"))
            title_label.pack(pady=(0, 10))

            # 进度条 - 根据对话框宽度调整长度
            progress_var = tk.DoubleVar()
            progress_bar_length = width - 50  # 根据对话框宽度调整进度条长度
            progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100, length=progress_bar_length)
            progress_bar.pack(pady=5)

            # 进度文本
            status_var = tk.StringVar()
            status_var.set("准备下载...")
            status_label = ttk.Label(frame, textvariable=status_var)
            status_label.pack(pady=5)

            # 进度百分比
            percent_var = tk.StringVar()
            percent_var.set("0%")
            percent_label = ttk.Label(frame, textvariable=percent_var)
            percent_label.pack()

            # 取消按钮
            cancel_button = ttk.Button(frame, text="取消", command=progress_dialog.destroy)
            cancel_button.pack(pady=10)

            # 准备下载路径
            if is_exe_update:
                # 获取当前程序路径
                current_exe = sys.executable
                # 如果是打包后的程序，使用当前程序路径
                if current_exe.endswith('.exe'):
                    # 生成新的文件名，保留原文件名的格式
                    new_exe = _get_updated_filename(os.path.basename(current_exe), latest_version)
                    download_path = os.path.join(os.path.dirname(current_exe), new_exe)
                else:
                    # 如果不是.exe，使用临时目录
                    download_path = os.path.join(tempfile.gettempdir(), f"BOM_Comparer_v{latest_version}.exe")
            else:
                # 如果是源代码包，使用临时目录
                download_path = os.path.join(tempfile.gettempdir(), f"BOM_Comparer_v{latest_version}.zip")

            # 进度回调函数
            def update_progress_callback(downloaded, total, progress):
                progress_var.set(progress * 100)
                percent_var.set(f"{progress * 100:.1f}%")
                status_var.set(f"已下载: {downloaded / 1024 / 1024:.2f} MB / {total / 1024 / 1024:.2f} MB")

            # 状态回调函数
            def status_callback(message):
                status_var.set(message)

            # 在新线程中下载
            def download_thread():
                try:
                    # 下载文件
                    success = download_with_resume(download_url, download_path,
                                                 update_progress_callback, status_callback)

                    # 如果下载成功
                    if success:
                        # 关闭进度对话框
                        progress_dialog.destroy()

                        # 显示下载完成对话框
                        if is_exe_update:
                            # 如果是可执行文件，询问用户是否关闭当前程序并运行新版本
                            if messagebox.askyesno("更新完成",
                                                f"新版本 {latest_version} 已下载完成。\n\n"
                                                f"是否关闭当前程序并运行新版本？"):
                                # 启动新版本并关闭当前程序
                                subprocess.Popen([download_path])
                                self.root.quit()
                                self.root.destroy()
                                sys.exit(0)
                        else:
                            # 如果是源代码包，提示用户下载完成
                            messagebox.showinfo("下载完成",
                                             f"新版本 {latest_version} 已下载到:\n{download_path}")
                    else:
                        # 如果下载失败，显示错误消息
                        messagebox.showerror("下载失败",
                                         f"下载新版本 {latest_version} 失败。\n"
                                         f"请稍后重试或访问官方网站手动下载。")
                except Exception as e:
                    # 如果发生异常，显示错误消息
                    messagebox.showerror("下载错误",
                                     f"下载过程中发生错误: {str(e)}")

            # 所有内容创建完成后再显示对话框
            progress_dialog.update_idletasks()  # 确保所有内容已经布局完成
            progress_dialog.deiconify()  # 显示对话框

            # 启动下载线程
            threading.Thread(target=download_thread, daemon=True).start()

        except Exception as e:
            messagebox.showerror("更新错误", f"准备更新时发生错误: {str(e)}")

    def setup_styles(self):
        """设置自定义样式，模拟macOS/Figma风格"""
        # 定义字体
        self.default_font = Font(family="Arial", size=10)
        self.title_font = Font(family="Arial", size=12, weight="bold")
        self.heading_font = Font(family="Arial", size=14, weight="bold")

        # 获取适合的等宽字体
        try:
            available_fonts = list(tk_font.families())
            code_font_name = "Consolas"
            if code_font_name not in available_fonts:
                for font in ["Courier New", "Courier", "Monospace", "DejaVu Sans Mono"]:
                    if font in available_fonts:
                        code_font_name = font
                        break
                else:
                    code_font_name = "TkFixedFont"  # tkinter的默认等宽字体
        except:
            code_font_name = "Consolas"

        self.code_font = Font(family=code_font_name, size=10)

        style = ttk.Style()

        # 配置基础样式
        style.configure("TFrame", background="#f5f5f7")
        style.configure("TLabel", background="#f5f5f7", font=self.default_font)
        style.configure("TButton", font=self.default_font, padding=6)

        # 标题标签样式
        style.configure("Title.TLabel", font=self.heading_font, foreground="#1d1d1f", background="#f5f5f7")

        # 现代按钮样式
        style.configure("Modern.TButton",
                        font=self.default_font,
                        background="#0071e3",
                        foreground="white",
                        padding=(12, 6),
                        relief="flat")
        style.map("Modern.TButton",
                  background=[('active', '#0077ed'), ('pressed', '#005bbd')],
                  relief=[('pressed', 'flat'), ('!pressed', 'flat')])

        # 次要按钮样式
        style.configure("Secondary.TButton",
                        font=self.default_font,
                        background="#e6e6e6",
                        foreground="#1d1d1f",
                        padding=(12, 6),
                        relief="flat")
        style.map("Secondary.TButton",
                  background=[('active', '#d9d9d9'), ('pressed', '#cccccc')],
                  relief=[('pressed', 'flat'), ('!pressed', 'flat')])

        # 进度条样式
        style.configure("Modern.Horizontal.TProgressbar",
                        background="#0071e3",
                        troughcolor="#e6e6e6",
                        borderwidth=0,
                        thickness=6)

        # 自定义LabelFrame样式
        style.configure("Card.TLabelframe",
                        background="white",
                        relief="flat",
                        borderwidth=0,
                        padding=15)
        style.configure("Card.TLabelframe.Label",
                        font=self.title_font,
                        background="white",
                        foreground="#1d1d1f")

    def create_rounded_frame(self, parent, **kwargs):
        """创建圆角边框的框架"""
        frame = ttk.Frame(parent, style="Card.TFrame", **kwargs)
        return frame

    def setup_ui(self):
        """设置UI组件"""
        main_container = ttk.Frame(self.root)
        main_container.pack(fill="both", expand=True, padx=15, pady=15)  # 减小内边距
        main_container.configure(style="TFrame")

        # 应用标题
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 8))  # 减小标题下方的间距

        title_label = ttk.Label(header_frame, text="BOM 对比工具", font=self.heading_font, foreground="#1d1d1f")
        title_label.pack(side="left")
        version_label = ttk.Label(header_frame, text="V1.4 | 小航 2025.4", foreground="#888888")
        version_label.pack(side="right", padx=(0, 3))  # 右侧添加小间距

        # 添加设置按钮
        self.settings_button = tk.Button(header_frame, text="设置", command=self.show_settings,
                                      bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                      relief="flat", padx=6, pady=1,  # 减小按钮内边距
                                      activebackground="#d9d9d9", activeforeground="#1d1d1f")
        self.settings_button.pack(side="right", padx=3)  # 减小按钮间距

        # 添加检查更新按钮
        self.update_button = tk.Button(header_frame, text="检查更新", command=self.check_updates_manually,
                                    bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                    relief="flat", padx=6, pady=1,
                                    activebackground="#d9d9d9", activeforeground="#1d1d1f")
        self.update_button.pack(side="right", padx=3)

        # 添加帮助按钮
        self.help_button = tk.Button(header_frame, text="帮助", command=self.show_help,
                                   bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                   relief="flat", padx=6, pady=1,  # 使用与设置按钮一致的内边距
                                   activebackground="#d9d9d9", activeforeground="#1d1d1f")
        self.help_button.pack(side="right", padx=3)  # 使用与设置按钮一致的外边距

        # 添加关于按钮
        self.about_button = tk.Button(header_frame, text="关于", command=self.show_about,
                                    bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                    relief="flat", padx=6, pady=1,  # 使用与设置按钮一致的内边距
                                    activebackground="#d9d9d9", activeforeground="#1d1d1f")
        self.about_button.pack(side="right", padx=3)  # 使用与设置按钮一致的外边距

        # 文件选择区域的卡片
        file_select_card = ttk.LabelFrame(main_container, text="文件选择")
        file_select_card.pack(fill="x", pady=2)  # 减小外边距
        file_select_card.configure(padding=3)  # 设置更小的内边距

        # 创建网格布局容器
        file_grid = ttk.Frame(file_select_card)
        file_grid.pack(fill="x", padx=5, pady=3)  # 减小内边距

        # 修改为水平布局：基准BOM(A)和对比BOM(B)并排显示
        # 基准BOM文件选择 - 左半部分
        ttk.Label(file_grid, text="基准BOM(A):", width=10, anchor="e").grid(row=0, column=0, padx=(0, 2), pady=2, sticky="e")
        self.file_a_entry = ttk.Entry(file_grid)
        self.file_a_entry.grid(row=0, column=1, padx=1, pady=2, sticky="ew")

        browse_a_button = ttk.Button(file_grid, text="浏览", command=self.select_file_a, width=5)
        browse_a_button.grid(row=0, column=2, padx=(1, 10), pady=2)

        # 对比BOM文件选择 - 右半部分
        ttk.Label(file_grid, text="对比BOM(B):", width=10, anchor="e").grid(row=0, column=3, padx=(10, 2), pady=2, sticky="e")
        self.file_b_entry = ttk.Entry(file_grid)
        self.file_b_entry.grid(row=0, column=4, padx=1, pady=2, sticky="ew")

        browse_b_button = ttk.Button(file_grid, text="浏览", command=self.select_file_b, width=5)
        browse_b_button.grid(row=0, column=5, padx=(1, 0), pady=2)

        # 使列可伸缩 - 两个输入框都可以伸缩
        file_grid.columnconfigure(1, weight=1)
        file_grid.columnconfigure(4, weight=1)

        # 添加BOM数据显示区域 - 增加高度
        data_card = ttk.LabelFrame(main_container, text="BOM数据")
        data_card.pack(fill="both", expand=True, pady=1)  # 从2减小到1
        data_card.configure(padding=1)  # 从2减小到1

        # 创建一个垂直分隔区域，上部显示BOM数据，下部显示对比结果
        main_paned = ttk.PanedWindow(data_card, orient=tk.VERTICAL)
        main_paned.pack(fill="both", expand=True, padx=2, pady=2)  # 从3减小到2

        # 上部区域 - BOM数据
        top_frame = ttk.Frame(main_paned)

        # 创建垂直分栏
        self.paned_window = ttk.PanedWindow(top_frame, orient=tk.VERTICAL)
        self.paned_window.pack(fill="both", expand=True)

        # BOM A 数据显示区域
        self.bom_a_frame = ttk.LabelFrame(self.paned_window, text="基准BOM(A)数据")
        self.paned_window.add(self.bom_a_frame, weight=1)

        # BOM B 数据显示区域
        self.bom_b_frame = ttk.LabelFrame(self.paned_window, text="对比BOM(B)数据")
        self.paned_window.add(self.bom_b_frame, weight=1)

        # 创建BOM A的表格
        self.create_bom_table(self.bom_a_frame, is_bom_a=True)

        # 创建BOM B的表格
        self.create_bom_table(self.bom_b_frame, is_bom_a=False)

        # 添加详细比较区域和结果区域到主分隔窗口
        main_paned.add(top_frame, weight=2)  # 上部占比更少，从3减小到2

        # 下部区域 - 详细比较和结果
        bottom_frame = ttk.Frame(main_paned)

        # 设置表格同步滚动
        self.setup_synchronized_scrolling()

        # 操作按钮框架
        button_card = ttk.Frame(bottom_frame)
        button_card.pack(fill="x", pady=2)  # 减小外部间距

        # 操作按钮区
        actions_frame = ttk.Frame(button_card)
        actions_frame.pack(fill="x", pady=0)  # 减小内部间距

        # 左侧主要操作按钮
        left_buttons = ttk.Frame(actions_frame)
        left_buttons.pack(side="left")

        # 使用标准样式但自定义外观
        self.compare_button = tk.Button(left_buttons, text="开始对比", command=self.start_compare,
                                      bg="#0071e3", fg="white", font=self.default_font,
                                      relief="flat", padx=8, pady=3,  # 减小按钮内部间距
                                      activebackground="#0077ed", activeforeground="white")
        self.compare_button.pack(side="left", padx=(0, 6))  # 减小按钮间距

        self.save_button = tk.Button(left_buttons, text="保存结果", command=self.save_result,
                                    bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                    relief="flat", padx=8, pady=3,  # 减小按钮内部间距
                                    activebackground="#d9d9d9", activeforeground="#1d1d1f")
        self.save_button.pack(side="left")

        # 右侧辅助按钮
        right_buttons = ttk.Frame(actions_frame)
        right_buttons.pack(side="right")

        # 结果显示区域
        result_card = ttk.LabelFrame(bottom_frame, text="对比结果")
        result_card.pack(fill="both", expand=True, pady=1)  # 从2减小到1
        result_card.configure(padding=3)  # 从5减小到3

        # 调整文本区域的内部填充
        result_inner = ttk.Frame(result_card)
        result_inner.pack(fill="both", expand=True, padx=2, pady=1)  # 减小内边距

        # 使用ScrolledText显示结果，设置现代感的外观
        self.result_text = scrolledtext.ScrolledText(
            result_inner,
            wrap=tk.WORD,
            height=35,  # 从30增加到35
            font=self.code_font,
            background="white",
            foreground="#333333",
            borderwidth=0,
            padx=8,  # 减小水平内边距
            pady=8   # 减小垂直内边距
        )
        self.result_text.pack(fill="both", expand=True)

        # 结果文本框样式设置
        self.result_text.tag_configure("clickable", foreground="#0066cc", underline=True)
        self.result_text.bind("<Double-Button-1>", self.on_result_double_click)

        # 添加到主分隔窗口
        main_paned.add(bottom_frame, weight=8)  # 下部占比更多，从7增加到8

        # 添加欢迎文本
        welcome_text = """欢迎使用BOM对比工具！

使用步骤:
1. 使用"浏览"按钮选择基准BOM和对比BOM文件
2. 点击"开始对比"按钮开始分析
3. 分析结果会显示在此区域
4. 可通过"保存结果"按钮将结果导出为文本文件

祝您使用愉快！
"""
        self.result_text.insert(tk.END, welcome_text)

        # 添加文本标签的悬停效果
        for button in [self.help_button, self.about_button, self.compare_button, self.save_button, self.settings_button, self.update_button]:
            button.bind("<Enter>", self.on_button_hover)
            button.bind("<Leave>", self.on_button_leave)

        # 进度显示卡片
        progress_card = ttk.LabelFrame(main_container, text="进度")
        progress_card.pack(fill="x", pady=3)  # 减小外部间距
        progress_card.configure(padding=5)  # 减小内边距

        # 进度条内部框架
        progress_inner = ttk.Frame(progress_card)
        progress_inner.pack(fill="x", padx=5, pady=3)  # 减小内边距

        # 进度条 - 使用标准进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_inner, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=2)  # 减小内边距

        # 进度文本和百分比显示
        progress_text_frame = ttk.Frame(progress_inner)
        progress_text_frame.pack(fill="x", padx=2, pady=(2, 0))  # 减小内边距

        self.progress_text = ttk.Label(progress_text_frame, text="就绪", foreground="#666666")
        self.progress_text.pack(side="left")

        self.progress_percent = ttk.Label(progress_text_frame, text="0%", foreground="#666666")
        self.progress_percent.pack(side="right")

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, foreground="#434343", anchor="w")
        status_bar.pack(fill="x", side="bottom", padx=20, pady=5)

    def on_button_hover(self, event):
        """鼠标悬停在按钮上的效果"""
        event.widget.configure(cursor="hand2")
        # 为tk.Button添加高亮效果
        if isinstance(event.widget, tk.Button):
            if event.widget.cget("bg") == "#0071e3":  # 主按钮
                event.widget.configure(bg="#0077ed")
            else:  # 次要按钮
                event.widget.configure(bg="#d9d9d9")

    def on_button_leave(self, event):
        """鼠标离开按钮的效果"""
        event.widget.configure(cursor="")
        # 恢复tk.Button的原始颜色
        if isinstance(event.widget, tk.Button):
            if event.widget.cget("fg") == "white":  # 主按钮
                event.widget.configure(bg="#0071e3")
            else:  # 次要按钮
                event.widget.configure(bg="#e6e6e6")

    def select_file_a(self):
        """选择基准BOM文件"""
        file_path = filedialog.askopenfilename(
            title="选择基准BOM文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")],
            initialdir=self.last_dir
        )
        if file_path:
            self.file_a_entry.delete(0, tk.END)
            self.file_a_entry.insert(0, file_path)
            self.last_dir = os.path.dirname(file_path)

            try:
                self.update_progress(0, "加载BOM A文件...")
                # 加载BOM数据并显示在表格中
                bom_data = self.comparer.load_bom(file_path)

                # 清空已有数据和列
                for item in self.bom_a_tree.get_children():
                    self.bom_a_tree.delete(item)

                # 重新配置列 - 动态创建以匹配原始数据
                columns = list(bom_data.columns)

                # 更新表格列配置
                self.bom_a_tree.configure(columns=columns)

                # 清除所有表头
                for col in self.bom_a_tree["columns"]:
                    self.bom_a_tree.heading(col, text="")

                # 设置新的表头
                for col in columns:
                    self.bom_a_tree.heading(col, text=col)
                    # 预设列宽 - 根据列名长度设置初始宽度，为Item和Quantity列设置较小的宽度
                    if col == 'Item' or col == 'Quantity':
                        width = 80  # 为Item和Quantity列设置固定宽度
                    else:
                        width = max(100, tk_font.Font().measure(str(col)) + 20)
                    self.bom_a_tree.column(col, width=width, anchor="center", stretch=True, minwidth=80)

                # 添加数据到表格 - 使用循环分批添加，避免大文件导致的卡顿
                total_rows = len(bom_data)
                batch_size = 200
                batches = (total_rows + batch_size - 1) // batch_size  # 计算需要多少批

                for batch in range(batches):
                    start_idx = batch * batch_size
                    end_idx = min((batch + 1) * batch_size, total_rows)
                    batch_data = bom_data.iloc[start_idx:end_idx]

                    for i, row in batch_data.iterrows():
                        values = [str(row.get(col, "")) for col in columns]
                        # 交替行颜色样式
                        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                        self.bom_a_tree.insert("", "end", values=values, tags=(tag,))

                    # 更新进度
                    progress = int((batch + 1) / batches * 100)
                    self.update_progress(progress, f"加载BOM A文件... ({end_idx}/{total_rows})")

                    # 允许GUI刷新
                    self.root.update_idletasks()

                # 优化列宽
                self.comparer.optimize_column_widths(self.bom_a_tree, bom_data, columns)

                # 保存BOM数据
                self.comparer.bom_a = bom_data

                # 如果两个BOM都已加载，同步列宽
                if hasattr(self.comparer, 'bom_b') and self.comparer.bom_b is not None:
                    self.sync_column_widths()

                self.update_progress(100, "BOM A文件加载完成")

            except Exception as e:
                error_message = str(e)
                # 用户取消选择不显示为错误
                if "用户取消了" in error_message:
                    self.update_progress(0, "用户取消了操作")
                    return

                self.show_error(f"加载BOM A文件失败: {error_message}")
                traceback.print_exc()

                # 重置进度条
                self.update_progress(0)

    def select_file_b(self):
        """选择对比BOM文件"""
        file_path = filedialog.askopenfilename(
            title="选择对比BOM文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")],
            initialdir=self.last_dir
        )
        if file_path:
            self.file_b_entry.delete(0, tk.END)
            self.file_b_entry.insert(0, file_path)
            self.last_dir = os.path.dirname(file_path)

            try:
                self.update_progress(0, "加载BOM B文件...")
                # 加载BOM数据并显示在表格中
                bom_data = self.comparer.load_bom(file_path)

                # 清空已有数据和列
                for item in self.bom_b_tree.get_children():
                    self.bom_b_tree.delete(item)

                # 重新配置列 - 动态创建以匹配原始数据
                columns = list(bom_data.columns)

                # 更新表格列配置
                self.bom_b_tree.configure(columns=columns)

                # 清除所有表头
                for col in self.bom_b_tree["columns"]:
                    self.bom_b_tree.heading(col, text="")

                # 设置新的表头
                for col in columns:
                    self.bom_b_tree.heading(col, text=col)
                    # 预设列宽 - 根据列名长度设置初始宽度，为Item和Quantity列设置较小的宽度
                    if col == 'Item' or col == 'Quantity':
                        width = 80  # 为Item和Quantity列设置固定宽度
                    else:
                        width = max(100, tk_font.Font().measure(str(col)) + 20)
                    self.bom_b_tree.column(col, width=width, anchor="center", stretch=True, minwidth=80)

                # 添加数据到表格 - 使用循环分批添加，避免大文件导致的卡顿
                total_rows = len(bom_data)
                batch_size = 200
                batches = (total_rows + batch_size - 1) // batch_size  # 计算需要多少批

                for batch in range(batches):
                    start_idx = batch * batch_size
                    end_idx = min((batch + 1) * batch_size, total_rows)
                    batch_data = bom_data.iloc[start_idx:end_idx]

                    for i, row in batch_data.iterrows():
                        values = [str(row.get(col, "")) for col in columns]
                        # 交替行颜色样式
                        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                        self.bom_b_tree.insert("", "end", values=values, tags=(tag,))

                    # 更新进度
                    progress = int((batch + 1) / batches * 100)
                    self.update_progress(progress, f"加载BOM B文件... ({end_idx}/{total_rows})")

                    # 允许GUI刷新
                    self.root.update_idletasks()

                # 优化列宽
                self.comparer.optimize_column_widths(self.bom_b_tree, bom_data, columns)

                # 保存BOM数据
                self.comparer.bom_b = bom_data

                # 如果两个BOM都已加载，同步列宽
                if hasattr(self.comparer, 'bom_a') and self.comparer.bom_a is not None:
                    self.sync_column_widths()

                self.update_progress(100, "BOM B文件加载完成")

            except Exception as e:
                error_message = str(e)
                # 用户取消选择不显示为错误
                if "用户取消了" in error_message:
                    self.update_progress(0, "用户取消了操作")
                    return

                self.show_error(f"加载BOM B文件失败: {error_message}")
                traceback.print_exc()

                # 重置进度条
                self.update_progress(0)

    def sync_column_widths(self):
        """同步两个BOM表格的列宽，以便更好地对比"""
        # 确保两个表都已加载
        if not hasattr(self, 'bom_a_tree') or not hasattr(self, 'bom_b_tree'):
            return

        # 获取两个表的列
        a_columns = self.bom_a_tree["columns"]
        b_columns = self.bom_b_tree["columns"]

        # 没有列则返回
        if not a_columns or not b_columns:
            return

        # 为特定列设置最大宽度限制
        column_max_widths = {
            'Item': 80,       # Item列宽度限制为80像素
            'Quantity': 80    # Quantity列宽度限制为80像素
        }

        # 处理两个表格共有的列
        common_columns = set()
        for col in a_columns:
            if col in b_columns:
                common_columns.add(col)

        # 遍历共有列，设置相同的宽度
        for col in common_columns:
            # 如果是Item或Quantity列，使用固定宽度
            if col in column_max_widths:
                fixed_width = column_max_widths[col]
                self.bom_a_tree.column(col, width=fixed_width)
                self.bom_b_tree.column(col, width=fixed_width)
            else:
                # 其他列使用最大宽度
                a_width = self.bom_a_tree.column(col, "width")
                b_width = self.bom_b_tree.column(col, "width")
                max_width = max(a_width, b_width)
                self.bom_a_tree.column(col, width=max_width)
                self.bom_b_tree.column(col, width=max_width)

    def update_progress(self, progress, message=""):
        """更新进度显示"""
        self.progress_var.set(progress)
        self.progress_percent.config(text=f"{int(progress)}%")

        if message:
            self.progress_text.config(text=message)
            self.status_var.set(message)
        self.root.update_idletasks()

    def start_compare(self):
        """开始比较两个BOM文件"""
        # 获取文件路径
        file_a = self.file_a_entry.get().strip()
        file_b = self.file_b_entry.get().strip()

        if not file_a or not file_b:
            self.show_error("请先选择两个BOM文件")
            return

        if not os.path.exists(file_a) or not os.path.exists(file_b):
            self.show_error("文件不存在")
            return

        # 禁用按钮，防止重复点击
        self.compare_button["state"] = "disabled"

        # 清空之前的结果
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state="disabled")

        # 重置进度条
        self.progress_var.set(0)

        # 显示进度条
        self.progress_bar.pack(fill="x", pady=5)

        # 显示"比较中"的状态
        self.status_var.set("比较中...")
        self.progress_text.config(text="比较中...")

        # 在单独的线程中执行比较操作，以避免阻塞GUI
        self.compare_thread = threading.Thread(target=self.run_compare)
        self.compare_thread.daemon = True  # 设置为守护线程，这样主程序退出时线程也会结束
        self.compare_thread.start()

    def run_compare(self):
        """在单独的线程中执行比较操作"""
        try:
            # 获取文件路径
            file_a = self.file_a_entry.get().strip()
            file_b = self.file_b_entry.get().strip()

            # 输出调试信息
            print(f"开始比较文件: \nA: {file_a}\nB: {file_b}")

            # 检查是否已经有加载好的BOM数据
            if hasattr(self.comparer, 'bom_a') and hasattr(self.comparer, 'bom_b') and \
               self.comparer.bom_a is not None and self.comparer.bom_b is not None:
                # 使用已加载的数据进行比较
                print("使用已加载的数据进行比较...")
                print(f"BOM A 数据行数: {len(self.comparer.bom_a)}")
                print(f"BOM B 数据行数: {len(self.comparer.bom_b)}")
                self.update_progress(10, "使用已加载的数据进行比较...")
                result = self.comparer.compare(self.comparer.bom_a, self.comparer.bom_b, is_dataframe=True)
            else:
                # 从文件加载数据进行比较
                print("从文件加载数据进行比较...")
                self.update_progress(10, "从文件加载数据...")
                result = self.comparer.compare(file_a, file_b)

            # 检查结果是否为空
            if not result or len(result.strip()) == 0:
                error_message = "比较结果为空，请检查BOM文件是否有内容"
                print(error_message)
                self.root.after(0, lambda: self.show_error(error_message))
                return

            print(f"比较完成，结果长度: {len(result)}")
            # 输出结果的前100个字符，帮助调试
            print(f"结果预览: {result[:100]}...")

            # 在主线程中显示结果
            self.root.after(0, lambda: self.show_result(result))

        except Exception as e:
            error_message = f"比较过程中出错: {str(e)}"
            print(f"错误详情: {error_message}")
            traceback.print_exc()
            # 在主线程中显示错误
            self.root.after(0, lambda: self.show_error(error_message))
        finally:
            # 恢复按钮状态
            self.root.after(0, lambda: self.compare_button.config(state="normal"))

    def show_result(self, result):
        """显示比较结果"""
        # 确保文本控件处于可编辑状态
        self.result_text.config(state="normal")

        # 清除旧内容并插入新结果
        self.result_text.delete(1.0, tk.END)

        # 打印调试信息
        print(f"显示结果，长度: {len(result)}")
        print(f"结果包含新增位号信息: {'新增位号' in result}")
        print(f"结果包含移除位号信息: {'移除位号' in result}")

        if not result:
            self.result_text.insert(tk.END, "未生成有效的比较结果，请检查输入文件")
            self.status_var.set("比较失败")
            self.progress_text.config(text="处理失败")
            self.result_text.config(state="disabled")
            return

        self.result_text.insert(tk.END, result)

        # 更新状态显示
        self.status_var.set("对比完成")
        self.progress_text.config(text="处理完成")

        # 为不同部分设置不同的文本颜色
        self.highlight_text()

        # 滚动到顶部
        self.result_text.see("1.0")

        # 保持文本可见但禁止编辑
        self.result_text.config(state="disabled")

        # 重新启用按钮
        self.compare_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.NORMAL)

        # 在可点击文本的末尾添加换行，避免后续文本还是可点击的
        self.result_text.insert(tk.END, "\n", "normal")

        # 同步两个表格的列宽以便更好地比较
        self.sync_column_widths()

    def highlight_text(self):
        """为报告中的不同部分应用不同颜色"""
        content = self.result_text.get(1.0, tk.END)

        # 重置文本颜色
        self.result_text.tag_configure("header", foreground="#0066cc", font=self.title_font)
        self.result_text.tag_configure("section", foreground="#333333", font=self.title_font)
        # 新增二级标题样式，使用粗体并稍微增大字号
        self.result_text.tag_configure("subsection", foreground="#555555", font=("Arial", 11, "bold"))
        self.result_text.tag_configure("added", foreground="#008800")
        self.result_text.tag_configure("removed", foreground="#cc0000")
        self.result_text.tag_configure("changed", foreground="#cc6600")
        self.result_text.tag_configure("warning", foreground="#cc6600")
        self.result_text.tag_configure("time", foreground="#666666")
        self.result_text.tag_configure("reference", foreground="#0066cc")

        # 查找并应用样式
        lines = content.split('\n')
        line_index = 1

        for line in lines:
            line_end = f"{line_index}.end"

            # 应用标题样式
            if line.startswith("==="):
                self.result_text.tag_add("header", f"{line_index}.0", line_end)
            # 应用章节样式
            elif re.match(r"^\d+\.\s+", line):
                self.result_text.tag_add("section", f"{line_index}.0", line_end)
            # 应用二级标题样式 - 添加对二级标题的识别，例如"位号变更[常规替换]:"这样的模式
            elif line.strip().startswith("    ") and ":" in line and not line.strip().startswith("    位号:"):
                if any(keyword in line for keyword in ["位号变更", "新增位号", "移除位号", "新增物料", "移除物料", "数量变更"]):
                    self.result_text.tag_add("subsection", f"{line_index}.0", line_end)
            # 应用位号信息样式
            elif "    位号:" in line:
                self.result_text.tag_add("reference", f"{line_index}.0", line_end)
            # 应用新增位号样式
            elif "新增位号:" in line or ("新增" in line and not "物料" in line):
                self.result_text.tag_add("added", f"{line_index}.0", line_end)
            # 应用移除位号样式
            elif "移除位号:" in line or ("移除" in line and not "物料" in line):
                self.result_text.tag_add("removed", f"{line_index}.0", line_end)
            # 应用新增物料样式
            elif "新增物料" in line or "新增:" in line:
                self.result_text.tag_add("added", f"{line_index}.0", line_end)
            # 应用移除物料样式
            elif "移除物料" in line or "移除:" in line:
                self.result_text.tag_add("removed", f"{line_index}.0", line_end)
            # 应用变更项样式
            elif "变更" in line or "→" in line:
                self.result_text.tag_add("changed", f"{line_index}.0", line_end)
            # 应用警告样式
            elif "警告" in line:
                self.result_text.tag_add("warning", f"{line_index}.0", line_end)
            # 应用时间信息样式
            elif "时间" in line:
                self.result_text.tag_add("time", f"{line_index}.0", line_end)

            line_index += 1

    def show_error(self, error_message):
        """显示错误信息"""
        self.status_var.set("对比失败")
        self.progress_text.config(text="处理失败")

        # 使用现代风格的错误对话框
        error_window = tk.Toplevel(self.root)
        error_window.title("错误")

        # 根据屏幕大小调整对话框尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 对于小屏幕设备，使用较小的对话框尺寸
        if screen_width <= 1366 or screen_height <= 768:
            dialog_width = min(350, int(screen_width * 0.5))
            dialog_height = min(180, int(screen_height * 0.3))
            min_width = 300
            min_height = 150
        else:
            dialog_width = 400
            dialog_height = 200
            min_width = 400
            min_height = 200

        error_window.geometry(f"{dialog_width}x{dialog_height}")
        error_window.minsize(min_width, min_height)
        error_window.configure(bg="white")

        # 设置模态
        error_window.transient(self.root)
        error_window.grab_set()

        # 错误图标和标题
        header_frame = ttk.Frame(error_window)
        header_frame.pack(fill="x", padx=20, pady=15)

        ttk.Label(header_frame, text="处理出错", font=self.heading_font).pack(side="left")

        # 错误详情
        detail_frame = ttk.Frame(error_window)
        detail_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # 只显示友好的错误信息，不显示技术细节
        error_text = scrolledtext.ScrolledText(detail_frame, wrap="word", height=4)
        error_text.pack(fill="both", expand=True)

        # 提取主要错误信息并进一步处理
        main_error = error_message.split('\n')[0] if '\n' in error_message else error_message

        # 特殊处理某些错误消息
        if "Missing optional dependency" in main_error and "xlrd" in main_error:
            main_error = "缺少读取Excel所需的库，请安装xlrd库:\npip install xlrd>=2.0.1"

        error_text.insert("1.0", main_error)
        error_text.configure(state="disabled")

        # 关闭按钮
        button_frame = ttk.Frame(error_window)
        button_frame.pack(fill="x", padx=20, pady=15)

        close_button = tk.Button(button_frame, text="关闭",
                                command=error_window.destroy,
                                bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                relief="flat", padx=12, pady=6,
                                activebackground="#d9d9d9", activeforeground="#1d1d1f")
        close_button.pack(side="right")

        # 确保窗口获取焦点
        error_window.focus_set()

        # 重新启用按钮
        self.compare_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.NORMAL)

    def _generate_default_filename(self):
        """生成默认的保存文件名"""
        # 获取当前日期和时间
        now = datetime.now()
        date_str = now.strftime("%Y%m%d_%H%M%S")

        # 生成简化的默认文件名：BOM对比结果_日期时间.txt
        default_filename = f"BOM对比结果_{date_str}"

        # 移除文件名中的非法字符
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            default_filename = default_filename.replace(char, '_')

        return default_filename

    def save_result(self):
        """保存对比结果"""
        content = self.result_text.get(1.0, tk.END)
        if not content.strip():
            messagebox.showinfo("提示", "没有可保存的结果")
            return

        # 生成默认文件名
        default_filename = self._generate_default_filename()

        # 使用last_dir作为默认保存路径，这个变量在选择B BOM文件时已更新为B BOM文件所在目录
        file_path = filedialog.asksaveasfilename(
            initialdir=self.last_dir,
            initialfile=default_filename,
            title="保存对比结果",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )

        if file_path:
            try:
                # 检查选择的文件类型
                if file_path.lower().endswith('.xlsx'):
                    self.save_as_excel(file_path, content)
                else:
                    # 原有的保存为文本文件的代码
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    self.status_var.set(f"结果已保存至: {file_path}")
                    # 现代成功消息
                    messagebox.showinfo("成功", "对比结果已成功保存到文件")
            except Exception as e:
                messagebox.showerror("错误", f"保存文件时出错: {str(e)}")

    def save_as_excel(self, file_path, content):
        """将对比结果保存为Excel文件

        Args:
            file_path (str): 保存文件的路径
            content (str): 对比结果文本内容
        """
        try:
            # 需要导入pandas库
            import pandas as pd
            # 引用全局os和subprocess模块
            global os, subprocess

            # 创建Excel写入器
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 将文本内容按行分割
                lines = content.strip().split('\n')

                # 创建数据列表，用于存储行数据
                data = []

                # 当前处理的部分标题
                current_section = ""

                # 处理每一行
                for line in lines:
                    # 跳过空行
                    if not line.strip():
                        # 添加空行
                        data.append({"内容": ""})
                        continue

                    # 检查是否为标题行（通常是以===或者大标题样式显示的）
                    if line.startswith('==') or line.startswith('--') or line.endswith('==') or line.endswith('--'):
                        continue

                    # 检查是否为分节标题
                    if line.startswith('【') and line.endswith('】'):
                        # 添加一个空行，使标题与上下内容分开
                        if data and data[-1]["内容"] != "":
                            data.append({"内容": ""})

                        current_section = line
                        data.append({"内容": line})
                    # 正常内容行
                    else:
                        data.append({"内容": line})

                # 创建DataFrame并写入Excel
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name="BOM对比结果", index=False)

                # 获取工作表并调整列宽
                workbook = writer.book
                worksheet = writer.sheets['BOM对比结果']

                # 设置列宽 - 只有一列，所以设置A列宽度
                worksheet.column_dimensions['A'].width = 100

            self.status_var.set(f"结果已保存至Excel文件: {file_path}")
            messagebox.showinfo("成功", "对比结果已成功保存为Excel文件")

            # 询问是否打开文件
            if messagebox.askyesno("保存成功", f"结果已保存为Excel文件: {os.path.basename(file_path)}\n\n是否打开文件？"):
                # 尝试打开文件
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # macOS 和 Linux
                    if os.path.exists('/usr/bin/open'):  # macOS
                        subprocess.call(['open', file_path])
                    else:  # Linux
                        subprocess.call(['xdg-open', file_path])

        except ImportError:
            messagebox.showerror("错误", "保存为Excel需要安装pandas库。请使用pip install pandas进行安装。")
        except Exception as e:
            messagebox.showerror("错误", f"保存Excel文件时出错: {str(e)}")

    def show_about(self):
        """显示关于信息"""
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")

        # 根据屏幕大小调整对话框尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 对于小屏幕设备，使用较小的对话框尺寸
        if screen_width <= 1366 or screen_height <= 768:
            dialog_width = min(350, int(screen_width * 0.5))
            dialog_height = min(350, int(screen_height * 0.5))
        else:
            dialog_width = 400
            dialog_height = 400

        about_window.resizable(False, False)
        about_window.configure(bg="white")

        # 设置模态
        about_window.transient(self.root)
        about_window.grab_set()

        # 窗口居中显示
        self.center_window(about_window, dialog_width, dialog_height)

        # 应用标志和名称
        header_frame = ttk.Frame(about_window)
        header_frame.pack(fill="x", padx=30, pady=20)

        title_label = ttk.Label(header_frame, text="BOM对比工具", font=("Arial", 18, "bold"), foreground="#0071e3", background="white")
        title_label.pack()

        version_label = ttk.Label(header_frame, text="版本: 1.4", font=("Arial", 11), background="white")
        version_label.pack(pady=5)

        # 分隔线
        separator = ttk.Separator(about_window, orient="horizontal")
        separator.pack(fill="x", padx=20, pady=10)

        # 信息区域
        info_frame = ttk.Frame(about_window, style="TFrame")
        info_frame.pack(fill="both", expand=True, padx=30, pady=10)

        # 关于文本
        about_text = """BOM对比工具是一个用于比较两个BOM文件差异的专业工具。

功能特点:
• 智能识别和匹配物料
• 详细的差异分析报告
• 位号变更分类显示
• 物料数量变更跟踪
• 优化的报告格式
• 多线程处理，响应流畅

作者: 小航
发布日期: 2025年4月

技术支持联系方式:
微信: XiaoHang_Sky
"""

        text_widget = tk.Text(info_frame, wrap="word", highlightthickness=0,
                             relief="flat", bg="white", height=16, width=40)
        text_widget.pack(fill="both", expand=True)
        text_widget.insert("1.0", about_text)
        text_widget.configure(state="disabled")

    def show_help(self):
        """显示帮助信息"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用帮助")

        # 根据屏幕大小调整对话框尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 对于小屏幕设备，使用较小的对话框尺寸
        if screen_width <= 1366 or screen_height <= 768:
            dialog_width = min(500, int(screen_width * 0.7))
            dialog_height = min(450, int(screen_height * 0.7))
            min_width = 350
            min_height = 250
        else:
            dialog_width = 550
            dialog_height = 500
            min_width = 400
            min_height = 300

        help_window.minsize(min_width, min_height)
        help_window.configure(bg="white")

        # 设置模态
        help_window.transient(self.root)
        help_window.grab_set()

        # 窗口居中显示
        self.center_window(help_window, dialog_width, dialog_height)

        # 标题
        header_frame = ttk.Frame(help_window, style="TFrame")
        header_frame.pack(fill="x", padx=30, pady=20)

        title_label = ttk.Label(header_frame, text="使用指南", font=("Arial", 16, "bold"), foreground="#333333", background="white")
        title_label.pack(anchor="w")

        # 内容区域
        content_frame = ttk.Frame(help_window, style="TFrame")
        content_frame.pack(fill="both", expand=True, padx=30, pady=10)

        help_text = tk.Text(content_frame, wrap="word", highlightthickness=0,
                           relief="flat", bg="white", padx=0, pady=0)
        help_text.pack(fill="both", expand=True)

        # 帮助内容
        text = """BOM对比工具使用指南

基本步骤:
1. 选择基准BOM(A)和对比BOM(B)文件
   - 使用"浏览"按钮选择Excel格式的BOM文件
   - 支持的格式: .xlsx, .xls

2. 点击"开始对比"按钮
   - 系统将自动分析两个文件的差异
   - 进度条显示处理进度

3. 查看对比结果
   - 结果将显示在主窗口中
   - 报告包含以下主要部分:
     • 基本信息 - 文件统计数据和处理时间
     • 物料变更汇总 - 新增、移除和数量变更的物料
     • 位号变动详情 - 新增、移除和变更的位号
     • 位号变更 - 按不同类型分类显示([常规替换]、[新物料替换]、[物料整合])
     • 数量变更 - 物料数量变化及关联位号

4. 保存结果(可选)
   - 点击"保存结果"将分析报告保存为文本文件

报告格式说明:
• 使用"→"箭头标识变更前后的状态
• 位号变更按三种类型分类显示
• 各分类之间添加空行提高可读性
• 物料数量变更显示变更前后的数量和关联位号

设置选项:
• 在报告中显示物料料号(MPN) - 控制是否在报告中显示制造商料号
• 替代料合并显示 - 同一位号的多个替代料在一行中显示

注意事项:
• 程序支持自动识别BOM文件的表头，但需保证基本字段齐全
• 对于大型BOM文件，处理可能需要一些时间

技术支持联系方式:
微信: XiaoHang_Sky
"""
        help_text.insert("1.0", text)

        # 配置标题样式
        help_text.tag_configure("title", font=("Arial", 12, "bold"), foreground="#0071e3")
        help_text.tag_configure("subtitle", font=("Arial", 11, "bold"), foreground="#333333")
        help_text.tag_configure("step", font=("Arial", 10, "bold"))

        # 应用样式
        help_text.tag_add("title", "1.0", "1.end")

        for line_num, line in enumerate(text.split('\n')):
            line_idx = line_num + 1  # 1-indexed
            if line.endswith(':'):
                help_text.tag_add("subtitle", f"{line_idx}.0", f"{line_idx}.end")
            elif re.match(r"^\d+\.\s+", line):
                help_text.tag_add("step", f"{line_idx}.0", f"{line_idx}.end")

        help_text.configure(state="disabled")  # 设为只读

    def center_window(self, window, width, height):
        """将窗口居中显示在主窗口上"""
        # 获取主窗口的位置和尺寸
        master_x = self.root.winfo_x()
        master_y = self.root.winfo_y()
        master_width = self.root.winfo_width()
        master_height = self.root.winfo_height()

        # 计算弹出窗口应该在的位置
        x = master_x + (master_width - width) // 2
        y = master_y + (master_height - height) // 2

        # 确保窗口位置不为负值
        x = max(0, x)
        y = max(0, y)

        # 设置窗口位置和大小
        window.geometry(f"{width}x{height}+{x}+{y}")

    def show_tooltip(self, event, text):
        """显示工具提示"""
        x, y, _, _ = event.widget.bbox("insert")
        x += event.widget.winfo_rootx() + 25
        y += event.widget.winfo_rooty() + 25

        # 创建带圆角的工具提示
        self.tooltip = tk.Toplevel(event.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(self.tooltip, text=text, background="#f7f7f7", foreground="#333333",
                        relief="solid", borderwidth=1, padx=10, pady=5,
                        font=("Arial", 9))
        label.pack()

    def hide_tooltip(self, event):
        """隐藏工具提示"""
        if hasattr(self, 'tooltip'):
            self.tooltip.destroy()

    def show_settings(self):
        """显示设置对话框"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("BOM对比工具 - 设置")

        # 根据屏幕大小调整设置窗口尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 对于小屏幕设备，使用较小的窗口尺寸
        if screen_width <= 1366 or screen_height <= 768:
            dialog_width = min(500, int(screen_width * 0.7))
            dialog_height = min(450, int(screen_height * 0.7))
        else:
            dialog_width = 550
            dialog_height = 500

        settings_window.resizable(True, True)
        settings_window.transient(self.root)  # 设置为主窗口的子窗口
        settings_window.grab_set()  # 模态窗口

        # 居中显示
        self.center_window(settings_window, dialog_width, dialog_height)

        # 标题
        title_frame = ttk.Frame(settings_window)
        title_frame.pack(fill="x", padx=20, pady=10)

        ttk.Label(title_frame, text="BOM对比工具设置", font=self.title_font).pack(side="left")

        # 创建选项卡控件
        tab_control = ttk.Notebook(settings_window)
        tab_control.pack(fill="both", expand=True, padx=20, pady=10)

        # 表头映射选项卡
        headers_tab = ttk.Frame(tab_control)
        tab_control.add(headers_tab, text="表头映射")

        # 报告设置选项卡
        report_tab = ttk.Frame(tab_control)
        tab_control.add(report_tab, text="报告设置")

        # 表头映射内容
        mapping_frame = ttk.Frame(headers_tab)
        mapping_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 字段列表容器
        fields_frame = ttk.Frame(mapping_frame)
        fields_frame.pack(fill="both", expand=True)

        # 存储每个字段的Entry控件
        self.field_entries = {}

        # 添加每个标准字段和对应的映射输入框
        row = 0
        for field, aliases in self.comparer.field_mappings.items():
            field_frame = ttk.Frame(fields_frame)
            field_frame.grid(row=row, column=0, sticky="ew", pady=5)

            # 字段标签
            ttk.Label(field_frame, text=f"{field}:", width=10).grid(row=0, column=0, sticky="w")

            # 字段别名输入框
            entry = ttk.Entry(field_frame, width=50)
            entry.grid(row=0, column=1, sticky="ew", padx=5)
            # 将别名列表转换为逗号分隔的字符串并设置为初始值
            entry.insert(0, ', '.join(aliases))

            # 将Entry引用存储在字典中
            self.field_entries[field] = entry

            # 添加说明标签
            if field == 'P/N':
                ttk.Label(field_frame, text="料号", foreground="#777").grid(row=0, column=2, sticky="w")
            elif field == 'Reference':
                ttk.Label(field_frame, text="位号", foreground="#777").grid(row=0, column=2, sticky="w")
            elif field == 'Description':
                ttk.Label(field_frame, text="描述", foreground="#777").grid(row=0, column=2, sticky="w")
            elif field == 'MPN':
                ttk.Label(field_frame, text="厂家料号", foreground="#777").grid(row=0, column=2, sticky="w")

            row += 1

        # 添加说明
        help_text = "说明：每个字段可以有多个别名，以逗号分隔。系统将自动匹配表头中的这些关键字。"
        ttk.Label(fields_frame, text=help_text, wraplength=500, foreground="#555", justify="left").grid(row=row, column=0, columnspan=3, sticky="w", pady=10)

        # 报告设置内容
        report_frame = ttk.Frame(report_tab)
        report_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 显示MPN选项
        self.show_mpn_var = tk.BooleanVar(value=self.comparer.show_mpn_in_report)
        mpn_check = ttk.Checkbutton(report_frame, text="在报告中显示物料料号(MPN)信息", variable=self.show_mpn_var)
        mpn_check.pack(anchor=tk.W, pady=5)

        # 添加MPN显示说明
        mpn_help_text = "说明：取消勾选此项将在报告中隐藏物料料号(MPN)信息，使报告更简洁。"
        ttk.Label(report_frame, text=mpn_help_text, wraplength=500, foreground="#555", justify="left").pack(anchor=tk.W, pady=5)

        # 底部按钮区域
        button_frame = ttk.Frame(settings_window)
        button_frame.pack(fill="x", padx=20, pady=15)

        # 保存按钮
        save_button = tk.Button(button_frame, text="保存设置", command=lambda: self.save_settings(settings_window),
                             bg="#0071e3", fg="white", font=self.default_font,
                             relief="flat", padx=10, pady=4,
                             activebackground="#0077ed", activeforeground="white")
        save_button.pack(side="right", padx=5)

        # 取消按钮
        cancel_button = tk.Button(button_frame, text="取消", command=settings_window.destroy,
                               bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                               relief="flat", padx=10, pady=4,
                               activebackground="#d9d9d9", activeforeground="#1d1d1f")
        cancel_button.pack(side="right", padx=5)

        # 恢复默认按钮
        default_button = tk.Button(button_frame, text="恢复默认", command=self.reset_default_mappings,
                                bg="#e6e6e6", fg="#1d1d1f", font=self.default_font,
                                relief="flat", padx=10, pady=4,
                                activebackground="#d9d9d9", activeforeground="#1d1d1f")
        default_button.pack(side="left", padx=5)

    def save_settings(self, window):
        """保存设置并关闭窗口"""
        # 从输入框中获取新的字段映射
        new_field_mappings = {}
        for field, entry in self.field_entries.items():
            # 获取输入框中的文本，将其拆分为别名列表
            aliases_text = entry.get().strip()
            if aliases_text:
                # 根据逗号拆分，并移除每个别名的前后空格
                aliases = [alias.strip() for alias in aliases_text.split(',')]
                # 过滤掉空字符串
                aliases = [alias for alias in aliases if alias]
                new_field_mappings[field] = aliases
            else:
                # 如果输入为空，使用默认值
                default_mappings = {
                    'Item': ['Item', 'item', '序号', 'Number'],
                    'P/N': ['P/N', '料号', '物料编码', '物料编号', 'Part Number', '型号'],
                    'Reference': ['Reference', 'Ref', 'ref', '位号'],
                    'Description': ['Description', '描述', '物料描述'],
                    'MPN': ['Manufacturer P/N', 'MPN', '制造商料号', '厂家料号', '生产商料号']
                }
                new_field_mappings[field] = default_mappings.get(field, [])

        # 更新比较器的字段映射
        self.comparer.set_field_mappings(new_field_mappings)

        # 保存MPN显示设置
        self.comparer.show_mpn_in_report = self.show_mpn_var.get()

        # 保存到配置文件
        self.save_config_to_file()

        # 直接关闭窗口，不显示确认对话框
        window.destroy()

    def reset_default_mappings(self):
        """重置为默认的字段映射"""
        default_mappings = {
            'Item': ['Item', 'item', '序号', 'Number'],
            'P/N': ['P/N', '料号', '物料编码', '物料编号', 'Part Number', '型号'],
            'Reference': ['Reference', 'Ref', 'ref', '位号'],
            'Description': ['Description', '描述', '物料描述'],
            'MPN': ['Manufacturer P/N', 'MPN', '制造商料号', '厂家料号', '生产商料号']
        }

        # 重置输入框的内容
        for field, entry in self.field_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, ', '.join(default_mappings.get(field, [])))

        # 显示消息
        messagebox.showinfo("默认设置已恢复", "表头关键字已重置为默认值，请点击保存以应用更改")

    def save_config_to_file(self):
        """保存配置到文件"""
        try:
            # 获取程序路径（支持打包为exe的情况）
            if getattr(sys, 'frozen', False):
                script_dir = os.path.dirname(sys.executable)
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))

            # 配置文件路径（保存在程序目录）
            config_file = os.path.join(script_dir, "config.json")

            # 更新配置数据
            config_data = {
                "field_mappings": self.comparer.field_mappings,
                "show_mpn_in_report": self.comparer.show_mpn_in_report,
                "last_dir": self.last_dir
            }

            # 保存到文件
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)

            print(f"配置已保存到 {config_file}")
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False

    def load_config_from_file(self):
        """从文件加载配置"""
        try:
            # 获取程序路径（支持打包为exe的情况）
            if getattr(sys, 'frozen', False):
                script_dir = os.path.dirname(sys.executable)
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                if not script_dir:  # 如果获取不到路径，则使用当前工作目录
                    script_dir = os.getcwd()

            # 配置文件路径（从程序目录加载）
            config_file = os.path.join(script_dir, "config.json")

            # 读取配置文件
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                print(f"配置已从 {config_file} 加载")
            else:
                print(f"未找到配置文件 {config_file}，使用默认设置")
                return

            # 设置字段映射
            if "field_mappings" in config_data:
                # 定义有效的字段列表
                valid_fields = ['Item', 'P/N', 'Reference', 'Description', 'MPN']
                new_field_mappings = {}

                # 处理配置文件中的字段映射
                for field, aliases in config_data["field_mappings"].items():
                    # 如果是有效字段，则保留
                    if field in valid_fields:
                        new_field_mappings[field] = aliases

                # 确保所有必要字段都存在
                for field in valid_fields:
                    if field not in new_field_mappings:
                        new_field_mappings[field] = self.comparer.field_mappings.get(field, [])

                # 更新比较器的字段映射
                self.comparer.set_field_mappings(new_field_mappings)

            # 设置MPN显示选项
            if "show_mpn_in_report" in config_data:
                self.comparer.show_mpn_in_report = config_data["show_mpn_in_report"]

            # 设置最后打开的目录
            if "last_dir" in config_data:
                self.last_dir = config_data["last_dir"]

        except Exception as e:
            print(f"加载配置时出错: {str(e)}")
            # 这里不弹出错误消息，因为这不是关键功能

    def on_close(self):
        # 在关闭窗口前保存配置
        self.save_config_to_file()
        self.root.destroy()

    def clean_config_files(self):
        """清理配置文件中的无效字段"""
        try:
            # 获取程序所在目录的绝对路径
            script_dir = os.path.dirname(os.path.abspath(__file__))

            # 配置文件路径
            config_file = os.path.join(script_dir, "config.json")

            # 检查并清理程序目录下的配置文件
            if os.path.exists(config_file):
                self._clean_config_file(config_file)

            # 尝试清理可能存在的用户主目录中的旧配置文件（一次性迁移）
            user_config_file = os.path.join(os.path.expanduser("~"), ".bom_comparer", "config.json")
            if os.path.exists(user_config_file):
                try:
                    # 读取用户目录的配置
                    with open(user_config_file, 'r', encoding='utf-8') as f:
                        user_config = json.load(f)

                    # 清理后保存到程序目录
                    self._clean_config_file(user_config_file)
                    os.remove(user_config_file)
                    print(f"已移除旧配置文件: {user_config_file}")
                except:
                    pass

        except Exception as e:
            print(f"清理配置文件时出错: {str(e)}")

    def _clean_config_file(self, file_path):
        """清理指定的配置文件"""
        try:
            # 读取配置文件
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            # 如果包含字段映射
            if "field_mappings" in config_data:
                # 定义有效的字段列表
                valid_fields = ['Item', 'P/N', 'Reference', 'Description', 'MPN']
                new_field_mappings = {}

                # 处理配置文件中的字段映射
                for field, aliases in config_data["field_mappings"].items():
                    # 如果是有效字段，则保留
                    if field in valid_fields:
                        new_field_mappings[field] = aliases

                # 更新配置数据
                config_data["field_mappings"] = new_field_mappings

                # 保存更新后的配置
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, ensure_ascii=False, indent=4)

                print(f"已清理配置文件: {file_path}")
        except Exception as e:
            print(f"清理配置文件 {file_path} 时出错: {str(e)}")

    def create_bom_table(self, parent_frame, is_bom_a=True):
        """创建BOM数据显示表格"""
        # 创建表格框架
        table_frame = ttk.Frame(parent_frame)
        table_frame.pack(fill="both", expand=True, padx=3, pady=3)  # 从5减小到3

        # 创建表头 - 初始不指定列，将在加载数据时动态设置
        columns = []

        # 创建Treeview组件
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse", height=8)  # 设置为browse模式，确保单击选择功能正常

        # 确保树视图的样式设置正确，支持选择高亮
        style = ttk.Style()
        style.configure("Treeview",
                        background="#ffffff",  # 正常背景色
                        fieldbackground="#ffffff",  # 字段背景色
                        foreground="#000000")  # 前景色（文字颜色）

        # 配置选中项的样式
        style.map("Treeview",
                  background=[("selected", "#0078d7")],  # 选中项的背景色
                  foreground=[("selected", "#ffffff")])  # 选中项的文字颜色

        # 绑定单击事件
        tree.bind('<<TreeviewSelect>>', lambda event: self.comparer.on_tree_select(event, tree, is_bom_a))

        # 绑定双击事件 - 用于搜索结果
        tree.bind('<Double-1>', lambda event: self.comparer.on_tree_double_click(event, tree, is_bom_a))

        # 创建滚动条
        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # 放置表格和滚动条
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        # 保存引用
        if is_bom_a:
            self.bom_a_tree = tree
            self.bom_a_xscroll = scrollbar_x
            self.bom_a_yscroll = scrollbar_y
        else:
            self.bom_b_tree = tree
            self.bom_b_xscroll = scrollbar_x
            self.bom_b_yscroll = scrollbar_y

        return tree

    def is_valid_reference(self, ref):
        """检查是否为有效位号（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        # 在A中查找该位号
        for _, row in self.comparer.bom_a.iterrows():
            refs = str(row.get('Reference', ''))
            if ref in refs.split(',') or ref in refs.split():
                return True

        # 在B中查找该位号
        for _, row in self.comparer.bom_b.iterrows():
            refs = str(row.get('Reference', ''))
            if ref in refs.split(',') or ref in refs.split():
                return True

        return False

    def is_valid_pn(self, pn):
        """检查是否为有效料号（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        # 打印调试信息
        print(f"检查料号是否有效: {pn}")

        # 在A中查找该料号，使用精确匹配
        found_a = (self.comparer.bom_a['P/N'].astype(str) == pn).any()
        if not found_a:
            # 如果精确匹配失败，尝试模糊匹配
            found_a = self.comparer.bom_a['P/N'].astype(str).str.contains(pn, regex=False).any()

        # 在B中查找该料号，使用精确匹配
        found_b = (self.comparer.bom_b['P/N'].astype(str) == pn).any()
        if not found_b:
            # 如果精确匹配失败，尝试模糊匹配
            found_b = self.comparer.bom_b['P/N'].astype(str).str.contains(pn, regex=False).any()

        result = found_a or found_b
        print(f"料号{pn}在BOM中{'存在' if result else '不存在'}")
        return result

    def is_valid_mpn(self, mpn):
        """检查是否为有效MPN（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        if 'MPN' not in self.comparer.bom_a.columns or 'MPN' not in self.comparer.bom_b.columns:
            return False

        # 在A中查找该MPN
        found_a = self.comparer.bom_a['MPN'].astype(str).str.contains(mpn, na=False).any()
        # 在B中查找该MPN
        found_b = self.comparer.bom_b['MPN'].astype(str).str.contains(mpn, na=False).any()

        return found_a or found_b

    def on_result_double_click(self, event):
        """处理结果文本区域的双击事件，在BOM数据区高亮显示对应位号或物料的数据行"""
        # 确保BOM数据已加载
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b') or \
           self.comparer.bom_a is None or self.comparer.bom_b is None:
            self.show_error("请先加载BOM文件")
            return

        # 先清除之前的高亮显示
        self.clear_tree_highlights()

        # 清除所有树视图的选择状态，避免同时存在蓝色选择和黄色高亮
        if hasattr(self, 'bom_a_tree'):
            self.bom_a_tree.selection_remove(self.bom_a_tree.selection())
        if hasattr(self, 'bom_b_tree'):
            self.bom_b_tree.selection_remove(self.bom_b_tree.selection())

        # 获取光标位置的行和列
        index = self.result_text.index(f"@{event.x},{event.y}")
        line_info = index.split(".")
        line_num = int(line_info[0])
        col = int(line_info[1])

        # 获取当前行的文本
        line_start = f"{line_num}.0"
        line_end = f"{line_num}.end"
        line_text = self.result_text.get(line_start, line_end)

        print(f"双击的行文本: {line_text}")
        print(f"双击的位置: 行={line_num}, 列={col}")

        # 检查是否为新增、移除或变更位号的行
        is_added_ref = "新增:" in line_text or "新增位号:" in line_text
        is_removed_ref = "移除:" in line_text or "移除位号:" in line_text
        is_changed_ref = "变更:" in line_text
        is_positions = "位号:" in line_text  # 位号列表行
        is_alternative = "替代料:" in line_text or "替代物料:" in line_text  # 替代料行

        # 查找位号、物料编码、型号等信息
        # 尝试匹配位号模式 (XX123, R123, C123, U123等)
        ref_pattern = r'([A-Z]{1,2}\d+)'
        pn_pattern = r'(MPE\d+|[A-Z0-9]{5,})'  # 匹配料号，特别处理MPE开头的料号
        mpn_pattern = r'(RC\d+[A-Z0-9\-]+|GRM\d+[A-Z0-9\-]+|SY\d+[A-Z0-9\-]+|UM\d+[A-Z0-9\-]+|ME\d+[A-Z0-9\-]+)'  # 匹配MPN

        found_ref = None
        found_pn = None
        found_mpn = None

        # 特殊处理替代料行 - 高亮双击的特定料号
        if is_alternative:
            print("检测到替代料行")
            # 提取所有料号
            pn_matches = re.findall(pn_pattern, line_text)

            # 查找最接近光标位置的料号
            closest_pn = None
            min_distance = float('inf')

            for pn in pn_matches:
                if self.is_valid_pn(pn):
                    # 查找料号在文本中的位置
                    start_idx = line_text.find(pn)
                    if start_idx != -1:
                        # 计算与光标的距离
                        distance = abs(start_idx - col)
                        print(f"料号: {pn}, 位置: {start_idx}, 距离光标: {distance}")
                        if distance < min_distance:
                            min_distance = distance
                            closest_pn = pn

            # 如果找到了最接近的料号，使用它
            if closest_pn and min_distance < 15:  # 使用稍大的距离阈值
                found_pn = closest_pn
                print(f"在替代料行中找到最接近光标的料号: {found_pn}, 距离: {min_distance}")

                # 直接高亮这个料号
                self.clear_tree_highlights()  # 确保清除之前的高亮

                # 直接在A树和B树中查找并高亮这个料号
                found_in_a = False
                found_in_b = False

                # 获取所有行
                for item in self.bom_a_tree.get_children():
                    values = self.bom_a_tree.item(item, 'values')
                    columns = self.bom_a_tree['columns']

                    # 查找物料编号列
                    pn_col_idx = -1
                    for i, col in enumerate(columns):
                        if col == 'P/N':
                            pn_col_idx = i
                            break

                    # 如果找到了物料编号列
                    if pn_col_idx >= 0 and pn_col_idx < len(values):
                        cell_value = str(values[pn_col_idx]).strip()

                        # 精确匹配料号
                        if cell_value == found_pn:
                            # 高亮显示
                            tag_name = f"material_highlight_bom_a_{item}_{found_pn}"
                            if not hasattr(self, 'highlight_tags'):
                                self.highlight_tags = set()
                            self.highlight_tags.add(tag_name)
                            self.bom_a_tree.item(item, tags=(tag_name,))
                            self.bom_a_tree.tag_configure(tag_name, background='#FFFF99')
                            self.bom_a_tree.see(item)
                            found_in_a = True

                # 同样在B树中查找
                for item in self.bom_b_tree.get_children():
                    values = self.bom_b_tree.item(item, 'values')
                    columns = self.bom_b_tree['columns']

                    # 查找物料编号列
                    pn_col_idx = -1
                    for i, col in enumerate(columns):
                        if col == 'P/N':
                            pn_col_idx = i
                            break

                    # 如果找到了物料编号列
                    if pn_col_idx >= 0 and pn_col_idx < len(values):
                        cell_value = str(values[pn_col_idx]).strip()

                        # 精确匹配料号
                        if cell_value == found_pn:
                            # 高亮显示
                            tag_name = f"material_highlight_bom_b_{item}_{found_pn}"
                            if not hasattr(self, 'highlight_tags'):
                                self.highlight_tags = set()
                            self.highlight_tags.add(tag_name)
                            self.bom_b_tree.item(item, tags=(tag_name,))
                            self.bom_b_tree.tag_configure(tag_name, background='#FFFF99')
                            self.bom_b_tree.see(item)
                            found_in_b = True

                if found_in_a or found_in_b:
                    print(f"成功高亮显示替代料号: {found_pn}")
                    return
                else:
                    print(f"未能在任何树中找到并高亮替代料号: {found_pn}")

        # 特殊处理位号行 - 优先提取位号，然后查找对应物料信息
        if is_positions:
            # 提取位号列表
            refs = re.findall(ref_pattern, line_text)
            if refs:
                # 选取用户双击附近的位号
                closest_ref = None
                min_distance = float('inf')
                for ref in refs:
                    start_idx = line_text.find(ref)
                    if start_idx != -1:
                        distance = abs(start_idx - col)
                        if distance < min_distance:
                            min_distance = distance
                            closest_ref = ref

                if closest_ref and min_distance < 10:  # 认为距离小于10个字符是用户有意双击的
                    found_ref = closest_ref

        # 如果不是位号行或未找到最近位号，尝试常规匹配
        if not found_ref and not found_pn:
            # 查找位号
            ref_matches = re.findall(ref_pattern, line_text)
            for ref in ref_matches:
                # 检查是否为有效位号（在A或B中存在）
                if self.is_valid_reference(ref):
                    # 计算与光标的距离
                    start_idx = line_text.find(ref)
                    if start_idx != -1 and abs(start_idx - col) < 10:  # 只选择接近光标的位号
                        found_ref = ref
                        break

        # 如果找不到位号，尝试查找料号
        if not found_ref and not found_pn:
            # 提取整行中的料号信息
            pn_matches = re.findall(pn_pattern, line_text)

            # 查找最接近光标位置的料号
            closest_pn = None
            min_distance = float('inf')

            for pn in pn_matches:
                if self.is_valid_pn(pn):
                    # 查找料号在文本中的位置
                    start_idx = line_text.find(pn)
                    if start_idx != -1:
                        # 计算与光标的距离
                        distance = abs(start_idx - col)
                        if distance < min_distance:
                            min_distance = distance
                            closest_pn = pn

            # 如果找到了最接近的料号，使用它
            if closest_pn and min_distance < 15:  # 使用稍大的距离阈值
                found_pn = closest_pn

        # 如果找不到料号，尝试查找MPN
        if not found_ref and not found_pn:
            # 提取整行中的MPN信息
            mpn_matches = re.findall(mpn_pattern, line_text)

            # 查找最接近光标位置的MPN
            closest_mpn = None
            min_distance = float('inf')

            for mpn in mpn_matches:
                if self.is_valid_mpn(mpn):
                    # 查找MPN在文本中的位置
                    start_idx = line_text.find(mpn)
                    if start_idx != -1:
                        # 计算与光标的距离
                        distance = abs(start_idx - col)
                        if distance < min_distance:
                            min_distance = distance
                            closest_mpn = mpn

            # 如果找到了最接近的MPN，使用它
            if closest_mpn and min_distance < 15:  # 使用稍大的距离阈值
                found_mpn = closest_mpn

        # 显示找到的信息
        if found_ref:
            # 根据上下文，决定是普通位号比较还是变更位号比较
            if is_added_ref:
                # 对于新增位号，只有B中有该位号
                info_b, _ = self.find_reference_info(self.comparer.bom_b, found_ref)
                if info_b:
                    actual_pn_b = info_b.get('P/N', '')
                    print(f"位号{found_ref}在B BOM中的料号: {actual_pn_b}")
                    # 查找A中是否有相同料号的物料
                    info_a = self.find_pn_info(self.comparer.bom_a, actual_pn_b)
                    # 高亮显示
                    self.highlight_reference_in_trees(found_ref, None, actual_pn_b)
                    return
            elif is_removed_ref:
                # 对于移除位号，只有A中有该位号
                info_a, _ = self.find_reference_info(self.comparer.bom_a, found_ref)
                if info_a:
                    actual_pn_a = info_a.get('P/N', '')
                    print(f"位号{found_ref}在A BOM中的料号: {actual_pn_a}")
                    # 查找B中是否有相同料号的物料
                    info_b = self.find_pn_info(self.comparer.bom_b, actual_pn_a)
                    # 高亮显示
                    self.highlight_reference_in_trees(found_ref, actual_pn_a)
                    return

            # 获取位号在A和B中的信息
            info_a, _ = self.find_reference_info(self.comparer.bom_a, found_ref)
            info_b, _ = self.find_reference_info(self.comparer.bom_b, found_ref)

            # 如果位号在A中存在
            if info_a:
                actual_pn_a = info_a.get('P/N', '')
                print(f"位号{found_ref}在A BOM中的料号: {actual_pn_a}")
                # 如果位号在B中也存在
                if info_b:
                    actual_pn_b = info_b.get('P/N', '')
                    print(f"位号{found_ref}在B BOM中的料号: {actual_pn_b}")
                    # 高亮显示，使用实际料号
                    self.highlight_reference_in_trees(found_ref, actual_pn_a, actual_pn_b)
                else:
                    # 只在A中存在，高亮A中的料号
                    self.highlight_reference_in_trees(found_ref, actual_pn_a)
            # 如果位号只在B中存在
            elif info_b:
                actual_pn_b = info_b.get('P/N', '')
                print(f"位号{found_ref}在B BOM中的料号: {actual_pn_b}")
                # 高亮B中的料号
                self.highlight_reference_in_trees(found_ref, None, actual_pn_b)
            else:
                # 如果在A和B中都找不到位号信息，使用常规高亮
                self.highlight_reference_in_trees(found_ref)
        elif found_pn:
            # 双击料号时，高亮对应的特定料号（直接使用双击的料号）
            print(f"高亮显示用户双击的料号: {found_pn}")
            self.highlight_material_in_both_trees(found_pn)
        elif found_mpn:
            # 查找包含该MPN的物料并高亮显示
            info_a = self.find_mpn_info(self.comparer.bom_a, found_mpn)
            info_b = self.find_mpn_info(self.comparer.bom_b, found_mpn)

            if info_a:
                self.highlight_material_in_tree(self.bom_a_tree, info_a.get('P/N', ''))
            if info_b:
                self.highlight_material_in_tree(self.bom_b_tree, info_b.get('P/N', ''))
        else:
            # 恢复光标位置的选区，查找更多匹配可能
            word_start = self.result_text.index(f"@{event.x},{event.y} wordstart")
            word_end = self.result_text.index(f"@{event.x},{event.y} wordend")
            selected_word = self.result_text.get(word_start, word_end)

            if selected_word and len(selected_word) > 2:
                self.display_text_search_details(selected_word)

    def clear_tree_highlights(self):
        """清除所有树视图中的高亮显示"""
        print("清除所有高亮显示")

        # 清除BOM A树视图中的高亮显示
        for item in self.bom_a_tree.get_children():
            self.bom_a_tree.item(item, tags=())

        # 清除BOM B树视图中的高亮显示
        for item in self.bom_b_tree.get_children():
            self.bom_b_tree.item(item, tags=())

        # 清除所有标签配置 - 使用存储的标签列表
        if hasattr(self, 'highlight_tags'):
            for tag in self.highlight_tags:
                try:
                    if tag.startswith('ref_highlight_bom_a_'):
                        self.bom_a_tree.tag_configure(tag, background='')
                    elif tag.startswith('ref_highlight_bom_b_'):
                        self.bom_b_tree.tag_configure(tag, background='')
                except Exception as e:
                    print(f"清除标签配置时出错: {e}")

            # 清除标签列表
            self.highlight_tags.clear()

        # 恢复原始值
        if hasattr(self, 'original_values') and self.original_values:
            print(f"恢复原始值，共{len(self.original_values)}个值")
            for key, value in list(self.original_values.items()):
                try:
                    # 解析键值 - 新的键值格式是 id(tree).item.Reference.ref
                    parts = key.split('.')
                    if len(parts) >= 4:  # 新格式
                        tree_id = parts[0]
                        item_id = parts[1]
                        column = parts[2]

                        # 确定树视图
                        tree = None
                        if id(self.bom_a_tree) == int(tree_id):
                            tree = self.bom_a_tree
                        elif id(self.bom_b_tree) == int(tree_id):
                            tree = self.bom_b_tree

                        if tree and item_id in tree.get_children():
                            # 恢复原始值
                            tree.set(item_id, column, value)
                            print(f"恢复原始值: {item_id}.{column} = {value}")
                    elif len(parts) >= 3:  # 兼容旧格式
                        tree_name = parts[0]
                        item_id = parts[1]
                        column = parts[2]

                        # 确定树视图
                        tree = None
                        if 'bom_a_tree' in tree_name:
                            tree = self.bom_a_tree
                        elif 'bom_b_tree' in tree_name:
                            tree = self.bom_b_tree

                        if tree and item_id in tree.get_children():
                            # 恢复原始值
                            tree.set(item_id, column, value)
                            print(f"恢复原始值: {item_id}.{column} = {value}")
                except Exception as e:
                    print(f"恢复原始值时出错: {e}")

            # 清除存储的原始值
            self.original_values.clear()

    def is_valid_reference(self, ref):
        """检查是否为有效位号（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        # 在A中查找该位号
        for _, row in self.comparer.bom_a.iterrows():
            refs = str(row.get('Reference', ''))
            if ref in refs.split(',') or ref in refs.split():
                return True

        # 在B中查找该位号
        for _, row in self.comparer.bom_b.iterrows():
            refs = str(row.get('Reference', ''))
            if ref in refs.split(',') or ref in refs.split():
                return True

        return False

    def is_valid_pn(self, pn):
        """检查是否为有效料号（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        # 打印调试信息
        print(f"检查料号是否有效: {pn}")

        # 在A中查找该料号，使用精确匹配
        found_a = (self.comparer.bom_a['P/N'].astype(str) == pn).any()
        if not found_a:
            # 如果精确匹配失败，尝试模糊匹配
            found_a = self.comparer.bom_a['P/N'].astype(str).str.contains(pn, regex=False).any()

        # 在B中查找该料号，使用精确匹配
        found_b = (self.comparer.bom_b['P/N'].astype(str) == pn).any()
        if not found_b:
            # 如果精确匹配失败，尝试模糊匹配
            found_b = self.comparer.bom_b['P/N'].astype(str).str.contains(pn, regex=False).any()

        result = found_a or found_b
        print(f"料号{pn}在BOM中{'存在' if result else '不存在'}")
        return result

    def is_valid_mpn(self, mpn):
        """检查是否为有效MPN（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        if 'MPN' not in self.comparer.bom_a.columns or 'MPN' not in self.comparer.bom_b.columns:
            return False

        # 在A中查找该MPN
        found_a = self.comparer.bom_a['MPN'].astype(str).str.contains(mpn, na=False).any()
        # 在B中查找该MPN
        found_b = self.comparer.bom_b['MPN'].astype(str).str.contains(mpn, na=False).any()

        return found_a or found_b

    def highlight_reference_in_trees(self, ref, pn_a=None, pn_b=None):
        """在树视图中高亮显示包含指定位号的行，并标红位号字段

        Args:
            ref: 要高亮的位号
            pn_a: 可选的A BOM中的料号，用于精确匹配
            pn_b: 可选的B BOM中的料号，用于精确匹配
        """
        print(f"高亮显示位号: {ref}")
        print(f"指定的料号: A={pn_a}, B={pn_b}")

        # 先彻底清除之前的高亮，确保之前的标记被清除
        self.clear_tree_highlights()

        # 在BOM A中查找位号
        info_a, _ = self.find_reference_info(self.comparer.bom_a, ref)
        # 在BOM B中查找位号
        info_b, _ = self.find_reference_info(self.comparer.bom_b, ref)

        # 如果两个BOM中都没有找到该位号，则返回
        if info_a is None and info_b is None:
            print(f"未找到位号: {ref}")
            return

        # 初始化标签列表，如果不存在
        if not hasattr(self, 'highlight_tags'):
            self.highlight_tags = set()

        # 在BOM A中高亮显示对应行并标红位号
        if info_a is not None:
            print(f"在BOM A中找到位号: {ref}")
            # 如果没有指定料号，使用位号对应的料号
            if pn_a is None:
                pn_a = info_a.get('P/N', '')
            print(f"在BOM A中使用料号: {pn_a}")
            self.highlight_reference_cell(self.bom_a_tree, ref, pn_a)

        # 在BOM B中高亮显示对应行并标红位号
        if info_b is not None:
            print(f"在BOM B中找到位号: {ref}")
            # 如果没有指定料号，使用位号对应的料号
            if pn_b is None:
                pn_b = info_b.get('P/N', '')
            print(f"在BOM B中使用料号: {pn_b}")
            self.highlight_reference_cell(self.bom_b_tree, ref, pn_b)

        # 处理位号移除或新增的情况
        if info_a is not None and info_b is None:
            # 位号在A中存在但在B中不存在，说明该位号被移除
            print(f"位号{ref}在B中被移除，尝试高亮对应的物料行")
            self.highlight_corresponding_material_row(ref, info_a, 'A_to_B')
        elif info_a is None and info_b is not None:
            # 位号在A中不存在但在B中存在，说明该位号是新增的
            print(f"位号{ref}在B中是新增的，尝试高亮对应的物料行")
            self.highlight_corresponding_material_row(ref, info_b, 'B_to_A')
        elif info_a is not None and info_b is not None:
            # 位号在A和B中都存在，检查是否物料变更
            pn_a = info_a.get('P/N', '')
            pn_b = info_b.get('P/N', '')
            if pn_a != pn_b:
                print(f"位号{ref}的物料发生变更: {pn_a} -> {pn_b}")

    def highlight_corresponding_material_row(self, ref, info, direction):
        """高亮显示与指定位号对应的物料行

        Args:
            ref: 位号
            info: 位号对应的物料信息
            direction: 方向，'A_to_B'表示从原始BOM到新BOM，'B_to_A'表示从新BOM到原始BOM
        """
        # 获取物料编号
        pn = info.get('P/N', '')
        if not pn:
            print(f"无法获取位号{ref}对应的物料编号")
            return

        # 确定目标树和源树
        if direction == 'A_to_B':
            # 从原始BOM到新BOM，在B树中查找相同物料
            target_tree = self.bom_b_tree
            # source_tree变量在这里没有使用，可以删除
            target_bom = self.comparer.bom_b
        else:  # 'B_to_A'
            # 从新BOM到原始BOM，在A树中查找相同物料
            target_tree = self.bom_a_tree
            # source_tree变量在这里没有使用，可以删除
            target_bom = self.comparer.bom_a

        # 在目标树中查找相同物料编号的行
        found = False
        for item in target_tree.get_children():
            values = target_tree.item(item, 'values')
            columns = target_tree['columns']

            # 查找物料编号列
            pn_col_idx = -1
            for i, col in enumerate(columns):
                if col == 'P/N':
                    pn_col_idx = i
                    break

            # 如果找到了物料编号列
            if pn_col_idx >= 0 and pn_col_idx < len(values):
                cell_value = str(values[pn_col_idx])

                # 检查是否与目标物料编号匹配
                if cell_value == pn:
                    # 定义一个特殊的标签，用于标记这一行
                    tree_name = "bom_a" if target_tree == self.bom_a_tree else "bom_b"
                    tag_name = f"material_highlight_{tree_name}_{item}_{pn}"

                    # 初始化标签列表，如果不存在
                    if not hasattr(self, 'highlight_tags'):
                        self.highlight_tags = set()

                    # 将标签添加到列表中
                    self.highlight_tags.add(tag_name)

                    # 将这个标签应用于该行
                    target_tree.item(item, tags=(tag_name,))

                    # 配置这个标签的样式，将背景设置为黄色
                    target_tree.tag_configure(tag_name, background='#FFFF99')  # 黄色背景

                    # 滚动到该行
                    target_tree.see(item)
                    found = True
                    print(f"在{tree_name}中找到并高亮显示了物料 {pn}")

        if not found:
            print(f"在目标树中未找到物料 {pn}")

            # 尝试查找替代料或相似物料
            # 直接检查当前料号的替代料
            alt_found = False
            if pn in self.comparer.alternative_map:
                # 如果当前料号有替代料
                alt_pns = self.comparer.alternative_map[pn]
                for alt_pn in alt_pns:
                    # 在目标BOM中查找该替代料
                    if alt_pn in target_bom['P/N'].values:
                        print(f"在目标树中找到替代料 {alt_pn}")
                        # 高亮显示该替代料
                        self.highlight_material_in_tree(target_tree, alt_pn)
                        alt_found = True
                        break

            # 如果没有找到替代料，检查是否有其他料号将当前料号作为替代料
            if not alt_found:
                for other_pn, alt_pns in self.comparer.alternative_map.items():
                    if pn in alt_pns:
                        # 如果当前料号是其他料号的替代料
                        # 先检查主料号
                        if other_pn in target_bom['P/N'].values:
                            print(f"在目标树中找到主料号 {other_pn}")
                            self.highlight_material_in_tree(target_tree, other_pn)
                            alt_found = True
                            break

                        # 再检查其他替代料
                        for alt_pn in alt_pns:
                            if alt_pn != pn and alt_pn in target_bom['P/N'].values:
                                print(f"在目标树中找到替代料 {alt_pn}")
                                self.highlight_material_in_tree(target_tree, alt_pn)
                                alt_found = True
                                break

                        if alt_found:
                            break

    def highlight_reference_cell(self, tree, ref, pn):
        """在树视图中高亮显示包含指定位号的行，并标红位号字段

        Args:
            tree: 树视图对象
            ref: 要高亮的位号
            pn: 物料编号，用于精确匹配行
        """
        print(f"在树中高亮显示位号: {ref}, 料号: {pn}")
        found = False

        # 清除树视图的选择状态，避免蓝色高亮与黄色高亮同时存在
        tree.selection_remove(tree.selection())

        # 获取所有行
        for item in tree.get_children():
            # 获取行数据
            values = tree.item(item, 'values')
            columns = tree['columns']

            # 查找物料编号列和位号列
            pn_col_idx = -1
            ref_col_idx = -1
            for i, col in enumerate(columns):
                if col == 'P/N':
                    pn_col_idx = i
                elif col == 'Reference':
                    ref_col_idx = i

            # 如果找到了物料编号列和位号列
            if pn_col_idx >= 0 and pn_col_idx < len(values) and ref_col_idx >= 0 and ref_col_idx < len(values):
                cell_pn = str(values[pn_col_idx]).strip()
                cell_ref = str(values[ref_col_idx]).strip()

                # 检查位号是否匹配
                ref_match = False
                if ref in cell_ref.split(',') or ref in cell_ref.split():
                    ref_match = True

                # 检查料号是否匹配
                pn_match = (cell_pn == pn)

                # 如果位号和料号都匹配
                if ref_match and pn_match:
                    print(f"找到匹配的位号和料号: {ref}, {pn}")
                    found = True

                    # 定义一个特殊的标签，用于标记这一行
                    tree_name = "bom_a" if tree == self.bom_a_tree else "bom_b"
                    tag_name = f"ref_highlight_{tree_name}_{item}_{ref}"

                    # 初始化标签列表，如果不存在
                    if not hasattr(self, 'highlight_tags'):
                        self.highlight_tags = set()

                    # 将标签添加到列表中
                    self.highlight_tags.add(tag_name)

                    # 将这个标签应用于该行
                    tree.item(item, tags=(tag_name,))

                    # 配置这个标签的样式，使用黄色背景
                    tree.tag_configure(tag_name, background='#FFFF99')

                    # 滚动到该行
                    tree.see(item)
                    break

        # 如果没有找到，确保清除所有选择，避免用户困惑
        if not found:
            print(f"未找到匹配的位号和料号: {ref}, {pn}")
            tree.selection_remove(tree.selection())

        return found

    def highlight_material_in_tree(self, tree, pn):
        """在指定的树视图中高亮显示指定物料编号的行"""
        print(f"在树中查找并高亮料号: {pn}")
        found = False

        # 清除树视图的选择状态，避免蓝色高亮与黄色高亮同时存在
        tree.selection_remove(tree.selection())

        # 获取所有行
        for item in tree.get_children():
            # 获取行数据
            values = tree.item(item, 'values')
            columns = tree['columns']

            # 查找物料编号列
            pn_col_idx = -1
            for i, col in enumerate(columns):
                if col == 'P/N':
                    pn_col_idx = i
                    break

            # 如果找到了物料编号列
            if pn_col_idx >= 0 and pn_col_idx < len(values):
                cell_value = str(values[pn_col_idx]).strip()

                # 打印调试信息
                # print(f"比较: '{cell_value}' vs '{pn}'")

                # 精确匹配料号
                if cell_value == pn:
                    print(f"找到精确匹配的料号: {pn}")
                    found = True

                    # 定义一个特殊的标签，用于标记这一行
                    tree_name = "bom_a" if tree == self.bom_a_tree else "bom_b"
                    tag_name = f"material_highlight_{tree_name}_{item}_{pn}"

                    # 初始化标签列表，如果不存在
                    if not hasattr(self, 'highlight_tags'):
                        self.highlight_tags = set()

                    # 将标签添加到列表中
                    self.highlight_tags.add(tag_name)

                    # 将这个标签应用于该行
                    tree.item(item, tags=(tag_name,))

                    # 配置这个标签的样式，加粗字体并使用黄色背景，以便更明显
                    tree.tag_configure(tag_name, background='#FFFF99', font=('Arial', 10, 'bold'))

                    # 取消任何现有选择，不使用蓝色高亮
                    tree.selection_remove(tree.selection())

                    # 不要设置选择，避免蓝色高亮
                    # tree.selection_set(item)

                    # 滚动到该行
                    tree.see(item)

                    # 仅将焦点设置到该行，但不选择它
                    # tree.focus(item)

        # 如果没有找到，确保清除所有选择，避免用户困惑
        if not found:
            tree.selection_remove(tree.selection())

        return found

    def highlight_material_in_both_trees(self, pn):
        """在两个BOM树中高亮显示指定物料编号的行"""
        print(f"在两个BOM树中高亮显示物料: {pn}")

        # 先清除之前的高亮显示
        self.clear_tree_highlights()

        # 确保清除所有树视图的选择状态，避免同时存在蓝色选择和黄色高亮
        if hasattr(self, 'bom_a_tree'):
            self.bom_a_tree.selection_remove(self.bom_a_tree.selection())
        if hasattr(self, 'bom_b_tree'):
            self.bom_b_tree.selection_remove(self.bom_b_tree.selection())

        # 在A树中高亮显示
        found_a = self.highlight_material_in_tree(self.bom_a_tree, pn)
        if found_a:
            print(f"在BOM A中找到并高亮显示了物料 {pn}")

        # 在B树中高亮显示
        found_b = self.highlight_material_in_tree(self.bom_b_tree, pn)
        if found_b:
            print(f"在BOM B中找到并高亮显示了物料 {pn}")

        # 如果两个BOM中都没有找到该物料，输出提示信息
        if not found_a and not found_b:
            print(f"在两个BOM中都未找到物料 {pn}")

        return found_a or found_b

    def highlight_row_in_tree(self, tree, search_column, search_value, tag):
        """在树视图中高亮显示符合条件的行

        Args:
            tree: 树视图对象
            search_column: 要搜索的列名
            search_value: 要搜索的值
            tag: 要应用的标签名

        Returns:
            bool: 是否找到并高亮了符合条件的行
        """
        print(f"在树中查找并高亮: 列={search_column}, 值={search_value}")
        found = False

        # 获取所有行
        for item in tree.get_children():
            # 获取行数据
            values = tree.item(item, 'values')
            columns = tree['columns']

            # 查找目标列
            col_idx = -1
            for i, col in enumerate(columns):
                if col == search_column:
                    col_idx = i
                    break

            # 如果找到了目标列
            if col_idx >= 0 and col_idx < len(values):
                cell_value = str(values[col_idx]).strip()

                # 检查是否匹配
                if cell_value == search_value:
                    print(f"找到匹配的行: {search_value}")
                    found = True

                    # 将标签应用于该行
                    tree.item(item, tags=(tag,))

                    # 配置标签样式
                    tree.tag_configure(tag, background='#FFFF99')

                    # 滚动到该行
                    tree.see(item)
                    break

        return found

    def is_valid_pn(self, pn):
        """检查是否为有效料号（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        # 在A中查找该料号
        found_a = self.comparer.bom_a['P/N'].astype(str).str.contains(pn).any()
        # 在B中查找该料号
        found_b = self.comparer.bom_b['P/N'].astype(str).str.contains(pn).any()

        return found_a or found_b

    def is_valid_mpn(self, mpn):
        """检查是否为有效MPN（在BOM A或B中存在）"""
        if not hasattr(self.comparer, 'bom_a') or not hasattr(self.comparer, 'bom_b'):
            return False

        if 'MPN' not in self.comparer.bom_a.columns or 'MPN' not in self.comparer.bom_b.columns:
            return False

        # 在A中查找该MPN
        found_a = self.comparer.bom_a['MPN'].astype(str).str.contains(mpn, na=False).any()
        # 在B中查找该MPN
        found_b = self.comparer.bom_b['MPN'].astype(str).str.contains(mpn, na=False).any()

        return found_a or found_b

    def find_reference_info(self, bom_data, ref):
        """在BOM数据中查找位号对应的信息，返回物料信息及位号位置"""
        if bom_data is None or 'Reference' not in bom_data.columns:
            return None, None

        for _, row in bom_data.iterrows():
            refs = str(row.get('Reference', ''))

            # 分割位号字符串为列表
            ref_list = []
            if ',' in refs:
                ref_list = [r.strip() for r in refs.split(',')]
            else:
                ref_list = [r.strip() for r in refs.split()]

            if ref in ref_list:
                # 返回匹配的物料信息和位号在列表中的位置
                return row.to_dict(), ref_list.index(ref)

        return None, None

    def find_pn_info(self, bom_data, pn):
        """在BOM数据中查找料号对应的信息"""
        if bom_data is None or 'P/N' not in bom_data.columns:
            return None

        # 只进行精确匹配，不再进行模糊匹配
        for _, row in bom_data.iterrows():
            if str(row.get('P/N', '')).strip() == pn:
                return row.to_dict()

        return None

    def find_mpn_info(self, bom_data, mpn):
        """在BOM数据中查找MPN对应的信息"""
        if bom_data is None or 'MPN' not in bom_data.columns:
            return None

        # 只进行精确匹配，不再进行模糊匹配
        for _, row in bom_data.iterrows():
            if str(row.get('MPN', '')).strip() == mpn:
                return row.to_dict()

        return None

    def setup_synchronized_scrolling(self):
        """设置两个BOM表格的同步滚动"""
        # 标记是否正在处理同步，防止无限递归
        self.syncing_x = False
        self.syncing_y = False

        # 重新定义滚动函数为类的方法
        def sync_a_to_b_xview(*args):
            if not self.syncing_x:
                self.syncing_x = True
                self.bom_b_tree.xview(*args)
                self.syncing_x = False

        def sync_b_to_a_xview(*args):
            if not self.syncing_x:
                self.syncing_x = True
                self.bom_a_tree.xview(*args)
                self.syncing_x = False

        def sync_a_to_b_yview(*args):
            if not self.syncing_y:
                self.syncing_y = True
                self.bom_b_tree.yview(*args)
                self.syncing_y = False

        def sync_b_to_a_yview(*args):
            if not self.syncing_y:
                self.syncing_y = True
                self.bom_a_tree.yview(*args)
                self.syncing_y = False

        # 配置BOM表格的水平滚动条
        self.bom_a_xscroll.config(command=lambda *args: [self.bom_a_tree.xview(*args), sync_a_to_b_xview(*args)])
        self.bom_b_xscroll.config(command=lambda *args: [self.bom_b_tree.xview(*args), sync_b_to_a_xview(*args)])

        # 配置BOM表格的垂直滚动条
        self.bom_a_yscroll.config(command=lambda *args: [self.bom_a_tree.yview(*args), sync_a_to_b_yview(*args)])
        self.bom_b_yscroll.config(command=lambda *args: [self.bom_b_tree.yview(*args), sync_b_to_a_yview(*args)])

        # 处理Windows上的鼠标滚轮事件
        def on_mousewheel_win(event, tree):
            if event.delta > 0:
                tree.yview_scroll(-2, "units")  # 向上滚动
            else:
                tree.yview_scroll(2, "units")   # 向下滚动

        # 处理Linux/macOS上的鼠标滚轮事件
        def on_mousewheel_linux_up(_, tree):
            tree.yview_scroll(-2, "units")  # 向上滚动

        def on_mousewheel_linux_down(_, tree):
            tree.yview_scroll(2, "units")  # 向下滚动

        # Linux/macOS鼠标滚轮事件绑定（向下滚动）
        self.bom_a_tree.bind("<Button-5>", lambda event: [on_mousewheel_linux_down(event, self.bom_a_tree),
                                                        sync_a_to_b_yview("moveto", self.bom_a_tree.yview()[0])])
        self.bom_b_tree.bind("<Button-5>", lambda event: [on_mousewheel_linux_down(event, self.bom_b_tree),
                                                        sync_b_to_a_yview("moveto", self.bom_b_tree.yview()[0])])

        # Windows鼠标滚轮事件绑定
        self.bom_a_tree.bind("<MouseWheel>", lambda event: [on_mousewheel_win(event, self.bom_a_tree),
                                                          sync_a_to_b_yview("moveto", self.bom_a_tree.yview()[0])])
        self.bom_b_tree.bind("<MouseWheel>", lambda event: [on_mousewheel_win(event, self.bom_b_tree),
                                                          sync_b_to_a_yview("moveto", self.bom_b_tree.yview()[0])])

        # Linux/macOS鼠标滚轮事件绑定（向上滚动）
        self.bom_a_tree.bind("<Button-4>", lambda event: [on_mousewheel_linux_up(event, self.bom_a_tree),
                                                        sync_a_to_b_yview("moveto", self.bom_a_tree.yview()[0])])
        self.bom_b_tree.bind("<Button-4>", lambda event: [on_mousewheel_linux_up(event, self.bom_b_tree),
                                                        sync_b_to_a_yview("moveto", self.bom_b_tree.yview()[0])])

class FieldMappingDialog:
    """字段映射选择对话框，用于手动选择BOM文件中的字段映射"""

    def __init__(self, parent, title, message, candidates, field_type=None, default_selection=None):
        # default_selection参数用于设置默认选中的项目，当前未使用但保留以便未来扩展
        """
        初始化字段映射对话框

        Args:
            parent: 父窗口
            title: 对话框标题
            message: 提示信息
            candidates: 候选字段列表
            field_type: 字段类型，用于显示不同的提示信息
            default_selection: 默认选中的项
        """
        self.result = None
        self.candidates = candidates
        self.field_type = field_type

        # 创建顶层窗口
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.transient(parent)  # 设置为父窗口的临时窗口
        self.dialog.grab_set()  # 模态对话框
        self.dialog.focus_set()  # 获取焦点

        # 窗口大小和位置
        window_width = 450
        window_height = 450
        screen_width = parent.winfo_screenwidth()
        screen_height = parent.winfo_screenheight()
        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)
        self.dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.dialog.resizable(False, False)

        # 设置图标
        try:
            self.dialog.iconbitmap("icon.ico")
        except:
            pass

        # 提示信息标签
        message_frame = ttk.Frame(self.dialog, padding=10)
        message_frame.pack(fill="x")

        message_label = ttk.Label(message_frame, text=message, wraplength=430, justify="left")
        message_label.pack(fill="x")

        # 根据字段类型显示不同的提示信息
        hint_text = ""
        if field_type == "Reference":
            hint_text = "(位号列通常包含如 R1, C2, U3 等器件编号)"
        elif field_type == "P/N":
            hint_text = "(物料编号列通常包含物料的唯一标识符号)"
        elif field_type == "Description":
            hint_text = "(描述列通常包含物料的详细文字说明)"
        elif field_type == "MPN":
            hint_text = "(制造商料号列通常包含厂家的物料编号)"
        elif field_type == "Item":
            hint_text = "(序号列通常包含数字序列，如1, 2, 3...)"

        hint_label = ttk.Label(message_frame, text=hint_text,
                              foreground="gray", wraplength=430, justify="left")
        hint_label.pack(fill="x")

        # 搜索框
        search_frame = ttk.Frame(self.dialog, padding=(10, 5))
        search_frame.pack(fill="x")

        search_label = ttk.Label(search_frame, text="搜索:")
        search_label.pack(side="left")

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_candidates)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side="left", padx=(5, 0), fill="x", expand=True)

        # 清除按钮
        clear_button = ttk.Button(search_frame, text="X", width=3, command=self.clear_search)
        clear_button.pack(side="left", padx=(5, 0))

        # 创建列表框
        list_frame = ttk.Frame(self.dialog, padding=10)
        list_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(list_frame, selectmode="single", font=("Consolas", 10))
        self.listbox.pack(side="left", fill="both", expand=True)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # 填充候选项并标记可能匹配的项
        self.populate_listbox()

        # 记住设置复选框
        option_frame = ttk.Frame(self.dialog, padding=(10, 5))
        option_frame.pack(fill="x")

        self.remember_var = tk.BooleanVar(value=True)
        remember_checkbox = ttk.Checkbutton(option_frame, text="记住我的选择",
                                           variable=self.remember_var)
        remember_checkbox.pack(anchor="w")

        # 按钮框架
        button_frame = ttk.Frame(self.dialog, padding=10)
        button_frame.pack(fill="x")

        cancel_button = ttk.Button(button_frame, text="取消", command=self.on_cancel)
        cancel_button.pack(side="right", padx=(5, 0))

        ok_button = ttk.Button(button_frame, text="确定", command=self.on_ok)
        ok_button.pack(side="right")

        # 绑定事件
        self.listbox.bind("<Double-1>", lambda _: self.on_ok())
        self.dialog.bind("<Return>", lambda _: self.on_ok())
        self.dialog.bind("<Escape>", lambda _: self.on_cancel())
        search_entry.focus_set()  # 将焦点设置到搜索框

        # 使对话框居中
        self.dialog.update_idletasks()
        self.dialog.deiconify()

        # 等待窗口关闭
        parent.wait_window(self.dialog)

    def populate_listbox(self):
        """填充列表框并高亮可能匹配的项"""
        self.listbox.delete(0, tk.END)  # 清空列表框

        # 获取过滤后的候选项
        filtered_candidates = self.get_filtered_candidates()

        # 确定可能匹配的关键词
        keywords = []
        if self.field_type == "Reference":
            keywords = ["ref", "位号", "reference"]
        elif self.field_type == "P/N":
            keywords = ["p/n", "料号", "物料", "编码", "编号", "part"]
        elif self.field_type == "Description":
            keywords = ["desc", "描述", "specification"]
        elif self.field_type == "MPN":
            keywords = ["mpn", "制造商", "厂家", "manufacturer"]
        elif self.field_type == "Item":
            keywords = ["item", "序号", "number", "序列"]

        # 填充列表框
        for idx, candidate in enumerate(filtered_candidates):
            self.listbox.insert(tk.END, candidate)

            # 检查是否可能匹配
            is_match = False
            candidate_lower = candidate.lower()
            for keyword in keywords:
                if keyword in candidate_lower:
                    is_match = True
                    break

            # 设置背景色以高亮可能匹配的项
            if is_match:
                self.listbox.itemconfig(idx, background="#e6f3ff")

    def get_filtered_candidates(self):
        """根据搜索文本过滤候选项"""
        search_text = self.search_var.get().lower().strip()
        if not search_text:
            return self.candidates

        return [c for c in self.candidates if search_text in c.lower()]

    def filter_candidates(self, *_):
        """当搜索文本变化时过滤候选项"""
        self.populate_listbox()

    def clear_search(self):
        """清除搜索文本"""
        self.search_var.set("")
        self.populate_listbox()

    def on_ok(self):
        """确定按钮回调函数"""
        selection = self.listbox.curselection()
        if selection:
            self.result = {
                'field': self.listbox.get(selection[0]),
                'remember': self.remember_var.get()
            }
            self.dialog.destroy()
        else:
            messagebox.showwarning("提示", "请选择一个字段")

    def on_cancel(self):
        """取消按钮回调函数"""
        self.dialog.destroy()

# 更新检测相关函数
def check_for_updates(current_version):
    """
    检查GitHub上是否有新版本

    Args:
        current_version: 当前版本号

    Returns:
        tuple: (是否有更新, 最新版本, 下载链接, 更新日志, 是否为exe更新)
    """
    try:
        print(f"检查更新，当前版本: {current_version}")

        # 设置请求头，避免API限制
        headers = {
            "User-Agent": "BOM-Comparer-Update-Checker"
        }

        # 添加超时设置，避免长时间等待
        response = requests.get(GITHUB_API_URL, headers=headers, timeout=DOWNLOAD_TIMEOUT)

        if response.status_code == 200:
            data = response.json()
            latest_version = data["tag_name"].lstrip("v")
            print(f"发现版本: {latest_version}")

            # 使用packaging.version进行版本比较
            if pkg_version.parse(latest_version) > pkg_version.parse(current_version):
                print(f"发现新版本: {latest_version}")

                # 查找exe资源文件
                download_url = ""
                is_exe_update = False

                for asset in data.get("assets", []):
                    if asset["name"].endswith(".exe"):
                        download_url = asset["browser_download_url"]
                        is_exe_update = True
                        print(f"找到exe更新: {asset['name']}")
                        break

                # 如果没有资源文件，使用源代码下载链接
                if not download_url:
                    download_url = data["zipball_url"]
                    print("使用源代码链接作为备用")

                # 获取更新日志
                changelog = data["body"] if "body" in data else "无可用的更新日志"

                return True, latest_version, download_url, changelog, is_exe_update

        # 如果没有新版本或请求失败
        return False, current_version, "", "", False
    except Exception as e:
        print(f"检查更新失败: {str(e)}")
        return False, current_version, "", "", False

def download_with_resume(url, dest_file, progress_callback=None, status_callback=None):
    """
    支持断点续传的下载函数

    Args:
        url: 下载链接
        dest_file: 目标文件路径
        progress_callback: 进度回调函数，接收三个参数(已下载大小, 总大小, 进度百分比)
        status_callback: 状态回调函数，接收一个参数(状态消息)

    Returns:
        bool: 下载是否成功
    """
    try:
        # 设置请求头
        headers = {
            "User-Agent": "BOM-Comparer-Updater"
        }

        # 获取文件大小
        response = requests.head(url, headers=headers, timeout=DOWNLOAD_TIMEOUT)
        file_size = int(response.headers.get('content-length', 0))

        # 已下载的大小
        downloaded = 0

        # 检查是否存在部分下载的文件
        if os.path.exists(dest_file):
            downloaded = os.path.getsize(dest_file)

            # 如果文件已经下载完成，直接返回成功
            if downloaded == file_size:
                if status_callback:
                    status_callback("文件已存在，跳过下载")
                return True

            # 如果文件大小不匹配，设置断点续传的请求头
            if downloaded < file_size:
                headers['Range'] = f'bytes={downloaded}-'
                if status_callback:
                    status_callback(f"继续下载，已完成: {downloaded/file_size*100:.1f}%")
            else:
                # 文件大小超过预期，可能是损坏的，重新下载
                downloaded = 0
                if status_callback:
                    status_callback("文件可能损坏，重新下载")

        # 打开文件，如果是断点续传则追加，否则覆盖
        mode = 'ab' if downloaded > 0 else 'wb'

        # 重试计数器
        retries = 0

        while retries < DOWNLOAD_MAX_RETRIES:
            try:
                # 发起请求
                response = requests.get(url, headers=headers, stream=True, timeout=DOWNLOAD_TIMEOUT)

                # 检查响应状态
                if response.status_code not in [200, 206]:
                    raise Exception(f"下载失败，HTTP状态码: {response.status_code}")

                # 获取文件总大小
                if 'content-length' in response.headers:
                    file_size = int(response.headers['content-length']) + downloaded

                # 打开文件并写入
                with open(dest_file, mode) as f:
                    for chunk in response.iter_content(chunk_size=DOWNLOAD_CHUNK_SIZE):
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)

                            # 更新进度
                            if progress_callback and file_size > 0:
                                progress = downloaded / file_size
                                progress_callback(downloaded, file_size, progress)

                # 下载完成
                if status_callback:
                    status_callback("下载完成")
                return True

            except (requests.exceptions.RequestException, IOError) as e:
                retries += 1
                if status_callback:
                    status_callback(f"下载出错，正在重试 ({retries}/{DOWNLOAD_MAX_RETRIES}): {str(e)}")

                # 如果不是最后一次重试，等待一段时间再重试
                if retries < DOWNLOAD_MAX_RETRIES:
                    time.sleep(2 * retries)  # 指数退避

        # 超过最大重试次数
        if status_callback:
            status_callback("下载失败，超过最大重试次数")
        return False

    except Exception as e:
        if status_callback:
            status_callback(f"下载过程中发生错误: {str(e)}")
        print(f"下载错误: {str(e)}")
        return False

def show_update_notification(parent, current_version, latest_version, changelog, download_url, is_exe_update):
    """
    显示更新通知对话框

    Args:
        parent: 父窗口
        current_version: 当前版本
        latest_version: 最新版本
        changelog: 更新日志
        download_url: 下载链接
        is_exe_update: 是否为exe更新

    Returns:
        bool: 用户是否选择更新
    """
    # 获取屏幕尺寸
    screen_width = parent.winfo_screenwidth()
    screen_height = parent.winfo_screenheight()

    # 根据屏幕大小调整对话框尺寸
    if screen_width <= 1366 or screen_height <= 768:  # 小屏幕设备
        width = min(450, int(screen_width * 0.6))
        height = min(350, int(screen_height * 0.6))
    else:  # 大屏幕设备
        width = 500
        height = 400

    # 计算居中位置
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # 创建对话框
    dialog = tk.Toplevel(parent)
    dialog.title("发现新版本")
    dialog.transient(parent)  # 设置为父窗口的临时窗口

    # 先设置位置和大小，再显示窗口
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    dialog.resizable(False, False)
    dialog.withdraw()  # 先隐藏窗口

    # 设置为模态对话框
    dialog.grab_set()  # 模态对话框

    # 创建内容框架
    frame = ttk.Frame(dialog, padding=10)
    frame.pack(fill="both", expand=True)

    # 标题
    title_label = ttk.Label(frame, text=f"发现新版本: {latest_version}", font=("微软雅黑", 12, "bold"))
    title_label.pack(pady=(0, 10))

    # 版本信息
    version_frame = ttk.Frame(frame)
    version_frame.pack(fill="x", pady=5)

    ttk.Label(version_frame, text="当前版本:", width=10).pack(side="left")
    ttk.Label(version_frame, text=current_version).pack(side="left", padx=5)

    ttk.Label(version_frame, text="最新版本:", width=10).pack(side="left", padx=(10, 0))
    ttk.Label(version_frame, text=latest_version, foreground="#0066cc").pack(side="left", padx=5)

    # 更新日志
    ttk.Label(frame, text="更新内容:", anchor="w").pack(fill="x", pady=(10, 5))

    # 创建滚动文本框显示更新日志
    changelog_frame = ttk.Frame(frame)
    changelog_frame.pack(fill="both", expand=True, pady=5)

    changelog_text = scrolledtext.ScrolledText(changelog_frame, wrap="word", height=10)
    changelog_text.pack(fill="both", expand=True)
    changelog_text.insert("1.0", changelog)
    changelog_text.config(state="disabled")  # 设置为只读

    # 更新类型信息
    update_type = "可执行文件(.exe)" if is_exe_update else "源代码包(.zip)"
    ttk.Label(frame, text=f"更新类型: {update_type}", foreground="#666666").pack(anchor="w", pady=5)

    # 按钮区域
    button_frame = ttk.Frame(frame)
    button_frame.pack(fill="x", pady=(10, 0))

    # 用户选择结果
    result = [False]  # 使用列表存储结果，以便在回调中修改

    # 更新按钮
    update_button = ttk.Button(button_frame, text="立即更新",
                             command=lambda: [result.append(True), dialog.destroy()])
    update_button.pack(side="right", padx=5)

    # 取消按钮
    cancel_button = ttk.Button(button_frame, text="稍后提醒",
                             command=dialog.destroy)
    cancel_button.pack(side="right", padx=5)

    # 所有内容创建完成后再显示对话框
    dialog.update_idletasks()  # 确保所有内容已经布局完成
    dialog.deiconify()  # 显示对话框

    # 等待对话框关闭
    parent.wait_window(dialog)

    # 返回用户选择
    return len(result) > 1 and result[1]

def _get_updated_filename(original_filename, new_version):
    """
    智能处理文件名中的版本号

    Args:
        original_filename: 原始文件名
        new_version: 新版本号

    Returns:
        str: 更新后的文件名
    """
    # 常见的版本号模式：
    # 1. BOM_Comparer_v1.2.exe
    # 2. BOM_Comparer-v1.2.exe
    # 3. BOM_Comparer_1.2.exe
    # 4. BOM_Comparer-1.2.exe
    # 5. BOM_Comparer v1.2.exe
    # 6. BOM_Comparer 1.2.exe

    # 定义版本号模式
    version_patterns = [
        r'(_v)([0-9]+\.[0-9]+(?:\.[0-9]+)?)',  # BOM_Comparer_v1.2.exe
        r'(-v)([0-9]+\.[0-9]+(?:\.[0-9]+)?)',  # BOM_Comparer-v1.2.exe
        r'(_)([0-9]+\.[0-9]+(?:\.[0-9]+)?)',    # BOM_Comparer_1.2.exe
        r'(-)([0-9]+\.[0-9]+(?:\.[0-9]+)?)',    # BOM_Comparer-1.2.exe
        r'( v)([0-9]+\.[0-9]+(?:\.[0-9]+)?)',   # BOM_Comparer v1.2.exe
        r'( )([0-9]+\.[0-9]+(?:\.[0-9]+)?)'     # BOM_Comparer 1.2.exe
    ]

    # 尝试匹配每一种模式
    for pattern in version_patterns:
        match = re.search(pattern, original_filename)
        if match:
            # 找到版本号，替换为新版本
            prefix = match.group(1)  # 分隔符（_v, -v, _, -, 等）
            return re.sub(pattern, f"{prefix}{new_version}", original_filename)

    # 如果没有找到版本号模式，返回原始文件名
    return original_filename

def main():
    """主函数"""
    try:
        # 创建主窗口
        root = tk.Tk()

        # 设置应用图标
        try:
            root.iconbitmap("icon.ico")
        except:
            pass  # 图标不存在时忽略

        # 获取屏幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 计算居中位置 - 考虑任务栏高度
        taskbar_height = 40  # 估计任务栏高度
        available_height = screen_height - taskbar_height

        # 设置默认窗口尺寸
        default_width = 1200  # 默认宽度
        default_height = 900  # 默认高度

        # 如果窗口高度大于可用高度，调整窗口高度
        if default_height > available_height:
            default_height = available_height

        # 计算位置 - 在可用区域内居中
        x = (screen_width - default_width) // 2
        y = (available_height - default_height) // 2

        # 确保坐标不为负
        x = max(0, x)
        y = max(0, y)

        # 设置初始窗口尺寸和位置
        root.geometry(f"{default_width}x{default_height}+{x}+{y}")

        # 打印调试信息
        print(f"屏幕尺寸: {screen_width}x{screen_height}")
        print(f"任务栏高度: {taskbar_height}")
        print(f"可用高度: {available_height}")
        print(f"窗口尺寸: {default_width}x{default_height}")
        print(f"窗口位置: +{x}+{y}")

        # 创建应用
        app = BOMComparerGUI(root)

        # 启动后自动检查更新
        threading.Thread(target=app.check_updates_on_startup, daemon=True).start()

        root.mainloop()
    except Exception as e:
        import traceback
        error_msg = f"发生错误: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)

        # 尝试写入错误日志
        try:
            with open("error_log.txt", "w", encoding="utf-8") as f:
                f.write(error_msg)
        except:
            pass

if __name__ == "__main__":
    main()