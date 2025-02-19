"""
表格图片生成器模块
用于将数据转换为格式化的表格图片，支持多种样式和配置选项
"""

from PIL import Image, ImageDraw, ImageFont
from typing import Dict, List, Union, Tuple, Optional
import json
import os
import time
import random
import tempfile
from datetime import datetime

# 修改 PIL 的最大图片尺寸限制
Image.MAX_IMAGE_PIXELS = None

class Cell:
    """
    表格单元格类
    
    Attributes:
        text: 单元格文本内容
        rowspan: 单元格跨行数
        colspan: 单元格跨列数
    """
    def __init__(self, text: str, rowspan: int = 1, colspan: int = 1):
        self.text = text
        self.rowspan = rowspan
        self.colspan = colspan

class TableImageGenerator:
    """
    表格图片生成器类
    
    支持生成包含以下特性的表格图片：
    - 多级表头
    - 单元格合并
    - 自定义样式
    - 条件格式
    - 数据格式化
    
    Attributes:
        font_path: 字体文件路径配置
        cell_width: 单元格宽度
        cell_height: 单元格高度
        font_size: 字体大小
        padding: 单元格内边距
        styles: 表格样式配置
        color_mapping: 状态颜色映射
        fonts: 字体对象
    """
    
    def __init__(self, font_path=None):
        """
        初始化表格图片生成器
        
        Args:
            font_path: 字体路径配置字典，包含 'regular' 和 'bold' 两种字体路径
                      默认使用系统字体
        """
        # 字体配置
        self.font_path = font_path or {
            'regular': os.path.join(".", "font", "PingFangSC-Regular.otf"),
            'bold': os.path.join(".", "font", "PingFangSC-Semibold.otf")
        }
        
        # 基础参数
        self.cell_width = None  # 将在 _calculate_table_size 中动态计算
        self.cell_height = 60   # 单元格高度
        self.font_size = 24     # 字体大小
        self.padding = 20       # 内边距
        
        # 样式配置
        self.styles = {
            'header_color': '#94A3B8',              # 表头背景色
            'header_text_color': '#FFFFFF',         # 表头文字颜色
            'row_colors': ['#FFFFFF', '#F9FAFB'],   # 行交替颜色
            'border_color': '#E5E7EB',              # 边框颜色
            'summary_color': '#E2E8F0',             # 汇总行颜色
            'text_color': '#111827',                # 文字颜色
            'empty_text_color': '#9CA3AF'           # 空值文字颜色
        }
        
        # 状态颜色映射
        self.color_mapping = {
            '绿灯': '#059669',  # 绿色
            '红灯': '#DC2626',  # 红色
            '黄灯': '#D97706'   # 黄色
        }
        
        # 初始化字体
        try:
            self.fonts = {
                'regular': ImageFont.truetype(self.font_path['regular'], self.font_size),
                'bold': ImageFont.truetype(self.font_path['bold'], self.font_size)
            }
        except:
            self.fonts = {
                'regular': ImageFont.load_default(),
                'bold': ImageFont.load_default()
            }

    def _hex_to_rgb(self, hex_color: str) -> tuple:
        """
        将十六进制颜色转换为RGB元组
        
        Args:
            hex_color: 十六进制颜色代码，如 '#FFFFFF'
            
        Returns:
            RGB颜色元组 (r, g, b)
        """
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def _process_value(self, value: str, replace_zero: bool = False, format_type: str = None) -> str:
        """
        处理单元格值，包括零值、空值和格式化处理
        
        Args:
            value: 原始值
            replace_zero: 是否将0替换为'-'
            format_type: 格式化类型，支持 'to_format'(完整时间) 和 'to_day'(月日)
            
        Returns:
            处理后的字符串值
        """
        if value is None or value == '':
            return '-'
        
        # 根据不同的格式化类型处理
        if format_type:
            try:
                timestamp = int(value)
                if format_type == 'to_format':
                    return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
                elif format_type == 'to_day':
                    return datetime.fromtimestamp(timestamp).strftime('%m月%d日')
            except (ValueError, TypeError):
                return str(value)
        
        # 处理零值
        if replace_zero:
            try:
                if float(value) == 0:
                    return '-'
            except (ValueError, TypeError):
                pass
        
        return str(value)

    def _calculate_table_size(self, headers: List[List[Optional[Cell]]], data: List[List[str]]) -> Tuple[int, int]:
        """计算表格的总宽度和高度"""
        # 计算表头的实际列数
        header_cols = 0
        for row in headers:
            curr_cols = 0
            for cell in row:
                if cell is not None:
                    curr_cols += cell.colspan
            header_cols = max(header_cols, curr_cols)
        
        # 取表头列数和数据列数的最大值
        total_cols = max(header_cols, max(len(row) for row in data))
        
        # 计算目标总宽度（考虑边距）
        target_total_width = 1920 - 2 * 40  # 屏幕宽度减去左右边距
        
        # 动态计算单元格宽度
        MIN_CELL_WIDTH = 120  # 最小单元格宽度
        MAX_CELL_WIDTH = 240  # 最大单元格宽度
        
        # 计算建议的单元格宽度
        suggested_width = target_total_width // total_cols
        
        # 根据列数调整单元格宽度
        if suggested_width > MAX_CELL_WIDTH:
            self.cell_width = MAX_CELL_WIDTH
        elif suggested_width < MIN_CELL_WIDTH:
            self.cell_width = MIN_CELL_WIDTH
        else:
            self.cell_width = suggested_width
        
        # 计算总宽度和高度
        total_width = total_cols * self.cell_width + 1
        total_height = (len(headers) + len(data)) * self.cell_height + 1
        
        return total_width, total_height

    def _draw_cell(self, draw: ImageDraw, x: int, y: int, cell: Union[Cell, str], 
                  is_header: bool = False, row_idx: int = 0, 
                  color_column: str = '', column_name: str = '',
                  replace_zero: bool = False, highlight: bool = False):
        """绘制单个单元格"""
        if isinstance(cell, str):
            cell = Cell(cell)
            
        # 计算合并后的单元格大小
        width = self.cell_width * cell.colspan
        height = self.cell_height * cell.rowspan
        
        # 处理单元格值
        cell.text = self._process_value(cell.text, replace_zero=replace_zero)
        
        # 设置单元格背景色和边框颜色
        if is_header:
            bg_color = self._hex_to_rgb(self.styles['header_color'])
            text_color = self._hex_to_rgb(self.styles['header_text_color'])
            font = self.fonts['bold']
        else:
            if highlight:
                bg_color = self._hex_to_rgb(self.styles['summary_color'])
            else:
                bg_color = self._hex_to_rgb(self.styles['row_colors'][row_idx % 2])
            
            # 设置文字颜色
            if cell.text == '-':
                text_color = self._hex_to_rgb(self.styles['empty_text_color'])
            elif column_name == color_column and cell.text in self.color_mapping:
                text_color = self._hex_to_rgb(self.color_mapping[cell.text])
            else:
                text_color = self._hex_to_rgb(self.styles['text_color'])
            
            font = self.fonts['regular']
        
        border_color = self._hex_to_rgb(self.styles['border_color'])
        
        # 绘制单元格背景
        draw.rectangle(
            [(x, y), (x + width, y + height)],
            fill=bg_color
        )
        
        # 绘制加粗的边框
        for i in range(2):  # 画2次边框来加粗
            draw.rectangle(
                [(x+i, y+i), (x + width-i, y + height-i)],
                outline=border_color
            )
        
        # 计算文本位置使其居中
        bbox = draw.textbbox((0, 0), cell.text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        # 使用负的偏移量使文本稍微上移
        vertical_offset = height * -0.05  # -5% 的单元格高度作为偏移
        text_y = y + (height - text_height) / 2 + vertical_offset
        text_x = x + (width - text_width) / 2
        
        # 绘制文本
        draw.text((text_x, text_y), cell.text, fill=text_color, font=font)

    def create_table_image(
        self, data, output_file, columns_order=None, banner_path=None, 
        banner_text=None, color_column='', multi_columns=None, 
        column_display=None, replace_zero=False, 
        highlight_rules={}
    ):
        """
        创建表格图片
        Args:
            data: 表格数据
            output_file: 输出文件路径
            columns_order: 列顺序
            banner_path: banner图片路径
            banner_text: banner文字
            color_column: 需要着色的列名
            multi_columns: 多级表头配置
            column_display: 列显示名称映射
            replace_zero: 替换占位（将数字 0 替换为 -）
            highlight_rules: 高亮规则，格式为 {'列名': '关键字'} 
        Returns:
            str: 生成的图片路径
        """
        # try:
        # 调整 DPI 到更合理的值
        dpi = 150  # 从300降到150，减小图片尺寸
        
        # 创建临时文件路径
        temp_table = os.path.join(tempfile.gettempdir(), 'temp_table.png')
        temp_banner = os.path.join(tempfile.gettempdir(), 'temp_banner.png')
        
        # 使用时间戳作为文件名
        random_value = random.randint(1000, 9999)
        file_path = os.path.join(output_file, f"{int(time.time())}_{random_value}_table.png")
        
        # 构建表格数据结构
        table_data = self._build_table_data(data, columns_order, multi_columns, column_display, replace_zero)
        
        # 生成表格图片
        table_image = self._create_table(
            json.dumps(table_data),
            color_column=color_column,
            replace_zero=replace_zero,
            highlight_rules=highlight_rules,
            dpi=dpi  # 使用较小的DPI
        )
        table_image.save(temp_table)
        
        # 处理banner
        if banner_path and os.path.exists(banner_path):
            self._create_banner_image(banner_path, banner_text, temp_banner)
            self._merge_images(temp_banner, temp_table, file_path)
        else:
            table_image.save(file_path)
        
        return file_path
            
        # except Exception as e:
        #     print(f"Error: {str(e)}")
        #     return False
            
        # finally:
        #     # 清理临时文件
        #     for temp_file in [temp_table, temp_banner]:
        #         if os.path.exists(temp_file):
        #             os.remove(temp_file)

    def _build_table_data(self, data, columns_order, multi_columns, column_display, replace_zero=False):
        """构建表格数据结构"""
        # 检查哪些列有数据
        valid_columns = set()
        if columns_order:
            for row in data:
                for col in columns_order:
                    if row.get(col) not in [None, '', '-', 0, '0', 0.0, '0.0']:
                        valid_columns.add(col)
        
        # 过滤掉无数据的列
        filtered_columns = [col for col in columns_order if col in valid_columns] if columns_order else None
        
        # 构建表头
        headers = []
        if multi_columns:
            current_position = 0
            for layer in multi_columns:
                header_row = []
                current_col = 0
                for header, span in layer.items():
                    # 计算当前表头下有多少有效列
                    valid_span = 0
                    for i in range(span):
                        if current_position + i < len(filtered_columns):
                            valid_span += 1
                    
                    if valid_span > 0:  # 只添加有效数据的表头
                        # 添加当前表头
                        header_row.append({
                            "text": header,
                            "colspan": valid_span,
                            "rowspan": 1
                        })
                        
                        # 为合并的列添加 null
                        for _ in range(valid_span - 1):
                            header_row.append(None)
                        
                        current_col += valid_span
                    
                    current_position += span
                
                if header_row:  # 只添加非空的行
                    headers.append(header_row)
            
            # 如果有 column_display，添加最后一行表头
            if column_display and filtered_columns:
                last_row = []
                for col in filtered_columns:
                    # 只有非格式化类型的才替换显示名称
                    display_value = column_display.get(col)
                    display_name = col if display_value in ['to_day', 'to_format'] else column_display.get(col, col)
                    last_row.append({
                        "text": display_name,
                        "colspan": 1,
                        "rowspan": 1
                    })
                headers.append(last_row)
        else:
            # 添加单层表头
            header_row = []
            for col in filtered_columns:
                # 只有非格式化类型的才替换显示名称
                display_value = column_display.get(col)
                display_name = col if display_value in ['to_day', 'to_format'] else column_display.get(col, col)
                header_row.append({
                    "text": display_name,
                    "colspan": 1,
                    "rowspan": 1
                })
            headers.append(header_row)
        
        # 构建数据行
        rows = []
        for row in data:
            if filtered_columns:
                row_data = []
                for col in filtered_columns:
                    value = row.get(col, '')
                    # 获取格式化类型
                    format_type = column_display.get(col) if column_display else None
                    processed_value = self._process_value(value, replace_zero, format_type)
                    row_data.append(processed_value)
            else:
                row_data = [str(val) for val in row.values()]
            rows.append(row_data)
        
        return {
            "headers": headers,
            "data": rows
        }

    def _create_table(self, json_data: str, color_column: str = '', 
                    replace_zero: bool = False, highlight_rules: Dict = None,
                    dpi: int = 300) -> Image:
        """
        从JSON数据创建表格图片
        Args:
            json_data: JSON格式的表格数据
            color_column: 需要应用颜色映射的列名
            replace_zero: 是否将0替换为-
            highlight_rules: 高亮规则，格式为 {'列名': '关键字'}
            dpi: 图片分辨率，默认300
        """
        data = json.loads(json_data)
        
        # 转换headers数据为Cell对象
        raw_headers = data.get("headers", [])
        headers = []
        for row in raw_headers:
            header_row = []
            for cell in row:
                if isinstance(cell, dict):
                    header_row.append(Cell(
                        cell.get("text", ""),
                        cell.get("rowspan", 1),
                        cell.get("colspan", 1)
                    ))
                elif isinstance(cell, str):
                    header_row.append(Cell(cell) if cell else None)
                else:
                    header_row.append(None)
            headers.append(header_row)
            
        rows = data.get("data", [])
        
        # 计算表格尺寸
        base_width, base_height = self._calculate_table_size(headers, rows)
        
        # 根据DPI调整实际图片大小
        scale_factor = dpi / 72  # 72是基准DPI
        width = int(base_width * scale_factor)
        height = int(base_height * scale_factor)
        
        # 创建高分辨率空白图片
        image = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(image)
        
        # 调整字体大小和单元格大小
        original_cell_width = self.cell_width
        original_cell_height = self.cell_height
        original_font_size = self.font_size
        
        self.cell_width = int(self.cell_width * scale_factor)
        self.cell_height = int(self.cell_height * scale_factor)
        self.font_size = int(self.font_size * scale_factor)
        
        # 重新加载更大尺寸的字体
        try:
            self.fonts = {
                'regular': ImageFont.truetype(self.font_path['regular'], self.font_size),
                'bold': ImageFont.truetype(self.font_path['bold'], self.font_size)
            }
        except:
            self.fonts = {
                'regular': ImageFont.load_default(),
                'bold': ImageFont.load_default()
            }
        
        # 创建已绘制单元格的跟踪矩阵
        drawn_cells = [[False] * (width // self.cell_width) for _ in range(len(headers))]
        
        # 绘制表头
        for row_idx, header_row in enumerate(headers):
            for col_idx, cell in enumerate(header_row):
                if cell is None or drawn_cells[row_idx][col_idx]:
                    continue
                
                # 标记合并范围
                for r in range(cell.rowspan):
                    for c in range(cell.colspan):
                        if row_idx + r < len(drawn_cells) and col_idx + c < len(drawn_cells[0]):
                            drawn_cells[row_idx + r][col_idx + c] = True
                
                self._draw_cell(
                    draw,
                    col_idx * self.cell_width,
                    row_idx * self.cell_height,
                    cell,
                    is_header=True
                )
        
        # 绘制数据行
        highlight_rules = highlight_rules or {}
        for row_idx, row in enumerate(rows):
            # 检查是否需要高亮
            should_highlight = False
            for col_name, keyword in highlight_rules.items():
                try:
                    col_idx = next(i for i, cell in enumerate(headers[-1]) 
                                 if cell and cell.text == col_name)
                    if str(keyword) in str(row[col_idx]):
                        should_highlight = True
                        break
                except (StopIteration, IndexError):
                    continue
            
            y = len(headers) * self.cell_height
            for col_idx, cell_text in enumerate(row):
                # 获取列名
                col_name = next((cell.text for cell in headers[-1] 
                               if cell and headers[-1].index(cell) == col_idx), '')
                
                self._draw_cell(
                    draw,
                    col_idx * self.cell_width,
                    y + row_idx * self.cell_height,
                    cell_text,
                    is_header=False,
                    row_idx=row_idx,
                    color_column=color_column,
                    column_name=col_name,
                    replace_zero=replace_zero,
                    highlight=should_highlight
                )
        
        # 恢复原始尺寸设置
        self.cell_width = original_cell_width
        self.cell_height = original_cell_height
        self.font_size = original_font_size
        try:
            self.fonts = {
                'regular': ImageFont.truetype(self.font_path['regular'], self.font_size),
                'bold': ImageFont.truetype(self.font_path['bold'], self.font_size)
            }
        except:
            self.fonts = {
                'regular': ImageFont.load_default(),
                'bold': ImageFont.load_default()
            }
        
        return image

    def _create_banner_image(self, banner_path: str, banner_text: str, output_path: str):
        """
        创建包含banner和文字的图片
        Args:
            banner_path: banner图片路径
            banner_text: banner文字
            output_path: 输出路径
        """
        # 打开banner图片
        banner = Image.open(banner_path)
        
        # 固定宽度和边距
        target_width = 1920  # 固定宽度
        margin_top = 20
        margin_sides = 20
        text_height = 60 if banner_text else 0
        
        # 调整banner大小
        banner_width = target_width - 2 * margin_sides
        banner_height = int(banner.height * banner_width / banner.width)
        banner_resized = banner.resize((banner_width, banner_height))
        
        # 创建新图片
        total_height = margin_top + banner_height + text_height + margin_top
        combined = Image.new('RGB', (target_width, total_height), 'white')
        
        # 粘贴banner
        combined.paste(banner_resized, (margin_sides, margin_top))
        
        # 添加文字
        if banner_text:
            try:
                # 固定字体大小
                font_size = 30
                font = ImageFont.truetype(self.font_path['bold'], font_size)
            except:
                font = ImageFont.load_default()
            
            draw = ImageDraw.Draw(combined)
            
            # 固定文字位置
            text_y = margin_top + banner_height + 20
            
            # 计算文字位置使其居中
            bbox = draw.textbbox((0, 0), banner_text, font=font)
            text_width = bbox[2] - bbox[0]
            text_x = (target_width - text_width) // 2
            
            # 绘制文字
            draw.text((text_x, text_y), banner_text, font=font, fill=self.styles['text_color'])
        
        # 保存图片
        combined.save(output_path)

    def _merge_images(self, banner_image: str, table_image: str, output_path: str):
        """
        合并banner图片和表格图片
        Args:
            banner_image: banner图片路径
            table_image: 表格图片路径
            output_path: 输出路径
        """
        banner_img = Image.open(banner_image)
        table_img = Image.open(table_image)
        
        # 设置边距
        margin_sides = 40   # 左右边距
        margin_bottom = 40  # 底部边距
        
        # 调整表格图片宽度以匹配banner的内容区域
        target_width = banner_img.width - 2 * margin_sides
        if table_img.width != target_width:
            scale_factor = target_width / table_img.width
            new_height = int(table_img.height * scale_factor)
            table_img = table_img.resize((target_width, new_height), Image.Resampling.LANCZOS)
        
        # 创建新图片（包含边距）
        combined = Image.new(
            'RGB', 
            (banner_img.width, banner_img.height + table_img.height + margin_bottom), 
            'white'
        )
        
        # 粘贴banner（banner已经在_create_banner_image中处理了边距）
        combined.paste(banner_img, (0, 0))
        
        # 粘贴表格（添加左右边距）
        combined.paste(table_img, (margin_sides, banner_img.height))
        
        # 保存最终图片时使用最高质量设置
        combined.save(output_path, 'PNG', quality=100, optimize=False)