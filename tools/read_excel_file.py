"""
Read Excel File Tool - Excel文件读取工具
提供Excel文件读取和处理功能
"""

import logging
import os
from typing import Annotated, Optional, Dict, Any
from pydantic import Field
import pandas as pd

# 配置日志
logger = logging.getLogger("excel-mcp-server")

def read_excel_file(
    file_path: Annotated[str, Field(
        description="Excel 文件的完整路径（支持 .xlsx 和 .xls 格式）"
    )],
    sheet_name: Annotated[Optional[str], Field(
        default=None,
        description="要读取的工作表名称。如果不指定，将读取第一个工作表"
    )] = None,
    start_row: Annotated[int, Field(
        default=0,
        description="开始读取的行号（从0开始计数），默认为0"
    )] = 0,
    end_row: Annotated[Optional[int], Field(
        default=20,
        description="结束读取的行号（不包含该行），默认为20。如果为None则读取到文件末尾，但最多不超过start_row+100行"
    )] = 20
) -> Dict[str, Any]:
    """读取指定的 Excel 文件并返回结构化数据"""
    logger.info(f"[Tool] read_excel_file called with file_path: {file_path}, sheet_name: {sheet_name}, start_row: {start_row}, end_row: {end_row}")
    
    try:
        # 验证 start_row 参数
        if start_row < 0:
            error_msg = f"start_row 参数不能小于 0，当前值: {start_row}"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        # 处理 end_row 参数
        if end_row is None:
            # 如果 end_row 为 None，设置为 start_row + 100（最大读取100行）
            end_row = start_row + 100
        elif end_row <= start_row:
            error_msg = f"end_row ({end_row}) 必须大于 start_row ({start_row})"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        # 验证读取行数不超过100行
        rows_to_read = end_row - start_row
        if rows_to_read > 100:
            error_msg = f"读取行数不能超过 100 行，当前要读取 {rows_to_read} 行 (从第 {start_row} 行到第 {end_row-1} 行)"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        # 验证文件路径
        if not os.path.exists(file_path):
            error_msg = f"文件不存在: {file_path}"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        # 验证文件扩展名
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.xlsx', '.xls']:
            error_msg = f"不支持的文件格式: {file_ext}。仅支持 .xlsx 和 .xls 文件"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        logger.info(f"[API] Reading Excel file: {file_path} from row {start_row} to row {end_row-1}")
        
        # 读取 Excel 文件
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=start_row, nrows=rows_to_read)
            logger.info(f"[API] Successfully read sheet '{sheet_name}' from {file_path}, rows {start_row}-{end_row-1}")
        else:
            df = pd.read_excel(file_path, skiprows=start_row, nrows=rows_to_read)
            logger.info(f"[API] Successfully read first sheet from {file_path}, rows {start_row}-{end_row-1}")
        
        # 获取工作表信息
        excel_file = pd.ExcelFile(file_path)
        available_sheets = excel_file.sheet_names
        current_sheet = sheet_name if sheet_name else available_sheets[0]
        
        # 转换数据为 JSON 格式
        data_dict = df.to_dict('records')
        
        # 处理 NaN 值
        for record in data_dict:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = None
        
        result = {
            "success": True,
            "error": None,
            "data": {
                "file_path": file_path,
                "current_sheet": current_sheet,
                "available_sheets": available_sheets,
                "start_row": start_row,
                "end_row": end_row,
                "rows_read": len(df),
                "total_columns": len(df.columns),
                "columns": df.columns.tolist(),
                "records": data_dict,
                "max_rows_limit": 100,
                "truncated": len(df) == rows_to_read
            }
        }
        
        logger.info(f"[Tool] Successfully processed Excel file. Rows: {len(df)}, Columns: {len(df.columns)}")
        return result
        
    except FileNotFoundError:
        error_msg = f"文件未找到: {file_path}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
    except PermissionError:
        error_msg = f"没有权限访问文件: {file_path}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
    except ValueError as e:
        error_msg = f"工作表读取错误: {str(e)}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
    except pd.errors.EmptyDataError:
        error_msg = f"指定的行范围 ({start_row}-{end_row-1}) 超出了文件的数据范围，或该范围内没有数据"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
    except Exception as e:
        error_msg = f"读取 Excel 文件时发生未知错误: {str(e)}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
