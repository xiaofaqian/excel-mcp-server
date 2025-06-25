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
    max_rows: Annotated[int, Field(
        default=1000,
        description="最大读取行数，防止大文件导致性能问题。默认 1000 行"
    )] = 1000
) -> Dict[str, Any]:
    """读取指定的 Excel 文件并返回结构化数据"""
    logger.info(f"[Tool] read_excel_file called with file_path: {file_path}, sheet_name: {sheet_name}, max_rows: {max_rows}")
    
    try:
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
        
        logger.info(f"[API] Reading Excel file: {file_path}")
        
        # 读取 Excel 文件
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=max_rows)
            logger.info(f"[API] Successfully read sheet '{sheet_name}' from {file_path}")
        else:
            df = pd.read_excel(file_path, nrows=max_rows)
            logger.info(f"[API] Successfully read first sheet from {file_path}")
        
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
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": df.columns.tolist(),
                "records": data_dict,
                "max_rows_limit": max_rows,
                "truncated": len(df) == max_rows
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
    except Exception as e:
        error_msg = f"读取 Excel 文件时发生未知错误: {str(e)}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
