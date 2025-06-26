"""
Get Excel Summary Tool - Excel文件概览工具
提供Excel文件的简要信息，包括工作表统计和数据预览功能
"""

import logging
import os
from typing import Annotated, Optional, Dict, Any, List
from pydantic import Field
import pandas as pd

# 配置日志
logger = logging.getLogger("excel-mcp-server")

def get_excel_summary(
    file_path: Annotated[str, Field(
        description="Excel 文件的完整路径（支持 .xlsx 和 .xls 格式）"
    )],
    target_sheet: Annotated[Optional[str], Field(
        default=None,
        description="要预览数据的工作表名称。如果不指定，将预览第一个工作表"
    )] = None,
    preview_rows: Annotated[int, Field(
        default=10,
        description="预览数据的行数，默认 10 行，最多不能超过 20 行"
    )] = 10
) -> Dict[str, Any]:
    """获取Excel文件的简要信息，包括工作表统计和前几行数据预览"""
    logger.info(f"[Tool] get_excel_summary called with file_path: {file_path}, target_sheet: {target_sheet}, preview_rows: {preview_rows}")
    
    try:
        # 验证 preview_rows 参数
        if preview_rows > 20:
            error_msg = f"preview_rows 参数不能超过 20，当前值: {preview_rows}"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        if preview_rows <= 0:
            error_msg = f"preview_rows 参数必须大于 0，当前值: {preview_rows}"
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
        
        logger.info(f"[API] Analyzing Excel file: {file_path}")
        
        # 获取Excel文件信息
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        total_sheets = len(sheet_names)
        
        logger.info(f"[API] Found {total_sheets} sheets in Excel file")
        
        # 统计每个工作表的信息
        sheet_details = []
        total_rows_all_sheets = 0
        
        for sheet_name in sheet_names:
            try:
                # 只读取第一行来获取列数，避免读取大量数据
                df_info = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
                total_columns = len(df_info.columns)
                
                # 获取实际行数（不包括表头）
                df_full = pd.read_excel(file_path, sheet_name=sheet_name)
                total_rows = len(df_full)
                total_rows_all_sheets += total_rows
                
                sheet_details.append({
                    "sheet_name": sheet_name,
                    "total_rows": total_rows,
                    "total_columns": total_columns
                })
                
                logger.info(f"[API] Sheet '{sheet_name}': {total_rows} rows, {total_columns} columns")
                
            except Exception as e:
                logger.warning(f"[Warning] Could not analyze sheet '{sheet_name}': {str(e)}")
                sheet_details.append({
                    "sheet_name": sheet_name,
                    "total_rows": 0,
                    "total_columns": 0,
                    "error": f"无法分析工作表: {str(e)}"
                })
        
        # 确定要预览的工作表
        preview_sheet = target_sheet if target_sheet and target_sheet in sheet_names else sheet_names[0]
        
        # 验证目标工作表是否存在
        if target_sheet and target_sheet not in sheet_names:
            logger.warning(f"[Warning] Target sheet '{target_sheet}' not found, using first sheet '{sheet_names[0]}'")
        
        # 读取预览数据
        preview_data = None
        try:
            df_preview = pd.read_excel(file_path, sheet_name=preview_sheet, nrows=preview_rows)
            
            # 转换数据为 JSON 格式
            data_records = df_preview.to_dict('records')
            
            # 处理 NaN 值
            for record in data_records:
                for key, value in record.items():
                    if pd.isna(value):
                        record[key] = None
            
            preview_data = {
                "sheet_name": preview_sheet,
                "columns": df_preview.columns.tolist(),
                "preview_rows": len(df_preview),
                "data": data_records
            }
            
            logger.info(f"[API] Successfully generated preview for sheet '{preview_sheet}' with {len(df_preview)} rows")
            
        except Exception as e:
            logger.error(f"[Error] Failed to generate preview for sheet '{preview_sheet}': {str(e)}")
            preview_data = {
                "sheet_name": preview_sheet,
                "columns": [],
                "preview_rows": 0,
                "data": [],
                "error": f"无法预览数据: {str(e)}"
            }
        
        # 构建返回结果
        result = {
            "success": True,
            "error": None,
            "data": {
                "file_path": file_path,
                "file_summary": {
                    "total_sheets": total_sheets,
                    "total_rows_all_sheets": total_rows_all_sheets,
                    "sheet_details": sheet_details
                },
                "preview_data": preview_data,
                "settings": {
                    "requested_preview_rows": preview_rows,
                    "target_sheet": target_sheet,
                    "actual_preview_sheet": preview_sheet
                }
            }
        }
        
        logger.info(f"[Tool] Successfully generated Excel summary. Total sheets: {total_sheets}, Total rows: {total_rows_all_sheets}")
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
        error_msg = f"Excel文件格式错误: {str(e)}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
    except Exception as e:
        error_msg = f"分析 Excel 文件时发生未知错误: {str(e)}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
