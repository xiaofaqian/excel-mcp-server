"""
Search Excel Data Tool - Excel数据搜索工具
提供Excel文件中数据的搜索和过滤功能
"""

import logging
import os
from typing import Annotated, Optional, Dict, Any, Union
from pydantic import Field
import pandas as pd

# 配置日志
logger = logging.getLogger("excel-mcp-server")

def search_excel_data(
    file_path: Annotated[str, Field(
        description="Excel 文件的完整路径（支持 .xlsx 和 .xls 格式）"
    )],
    column_name: Annotated[str, Field(
        description="要搜索的列名"
    )],
    search_value: Annotated[Union[str, int, float], Field(
        description="要搜索的值（支持文本、数字等）"
    )],
    sheet_name: Annotated[Optional[str], Field(
        default=None,
        description="要搜索的工作表名称。如果不指定，将搜索第一个工作表"
    )] = None,
    match_type: Annotated[str, Field(
        default="exact",
        description="匹配类型：'exact'(精确匹配) 或 'contains'(包含匹配)"
    )] = "exact",
    max_results: Annotated[int, Field(
        default=50,
        description="最大返回结果数，默认 50 行，最多不能超过 100 行"
    )] = 50
) -> Dict[str, Any]:
    """在Excel文件中搜索指定列的数据并返回匹配的行"""
    logger.info(f"[Tool] search_excel_data called with file_path: {file_path}, column_name: {column_name}, search_value: {search_value}, match_type: {match_type}")
    
    try:
        # 验证 max_results 参数
        if max_results > 100:
            error_msg = f"max_results 参数不能超过 100，当前值: {max_results}"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        if max_results <= 0:
            error_msg = f"max_results 参数必须大于 0，当前值: {max_results}"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        # 验证 match_type 参数
        if match_type not in ["exact", "contains"]:
            error_msg = f"match_type 参数必须是 'exact' 或 'contains'，当前值: {match_type}"
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
        
        logger.info(f"[API] Searching Excel file: {file_path}")
        
        # 读取 Excel 文件
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            logger.info(f"[API] Successfully read sheet '{sheet_name}' from {file_path}")
        else:
            df = pd.read_excel(file_path)
            excel_file = pd.ExcelFile(file_path)
            sheet_name = excel_file.sheet_names[0]
            logger.info(f"[API] Successfully read first sheet '{sheet_name}' from {file_path}")
        
        # 验证列名是否存在
        if column_name not in df.columns:
            available_columns = df.columns.tolist()
            error_msg = f"列名 '{column_name}' 不存在。可用的列名: {available_columns}"
            logger.error(f"[Error] {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "data": None
            }
        
        logger.info(f"[API] Searching in column '{column_name}' for value '{search_value}' with match_type '{match_type}'")
        
        # 执行搜索
        if match_type == "exact":
            # 精确匹配
            mask = df[column_name] == search_value
        else:  # contains
            # 包含匹配（仅对字符串类型有效）
            if isinstance(search_value, str):
                # 将列转换为字符串进行搜索，处理NaN值
                mask = df[column_name].astype(str).str.contains(str(search_value), case=False, na=False)
            else:
                # 对于非字符串类型，使用精确匹配
                mask = df[column_name] == search_value
                logger.warning(f"[Warning] Contains match requested for non-string value '{search_value}', using exact match instead")
        
        # 获取匹配的行
        matched_df = df[mask]
        total_matches = len(matched_df)
        
        logger.info(f"[API] Found {total_matches} matches")
        
        # 限制返回结果数量
        if total_matches > max_results:
            matched_df = matched_df.head(max_results)
            logger.info(f"[API] Limited results to {max_results} rows")
        
        # 转换数据为 JSON 格式
        matched_records = matched_df.to_dict('records')
        
        # 处理 NaN 值
        for record in matched_records:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = None
        
        # 构建返回结果
        result = {
            "success": True,
            "error": None,
            "data": {
                "file_path": file_path,
                "sheet_name": sheet_name,
                "search_info": {
                    "column_name": column_name,
                    "search_value": search_value,
                    "match_type": match_type,
                    "total_matches": total_matches,
                    "returned_results": len(matched_records)
                },
                "columns": df.columns.tolist(),
                "matched_rows": matched_records
            }
        }
        
        logger.info(f"[Tool] Successfully completed search. Total matches: {total_matches}, Returned: {len(matched_records)}")
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
        error_msg = f"搜索 Excel 数据时发生未知错误: {str(e)}"
        logger.error(f"[Error] {error_msg}")
        return {
            "success": False,
            "error": error_msg,
            "data": None
        }
