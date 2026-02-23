# -*- coding: utf-8 -*-

"""核心提取模块"""

import os
import traceback
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd

from ifc_prop_getter import utils
from ifc_prop_getter.constants import SKIP_ENTITY_TYPES


def extract_properties(ifc_path, properties, include_globalid, include_name,
                       output_dir, base_filename, file_format, queue, stop_event):
    """工作线程函数：执行 IFC 实体扫描与属性提取"""
    try:
        queue.put({'type': 'log', 'message': f"开始处理文件: {ifc_path}"})

        # 阶段 1: 扫描
        queue.put({'type': 'status', 'message': "正在扫描 IFC 实体..."})

        try:
            ifc_file = ifcopenshell.open(ifc_path)
        except Exception as e:
            queue.put({'type': 'error', 'message': f"文件打开失败: {str(e)}"})
            return

        if stop_event.is_set(): return

        all_products = ifc_file.by_type("IfcProduct")
        total_count = len(all_products)
        queue.put({'type': 'log', 'message': f"共找到 {total_count} 个 IfcProduct 实例"})

        if total_count == 0:
            queue.put({'type': 'error', 'message': "文件中未找到任何 IfcProduct 实体"})
            return

        results = []

        # 阶段 2: 提取
        queue.put({'type': 'status', 'message': "正在提取属性..."})

        for element in all_products:
            if stop_event.is_set():
                queue.put({'type': 'log', 'message': "任务已被用户取消"})
                return

            if element.is_a() in SKIP_ENTITY_TYPES:
                continue

            try:
                psets = ifcopenshell.util.element.get_psets(element)
                row = {}
                has_valid = False

                for prop_name in properties:
                    value = utils.extract_property_from_psets(psets, prop_name)
                    if value not in (None, "N/A"):
                        has_valid = True
                    row[prop_name] = value

                if has_valid:
                    if include_globalid:
                        row["GlobalId"] = utils.safe_str(element.GlobalId)
                    if include_name:
                        row["Name"] = utils.safe_str(element.Name)
                    results.append(row)

            except Exception as e:
                queue.put({'type': 'log', 'message': f"警告: 构件 {element.GlobalId} 提取失败: {str(e)}"})
                continue

        queue.put({'type': 'log', 'message': f"提取完成，共获取 {len(results)} 个有效构件"})

        if not results:
            queue.put({'type': 'error', 'message': "未提取到任何有效数据"})
            return

        # 阶段 3: 写入
        queue.put({'type': 'status', 'message': f"正在写入 {file_format}..."})

        df = pd.DataFrame(results)

        cols = [c for c in properties if c in df.columns]
        if include_name and "Name" in df.columns:
            cols.insert(0, "Name")
        if include_globalid and "GlobalId" in df.columns:
            cols.insert(0, "GlobalId")
        df = df[cols]

        if stop_event.is_set(): return

        ext = "xlsx" if file_format == "Excel" else "csv"
        filename = utils.make_output_filename(base_filename, ext)
        filepath = os.path.join(output_dir, filename)

        try:
            if file_format == "Excel":
                utils.export_to_excel_with_format(df, filepath)
            else:
                df.to_csv(filepath, index=False, encoding='utf-8-sig')
        except Exception as e:
            queue.put({'type': 'error', 'message': f"写入文件失败: {str(e)}"})
            return

        queue.put({
            'type': 'complete',
            'filepath': filepath,
            'message': f"成功导出 {len(results)} 行数据"
        })

    except Exception as e:
        queue.put({'type': 'error', 'message': f"未捕获的异常: {str(e)}"})
        queue.put({'type': 'log', 'message': traceback.format_exc()})
    finally:
        queue.put({'type': 'finished'})