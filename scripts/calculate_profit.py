#!/usr/bin/env python3
"""
抖音销售利润计算器
严格按照原始计算逻辑脚本设计
"""
import pandas as pd
import os
import sys
import argparse
from datetime import datetime
from openpyxl import Workbook


def calculate_profit(sales_file: str, output_path: str = None) -> dict:
    """
    计算抖音销售利润
    
    Args:
        sales_file: 销售报表文件路径（CSV或Excel）
        output_path: 输出文件路径（可选）
    
    Returns:
        dict: 包含状态、输出路径、统计数据
    """
    # 获取成本表路径（永久目录）
    cost_file = r"D:\01A工作\抖音订单利润计算\cost_table.csv"
    
    if not os.path.exists(cost_file):
        return {"status": "ERROR", "message": f"成本表不存在: {cost_file}"}
    
    # 生成输出路径
    if output_path is None:
        timestamp = datetime.now().strftime("%m%d_%H%M")
        output_folder = "每日订单"
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, f"{timestamp}_抖音利润表.xlsx")
    
    try:
        # ========== 数据加载 ==========
        # 检测文件格式（根据扩展名或文件头）
        def detect_format(filepath):
            ext = os.path.splitext(filepath)[1].lower()
            if ext in ('.csv', '.txt'):
                return 'csv'
            elif ext in ('.xlsx', '.xls'):
                return 'excel'
            # 无扩展名时读文件头判断
            with open(filepath, 'rb') as f:
                header = f.read(4)
            if header[:3] == b'\xef\xbb\xbf' or header[:1] == b',':
                return 'csv'
            if header[:2] in (b'PK', b'\xd0\xcf'):
                return 'excel'
            return 'csv'  # 默认尝试CSV
        
        file_format = detect_format(sales_file)
        print(f"检测到文件格式：{file_format}")
        
        converters = {
            '主订单编号': lambda x: str(x).replace('—', '').replace('\t', ''),
            '子订单编号': lambda x: str(x).replace('—', '').replace('\t', '')
        }
        
        if file_format == 'csv':
            order_df = pd.read_csv(sales_file, converters=converters)
        else:
            order_df = pd.read_excel(
                sales_file,
                dtype={'主订单编号': str, '子订单编号': str},
                converters=converters
            )
        
        # 读取成本表
        cost_df = pd.read_csv(cost_file, encoding='utf-8-sig')
        
        print(f"成功读取订单数据，共 {len(order_df)} 条记录")
        print(f"成功读取成本数据，共 {len(cost_df)} 条记录")
        
        # ========== 数据清洗 ==========
        initial_count = len(order_df)
        
        # 清洗数值列（去除可能的制表符等）
        for col in ['订单应付金额', '商品数量', '平台实际承担优惠金额', '商家收入金额']:
            if col in order_df.columns:
                order_df[col] = pd.to_numeric(
                    order_df[col].astype(str).str.replace('\t', '').str.strip(),
                    errors='coerce'
                ).fillna(0)
        
        # 清洗订单状态列
        if '订单状态' in order_df.columns:
            order_df['订单状态'] = order_df['订单状态'].astype(str).str.replace('\t', '').str.strip()
        
        # 过滤规则：排除待支付、已关闭，商品数量<=10，订单金额<=350，排除¥0订单（退款订单）
        mask = (
            (~order_df["订单状态"].isin(["待支付", "已关闭"])) &
            (order_df["商品数量"] <= 10) &
            (order_df["订单应付金额"] <= 350) &
            (order_df["订单应付金额"] > 0)
        )
        clean_df = order_df[mask].copy()
        filtered_df = order_df[~mask].copy()
        filtered = initial_count - len(clean_df)
        print(f"数据清洗完成，剩余 {len(clean_df)} 条（过滤 {filtered} 条）")
        
        # ========== 数据合并 ==========
        # 清洗选购商品字段（去除制表符）
        clean_df["选购商品"] = clean_df["选购商品"].astype(str).str.replace('\t', '').str.strip()
        cost_df["选购商品"] = cost_df["选购商品"].astype(str).str.replace('\t', '').str.strip()
        
        # 第1步：精确匹配
        merged_data = clean_df.merge(cost_df, on="选购商品", how="left")
        
        # 第2步：模糊匹配（针对未匹配的百家布商品）
        # 规则：按产品类型+尺寸匹配，三件套必须包含"三件套"，区分单件和套装
        # 尺寸格式兼容：200cm*230cm、200*230、90cm*220cm、90*220 等
        def normalize_name(name):
            """标准化商品名：去除cm、空格、±5cm等干扰"""
            import re
            n = name.replace('cm', '').replace('CM', '').replace(' ', '')
            n = re.sub(r'[（(]±5[)）]', '', n)
            n = n.replace('±5', '')
            return n
        
        # (产品类型关键词, 尺寸关键词, 排除条件, 商品成本, 物流成本, 包装成本)
        fuzzy_rules = [
            # 夏凉被/被
            ("被", "200*230", [], 39, 4.5, 0.5),
            ("被", "150*200", [], 32, 4.5, 0.5),
            ("被", "90*220", [], 13.5, 2.5, 0.5),
            # 三件套（必须包含"三件套"）
            ("三件套", "200*230", [], 45.8, 3.6, 0.5),
            ("三件套", "230*250", [], 54.8, 4.5, 0.5),
            ("三件套", "90*220", [], 25.3, 2.5, 0.5),
            # 四件套（必须包含"四件套"）
            ("四件套", "230*250", [], 93.8, 4.5, 0.5),
            ("四件套", "200*230", [], 84.8, 3.6, 0.5),
            ("四件套", "90*220", [], 38.3, 2.5, 0.5),
            # 单床盖（排除三件套、四件套）
            ("床盖", "200*230", ["三件套", "四件套"], 34, 3.6, 0.5),
            ("床盖", "230*250", ["三件套", "四件套"], 43, 4.5, 0.5),
            ("床盖", "90*220", ["三件套", "四件套"], 13.5, 2.5, 0.5),
            # 枕套
            ("枕套", "48*74", [], 11.8, 1.7, 0.5),
            # 垫子
            ("垫", "90*220", [], 13.5, 2.5, 0.5),
        ]
        
        unmatched_mask = merged_data["商品成本"].isna() | (merged_data["商品成本"] == 0)
        if unmatched_mask.any():
            for idx in merged_data[unmatched_mask].index:
                raw_name = str(merged_data.loc[idx, "选购商品"])
                norm = normalize_name(raw_name)
                for type_kw, size_kw, excludes, cost_val, ship_val, pack_val in fuzzy_rules:
                    if type_kw in norm and size_kw in norm and not any(ex in norm for ex in excludes):
                        merged_data.loc[idx, "商品成本"] = cost_val
                        merged_data.loc[idx, "物流成本"] = ship_val
                        merged_data.loc[idx, "包装成本"] = pack_val
                        break
        
        # 填充仍未匹配的成本为0
        for col in ['商品成本', '物流成本', '包装成本']:
            if col in merged_data.columns:
                merged_data[col] = merged_data[col].fillna(0)
            else:
                merged_data[col] = 0
        
        # 检测未匹配成本的商品
        unmatched = merged_data[merged_data['商品成本'] == 0]['选购商品'].unique()
        if len(unmatched) > 0:
            print(f"\n⚠️ 以下 {len(unmatched)} 个商品在成本表中未匹配到：")
            for item in unmatched:
                print(f"  - {item}")
        
        # 订单编号验证
        invalid_parent = merged_data[~merged_data['主订单编号'].str.match(r'^\d{15,}$', na=False)]
        invalid_child = merged_data[~merged_data['子订单编号'].str.match(r'^\d+$', na=False)]
        
        if not invalid_parent.empty:
            print(f"⚠️ 发现 {len(invalid_parent)} 条异常主订单编号")
        if not invalid_child.empty:
            print(f"⚠️ 发现 {len(invalid_child)} 条异常子订单编号")
        
        # ========== 利润计算 ==========
        # 单利润 = 商家收入金额 * 0.95 - 商品数量 * (商品成本 + 物流成本 + 包装成本)
        merged_data["单利润"] = (
            (merged_data["商家收入金额"] * 0.95) -
            (merged_data["商品数量"] * merged_data["商品成本"]) -
            (merged_data["商品数量"] * merged_data["物流成本"]) -
            (merged_data["商品数量"] * merged_data["包装成本"])
        )
        
        # 计算各项成本总和
        merged_data["总商品成本"] = merged_data["商品数量"] * merged_data["商品成本"]
        merged_data["总物流成本"] = merged_data["商品数量"] * merged_data["物流成本"]
        merged_data["总包装成本"] = merged_data["商品数量"] * merged_data["包装成本"]
        
        avg_profit = merged_data["单利润"].mean()
        print(f"利润计算完成，平均利润 ¥{avg_profit:.2f}")
        
        # ========== 生成报表 ==========
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Sheet1: 利润总览
            profit_report = merged_data.groupby("选购商品").agg(
                总销量=("商品数量", "sum"),
                总商家收入=("商家收入金额", "sum"),
                总商品成本=("总商品成本", "sum"),
                总物流成本=("总物流成本", "sum"),
                总包装成本=("总包装成本", "sum"),
                总利润=("单利润", "sum")
            ).reset_index()
            
            profit_report = profit_report[[
                "选购商品", "总销量", "总商家收入",
                "总商品成本", "总物流成本",
                "总包装成本", "总利润"
            ]].sort_values("总利润", ascending=False)
            
            profit_report.to_excel(writer, sheet_name='利润总览', index=False)
            
            # Sheet2: 有效订单
            clean_df.to_excel(writer, sheet_name='有效订单', index=False)
            
            # Sheet3: 筛选订单
            filtered_df.to_excel(writer, sheet_name='筛选订单', index=False)
            
            # 设置文本格式（订单编号防止科学计数法）
            workbook = writer.book
            for sheetname in writer.sheets:
                sheet = workbook[sheetname]
                sheet.column_dimensions['A'].number_format = '@'
                sheet.column_dimensions['B'].number_format = '@'
        
        print(f"✅ 生成结果文件：{output_path}（包含3个工作表）")
        
        # ========== 统计汇总 ==========
        stats = {
            "status": "SUCCESS",
            "output_path": output_path,
            "total_orders": initial_count,
            "valid_orders": len(clean_df),
            "filtered_orders": filtered,
            "avg_profit": round(avg_profit, 2),
            "total_profit": round(merged_data["单利润"].sum(), 2),
            "total_sales": round(merged_data["商家收入金额"].sum(), 2),
        }
        
        if not profit_report.empty:
            max_profit_item = profit_report.loc[profit_report["总利润"].idxmax()]
            stats["best_product"] = max_profit_item["选购商品"]
            stats["best_profit"] = round(max_profit_item["总利润"], 2)
        
        print(f"\n📊 利润统计：")
        print(f"  总订单：{stats['total_orders']} 笔")
        print(f"  有效订单：{stats['valid_orders']} 笔")
        print(f"  过滤订单：{stats['filtered_orders']} 笔")
        print(f"  总商家收入：¥{stats['total_sales']:.2f}")
        print(f"  总利润：¥{stats['total_profit']:.2f}")
        print(f"  平均利润：¥{stats['avg_profit']:.2f}")
        if "best_product" in stats:
            print(f"  最高利润商品：{stats['best_product']}（¥{stats['best_profit']:.2f}）")
        
        return stats
        
    except Exception as e:
        error_msg = f"处理失败：{type(e).__name__}: {str(e)}"
        print(f"❌ {error_msg}")
        return {"status": "ERROR", "message": error_msg}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="抖音销售利润计算器")
    parser.add_argument("--sales", required=True, help="销售报表文件路径")
    parser.add_argument("--output", default=None, help="输出文件路径（可选）")
    args = parser.parse_args()
    
    result = calculate_profit(args.sales, args.output)
    if result["status"] == "ERROR":
        sys.exit(1)
