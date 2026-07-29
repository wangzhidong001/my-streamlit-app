"""验证输出结果文件"""
import pandas as pd
import os

OUTPUT_DIR = r'C:\Users\ruijie\Desktop\IDC数据文件\汇总结果'

print("=" * 60)
print("输出文件验证")
print("=" * 60)

# 1. 分析结果
print("\n=== IDC分析结果.xlsx ===")
analysis_path = os.path.join(OUTPUT_DIR, 'IDC分析结果.xlsx')
if os.path.exists(analysis_path):
    df = pd.read_excel(analysis_path)
    print(f"行数: {len(df)}")
    print(f"列: {list(df.columns)}")
    print(df.to_string())
else:
    print("文件不存在!")

# 2. 全产品数据 - 检查行业大类字段
print("\n=== IDC全产品数据.xlsx - 行业大类分析 ===")
combined_path = os.path.join(OUTPUT_DIR, 'IDC全产品数据.xlsx')
if os.path.exists(combined_path):
    df = pd.read_excel(combined_path)
    print(f"总行数: {len(df)}")
    print(f"列: {list(df.columns)}")
    
    # 检查行业大类字段
    if '行业大类' in df.columns:
        print("\n行业大类字段唯一值:")
        print(df['行业大类'].value_counts(dropna=False).head(20))
    
    # 检查实际数
    actual = df[df['实际/预测'] == '实际']
    print(f"\n实际数行数: {len(actual)}")
    if '行业大类' in actual.columns:
        print("实际数中行业大类分布:")
        print(actual['行业大类'].value_counts(dropna=False).head(20))
    
    # 检查通信相关
    if '行业大类' in df.columns:
        telecom = df[df['行业大类'].astype(str).str.contains('通信|Telecom|Telecommunication', case=False, na=False)]
        print(f"\n含'通信'的行数: {len(telecom)}")
    
    # 检查实际/预测和产品/行业字段
    print("\n产品/行业分布:")
    print(df['产品/行业'].value_counts(dropna=False))
    print("\n实际/预测分布:")
    print(df['实际/预测'].value_counts(dropna=False))
    
    # 检查二级产品分类
    if '二级产品分类' in df.columns:
        print("\n二级产品分类分布(top10):")
        print(df['二级产品分类'].value_counts(dropna=False).head(10))
else:
    print("文件不存在!")

# 3. 检查各产品文件
print("\n=== 各产品文件 ===")
for product in ['WLAN', 'Switch', 'Router', 'VCC']:
    path = os.path.join(OUTPUT_DIR, f'{product}.xlsx')
    if os.path.exists(path):
        df = pd.read_excel(path)
        print(f"\n{product}.xlsx: {len(df)} 行")
        print(f"  列: {list(df.columns)[:10]}...")
        if '实际/预测' in df.columns:
            print(f"  实际/预测分布: {dict(df['实际/预测'].value_counts(dropna=False))}")
        if '产品/行业' in df.columns:
            print(f"  产品/行业分布: {dict(df['产品/行业'].value_counts(dropna=False))}")

# 4. 检查图表HTML
chart_path = os.path.join(OUTPUT_DIR, 'IDC图表.html')
print(f"\n=== IDC图表.html ===")
print(f"存在: {os.path.exists(chart_path)}")
if os.path.exists(chart_path):
    print(f"大小: {os.path.getsize(chart_path)} bytes")
