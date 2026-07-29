"""诊断输出结果中的问题"""
import pandas as pd
import os

OUTPUT_DIR = r'C:\Users\ruijie\Desktop\IDC数据文件\汇总结果'

# 1. 检查全产品数据中的Quarter值
print("=" * 60)
print("1. 检查Quarter字段值")
print("=" * 60)
combined_path = os.path.join(OUTPUT_DIR, 'IDC全产品数据.xlsx')
df = pd.read_excel(combined_path)

# 检查实际数中2025年的Quarter分布
actual = df[df['实际/预测'] == '实际']
print("\n实际数中所有年份的Quarter分布:")
for year in sorted(actual['Year'].dropna().unique()):
    year_data = actual[actual['Year'] == year]
    quarters = year_data['Quarter'].dropna().unique()
    print(f"  {year}: quarters = {sorted(quarters)}")

# 2. 检查通信行业数据
print("\n" + "=" * 60)
print("2. 检查通信行业数据")
print("=" * 60)
telecom = actual[
    (actual['产品/行业'] == '行业口径') &
    (actual['行业大类'] == '通信')
]
print(f"通信行业实际数行数: {len(telecom)}")
if len(telecom) > 0:
    print("\n各年份通信行业数据:")
    for year in sorted(telecom['Year'].dropna().unique()):
        year_data = telecom[telecom['Year'] == year]
        total = year_data['Vendor Revenue (CNY M)'].sum()
        quarters = year_data['Quarter'].dropna().unique()
        dc_data = year_data[year_data['二级产品分类'] == 'Datacenter']
        dc_total = dc_data['Vendor Revenue (CNY M)'].sum()
        print(f"  {year}: 总计={total:.2f}, DC={dc_total:.2f}, quarters={sorted(quarters)}")
        
        # 检查锐捷数据
        ruijie = year_data[year_data['Vendor'].str.contains('Ruijie', case=False, na=False)]
        ruijie_dc = ruijie[ruijie['二级产品分类'] == 'Datacenter']
        ruijie_dc_rev = ruijie_dc['Vendor Revenue (CNY M)'].sum()
        print(f"    锐捷DC收入: {ruijie_dc_rev:.2f}")

# 3. 检查预测数据
print("\n" + "=" * 60)
print("3. 检查预测数据")
print("=" * 60)
forecast = df[df['实际/预测'] == '预测']
print(f"预测数总行数: {len(forecast)}")
if len(forecast) > 0:
    print("\n预测数据年份分布:")
    print(forecast['Year'].value_counts().sort_index())
    
    # 检查通信行业预测数
    telecom_fc = forecast[
        (forecast['产品/行业'] == '行业口径') &
        (forecast['行业大类'] == '通信')
    ]
    print(f"\n通信行业预测数行数: {len(telecom_fc)}")
    if len(telecom_fc) > 0:
        print("\n各年份通信行业预测数:")
        for year in sorted(telecom_fc['Year'].dropna().unique()):
            year_data = telecom_fc[telecom_fc['Year'] == year]
            total = year_data['Vendor Revenue (CNY M)'].sum()
            print(f"  {year}: 总计={total:.2f}")

# 4. 检查行业大类字段中的英文值
print("\n" + "=" * 60)
print("4. 检查行业大类字段英文值")
print("=" * 60)
print("\n行业大类所有唯一值:")
print(df['行业大类'].value_counts(dropna=False))

# 5. 检查Vendor字段中的锐捷
print("\n" + "=" * 60)
print("5. 检查Vendor字段")
print("=" * 60)
ruijie_vendors = df[df['Vendor'].str.contains('Ruijie|ruijie|锐捷', case=False, na=False)]['Vendor'].unique()
print(f"含锐捷的Vendor值: {ruijie_vendors}")
