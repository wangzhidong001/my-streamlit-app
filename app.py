"""
IDC 数据处理分析平台 - Streamlit 应用
功能：数据处理 / 结果查看 / 分析看板 / 参数设置
"""
import os
import sys
import io
import json
import contextlib
import pandas as pd
import streamlit as st

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import idc_processor as proc

# ============ 页面配置 ============
st.set_page_config(
    page_title="IDC 数据处理分析平台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============ 全局样式 ============
st.markdown("""
<style>
    .main-header {
        font-size: 2rem; font-weight: 700; color: #1f2937;
        padding-bottom: 0.5rem; border-bottom: 2px solid #e5e7eb;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: #f9fafb; border-radius: 8px; padding: 1rem;
        border-left: 4px solid #3b82f6;
    }
    .status-ok { color: #10b981; font-weight: 600; }
    .status-warn { color: #f59e0b; font-weight: 600; }
    .status-error { color: #ef4444; font-weight: 600; }
    div[data-testid="stMetric"] {
        background: #f9fafb; border-radius: 8px; padding: 0.75rem;
        border-left: 3px solid #3b82f6;
    }
</style>
""", unsafe_allow_html=True)

# ============ 配置管理 ============
def load_config():
    """加载配置"""
    config_path = os.path.join(proc.OUTPUT_DIR, 'app_config.json')
    default_config = {
        'data_dir': proc.DATA_DIR,
        'output_dir': proc.OUTPUT_DIR,
        'vat_rate': 0.13,
        'custom_params': [],
    }
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8-sig') as f:
            saved = json.load(f)
            default_config.update(saved)
    # 确保 custom_params 是列表
    if 'custom_params' not in default_config:
        default_config['custom_params'] = []
    return default_config


def save_config(config):
    """保存配置"""
    os.makedirs(proc.OUTPUT_DIR, exist_ok=True)
    config_path = os.path.join(proc.OUTPUT_DIR, 'app_config.json')
    with open(config_path, 'w', encoding='utf-8-sig') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ============ 动态参数管理 ============
def init_custom_params(config):
    """初始化 session_state 中的自定义参数"""
    if 'custom_params' not in st.session_state:
        saved = config.get('custom_params', [])
        # 兼容旧配置：为没有 year 字段的参数添加 year=None
        for p in saved:
            if 'year' not in p:
                p['year'] = None
        st.session_state.custom_params = saved.copy() if saved else []


def add_custom_param():
    """添加一个空白参数"""
    st.session_state.custom_params.append({'name': '', 'value': '', 'type': 'number', 'year': None})


def parse_year(year):
    """解析年份字段，兼容数值、字符串以及 '2026H'/'2026E' 等后缀。"""
    if year is None:
        return None
    if isinstance(year, (int, float)):
        return int(year)
    s = str(year).strip()
    # 去掉 H(实际半年)/E(预测) 等后缀
    s = s.rstrip('HhEe')
    if s.isdigit():
        return int(s)
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def get_available_years():
    """从分析结果中获取可用年份列表"""
    analysis_path = os.path.join(proc.OUTPUT_DIR, 'IDC分析结果.xlsx')
    if os.path.exists(analysis_path):
        df = pd.read_excel(analysis_path)
        if '年份' in df.columns:
            years = [parse_year(y) for y in df['年份'].dropna().unique()]
            years = [y for y in years if y is not None]
            return sorted(set(years))
    return [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]


def remove_custom_param(index):
    """删除指定索引的参数"""
    if 0 <= index < len(st.session_state.custom_params):
        st.session_state.custom_params.pop(index)


# ============ 侧边栏 ============
st.sidebar.markdown("## 📊 IDC 数据处理分析平台")
st.sidebar.caption("基于 Streamlit 的数据加工与分析工具")

config = load_config()

# 应用配置到处理器
proc.DATA_DIR = config['data_dir']
proc.OUTPUT_DIR = config['output_dir']
proc.BACKUP_DIR = os.path.join(proc.OUTPUT_DIR, '历史版本备份')
proc.HISTORY_FILE = os.path.join(proc.OUTPUT_DIR, '处理历史记录.json')
proc.SOURCE_HISTORY_FILE = os.path.join(proc.DATA_DIR, '处理历史.json')

page = st.sidebar.radio("功能导航", [
    "🔄 数据处理",
    "📋 结果查看",
    "📈 分析看板",
    "⚙️ 参数设置",
])

st.sidebar.markdown("---")
st.sidebar.markdown("**数据源目录**")
st.sidebar.code(config['data_dir'])
st.sidebar.markdown("**输出目录**")
st.sidebar.code(config['output_dir'])

# 检查输出是否存在
if os.path.exists(proc.OUTPUT_DIR):
    files = [f for f in os.listdir(proc.OUTPUT_DIR) if f.endswith('.xlsx')]
    st.sidebar.success(f"已有 {len(files)} 个结果文件")
else:
    st.sidebar.warning("尚未生成结果，请先处理数据")


# ============ 页面1: 数据处理 ============
if page == "🔄 数据处理":
    st.markdown('<div class="main-header">🔄 数据处理</div>', unsafe_allow_html=True)

    st.markdown("### 数据源状态")

    # 扫描文件
    files = proc.scan_excel_files()
    history = proc.load_history()
    new_files, all_filenames = proc.check_new_files(files, history)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("扫描到文件数", len(files))
    with col2:
        st.metric("新增/修改文件", len(new_files))
    with col3:
        st.metric("已处理文件数", len(history.get('processed_files', [])))

    # 显示新增文件
    if new_files:
        st.markdown("### 🆕 新增/修改的文件")
        new_df = pd.DataFrame([{
            '文件名': f['filename'],
            '修改时间': pd.Timestamp(f['mtime'], unit='s').strftime('%Y-%m-%d %H:%M')
        } for f in new_files])
        st.dataframe(new_df, use_container_width=True, hide_index=True)

    # 显示所有文件
    with st.expander(f"查看全部文件（{len(files)} 个）", expanded=False):
        all_df = pd.DataFrame([{
            '文件名': f['filename'],
            '修改时间': pd.Timestamp(f['mtime'], unit='s').strftime('%Y-%m-%d %H:%M'),
            '状态': '🆕 新增' if f in new_files else '✅ 已处理'
        } for f in files])
        st.dataframe(all_df, use_container_width=True, hide_index=True)

    # 文件分类预览
    st.markdown("### 📂 文件分类预览")
    if files:
        classified = proc.classify_all_files(files)
        class_data = []
        for key, info in classified.items():
            class_data.append({
                '产品': info['product'] or '—',
                '数据类型': {'actual': '实际数', 'forecast': '预测数'}.get(info['data_type'], '—'),
                '口径': {'product': '产品口径', 'industry': '行业口径'}.get(info['perspective'], '—'),
                '时间段': info['period'] or '—',
                '文件名': os.path.basename(info['filename'])
            })
        st.dataframe(pd.DataFrame(class_data), use_container_width=True, hide_index=True)

    # ========== 一键数据处理 ==========
    combined_file_path = os.path.join(proc.OUTPUT_DIR, 'IDC全产品数据.xlsx')

    st.markdown("### 🚀 执行数据处理")
    st.caption("完整流程：扫描文件 → 历史对比 → 文件分类 → 实际数汇总 → 预测数汇总 → 全产品汇总 → 通信DC分析计算 → 图表生成 → 保存结果")

    if st.button("▶️ 一键执行数据处理", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="正在初始化...")
        output_buffer = io.StringIO()

        with contextlib.redirect_stdout(output_buffer):
            try:
                # ---- 第一阶段：IDC明细数据处理 ----
                progress_bar.progress(5, text="步骤1: 扫描文件...")
                files = proc.scan_excel_files()

                progress_bar.progress(10, text="步骤2: 历史对比...")
                history = proc.load_history()
                new_files, all_filenames = proc.check_new_files(files, history)

                progress_bar.progress(15, text="步骤3-4: 文件分类...")
                classified = proc.classify_all_files(files)

                progress_bar.progress(25, text="步骤5: 实际数汇总...")
                actual_results = proc.process_actual_data(classified)

                progress_bar.progress(35, text="步骤6: 预测数汇总...")
                forecast_results = proc.process_forecast_data(classified, actual_results)

                progress_bar.progress(45, text="步骤7: 全产品汇总...")
                combined = proc.combine_all_products(forecast_results)

                progress_bar.progress(50, text="保存IDC明细结果...")
                os.makedirs(proc.OUTPUT_DIR, exist_ok=True)
                proc.save_product_files(forecast_results, proc.OUTPUT_DIR)
                proc.save_combined_file(combined, proc.OUTPUT_DIR)

                # 更新历史
                from datetime import datetime
                new_history = {
                    'last_processed': datetime.now().isoformat(),
                    'processed_files': [{'filename': f['filename'], 'mtime': f['mtime']} for f in files]
                }
                proc.save_history(new_history)

                # ---- 第二阶段：通信DC数据加工 ----
                progress_bar.progress(60, text="步骤8: 通信DC分析计算...")
                # 注意：combined 在 save_combined_file 后可能被污染，需要 .copy() 确保数据独立性
                analysis = proc.calculate_analysis_fixed(
                    combined.copy(),
                    vat_rate=config['vat_rate'],
                )

                if len(analysis) == 0:
                    progress_bar.empty()
                    st.warning("⚠️ 未生成分析结果，请检查数据是否包含行业口径/通信数据。")
                else:
                    progress_bar.progress(80, text="步骤9: 图表生成...")
                    chart_data = proc.generate_chart_data(analysis)

                    progress_bar.progress(90, text="步骤10: 保存分析结果...")
                    proc.save_analysis_file(analysis, proc.OUTPUT_DIR)
                    proc.save_chart_html(chart_data, proc.OUTPUT_DIR)
                    st.session_state['analysis_df_original'] = analysis.copy()
                    st.session_state['analysis_df_simulated'] = analysis.copy()

                    progress_bar.progress(100, text="完成")
                    progress_bar.empty()
                    st.success("✅ 数据处理完成！已生成IDC明细数据、分析结果和图表。")

            except Exception as e:
                progress_bar.empty()
                st.error(f"❌ 处理失败: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

        # 显示处理日志
        log_text = output_buffer.getvalue()
        if log_text:
            with st.expander("📋 处理日志", expanded=True):
                st.code(log_text, language='text')

    # 显示上次处理时间
    if history.get('last_processed'):
        st.info(f"上次处理时间: {history['last_processed']}")


# ============ 页面2: 结果查看 ============
elif page == "📋 结果查看":
    st.markdown('<div class="main-header">📋 结果查看</div>', unsafe_allow_html=True)

    # 检查输出文件
    if not os.path.exists(proc.OUTPUT_DIR):
        st.warning("⚠️ 尚无处理结果，请先到「数据处理」页面执行处理")
        st.stop()

    result_files = [f for f in os.listdir(proc.OUTPUT_DIR) if f.endswith('.xlsx')]
    if not result_files:
        st.warning("⚠️ 尚无处理结果，请先到「数据处理」页面执行处理")
        st.stop()

    # 文件选择
    st.markdown("### 选择数据文件")
    selected_file = st.selectbox("文件", result_files, label_visibility="collapsed")

    if selected_file:
        file_path = os.path.join(proc.OUTPUT_DIR, selected_file)
        df = pd.read_excel(file_path)

        # 概览指标
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("总行数", f"{len(df):,}")
        with col2:
            st.metric("总列数", len(df.columns))
        with col3:
            if '实际/预测' in df.columns:
                actual_count = len(df[df['实际/预测'] == '实际'])
                st.metric("实际数行数", f"{actual_count:,}")
            else:
                st.metric("实际数行数", "—")
        with col4:
            if '实际/预测' in df.columns:
                forecast_count = len(df[df['实际/预测'] == '预测'])
                st.metric("预测数行数", f"{forecast_count:,}")
            else:
                st.metric("预测数行数", "—")

        # 筛选器
        st.markdown("### 🔍 数据筛选")
        filter_cols = st.columns(4)

        filtered_df = df.copy()

        with filter_cols[0]:
            if '产品/行业' in df.columns:
                persp_options = ['全部'] + list(df['产品/行业'].dropna().unique())
                persp = st.selectbox("产品/行业", persp_options)
                if persp != '全部':
                    filtered_df = filtered_df[filtered_df['产品/行业'] == persp]

        with filter_cols[1]:
            if '实际/预测' in df.columns:
                type_options = ['全部'] + list(df['实际/预测'].dropna().unique())
                dtype = st.selectbox("实际/预测", type_options)
                if dtype != '全部':
                    filtered_df = filtered_df[filtered_df['实际/预测'] == dtype]

        with filter_cols[2]:
            if 'Year' in df.columns:
                year_options = ['全部'] + sorted([str(y) for y in df['Year'].dropna().unique()])
                year = st.selectbox("年份", year_options)
                if year != '全部':
                    filtered_df = filtered_df[filtered_df['Year'].astype(str) == year]

        with filter_cols[3]:
            if 'Technology' in df.columns:
                tech_options = ['全部'] + list(df['Technology'].dropna().unique())
                tech = st.selectbox("产品", tech_options)
                if tech != '全部':
                    filtered_df = filtered_df[filtered_df['Technology'] == tech]

        # 搜索框
        search = st.text_input("🔍 搜索（支持任意列内容）", placeholder="输入关键词...")
        if search:
            mask = filtered_df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)
            filtered_df = filtered_df[mask]

        # 数据表
        st.markdown(f"### 📊 数据表（{len(filtered_df):,} 行）")

        # 分页
        page_size = st.selectbox("每页显示行数", [50, 100, 200, 500], index=1)
        total_pages = max(1, (len(filtered_df) + page_size - 1) // page_size)
        page_num = st.number_input("页码", 1, total_pages, 1)
        start_idx = (page_num - 1) * page_size
        end_idx = start_idx + page_size

        st.dataframe(
            filtered_df.iloc[start_idx:end_idx],
            use_container_width=True,
            hide_index=True,
            height=500
        )

        # 导出
        st.markdown("### 💾 导出")
        col1, col2 = st.columns([1, 4])
        with col1:
            export_format = st.selectbox("格式", ["Excel", "CSV"])
        if st.button("📥 导出筛选结果"):
            if export_format == "Excel":
                buf = io.BytesIO()
                filtered_df.to_excel(buf, index=False, engine='openpyxl')
                st.download_button(
                    "下载 Excel", buf.getvalue(),
                    file_name=f"筛选结果_{selected_file}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "下载 CSV", csv,
                    file_name=f"筛选结果_{selected_file.replace('.xlsx', '.csv')}",
                    mime="text/csv"
                )


# ============ 页面3: 分析看板 ============
elif page == "📈 分析看板":
    st.markdown('<div class="main-header">📈 分析看板</div>', unsafe_allow_html=True)

    analysis_path = os.path.join(proc.OUTPUT_DIR, 'IDC分析结果.xlsx')
    if not os.path.exists(analysis_path):
        st.warning("⚠️ 尚无分析结果，请先执行数据处理")
        st.stop()

    # 使用 session_state 保存模拟后的数据，避免 rerun 后丢失
    if 'analysis_df_original' not in st.session_state:
        st.session_state.analysis_df_original = pd.read_excel(analysis_path)
    if 'analysis_df_simulated' not in st.session_state:
        st.session_state.analysis_df_simulated = st.session_state.analysis_df_original.copy()

    # 检测是否需要重置（数据源更新后原始文件变化）
    current_mtime = os.path.getmtime(analysis_path)
    if 'analysis_file_mtime' not in st.session_state:
        st.session_state.analysis_file_mtime = current_mtime
    elif current_mtime != st.session_state.analysis_file_mtime:
        # 原始文件已更新，重置模拟数据
        st.session_state.analysis_df_original = pd.read_excel(analysis_path)
        st.session_state.analysis_df_simulated = st.session_state.analysis_df_original.copy()
        st.session_state.analysis_file_mtime = current_mtime

    analysis_df = st.session_state.analysis_df_simulated.copy()

    # 参数模拟
    st.markdown("### 🔧 参数模拟")

    custom_params = config.get('custom_params', [])
    available_years = get_available_years()

    # 内置参数 + 自定义参数调节控件
    sim_values = {}  # 存储所有模拟参数的当前值

    # 增值税率（全局）
    sim_col1, sim_col2 = st.columns(2)
    with sim_col1:
        sim_vat = st.slider("增值税率（全局）", 0.01, 0.25, config['vat_rate'], 0.01, format="%.2f")
        sim_values['增值税率'] = sim_vat

    # 自定义参数按年份分组展示
    if custom_params:
        st.markdown("#### ➕ 自定义参数调节（按年份）")

        # 分组：全局参数 + 各年份参数
        global_params = [p for p in custom_params if p.get('year') is None]
        year_params_map = {}
        for y in available_years:
            y_params = [p for p in custom_params if p.get('year') == y]
            if y_params:
                year_params_map[y] = y_params

        # 先展示全局参数
        if global_params:
            st.markdown("**📌 全局参数（适用于所有年份）**")
            sim_cols = st.columns(min(len(global_params), 4))
            for idx, param in enumerate(global_params):
                with sim_cols[idx % len(sim_cols)]:
                    pname = param.get('name', f'参数{idx+1}')
                    pval = param.get('value', 0)
                    if param.get('type') == 'number':
                        try:
                            pval = float(pval)
                        except:
                            pval = 0.0
                        abs_val = abs(pval)
                        if abs_val == 0:
                            min_v, max_v, step = -100, 100, 1
                        elif abs_val < 1:
                            min_v, max_v, step = 0, max(pval * 3, 1), 0.01
                        elif abs_val < 100:
                            min_v, max_v, step = 0, max(pval * 3, 100), 0.1
                        else:
                            min_v, max_v, step = 0, max(pval * 3, 1000), 1
                        sim_values[(pname, None)] = {
                            'value': st.slider(
                                pname, float(min_v), float(max_v), float(pval),
                                float(step), key=f"sim_{pname}"
                            ),
                            'year': None
                        }
                    else:
                        sim_values[(pname, None)] = {
                            'value': st.text_input(pname, str(pval), key=f"sim_{pname}"),
                            'year': None
                        }

        # 再展示各年份参数
        for y in available_years:
            if y in year_params_map:
                y_params = year_params_map[y]
                st.markdown(f"**📅 {y}年专属参数**")
                y_cols = st.columns(min(len(y_params), 4))
                for idx, param in enumerate(y_params):
                    with y_cols[idx % len(y_cols)]:
                        pname = param.get('name', f'参数{idx+1}')
                        pval = param.get('value', 0)
                        label = f"{pname}"
                        if param.get('type') == 'number':
                            try:
                                pval = float(pval)
                            except:
                                pval = 0.0
                            abs_val = abs(pval)
                            if abs_val == 0:
                                min_v, max_v, step = -100, 100, 1
                            elif abs_val < 1:
                                min_v, max_v, step = 0, max(pval * 3, 1), 0.01
                            elif abs_val < 100:
                                min_v, max_v, step = 0, max(pval * 3, 100), 0.1
                            else:
                                min_v, max_v, step = 0, max(pval * 3, 1000), 1
                            sim_values[(pname, y)] = {
                                'value': st.slider(
                                    label, float(min_v), float(max_v), float(pval),
                                    float(step), key=f"sim_{y}_{pname}"
                                ),
                                'year': y
                            }
                        else:
                            sim_values[(pname, y)] = {
                                'value': st.text_input(label, str(pval), key=f"sim_{y}_{pname}"),
                                'year': y
                            }

    with sim_col2:
        if st.button("🔄 重新计算", type="primary"):
            # 从原始数据重新计算，确保每次基于原始值
            work_df = st.session_state.analysis_df_original.copy()

            # === 记录调整前的数据（按年份存一份快照） ===
            before_by_year = {}
            for i in range(len(work_df)):
                y_str = str(work_df.iloc[i]['年份'])
                before_by_year[y_str] = {
                    '份额': work_df.iloc[i].get('锐捷DC份额'),
                    '锐捷DC收入': work_df.iloc[i].get('锐捷DC收入'),
                    '开票金额': work_df.iloc[i].get('锐捷开票金额'),
                }
            affected_years = set()

            # 1. 增值税率影响：开票金额、开票同比变动（全局）
            work_df['增值税率'] = round(sim_vat, 2)
            ruijie_rev = work_df['锐捷DC收入']
            work_df['锐捷开票金额'] = (ruijie_rev * (1 + sim_vat)).round(2)

            # 2. 自定义参数影响（按年份条件应用）
            # 收集所有"锐捷DC份额"类参数，按年份分组
            share_params = {}  # year -> new_share_value, None表示全局
            for key, pmeta in sim_values.items():
                # key 可能是 (pname, year) 元组（新结构），也可能是 pname 字符串（兼容）
                if isinstance(key, tuple):
                    pname, pyear = key
                else:
                    pname = key
                    pyear = pmeta.get('year') if isinstance(pmeta, dict) else None
                if '份额' in pname or 'share' in pname.lower():
                    pval = pmeta['value'] if isinstance(pmeta, dict) else pmeta
                    share_params[pyear] = float(pval)

            # 应用份额参数（按年份）
            if share_params:
                for i in range(len(work_df)):
                    row_year = parse_year(work_df.iloc[i]['年份'])
                    y_str = str(work_df.iloc[i]['年份'])
                    # 优先使用年份专属参数，其次使用全局参数
                    if row_year in share_params:
                        new_share = share_params[row_year]
                    elif None in share_params:
                        new_share = share_params[None]
                    else:
                        continue
                    # 标记该年份受影响
                    affected_years.add(y_str)

                    work_df.loc[work_df.index[i], '锐捷DC份额'] = round(new_share, 2)
                    # 重新计算：锐捷DC收入 = 通信DC容量 × 份额%
                    dc_capacity = work_df.iloc[i]['通信DC容量']
                    new_rev = round(dc_capacity * new_share / 100, 2)
                    work_df.loc[work_df.index[i], '锐捷DC收入'] = new_rev
                    # 重新计算开票金额
                    work_df.loc[work_df.index[i], '锐捷开票金额'] = round(new_rev * (1 + sim_vat), 2)

                # 重新计算竞争力指数（所有行）
                for i in range(len(work_df)):
                    if i > 0:
                        prev_share = work_df.iloc[i-1]['锐捷DC份额']
                        curr_share = work_df.iloc[i]['锐捷DC份额']
                        if prev_share is not None and prev_share > 0:
                            work_df.loc[work_df.index[i], '竞争力指数'] = round(curr_share / prev_share, 2)

            # 增值税率单独影响开票金额（所有年份开票金额都被重算过）
            # 如果没有份额参数，那至少所有行的开票金额都受增值税率影响
            if not affected_years:
                for i in range(len(work_df)):
                    y_str = str(work_df.iloc[i]['年份'])
                    before = before_by_year.get(y_str, {})
                    after_inv = work_df.iloc[i].get('锐捷开票金额')
                    if before.get('开票金额') != after_inv:
                        affected_years.add(y_str)

            # 重新计算开票同比变动（所有行）
            for i in range(len(work_df)):
                if i > 0:
                    prev_inv = work_df.iloc[i-1]['锐捷开票金额']
                    curr_inv = work_df.iloc[i]['锐捷开票金额']
                    if prev_inv is not None and prev_inv > 0:
                        work_df.loc[work_df.index[i], '开票同比变动'] = round(
                            (curr_inv - prev_inv) / prev_inv * 100, 2
                        )

            # === 生成调整记录 ===
            if 'adjust_log' not in st.session_state:
                st.session_state.adjust_log = []

            # 取本次执行时间戳作为批次标识
            from datetime import datetime
            batch_ts = datetime.now().strftime("%H:%M:%S")

            # 收集受影响的行（按结果表中的顺序）
            for i in range(len(work_df)):
                y_str = str(work_df.iloc[i]['年份'])
                if y_str not in affected_years:
                    continue
                before = before_by_year.get(y_str, {})
                after_share = work_df.iloc[i].get('锐捷DC份额')
                after_rev = work_df.iloc[i].get('锐捷DC收入')
                after_inv = work_df.iloc[i].get('锐捷开票金额')
                entry = {
                    '批次': batch_ts,
                    '年份': y_str,
                    '调整前份额': before.get('份额'),
                    '调整后份额': after_share,
                    '调整前锐捷DC收入': before.get('锐捷DC收入'),
                    '调整后锐捷DC收入': after_rev,
                    '调整前开票金额': before.get('开票金额'),
                    '调整后开票金额': after_inv,
                }
                st.session_state.adjust_log.append(entry)

            # 保存到 session_state
            st.session_state.analysis_df_simulated = work_df
            st.success(f"✅ 已根据当前参数重新计算，本次共记录 {len(affected_years)} 个年份的调整。")
            st.rerun()

    # 重置按钮
    col_r1, col_r2 = st.columns([1, 1])
    with col_r1:
        if st.button("↩️ 恢复原始数据"):
            st.session_state.analysis_df_simulated = st.session_state.analysis_df_original.copy()
            st.success("✅ 已恢复原始数据")
            st.rerun()
    with col_r2:
        if st.button("🗑 清空调整记录"):
            if 'adjust_log' in st.session_state:
                st.session_state.adjust_log = []
            st.success("✅ 已清空调整记录")
            st.rerun()

    # 概览指标
    st.markdown("### 📝 指标调整记录")
    actual_data = analysis_df[analysis_df['数据类型'] == '实际']
    forecast_data = analysis_df[analysis_df['数据类型'] == '预测']

    # 显示调整记录（如果存在）
    if 'adjust_log' in st.session_state and st.session_state.adjust_log:
        log_df = pd.DataFrame(st.session_state.adjust_log)
        # 按用户指定顺序排列列，并去掉批次等辅助列
        display_cols = [
            '年份', '调整前份额', '调整后份额',
            '调整前锐捷DC收入', '调整后锐捷DC收入',
            '调整前开票金额', '调整后开票金额'
        ]
        # 只保留存在的列
        display_cols = [c for c in display_cols if c in log_df.columns]
        log_df = log_df[display_cols].copy()
        # 格式化百分比和金额
        pct_cols_log = ['调整前份额', '调整后份额']
        for c in pct_cols_log:
            if c in log_df.columns:
                log_df[c] = log_df[c].apply(lambda x: f"{x}%" if pd.notna(x) else "—")
        money_cols = ['调整前锐捷DC收入', '调整后锐捷DC收入', '调整前开票金额', '调整后开票金额']
        for c in money_cols:
            if c in log_df.columns:
                log_df[c] = log_df[c].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "—")
        st.dataframe(log_df, use_container_width=True, hide_index=True)
        st.caption(f"累计 {len(log_df)} 条调整记录。")
    else:
        st.info("ℹ️ 暂无调整记录。调整参数并点击「重新计算」后，这里会记录参数调整前后的对比。")

    # 图表
    st.markdown("### 📊 组合图：通信DC容量 & 锐捷DC份额")
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(x=analysis_df['年份'], y=analysis_df['通信DC容量'],
               name='通信DC容量', marker_color='#5470c6'),
        secondary_y=False
    )
    fig.add_trace(
        go.Bar(x=analysis_df['年份'], y=analysis_df['锐捷DC收入'],
               name='锐捷DC收入', marker_color='#91cc75'),
        secondary_y=False
    )
    fig.add_trace(
        go.Scatter(x=analysis_df['年份'], y=analysis_df['锐捷DC份额'],
                   name='锐捷DC份额', line=dict(color='#ee6666', width=3),
                   mode='lines+markers'),
        secondary_y=True
    )
    fig.update_layout(
        title="通信DC容量与锐捷DC份额趋势",
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=450
    )
    fig.update_yaxes(title_text="容量 (CNY M)", secondary_y=False)
    fig.update_yaxes(title_text="份额 (%)", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

    # 分析数据表
    st.markdown("### 📋 分析结果明细")

    display_df = analysis_df.copy()
    # 移除锐捷DC容量列（不再显示）
    if '锐捷DC容量' in display_df.columns:
        display_df = display_df.drop(columns=['锐捷DC容量'])
    pct_cols = ['DC占全产品比例', '通信DC容量增速', '锐捷DC份额', '开票同比变动']
    for col in pct_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(
                lambda x: f"{x}%" if pd.notna(x) else "—"
            )

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # 趋势分析
    st.markdown("### 📈 趋势分析")
    if len(actual_data) >= 2:
        col1, col2 = st.columns(2)
        with col1:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(x=actual_data['年份'], y=actual_data['通信全产品容量'],
                                  name='全产品容量', marker_color='#3b82f6'))
            fig2.add_trace(go.Bar(x=actual_data['年份'], y=actual_data['通信DC容量'],
                                  name='DC容量', marker_color='#10b981'))
            fig2.update_layout(title="实际数：全产品容量 vs DC容量", height=350,
                              barmode='group')
            st.plotly_chart(fig2, use_container_width=True)
        with col2:
            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(x=analysis_df['年份'], y=analysis_df['锐捷DC收入'],
                                      name='锐捷DC收入', line=dict(color='#f59e0b', width=2),
                                      mode='lines+markers'))
            fig3.add_trace(go.Scatter(x=analysis_df['年份'], y=analysis_df['锐捷开票金额'],
                                      name='锐捷开票金额', line=dict(color='#ef4444', width=2),
                                      mode='lines+markers'))
            fig3.update_layout(title="锐捷DC收入与开票金额", height=350)
            st.plotly_chart(fig3, use_container_width=True)

    # 竞争力指数看板
    st.markdown("### 📈 竞争力指数看板")
    if len(actual_data) >= 2 and '竞争力指数' in actual_data.columns:
        fig_comp = go.Figure()
        fig_comp.add_trace(go.Scatter(
            x=actual_data['年份'],
            y=actual_data['竞争力指数'],
            name='竞争力指数',
            line=dict(color='#8b5cf6', width=3),
            mode='lines+markers',
            marker=dict(size=10, symbol='diamond')
        ))
        # 添加基准线 1.0（份额持平线）
        fig_comp.add_hline(
            y=1.0, line_dash="dash", line_color="#9ca3af",
            annotation_text="份额持平线 (1.0)", annotation_position="right"
        )
        fig_comp.update_layout(
            title="实际数年份：竞争力指数趋势",
            xaxis_title="年份",
            yaxis_title="竞争力指数",
            hovermode='x unified',
            height=400,
            yaxis=dict(
                tickformat=".2f",
                zeroline=True,
                zerolinecolor="#e5e7eb"
            )
        )
        st.plotly_chart(fig_comp, use_container_width=True)

        # 竞争力指数解读
        latest_comp = actual_data.iloc[-1]['竞争力指数']
        prev_comp = actual_data.iloc[-2]['竞争力指数'] if len(actual_data) >= 2 else None
        comp_delta = round(latest_comp - prev_comp, 2) if prev_comp else 0

        comp_cols = st.columns(3)
        with comp_cols[0]:
            st.metric("最新年竞争力指数", f"{latest_comp:.2f}",
                      delta=f"{comp_delta:+.2f}", delta_color="inverse")
        with comp_cols[1]:
            trend = "份额提升" if latest_comp > 1 else ("份额下降" if latest_comp < 1 else "份额持平")
            st.metric("趋势判断", trend)
        with comp_cols[2]:
            max_comp = actual_data['竞争力指数'].max()
            max_year = actual_data[actual_data['竞争力指数'] == max_comp]['年份'].values[0]
            st.metric("历史最高", f"{max_comp:.2f}", delta=f"{max_year}年")
    else:
        st.info("⚠️ 实际数数据不足，暂无法展示竞争力指数")


# ============ 页面4: 参数设置 ============
elif page == "⚙️ 参数设置":
    st.markdown('<div class="main-header">⚙️ 参数设置</div>', unsafe_allow_html=True)

    st.markdown("### 路径配置")
    new_data_dir = st.text_input("数据源目录", config['data_dir'])
    new_output_dir = st.text_input("输出目录", config['output_dir'])

    st.markdown("### 计算参数")
    new_vat = st.number_input("增值税率", 0.01, 0.30, config['vat_rate'], 0.01, format="%.2f")

    # 动态自定义参数
    st.markdown("### ➕ 自定义计算参数")
    st.caption("点击下方按钮添加自定义参数，用于分析看板中的模拟计算。每个参数可绑定特定年份，绑定后仅影响该年份的数据。")

    init_custom_params(config)
    available_years = get_available_years()
    year_options = [None] + available_years
    year_labels = {None: "全部年份"}
    for y in available_years:
        year_labels[y] = str(y)

    # 渲染每个自定义参数
    for i, param in enumerate(st.session_state.custom_params):
        cols = st.columns([3, 2, 2, 2, 1])
        with cols[0]:
            param['name'] = st.text_input(
                "参数名称", param.get('name', ''),
                key=f"param_name_{i}", label_visibility="collapsed",
                placeholder="参数名称"
            )
        with cols[1]:
            ptype = st.selectbox(
                "类型", ['number', 'text'],
                index=0 if param.get('type', 'number') == 'number' else 1,
                key=f"param_type_{i}", label_visibility="collapsed"
            )
            param['type'] = ptype
        with cols[2]:
            if ptype == 'number':
                try:
                    default_val = float(param.get('value', 0))
                except:
                    default_val = 0.0
                param['value'] = st.number_input(
                    "值", value=default_val, step=0.01,
                    key=f"param_val_{i}", label_visibility="collapsed"
                )
            else:
                param['value'] = st.text_input(
                    "值", str(param.get('value', '')),
                    key=f"param_val_{i}", label_visibility="collapsed",
                    placeholder="参数值"
                )
        with cols[3]:
            current_year = param.get('year')
            year_idx = year_options.index(current_year) if current_year in year_options else 0
            selected_year = st.selectbox(
                "适用年份", year_options,
                index=year_idx,
                format_func=lambda x: year_labels.get(x, str(x)),
                key=f"param_year_{i}", label_visibility="collapsed"
            )
            param['year'] = selected_year
        with cols[4]:
            st.button("🗑️", key=f"param_del_{i}", on_click=remove_custom_param, args=(i,))

    col_add, _ = st.columns([1, 5])
    with col_add:
        st.button("➕ 添加参数", on_click=add_custom_param)

    st.markdown("### 产品配置")
    st.caption("产品关键词用于文件名识别，修改后影响文件分类")
    products_cfg = proc.PRODUCT_KEYWORDS.copy()
    edited_products = {}
    for product, keywords in products_cfg.items():
        kw_str = st.text_input(f"{product}", ", ".join(keywords))
        edited_products[product] = [k.strip() for k in kw_str.split(",") if k.strip()]

    # Sheet名配置
    st.markdown("### Sheet名映射")
    st.caption("Tracker文件中各产品对应的Sheet名称")
    sheet_cfg = proc.PRODUCT_SHEET_MAP.copy()
    edited_sheets = {}
    cols = st.columns(4)
    for i, (product, sheet) in enumerate(sheet_cfg.items()):
        with cols[i]:
            edited_sheets[product] = st.text_input(f"{product}", sheet, key=f"sheet_{product}")

    # 保存
    st.markdown("---")
    if st.button("💾 保存配置", type="primary"):
        # 过滤掉名称为空的参数
        valid_params = [p for p in st.session_state.custom_params if str(p.get('name', '')).strip()]
        new_config = {
            'data_dir': new_data_dir,
            'output_dir': new_output_dir,
            'vat_rate': new_vat,
            'custom_params': valid_params,
        }
        save_config(new_config)

        proc.DATA_DIR = new_data_dir
        proc.OUTPUT_DIR = new_output_dir
        proc.BACKUP_DIR = os.path.join(new_output_dir, '历史版本备份')
        proc.HISTORY_FILE = os.path.join(new_output_dir, '处理历史记录.json')
        proc.SOURCE_HISTORY_FILE = os.path.join(new_data_dir, '处理历史.json')
        proc.PRODUCT_KEYWORDS = edited_products
        proc.PRODUCT_SHEET_MAP = edited_sheets

        st.success("✅ 配置已保存！")
        st.info(f"共保存 {len(valid_params)} 个自定义参数，新配置将在下次数据处理时生效")

    # 重置历史记录
    st.markdown("### 🔄 历史记录管理")
    st.caption("清除历史记录后，下次处理将视为首次处理，重新处理所有文件")
    if st.button("🗑️ 清除处理历史记录"):
        history_path = proc.HISTORY_FILE
        if os.path.exists(history_path):
            os.remove(history_path)
            st.success("✅ 历史记录已清除")
        else:
            st.info("无历史记录文件")

    # 查看历史记录
    history = proc.load_history()
    if history.get('last_processed'):
        st.markdown("#### 当前历史记录")
        st.json({
            '上次处理时间': history['last_processed'],
            '已处理文件数': len(history.get('processed_files', []))
        })
