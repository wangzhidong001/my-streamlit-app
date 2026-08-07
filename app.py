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
    "💬 智能问答",
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

    # 调试：显示实际数据源路径（方便排查 40 vs 43 问题）
    st.caption(f"📂 实际数据源：`{proc.DATA_DIR}`  （来源：{proc.DATA_DIR_SOURCE}，共 {len(os.listdir(proc.DATA_DIR)) if os.path.isdir(proc.DATA_DIR) else '?'} 个条目）")

    # 扫描文件
    files = proc.scan_excel_files()
    if not os.path.isdir(proc.DATA_DIR):
        st.error(f"⚠️ 数据源目录不存在：`{proc.DATA_DIR}`\n\n请到「⚙️ 参数设置」中修改为正确路径，或确认桌面 `IDC数据文件` 文件夹未移动/重命名。")
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
                if combined is not None:
                    proc.save_combined_file(combined, proc.OUTPUT_DIR)
                else:
                    st.warning("⚠️ 全产品汇总为空，未生成合并文件。请检查处理日志确认各产品数据处理是否正常。")

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
                if combined is None or combined.empty:
                    st.warning("⚠️ 全产品汇总无数据，跳过通信DC分析计算。请检查上方处理日志确认各产品是否正常产出。")
                    analysis = pd.DataFrame()
                else:
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

    # 增值税率使用配置默认值（不在参数模拟区调整）
    sim_vat = config['vat_rate']
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

    # 生成带E后缀的显示年份（预测数据加E后缀）
    def _year_label(row):
        y_val = str(row['年份']).strip()
        dtype = str(row.get('数据类型', '')).strip()
        # 如果已经带E后缀则不再重复添加
        if y_val.endswith('E') or y_val.endswith('e'):
            return y_val
        if dtype == '预测':
            return y_val + 'E'
        return y_val
    chart_years = analysis_df.apply(_year_label, axis=1).tolist()

    # 图表
    st.markdown("### 📊 组合图：通信DC容量 & 锐捷DC份额")
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(x=chart_years, y=analysis_df['通信DC容量'],
               name='通信DC容量', marker_color='#5470c6'),
        secondary_y=False
    )
    fig.add_trace(
        go.Bar(x=chart_years, y=analysis_df['锐捷DC收入'],
               name='锐捷DC收入', marker_color='#91cc75'),
        secondary_y=False
    )
    fig.add_trace(
        go.Scatter(x=chart_years, y=analysis_df['锐捷DC份额'],
                   name='锐捷DC份额', line=dict(color='#ee6666', width=3),
                   mode='lines+markers'),
        secondary_y=True
    )
    fig.update_layout(
        title="通信DC容量与锐捷DC份额趋势（预测年份后缀为E）",
        hovermode='x unified',
        xaxis=dict(type='category'),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=450
    )
    fig.update_yaxes(title_text="容量 (CNY M)", secondary_y=False)
    fig.update_yaxes(title_text="份额 (%)", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)

    # 分析数据表
    st.markdown("### 📋 分析结果明细")

    display_df = analysis_df.copy()
    # 年份列：预测数据加E后缀
    display_df['年份'] = display_df.apply(_year_label, axis=1)
    # 移除锐捷DC容量列（不再显示）
    if '锐捷DC容量' in display_df.columns:
        display_df = display_df.drop(columns=['锐捷DC容量'])
    pct_cols = ['DC占全产品比例', '通信DC容量增速', '锐捷DC份额', '开票同比变动']
    for col in pct_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(
                lambda x: f"{x}%" if pd.notna(x) else "—"
            )
    # 增值税率：百分数值（13）→ 百分比（13%），直接加%后缀
    if '增值税率' in display_df.columns:
        display_df['增值税率'] = display_df['增值税率'].apply(
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
                              barmode='group', xaxis=dict(type='category'))
            st.plotly_chart(fig2, use_container_width=True)
        with col2:
            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(x=chart_years, y=analysis_df['锐捷DC收入'],
                                      name='锐捷DC收入', line=dict(color='#f59e0b', width=2),
                                      mode='lines+markers'))
            fig3.add_trace(go.Scatter(x=chart_years, y=analysis_df['锐捷开票金额'],
                                      name='锐捷开票金额', line=dict(color='#ef4444', width=2),
                                      mode='lines+markers'))
            fig3.update_layout(title="锐捷DC收入与开票金额（预测年份后缀为E）",
                               height=350, xaxis=dict(type='category'))
            st.plotly_chart(fig3, use_container_width=True)

    # 竞争力指数看板
    st.markdown("### 📈 竞争力指数看板")
    # 去掉半年数据（只保留纯年份格式，过滤 H1/H2/E 后缀）
    comp_data = actual_data.copy()
    import re as _re
    def _is_pure_year(y_val):
        s = str(y_val).strip()
        # 纯数字或纯年份格式（不含H、E等后缀）
        return s.isdigit() or (s.endswith('年') and s[:-1].isdigit())
    comp_data = comp_data[comp_data['年份'].apply(_is_pure_year)].reset_index(drop=True)
    # 过滤掉竞争力指数为空（无实际值）的年份
    if '竞争力指数' in comp_data.columns:
        comp_data = comp_data[comp_data['竞争力指数'].notna()].reset_index(drop=True)

    if len(comp_data) >= 2 and '竞争力指数' in comp_data.columns:
        # 年份转字符串，避免 Plotly 数值轴自动插值出 2022.5/2023.5 等半年刻度
        comp_years_str = comp_data['年份'].astype(str).tolist()
        fig_comp = go.Figure()
        fig_comp.add_trace(go.Scatter(
            x=comp_years_str,
            y=comp_data['竞争力指数'],
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
            title="实际数年份：竞争力指数趋势（不含半年）",
            xaxis_title="年份",
            xaxis=dict(
                type='category',
                tickmode='array',
                tickvals=comp_years_str,
                ticktext=comp_years_str,
                showgrid=False
            ),
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
        latest_comp = comp_data.iloc[-1]['竞争力指数']
        prev_comp = comp_data.iloc[-2]['竞争力指数'] if len(comp_data) >= 2 else None
        comp_delta = round(latest_comp - prev_comp, 2) if prev_comp else 0

        comp_cols = st.columns(3)
        with comp_cols[0]:
            st.metric("最新年竞争力指数", f"{latest_comp:.2f}",
                      delta=f"{comp_delta:+.2f}", delta_color="inverse")
        with comp_cols[1]:
            trend = "份额提升" if latest_comp > 1 else ("份额下降" if latest_comp < 1 else "份额持平")
            st.metric("趋势判断", trend)
        with comp_cols[2]:
            max_comp = comp_data['竞争力指数'].max()
            max_year = comp_data[comp_data['竞争力指数'] == max_comp]['年份'].values[0]
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


# ============ 页面5: 智能问答（多数据集版本） ============
elif page == "💬 智能问答":
    import re as _qa_re
    st.markdown('<div class="main-header">💬 智能问答</div>', unsafe_allow_html=True)

    # ============================================================
    # LLM 配置（智谱 glm-4-flash，OpenAI 兼容接口）
    # 作用：把用户的自然语言问题「归一化」成规则引擎能识别的规范问法，
    #       规则引擎/算数逻辑完全不动，只在最前面加一层翻译。
    # ============================================================
    import os as _os
    _ZHIPU_BASE = "https://open.bigmodel.cn/api/paas/v4"   # 用户提供的 base_url（已去掉结尾斜杠）
    _ZHIPU_MODEL = "glm-4-flash"

    if 'zhipu_api_key' not in st.session_state:
        st.session_state.zhipu_api_key = _os.getenv("ZHIPU_API_KEY", "")

    with st.expander("⚙️ 大模型(LLM)配置 — 用于自然语言问答", expanded=False):
        st.session_state.zhipu_api_key = st.text_input(
            "智谱 API Key（也可设置环境变量 ZHIPU_API_KEY）",
            value=st.session_state.zhipu_api_key,
            type="password",
            help="用于把口语化/同义的问题翻译成规则引擎可解析的查询；留空则跳过翻译、直接用原问题。"
        )
        st.caption(f"模型：{_ZHIPU_MODEL} ｜ 接口：{_ZHIPU_BASE}")

    def _build_schema_hint(datasets_map):
        """把已加载的数据集结构翻译成给大模型看的『说明书』"""
        lines = []
        for did, ds in datasets_map.items():
            cols = [str(c) for c in ds['df'].columns]
            yr = ""
            yc = ds.get('year_col')
            if yc and yc in ds['df'].columns:
                try:
                    yr = f"，年份范围 {ds['df'][yc].min()}~{ds['df'][yc].max()}"
                except Exception:
                    yr = ""
            lines.append(f"- 数据集[{did}]《{ds['name']}》类别={ds['category']}，列={cols}{yr}")
        return "\n".join(lines)

    def _llm_rewrite_question(question_text, datasets_map):
        """用大模型把自然语言问题改写成规则引擎易解析的规范问法；任何失败都回退到原问题。"""
        key = st.session_state.get('zhipu_api_key') or _os.getenv("ZHIPU_API_KEY", "")
        if not key:
            return question_text
        try:
            from openai import OpenAI
        except Exception:
            return question_text
        hint = _build_schema_hint(datasets_map)
        system = (
            "你是一个数据分析查询的『自然语言归一化器』。下面是可用的数据集及其列名。\n"
            "请把用户的中文问题，改写成一句【简洁、且只包含下面对齐词汇】的查询语句，"
            "以便后端的规则引擎能正确解析。\n\n"
            "**重要——根据问题类型选择指标词汇风格：**\n"
            "- 若问题提到具体产品（Switch/WLAN/Router/VCC/交换机/无线/路由器）或具体厂商（Cisco/华为/锐捷等），"
            "说明会查「全产品明细」表（列名为英文），此时收入类指标用「收入」或「营收」（不要用锐捷DC收入）。\n"
            "- 若问题是纯通信DC汇总（只提份额/容量/竞争力/税率/开票，无具体产品和厂商），"
            "说明会查「分析汇总」表（列名为中文），此时用「锐捷DC收入」「锐捷DC份额」等。\n\n"
            "规则：\n"
            "1) 指标只用这些词之一：收入/营收、人民币收入、美元收入、出货量、台数、份额、锐捷份额、"
            "锐捷DC收入、开票金额、DC容量、全产品容量、DC占比、DC容量增速、开票同比、竞争力指数、税率。\n"
            "2) 产品/技术只用：Switch/交换/交换机、WLAN/无线、Router/路由/路由器、VCC/虚拟客户端、通信DC。\n"
            "3) 厂商可用中英文名（锐捷/Ruijie、华为/Huawei、华三/H3C、思科/Cisco 等）。\n"
            "4) 聚合意图用词：排名/最大/最小/平均/总和/趋势/对比/列出/分布/各。\n"
            "5) 年份用 20xx 或 Q1~Q4，数据性质用 实际/预测。\n"
            "6) 只输出改写后的一句话，不要解释，不要多余内容。\n\n"
            "可用数据集：\n" + hint
        )
        try:
            client = OpenAI(api_key=key, base_url=_ZHIPU_BASE, timeout=20)
            resp = client.chat.completions.create(
                model=_ZHIPU_MODEL,
                temperature=0,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": question_text},
                ],
            )
            out = (resp.choices[0].message.content or "").strip()
            return out if out else question_text
        except Exception:
            return question_text

    # ============================================================
    # 阶段A：加载所有加工生成的结果表
    # ============================================================
    analysis_path = os.path.join(proc.OUTPUT_DIR, 'IDC分析结果.xlsx')
    combined_path = os.path.join(proc.OUTPUT_DIR, 'IDC全产品数据.xlsx')
    forecast_detail_path = os.path.join(proc.OUTPUT_DIR, '2026年预测数明细.xlsx')
    product_tables = {
        'Switch': os.path.join(proc.OUTPUT_DIR, 'Switch.xlsx'),
        'WLAN':   os.path.join(proc.OUTPUT_DIR, 'WLAN.xlsx'),
        'Router': os.path.join(proc.OUTPUT_DIR, 'Router.xlsx'),
        'VCC':    os.path.join(proc.OUTPUT_DIR, 'VCC.xlsx'),
    }

    # --- 1) 分析汇总类（分析看板/图表使用的汇总数据）---
    datasets = {}  # key: dataset_id, value: {name, df, category, description}

    # 分析结果表（通信DC汇总，用户参数调整后优先用 session 中的模拟版本）
    analysis_df = None
    use_simulated = False
    if 'analysis_df_simulated' in st.session_state and st.session_state.analysis_df_simulated is not None and len(st.session_state.analysis_df_simulated) > 0:
        analysis_df = st.session_state.analysis_df_simulated.copy()
        use_simulated = True
    elif os.path.exists(analysis_path):
        try:
            analysis_df = pd.read_excel(analysis_path, sheet_name=0)
        except Exception:
            analysis_df = None

    if analysis_df is not None and len(analysis_df) > 0:
        datasets['analysis'] = {
            'name': 'IDC分析结果（通信DC汇总）',
            'category': '分析汇总',
            'description': '通信DC维度汇总指标：份额、收入、开票、容量、竞争力指数、税率等（通信DC分析结果表）',
            'df': analysis_df,
            'year_col': '年份',
            'data_type_col': '数据类型',
        }

    # --- 2) 全产品明细类 ---
    if os.path.exists(combined_path):
        try:
            cdf = pd.read_excel(combined_path, sheet_name=0)
            if len(cdf) > 0:
                datasets['combined'] = {
                    'name': 'IDC全产品数据（全明细）',
                    'category': '全产品明细',
                    'description': '所有产品/行业/季度/厂商的全量明细，支持按产品/厂商/行业/部署方式等维度查询收入/出货量',
                    'df': cdf,
                    'year_col': 'Year',
                    'data_type_col': '实际/预测',
                }
        except Exception:
            pass

    if os.path.exists(forecast_detail_path):
        try:
            fdf = pd.read_excel(forecast_detail_path, sheet_name=0)
            if len(fdf) > 0:
                datasets['forecast2026'] = {
                    'name': '2026年预测数明细',
                    'category': '全产品明细',
                    'description': '2026年预测数据分行业/产品口径 + 各产品维度（Switch/WLAN/Router）',
                    'df': fdf,
                    'year_col': 'Year',
                    'data_type_col': '实际/预测',
                }
        except Exception:
            pass

    # --- 3) 单产品明细类（Switch/WLAN/Router/VCC 独立加工表）---
    for pt, ppath in product_tables.items():
        if os.path.exists(ppath):
            try:
                pdf = pd.read_excel(ppath, sheet_name=0)
                if len(pdf) > 0:
                    datasets['p_' + pt.lower()] = {
                        'name': f'{pt}产品加工数据',
                        'category': '单产品明细',
                        'description': f'{pt}产品独立加工结果：按年份/季度/厂商/部署/行业汇总收入与出货量',
                        'df': pdf,
                        'year_col': 'Year',
                        'data_type_col': '实际/预测',
                    }
            except Exception:
                pass

    # ============================================================
    # 阶段B：数据集选择面板 + 加载状态
    # ============================================================
    if len(datasets) == 0:
        st.warning("⚠️ 尚未找到任何加工结果。请先到「🔄 数据处理」页面执行「▶️ 一键执行数据处理」生成结果表后再使用问答。")
    else:
        # 数据状态摘要
        cat_count = {}
        for ds in datasets.values():
            c = ds['category']
            cat_count[c] = cat_count.get(c, 0) + 1
        status_parts = [f"{c}{n}张表" for c, n in cat_count.items()]
        st.caption("ℹ️ 已加载所有加工结果：" + "，".join(status_parts) + ("｜分析结果使用看板模拟数据" if use_simulated else "｜分析结果来自磁盘文件"))

        # 数据集展开面板（默认折叠，用户可查看已加载哪些表）
        with st.expander(f"📚 已加载数据集（共 {len(datasets)} 张结果表）- 点击查看详情", expanded=False):
            summary_rows = []
            for did, ds in datasets.items():
                dfc = ds['df']
                yc = ds['year_col']
                dtc = ds['data_type_col']
                y_min = dfc[yc].min() if yc in dfc.columns else '—'
                y_max = dfc[yc].max() if yc in dfc.columns else '—'
                num_cols = len(dfc.columns)
                actual_n = len(dfc[dfc[dtc] == '实际']) if dtc in dfc.columns else '—'
                forecast_n = len(dfc[dfc[dtc] == '预测']) if dtc in dfc.columns else '—'
                summary_rows.append({
                    '数据集ID': did,
                    '分类': ds['category'],
                    '名称': ds['name'],
                    '行数': f"{len(dfc):,}",
                    '列数': num_cols,
                    '年份范围': f"{y_min} ~ {y_max}",
                    '实际行数': str(actual_n),
                    '预测行数': str(forecast_n),
                })
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # ============================================================
        # 阶段C：列/指标/维度/产品 别名映射
        # ============================================================
        # 分析汇总类列别名（指标名关键词 -> 列名）
        ANALYSIS_COLUMN_ALIAS = {
            # 份额类
            '份额': '锐捷DC份额', '锐捷份额': '锐捷DC份额', '市场份额': '锐捷DC份额', '锐捷DC份额': '锐捷DC份额',
            # 收入/金额类
            '锐捷收入': '锐捷DC收入', '收入': '锐捷DC收入', '锐捷DC收入': '锐捷DC收入', 'DC收入': '锐捷DC收入',
            '开票金额': '锐捷开票金额', '开票': '锐捷开票金额', '锐捷开票': '锐捷开票金额', '锐捷开票金额': '锐捷开票金额',
            # 容量类
            '通信DC容量': '通信DC容量', 'DC容量': '通信DC容量', '容量': '通信DC容量',
            '通信全产品容量': '通信全产品容量', '全产品容量': '通信全产品容量',
            # 比例/增速类
            'DC占比': 'DC占全产品比例', 'DC占全产品比例': 'DC占全产品比例', '占比': 'DC占全产品比例',
            'DC容量增速': '通信DC容量增速', '容量增速': '通信DC容量增速', '通信DC容量增速': '通信DC容量增速',
            '开票同比': '开票同比变动', '开票同比变动': '开票同比变动', '开票增速': '开票同比变动',
            # 指数/税率类
            '竞争力指数': '竞争力指数', '竞争指数': '竞争力指数',
            '增值税率': '增值税率', '税率': '增值税率',
        }

        # 全产品/单产品明细类列别名
        DETAIL_COLUMN_ALIAS = {
            # 收入类：收入 / 营收 / 人民币 / CNY / Revenue / 金额 全部映射到 Vendor Revenue (CNY M)
            '收入': 'Vendor Revenue (CNY M)',
            '营收': 'Vendor Revenue (CNY M)',
            '人民币': 'Vendor Revenue (CNY M)',
            '人民币收入': 'Vendor Revenue (CNY M)',
            'CNY': 'Vendor Revenue (CNY M)',
            'CNY收入': 'Vendor Revenue (CNY M)',
            'Revenue': 'Vendor Revenue (CNY M)',
            '金额': 'Vendor Revenue (CNY M)',
            # 美元收入单独保留
            '美元收入': 'Vendor Revenue (USD M)',
            'USD收入': 'Vendor Revenue (USD M)',
            '美元': 'Vendor Revenue (USD M)',
            # 出货量/销量类
            '出货量': 'Units',
            '销量': 'Units',
            '台数': 'Units',
            # 维度类
            '产品': 'Technology',
            '技术': 'Technology',
            '厂商': 'Vendor',
            '品牌': 'Vendor',
            '供应商': 'Vendor',
            '部署': 'Deployment',
            '部署方式': 'Deployment',
            '行业': '行业大类',
            '行业大类': '行业大类',
            '细分行业': '行业细分',
            '行业细分': '行业细分',
            '国家': 'Country',
            '区域': 'Region',
            '季度': 'Quarter',
            '半年': 'Half Year',
            '产品类型': '二级产品分类',
            '二级分类': '二级产品分类',
            '产品大类': 'Product',
            '产品详情': 'Product Detail',
        }

        # 产品/技术关键词 -> 对应数据集ID优先级（全部从全产品明细中获取，去掉单产品表优先）
        PRODUCT_DATASET_PRIORITY = {
            'switch': ['combined', 'forecast2026'],
            '交换': ['combined', 'forecast2026'],
            '交换机': ['combined', 'forecast2026'],
            'wlan': ['combined', 'forecast2026'],
            'wifi': ['combined', 'forecast2026'],
            '无线': ['combined', 'forecast2026'],
            'router': ['combined', 'forecast2026'],
            '路由': ['combined', 'forecast2026'],
            '路由器': ['combined', 'forecast2026'],
            'vcc': ['combined', 'forecast2026'],
            '虚拟客户端': ['combined', 'forecast2026'],
            '虚拟': ['combined', 'forecast2026'],
            '通信dc': ['analysis'],
            'dc汇总': ['analysis'],
            '通信DC': ['analysis'],
            '锐捷': ['analysis'],
        }

        # ============================================================
        # Vendor（厂商）英文名 ↔ 中文名 双向映射 + 翻译函数
        # ============================================================
        VENDOR_CN_TO_EN = {
            '锐捷': 'Ruijie', '瑞捷': 'Ruijie', '锐捷网络': 'Ruijie',
            '华为': 'Huawei', '华三': 'H3C', '新华三': 'H3C',
            '思科': 'Cisco', '思杰': 'Cisco',
            '瞻博': 'Juniper', 'Juniper瞻博': 'Juniper', ' Juniper': 'Juniper',
            '阿鲁巴': 'Aruba', 'Aruba阿鲁巴': 'Aruba',
            '惠普': 'HP', '慧与': 'HPE', 'HPE': 'HPE',
            '飞塔': 'Fortinet', 'Fortinet飞塔': 'Fortinet',
            '威睿': 'VMware', 'VMware威睿': 'VMware',
            '浪潮': 'Inspur', '联想': 'Lenovo',
            '中兴': 'ZTE', '烽火': 'FiberHome', '大唐': 'Datang',
            '深信服': 'Sangfor', '天融信': 'Topsec', '绿盟': 'NSFOCUS',
            '奇安信': 'Qianxin', '启明星辰': 'Venustech',
            '迈普': 'Maipu', '星网锐捷': 'Ruijie',
            '博科': 'Brocade', 'Arista': 'Arista',
            '戴尔': 'Dell', 'EMC': 'Dell EMC',
            'Oracle': 'Oracle', '微软': 'Microsoft',
            'IBM': 'IBM', 'Redhat': 'RedHat', '红帽': 'RedHat',
            'SUSE': 'SUSE', 'Citrix': 'Citrix', '思杰系统': 'Citrix',
            'Palo Alto': 'Palo Alto', '派拓网络': 'Palo Alto',
            'CheckPoint': 'Check Point', '捷邦': 'Check Point',
            'F5': 'F5 Networks', 'F5 Networks': 'F5 Networks',
            'A10': 'A10 Networks', 'Radware': 'Radware',
        }
        VENDOR_EN_TO_CN = {
            'Ruijie': '锐捷', '锐捷': '锐捷', '锐捷网络': '锐捷',
            'Huawei': '华为', '华为': '华为',
            'H3C': '新华三', '新华三': '新华三', '华三': '新华三',
            'Cisco': '思科', '思科': '思科',
            'Juniper': '瞻博', '瞻博': '瞻博',
            'Aruba': '阿鲁巴', '阿鲁巴': '阿鲁巴',
            'HP': '惠普', 'HPE': '慧与', '惠普': '惠普', '慧与': '慧与',
            'Fortinet': '飞塔', '飞塔': '飞塔',
            'VMware': '威睿', 'Vmware': '威睿', '威睿': '威睿',
            'Inspur': '浪潮', '浪潮': '浪潮',
            'Lenovo': '联想', '联想': '联想',
            'ZTE': '中兴', '中兴': '中兴',
            'FiberHome': '烽火', '烽火': '烽火',
            'Datang': '大唐', '大唐': '大唐',
            'Sangfor': '深信服', '深信服': '深信服',
            'Topsec': '天融信', '天融信': '天融信',
            'NSFOCUS': '绿盟', '绿盟': '绿盟',
            'Qianxin': '奇安信', '奇安信': '奇安信',
            'Venustech': '启明星辰', '启明星辰': '启明星辰',
            'Maipu': '迈普', '迈普': '迈普',
            'Brocade': '博科', '博科': '博科',
            'Arista': 'Arista',
            'Dell': '戴尔', 'Dell EMC': '戴尔EMC', '戴尔': '戴尔',
            'Oracle': 'Oracle',
            'Microsoft': '微软', '微软': '微软',
            'IBM': 'IBM',
            'RedHat': '红帽', 'Redhat': '红帽', '红帽': '红帽',
            'SUSE': 'SUSE',
            'Citrix': '思杰', '思杰': '思杰',
            'Palo Alto': '派拓网络', '派拓网络': '派拓网络',
            'Check Point': '捷邦', 'CheckPoint': '捷邦', '捷邦': '捷邦',
            'F5 Networks': 'F5', 'F5': 'F5',
            'A10 Networks': 'A10', 'A10': 'A10',
            'Radware': 'Radware',
        }

        def _vendor_translate(v):
            """把厂商英文名翻译成中文，已中文则直接返回"""
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return "—"
            s = str(v).strip()
            if not s: return "—"
            # 先精确匹配
            if s in VENDOR_EN_TO_CN:
                return VENDOR_EN_TO_CN[s]
            # 去掉大小写差异匹配（如 vmware VMware）
            low = s.lower()
            for k, cn in VENDOR_EN_TO_CN.items():
                if k.lower() == low:
                    return cn
            # 子串包含（如 Ruijie Networks → 锐捷）
            for k, cn in VENDOR_EN_TO_CN.items():
                if len(k) >= 3 and (k.lower() in low or low in k.lower()):
                    return cn
            return s

        def _vendor_to_english_keywords(text):
            """从问题文本中抽取出厂商，返回英文名关键词列表（用于过滤Vendor列）"""
            results = []
            t_low = text.lower()
            # 中文名 → 英文名
            for cn, en in VENDOR_CN_TO_EN.items():
                if cn in text:
                    results.append(en)
                    results.append(cn)
            # 英文名（不区分大小写）→ 保留英文原样
            for en in VENDOR_EN_TO_CN.keys():
                if len(en) >= 2 and en.lower() in t_low:
                    results.append(en)
                    if en in VENDOR_EN_TO_CN:
                        results.append(VENDOR_EN_TO_CN[en])
            # 去重保序
            seen = set()
            uniq = []
            for r in results:
                if r and r not in seen:
                    seen.add(r); uniq.append(r)
            return uniq

        # ============================================================
        # 模糊列名匹配（关键词→列名），别名匹配失败时兜底
        # ============================================================
        FUZZY_COLUMN_RULES = [
            # 收入类（模糊命中：收入/金额/Revenue/CNY/人民币 → 优先 Vendor Revenue (CNY M)）
            {'keywords': ['收入', '营收', '人民币', 'CNY', 'revenue', '金额'],
             'target': 'Vendor Revenue (CNY M)', 'display': '人民币收入(百万元)'},
            # 美元收入
            {'keywords': ['美元', 'USD', '美金', 'usd'],
             'target': 'Vendor Revenue (USD M)', 'display': '美元收入(百万元)'},
            # 出货量/台数/销量
            {'keywords': ['出货量', '销量', '台数', 'units', '数量', '发货量'],
             'target': 'Units', 'display': '出货量'},
            # 份额类
            {'keywords': ['份额', 'share', '市占', '市场占比'],
             'target': '锐捷DC份额', 'display': '锐捷DC份额'},
            # 开票金额
            {'keywords': ['开票', '发票'],
             'target': '锐捷开票金额', 'display': '锐捷开票金额'},
            # 容量类
            {'keywords': ['容量'],
             'target': '通信DC容量', 'display': '通信DC容量'},
            # 占比
            {'keywords': ['占比', '比例', 'percent'],
             'target': 'DC占全产品比例', 'display': 'DC占全产品比例'},
            # 增速/同比
            {'keywords': ['增速', '增长率'],
             'target': '通信DC容量增速', 'display': '通信DC容量增速'},
            # 指数
            {'keywords': ['竞争力', '指数'],
             'target': '竞争力指数', 'display': '竞争力指数'},
            # 税率
            {'keywords': ['税率', '增值税'],
             'target': '增值税率', 'display': '增值税率'},
        ]

        def _fuzzy_match_column(text, available_cols):
            """当别名精确匹配未命中时，使用模糊规则匹配目标列"""
            t_low = text.lower()
            for rule in FUZZY_COLUMN_RULES:
                for kw in rule['keywords']:
                    if kw.lower() in t_low:
                        if rule['target'] in available_cols:
                            return rule['target'], rule['display']
                        # 其他列名中含 target 的变体也接受（比如分析汇总表里叫「锐捷DC收入」而非Vendor Revenue…）
                        for col in available_cols:
                            if rule['target'] in col or col in rule['target']:
                                return col, rule['display']
            return None, None

        # ============================================================
        # 辅助函数
        # ============================================================
        def _extract_years(text):
            years = set()
            for m in _qa_re.findall(r'(20\d{2})(?:[Ee年QqHh]?)?', text):
                years.add(int(m))
            for m in _qa_re.findall(r'(?<!\d)(\d{2})年', text):
                y = int(m)
                if 0 <= y <= 99:
                    years.add(2000 + y if y < 50 else 1900 + y)
            return sorted(years)

        def _extract_quarters(text):
            return sorted(set(_qa_re.findall(r'Q[1-4]', text, flags=_qa_re.I)))

        def _extract_keywords(text, alias_map):
            """从文本中按别名映射抽取出匹配的列/指标"""
            matched = []
            sorted_aliases = sorted(alias_map.keys(), key=lambda x: -len(x))
            for alias in sorted_aliases:
                if alias and alias in text:
                    matched.append((alias, alias_map[alias]))
            return matched  # list of (display_name, real_col)

        def _detect_agg(text):
            if any(k in text for k in ['最大', '最高', '最多', '峰值', 'max', 'MAX']):
                return 'max'
            if any(k in text for k in ['最小', '最低', '最少', 'min', 'MIN']):
                return 'min'
            if any(k in text for k in ['平均', '均值', 'avg', 'AVG', 'mean']):
                return 'mean'
            if any(k in text for k in ['总和', '合计', '总计', 'sum', 'SUM', '一共', '总共', '累计']):
                return 'sum'
            if any(k in text for k in ['计数', '多少条', '条数', '行数', 'count']):
                return 'count'
            if any(k in text for k in ['趋势', '变化', '走势', '同比', '增长', '下降']):
                return 'trend'
            if any(k in text for k in ['对比', '比较', '差别', '差异', 'vs', 'VS', '哪个', '分别']):
                return 'compare'
            if any(k in text for k in ['排名', '排第几', '排序', 'top', 'TOP', '前几']):
                return 'rank'
            if any(k in text for k in ['列出', '列举', '展示', '显示', '所有', '明细', '全部', '每']):
                return 'list'
            if any(k in text for k in ['分布', '占比结构', '结构', '分组', '各', '每个']):
                return 'group'
            return 'value'

        def _detect_data_type_filter(text):
            """返回 '实际' / '预测' / None"""
            if any(k in text for k in ['实际', '真实', '已发生', '历史']):
                return '实际'
            if any(k in text for k in ['预测', '预计', '未来', '预估', ' forecast', '预测数']):
                return '预测'
            return None

        def _match_dataset_by_product_keyword(text):
            """根据产品关键词返回推荐的数据集ID列表（按优先级）"""
            t_low = text.lower()
            for kw, priority in PRODUCT_DATASET_PRIORITY.items():
                if kw.lower() in t_low:
                    return [d for d in priority if d in datasets]
            return None

        # 收入类指标关键词（命中任意一个即判定为"收入类查询"，用于路由和字段映射）
        REVENUE_KEYWORDS = ['收入', '营收', '人民币', 'CNY', '金额', 'Revenue']

        def _is_revenue_query(text):
            """判断问题是否涉及收入类指标（不区分大小写）"""
            t_low = text.lower()
            for kw in REVENUE_KEYWORDS:
                if kw.lower() in t_low:
                    return True
            return False

        def _has_product_keyword(text):
            """判断问题是否包含产品类关键词（用于路由决策）"""
            t_low = text.lower()
            # 检查PRODUCT_DATASET_PRIORITY中的产品关键词（排除通信DC/锐捷这类汇总项）
            product_only_kws = [
                'switch', '交换', '交换机', 'wlan', 'wifi', '无线',
                'router', '路由', '路由器', 'vcc', '虚拟客户端', '虚拟'
            ]
            for kw in product_only_kws:
                if kw.lower() in t_low:
                    return True
            # 若combined数据集存在，再从Technology列动态匹配产品名
            if 'combined' in datasets:
                cdf = datasets['combined']['df']
                if 'Technology' in cdf.columns:
                    tech_values = cdf['Technology'].dropna().unique().tolist()
                    for tv in tech_values:
                        if str(tv).lower() in t_low:
                            return True
            return False

        def _resolve_dataset(question_text, detected_metric_type):
            """根据问题自动路由到最合适的数据集（全部产品查询优先走全产品明细）"""
            # 1) 用户显式提到数据集名
            for did, ds in datasets.items():
                if ds['name'] in question_text:
                    return did
            # 2) 产品类关键词 → 全部从全产品明细（combined）获取，按Technology字段过滤产品
            if _has_product_keyword(question_text):
                # 优先 combined（全产品明细），其次 forecast2026（预测明细）
                if 'combined' in datasets:
                    return 'combined'
                if 'forecast2026' in datasets:
                    return 'forecast2026'
            # 3) 收入类查询（无产品关键词但命中收入指标）→ 优先全产品明细
            if _is_revenue_query(question_text):
                if 'combined' in datasets:
                    return 'combined'
                if 'forecast2026' in datasets:
                    return 'forecast2026'
            # 4) 产品关键词路由兜底（保持原有匹配逻辑但优先级已全产品优先）
            pds = _match_dataset_by_product_keyword(question_text)
            if pds:
                return pds[0]
            # 5) 指标类型路由
            if detected_metric_type == 'analysis_metric':
                if 'analysis' in datasets:
                    return 'analysis'
            # 6) 默认：如果命中明细类指标则选全产品表，否则选 analysis 或 combined
            if detected_metric_type == 'detail_metric':
                if 'combined' in datasets:
                    return 'combined'
                if 'forecast2026' in datasets:
                    return 'forecast2026'
            # 7) 兜底优先级（combined 放在单产品表之前）
            fallback_order = ['analysis', 'combined', 'forecast2026', 'p_switch', 'p_wlan', 'p_router', 'p_vcc']
            for d in fallback_order:
                if d in datasets:
                    return d
            return list(datasets.keys())[0]

        def _fmt(v, col_name_hint):
            """通用数值格式化"""
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return "—"
            # 百分比类列
            pct_suffix = {'份额', '占比', '比例', '增速', '变动', '税率', '指数'}
            is_pct = any(s in str(col_name_hint) for s in pct_suffix)
            if is_pct:
                return f"{v}%"
            # 金额类列（人民币）
            money_suffix = {'CNY', '开票金额', '人民币', '收入'}
            if any(s in str(col_name_hint) for s in money_suffix) or (isinstance(v, float) and abs(v) >= 100):
                if isinstance(v, (int, float)) and abs(v) >= 1000:
                    return f"{v:,.2f}"
                if isinstance(v, float):
                    return f"{v:.2f}"
                return str(v)
            if isinstance(v, float):
                return f"{v:.2f}"
            return str(v)

        def _apply_basic_filters(df, ds_meta, years, dtype_filter, quarters, extra_dim_filter):
            """对明细类数据做基础过滤：年份、实际/预测、季度、维度过滤"""
            dfc = df.copy()
            yc = ds_meta['year_col']
            dtc = ds_meta['data_type_col']
            if years and yc in dfc.columns:
                dfc = dfc[dfc[yc].astype(int).isin(years)]
            if dtype_filter and dtc in dfc.columns:
                # 兼容实际/预测 或 实际数/预测数
                dfc = dfc[dfc[dtc].astype(str).str.contains(dtype_filter, na=False)]
            if quarters and 'Quarter' in dfc.columns:
                dfc = dfc[dfc['Quarter'].astype(str).str.upper().isin([q.upper() for q in quarters])]
            # 额外维度过滤（厂商/行业/部署等）
            for col_real, values in extra_dim_filter.items():
                if col_real in dfc.columns and values:
                    pat = "|".join([_qa_re.escape(v) for v in values])
                    dfc = dfc[dfc[col_real].astype(str).str.contains(pat, case=False, na=False)]
            return dfc

        def _build_answer_multi(question_text, datasets_map):
            """新版问答：多数据集回答"""
            q = question_text.strip()
            if not q:
                return "请输入问题后点击「提问」。", None, None

            years = _extract_years(q)
            quarters = _extract_quarters(q)
            dtype_filter = _detect_data_type_filter(q)
            agg = _detect_agg(q)

            # 抽取指标类匹配（分析类）
            analysis_matches = _extract_keywords(q, ANALYSIS_COLUMN_ALIAS)
            # 抽取明细类指标匹配
            detail_matches = _extract_keywords(q, DETAIL_COLUMN_ALIAS)
            # 抽取明细类维度过滤匹配（厂商/行业/部署等 → 用于过滤，不是数值指标）
            dim_only_keys = ['Vendor', 'Deployment', '行业大类', '行业细分', 'Country', 'Region', 'Technology', '二级产品分类', 'Product', 'Product Detail']
            dim_filter_values = {}  # col_real -> list of keywords
            metric_col_real = None
            metric_display = None

            # 指标优先：若同时命中分析类和明细类指标，看产品关键词选哪类数据集
            if analysis_matches:
                metric_display, metric_col_real = analysis_matches[0]
                detected_metric_type = 'analysis_metric'
            elif detail_matches:
                # 区分数值指标 vs 维度（Vendor/Deployment等非数值只做过滤，不算指标）
                for disp, real in detail_matches:
                    if real in dim_only_keys:
                        # 维度过滤关键词，记到 filter
                        if real not in dim_filter_values:
                            dim_filter_values[real] = []
                        dim_filter_values[real].append(disp)
                    else:
                        # 数值指标（收入/出货量等）
                        metric_display = disp
                        metric_col_real = real
                        detected_metric_type = 'detail_metric'
                        break
                if metric_col_real is None:
                    # 没找到数值指标 → 模糊匹配兜底，仍未命中则默认 人民币收入
                    fcol, fdisp = _fuzzy_match_column(q, list(DETAIL_COLUMN_ALIAS.values()))
                    if fcol:
                        metric_col_real = fcol
                        metric_display = fdisp if fdisp else '指标'
                    else:
                        metric_col_real = 'Vendor Revenue (CNY M)'
                        metric_display = '人民币收入'
                    detected_metric_type = 'detail_metric'
            else:
                # 没匹配到任何指标 → 先用模糊匹配尝试一次，仍未命中再判断帮助/概览意图
                fcol, fdisp = None, None
                # 先尝试明细类
                fcol, fdisp = _fuzzy_match_column(q, list(DETAIL_COLUMN_ALIAS.values()))
                if fcol:
                    metric_col_real = fcol
                    metric_display = fdisp
                    detected_metric_type = 'detail_metric'
                else:
                    # 再尝试分析汇总类
                    acol, adisp = _fuzzy_match_column(q, list(ANALYSIS_COLUMN_ALIAS.values()))
                    if acol:
                        metric_col_real = acol
                        metric_display = adisp
                        detected_metric_type = 'analysis_metric'
                # 模糊匹配仍未命中 → 再走帮助/概览或未识别提示
                if metric_col_real is None:
                    if agg in ['list', 'value', 'trend', 'group'] or any(k in q for k in ['有什么', '能问什么', '帮助', '怎么用', 'help']):
                        # 帮助 / 概览
                        if any(k in q for k in ['有什么', '能问什么', '帮助', '怎么用', 'help', '使用说明']):
                            help_msg = (
                                "### 🤖 可以查询这些加工结果表（按类别）：\n\n"
                                "**📊 分析汇总类（通信DC汇总指标）**：\n"
                                "  指标：锐捷份额 / 锐捷DC收入 / 开票金额 / DC容量 / 全产品容量 / DC占比 / DC容量增速 / 开票同比 / 竞争力指数 / 税率\n"
                                "  例：`2025年锐捷份额是多少？` `哪年收入最高？` `锐捷份额趋势？`\n\n"
                                "**📦 全产品明细类**（IDC全产品数据 / 2026年预测数明细）：\n"
                                "  指标：人民币收入 / 美元收入 / 出货量 / 台数\n"
                                "  维度过滤（可组合）：产品 / 厂商 / 行业 / 部署方式 / 季度 / 半年 / 国家 / 区域\n"
                                "  例：`2026年Switch锐捷人民币收入？` `Cisco 2025年Router收入？` `各厂商WLAN收入排名？`\n\n"
                                "**🧩 单产品明细类**（Switch / WLAN / Router / VCC 独立加工表）：\n"
                                "  例：`Switch各部署方式收入分布？` `VCC行业大类结构？` `2024Q1 WLAN出货量？`\n\n"
                                "**统计类**：最大/最高 / 最低 / 平均 / 总和 / 排名 / 趋势 / 对比 / 分布 / 列表明细\n\n"
                                f"**当前已加载数据集：{len(datasets_map)} 张**"
                            )
                            return help_msg, None, None
                        # 概览数据
                        overviews = []
                        for did, ds in datasets_map.items():
                            overviews.append(f"- **{ds['name']}**（{ds['category']}）：{len(ds['df']):,} 行 × {len(ds['df'].columns)} 列")
                        ov = "### 📚 已加载加工结果表概览\n" + "\n".join(overviews) + "\n\n👉 试试问：`2025年锐捷份额？` / `2026年Switch锐捷收入？` / `各厂商Router收入排名？`"
                        # 附带小表格预览（选combined前10条年份/产品/收入/厂商）
                        preview_df = None
                        if 'combined' in datasets_map:
                            cdf = datasets_map['combined']['df']
                            preview_cols = [c for c in ['Year', 'Technology', 'Vendor', 'Vendor Revenue (CNY M)', '实际/预测'] if c in cdf.columns]
                            if preview_cols:
                                preview_df = cdf[preview_cols].head(10).copy()
                                preview_df = preview_df.rename(columns={
                                    'Year': '年份', 'Technology': '产品', 'Vendor': '厂商',
                                    'Vendor Revenue (CNY M)': '人民币收入(百万元)', '实际/预测': '实际/预测'
                                })
                        return ov, preview_df, None
                    else:
                        # 其他未识别场景
                        return (
                            "❓ 未识别到具体指标名称。\n\n"
                            "请在问题中明确提到：\n"
                            "① **要查询的指标**（如：收入 / 出货量 / 份额 / 开票金额 / DC容量 等）\n"
                            "② **（可选）产品/技术范围**（如：Switch / WLAN / Router / VCC / 通信DC）\n"
                            "③ **（可选）过滤维度**（如：锐捷 / Cisco / 金融行业 / 企业部署 / 2025年 / Q1 / 实际 / 预测）"
                        ), None, None

            # 路由到数据集
            did = _resolve_dataset(q, detected_metric_type)
            ds = datasets_map[did]
            df = ds['df']

            # ===== 列名自动校正 =====
            # 解决 ANALYSIS/DETAIL 两套别名映射到不同列名、但路由到了"另一套"数据集的问题。
            # 例：用户问"收入"→ANALYSIS别名先命中→得到'锐捷DC收入'→却路由到了全产品明细表（列名为英文）
            actual_cols = list(df.columns)
            if metric_col_real and metric_col_real not in actual_cols:
                remapped = None
                # 1) 按原 display 名从 DETAIL 别名中找目标数据集里实际存在的列
                for alias, real_col in DETAIL_COLUMN_ALIAS.items():
                    if real_col in actual_cols and metric_display and (alias == metric_display or metric_display in alias or alias in metric_display):
                        remapped = (real_col, alias)
                        break
                # 2) 从 ANALYSIS 别名中找
                if not remapped:
                    for alias, real_col in ANALYSIS_COLUMN_ALIAS.items():
                        if real_col in actual_cols and metric_display and (alias == metric_display or metric_display in alias or alias in metric_display):
                            remapped = (real_col, alias)
                            break
                # 3) 兜底：模糊匹配——在实际列名中找与 display 名有包含关系的
                if not remapped and metric_display:
                    disp_low = metric_display.lower()
                    for ac in actual_cols:
                        ac_str = str(ac).lower()
                        if disp_low in ac_str or ac_str in disp_low:
                            remapped = (ac, metric_display)
                            break
                if remapped:
                    metric_col_real, metric_display = remapped

            # 分析汇总类 vs 明细类 分支处理
            if did == 'analysis' or ds['category'] == '分析汇总':
                # ===== 分析汇总类：沿用之前成熟的回答逻辑 =====
                yc = ds['year_col']
                dtc = ds['data_type_col']
                df_use = df.copy()
                if dtype_filter and dtc in df_use.columns:
                    df_use = df_use[df_use[dtc].astype(str).str.contains(dtype_filter, na=False)].reset_index(drop=True)
                if metric_col_real not in df_use.columns:
                    return f"⚠️ 数据集「{ds['name']}」中不存在指标列「{metric_display}」。", None, did
                df_valid = df_use[df_use[metric_col_real].notna()].reset_index(drop=True)
                if len(df_valid) == 0:
                    return f"⚠️ 数据集「{ds['name']}」中指标「{metric_display}」无有效值。", None, did

                ans_parts = [f"🔍 数据来源：**{ds['name']}**"]
                table_df = None

                def _alabel(r):
                    y = str(r[yc]).strip()
                    if str(r[dtc]).strip() == '预测' and not y.endswith('E'):
                        return y + 'E'
                    return y

                if agg == 'list':
                    table_df = df_valid[[yc, dtc, metric_col_real]].copy()
                    table_df.insert(0, '年份', table_df.apply(_alabel, axis=1))
                    table_df[metric_display] = table_df[metric_col_real].apply(lambda v: _fmt(v, metric_display))
                    table_df = table_df[['年份', dtc, metric_display]].rename(columns={dtc: '数据类型'})
                    ans_parts.append(f"📋 历年「{metric_display}」明细（{ds['name']}）")
                elif agg in ['max', 'min']:
                    bv = df_valid[metric_col_real].max() if agg == 'max' else df_valid[metric_col_real].min()
                    op = '最高（最大）' if agg == 'max' else '最低（最小）'
                    br = df_valid[df_valid[metric_col_real] == bv]
                    bys = [_alabel(r) for _, r in br.iterrows()]
                    ans_parts.append(
                        f"📈 「{metric_display}」**{op}**的年份：\n\n"
                        f"- 年份：**{'、'.join(bys)}**\n"
                        f"- 数值：**{_fmt(bv, metric_display)}**"
                    )
                    top5 = df_valid.sort_values(metric_col_real, ascending=(agg == 'min')).head(5)[[yc, dtc, metric_col_real]].copy()
                    top5.insert(0, '年份', top5.apply(_alabel, axis=1))
                    top5[metric_display] = top5[metric_col_real].apply(lambda v: _fmt(v, metric_display))
                    table_df = top5[['年份', metric_display]]
                elif agg in ['mean', 'sum', 'count']:
                    if agg == 'count':
                        v = len(df_valid); w = '行数'
                    else:
                        v = df_valid[metric_col_real].mean() if agg == 'mean' else df_valid[metric_col_real].sum()
                        w = '平均' if agg == 'mean' else '总和'
                    ans_parts.append(
                        f"📊 「{metric_display}」在 **{len(df_valid)}** 条有效数据中的**{w}**：\n\n"
                        f"👉 **{_fmt(v, metric_display)}**"
                    )
                elif agg == 'rank':
                    dfs = df_valid.sort_values(metric_col_real, ascending=False).reset_index(drop=True)
                    dfs.insert(0, '排名', range(1, len(dfs) + 1))
                    dfs.insert(0, '年份', dfs.apply(_alabel, axis=1))
                    dfs[metric_display] = dfs[metric_col_real].apply(lambda v: _fmt(v, metric_display))
                    table_df = dfs[['排名', '年份', metric_display]]
                    ans_parts.append(f"🏆 按「{metric_display}」从高到低排名：")
                elif agg == 'trend':
                    first = df_valid.iloc[0][metric_col_real]; last = df_valid.iloc[-1][metric_col_real]
                    fy = _alabel(df_valid.iloc[0]); ly = _alabel(df_valid.iloc[-1])
                    delta = last - first
                    pct = (delta / first * 100) if first and first != 0 else 0
                    drc = '📈 上升' if delta > 0 else ('📉 下降' if delta < 0 else '➡️ 持平')
                    ans_parts.append(
                        f"📊 「{metric_display}」趋势（{fy} → {ly}）：\n\n"
                        f"- 起始值（{fy}）：**{_fmt(first, metric_display)}**\n"
                        f"- 最新值（{ly}）：**{_fmt(last, metric_display)}**\n"
                        f"- 变化方向：**{drc}**，差值 **{_fmt(delta, metric_display)}**（{pct:+.2f}%）"
                    )
                    tdf = df_valid[[yc, dtc, metric_col_real]].copy()
                    tdf.insert(0, '年份', tdf.apply(_alabel, axis=1))
                    yoy = []; pv = None
                    for _, row in tdf.iterrows():
                        v = row[metric_col_real]
                        if pv is None or pv == 0: yoy.append('—')
                        else: yoy.append(f"{(v-pv)/pv*100:+.2f}%")
                        pv = v
                    tdf['同比'] = yoy
                    tdf[metric_display] = tdf[metric_col_real].apply(lambda v: _fmt(v, metric_display))
                    table_df = tdf[['年份', metric_display, '同比']]
                elif agg == 'compare':
                    if len(years) >= 2:
                        tdf = df_valid[df_valid[yc].astype(int).isin(years)]
                    else:
                        tdf = df_valid.tail(2)
                    if len(tdf) < 2:
                        ans_parts.append("⚠️ 可用于对比的有效数据不足 2 条，请在问题中明确对比年份。")
                    else:
                        vs = []; rows = []
                        for _, r in tdf.iterrows():
                            y = _alabel(r); v = r[metric_col_real]
                            vs.append((y, v)); rows.append({'年份': y, metric_display: _fmt(v, metric_display)})
                        table_df = pd.DataFrame(rows)
                        if len(vs) >= 2:
                            diff = vs[-1][1] - vs[0][1]
                            pct = (diff / vs[0][1] * 100) if vs[0][1] and vs[0][1] != 0 else 0
                            ans_parts.append(
                                f"⚖️ 「{metric_display}」对比：{vs[0][0]} vs {vs[-1][0]}\n\n"
                                f"- {vs[0][0]}：**{_fmt(vs[0][1], metric_display)}**\n"
                                f"- {vs[-1][0]}：**{_fmt(vs[-1][1], metric_display)}**\n"
                                f"- 差异：**{_fmt(diff, metric_display)}**（{pct:+.2f}%）"
                            )
                else:  # value
                    if years:
                        rows = []; found = False
                        for y in years:
                            mt = df_valid[df_valid[yc].astype(int) == y]
                            if len(mt) > 0:
                                found = True
                                for _, r in mt.iterrows():
                                    yy = _alabel(r); v = r[metric_col_real]
                                    rows.append({'年份': yy, '数据类型': r[dtc], metric_display: _fmt(v, metric_display)})
                                    ans_parts.append(f"✅ {yy}年「{metric_display}」 = **{_fmt(v, metric_display)}**")
                        if rows: table_df = pd.DataFrame(rows)
                        if not found:
                            ans_parts.append(f"⚠️ 年份 {years} 中未找到「{metric_display}」有效值。")
                    else:
                        r = df_valid.iloc[-1]
                        yy = _alabel(r); v = r[metric_col_real]
                        ans_parts.append(f"✅ 最新数据（{yy}年）：「{metric_display}」 = **{_fmt(v, metric_display)}**")
                        last3 = df_valid.tail(3).copy()
                        last3.insert(0, '年份', last3.apply(_alabel, axis=1))
                        last3[metric_display] = last3[metric_col_real].apply(lambda x: _fmt(x, metric_display))
                        table_df = last3[['年份', metric_display]]
                ans = "\n\n".join(ans_parts)
                return ans, table_df, did

            else:
                # ===== 明细类（全产品明细 / 单产品明细）=====
                yc = ds['year_col']
                dtc = ds['data_type_col']

                # 额外维度过滤（问题中提到的厂商/行业/部署等）
                # 先从问题文本中抽取 Vendor/行业/部署 关键词
                extra_dim = {}
                # Technology（产品）过滤：从数据集中的Technology列动态匹配产品名
                tech_values = set()
                q_low = q.lower()
                # 中英文产品名别名 → 标准Technology值（用于问题匹配）
                base_tech_map = {
                    'Switch': ['switch', '交换', '交换机'],
                    'WLAN': ['wlan', 'wifi', '无线', 'wi-fi'],
                    'Router': ['router', '路由', '路由器'],
                    'VCC': ['vcc', '虚拟客户端', '虚拟'],
                }
                # 1) 通过别名匹配
                for tech_name, aliases in base_tech_map.items():
                    if any(a in q_low for a in aliases):
                        tech_values.add(tech_name)
                # 2) 通过当前数据集Technology列的值精确匹配（支持数据集中的所有产品）
                if 'Technology' in df.columns:
                    all_techs = df['Technology'].dropna().unique().tolist()
                    for tv in all_techs:
                        tv_str = str(tv).strip()
                        if tv_str and tv_str.lower() in q_low:
                            tech_values.add(tv_str)
                if tech_values:
                    extra_dim['Technology'] = list(tech_values)
                # Vendor 过滤（使用全局双向映射 + 翻译函数，支持中英文任一输入）
                vendor_targets = _vendor_to_english_keywords(q)
                if 'Vendor' in df.columns and vendor_targets:
                    extra_dim['Vendor'] = list(set(vendor_targets))

                # 行业大类过滤关键词（金融/政府/教育/医疗/运营商/企业等）
                industry_kw_map = {
                    '金融': ['金融', 'Banking', 'Finance', 'Insurance'],
                    '银行': ['Banking', '银行'],
                    '政府': ['Government', '政府'],
                    '教育': ['Education', '教育'],
                    '医疗': ['Healthcare', 'Medical', '医疗', '医院'],
                    '运营商': ['Telecom', 'Service Provider', '运营商', '通信'],
                    '电信': ['Telecom', '电信'],
                    '企业': ['Enterprise', '企业'],
                    '商业': ['Retail', 'Wholesale', '商业', '零售'],
                }
                industry_targets = []
                for kw, variants in industry_kw_map.items():
                    if kw in q:
                        industry_targets.extend(variants)
                if '行业大类' in df.columns and industry_targets:
                    extra_dim['行业大类'] = list(set(industry_targets))

                # 部署方式过滤
                deploy_kw_map = {
                    '企业': ['Enterprise', '企业'],
                    '数据中心': ['Datacenter', 'Data Center', '数据中心', 'DC'],
                    '行业': ['Vertical Industry', '行业'],
                    '垂直行业': ['Vertical Industry', '垂直行业'],
                    '运营商': ['Service Provider', '运营商'],
                    '园区': ['Campus', '园区'],
                    '分布式': ['Distributed', '分布式'],
                }
                deploy_targets = []
                for kw, variants in deploy_kw_map.items():
                    if kw in q:
                        deploy_targets.extend(variants)
                if 'Deployment' in df.columns and deploy_targets:
                    extra_dim['Deployment'] = list(set(deploy_targets))

                # 合并用户手动从别名命中的维度过滤
                for col_real, values in dim_filter_values.items():
                    if col_real in extra_dim:
                        extra_dim[col_real] = list(set(extra_dim[col_real] + values))
                    else:
                        extra_dim[col_real] = values

                # 应用过滤
                dfc = _apply_basic_filters(df, ds, years, dtype_filter, quarters, extra_dim)
                if len(dfc) == 0:
                    filter_desc = []
                    if years: filter_desc.append(f"年份{years}")
                    if dtype_filter: filter_desc.append(f"{dtype_filter}数据")
                    if quarters: filter_desc.append(f"季度{quarters}")
                    for k, v in extra_dim.items(): filter_desc.append(f"{k}={v}")
                    return (
                        f"⚠️ 数据集「{ds['name']}」在当前过滤条件下无数据：{'，'.join(filter_desc) if filter_desc else '无过滤条件'}\n\n"
                        "请放宽条件或更换查询的产品/指标。"
                    ), None, did

                if metric_col_real not in dfc.columns:
                    return f"⚠️ 数据集「{ds['name']}」中不存在指标「{metric_display}」（列名：{metric_col_real}）。", None, did

                # 明细类回答需要先按年份汇总（同一年内有多季度/多行）
                # 构造 year_df: 按年份聚合
                year_agg = dfc.groupby(yc, dropna=False)[metric_col_real].sum(numeric_only=True).reset_index()
                year_agg = year_agg.rename(columns={yc: '年份_raw'})
                # 如果有实际/预测信息，拼到年份标签里
                if dtc in dfc.columns:
                    dtc_map = dfc.groupby(yc)[dtc].agg(lambda s: list(pd.Series(s).dropna().unique())).to_dict()
                else:
                    dtc_map = {}
                def _ylbl(r):
                    yr = int(r['年份_raw'])
                    lbl = str(yr)
                    types = dtc_map.get(yr, [])
                    # 若该年任何一个记录是预测，就标E（简化）
                    if any('预测' in str(t) for t in types) and not lbl.endswith('E'):
                        lbl += 'E'
                    return lbl
                year_agg['年份'] = year_agg.apply(_ylbl, axis=1)

                dfc_valid = dfc[dfc[metric_col_real].notna()]
                if len(dfc_valid) == 0 and agg not in ['count', 'list']:
                    return f"⚠️ 数据集「{ds['name']}」过滤后，指标「{metric_display}」无有效值。", None, did

                ans_parts = [f"🔍 数据来源：**{ds['name']}**"]
                filter_info = []
                if years: filter_info.append(f"年份={years}")
                if quarters: filter_info.append(f"季度={quarters}")
                if dtype_filter: filter_info.append(dtype_filter)
                for k, v in extra_dim.items(): filter_info.append(f"{k}={v}")
                if filter_info: ans_parts.append("🎛️ 过滤条件：" + "，".join(filter_info))
                table_df = None

                if agg == 'count':
                    v = len(dfc)
                    ans_parts.append(f"📊 当前过滤下共有 **{v:,}** 条明细记录。")
                    # 小预览
                    preview_cols = [c for c in [yc, 'Technology', 'Vendor', 'Deployment', '行业大类', metric_col_real] if c in dfc.columns]
                    if preview_cols:
                        table_df = dfc[preview_cols].head(10).copy()
                        table_df[metric_display] = table_df[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                elif agg == 'list':
                    # 列出按年份聚合的明细，或前20条原始明细（默认按年份聚合）
                    year_agg_sorted = year_agg.sort_values('年份_raw')
                    year_agg_sorted[metric_display] = year_agg_sorted[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                    table_df = year_agg_sorted[['年份', metric_display]].reset_index(drop=True)
                    ans_parts.append(f"📋 按年份聚合的「{metric_display}」明细（{ds['name']}，共 {len(table_df)} 年）")
                elif agg in ['max', 'min']:
                    # 按年份聚合后取极值
                    s = year_agg.sort_values('年份_raw').reset_index(drop=True)
                    if agg == 'max':
                        idx = s[metric_col_real].idxmax(); op = '最高（最大）'
                    else:
                        idx = s[metric_col_real].idxmin(); op = '最低（最小）'
                    row = s.loc[idx]
                    by = row['年份']; bv = row[metric_col_real]
                    ans_parts.append(
                        f"📈 按年份聚合后「{metric_display}」**{op}**：\n\n"
                        f"- 年份：**{by}**\n- 数值：**{_fmt(bv, metric_display)}**"
                    )
                    top5 = s.sort_values(metric_col_real, ascending=(agg == 'min')).head(5).copy()
                    top5[metric_display] = top5[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                    table_df = top5[['年份', metric_display]].reset_index(drop=True)
                elif agg in ['mean', 'sum']:
                    s = year_agg[metric_col_real]
                    vv = s.mean() if agg == 'mean' else s.sum()
                    w = '年平均' if agg == 'mean' else '按年份合计总和'
                    ans_parts.append(
                        f"📊 按年份聚合后，「{metric_display}」的**{w}**（{len(s)} 年）：\n\n👉 **{_fmt(vv, metric_display)}**"
                    )
                elif agg == 'rank':
                    # 排名：按年份 OR 按厂商 OR 按行业？优先厂商，如果有Vendor列则按厂商聚合
                    if 'Vendor' in dfc.columns and len(dfc['Vendor'].dropna().unique()) > 1:
                        g = dfc.groupby('Vendor', dropna=False)[metric_col_real].sum(numeric_only=True).reset_index()
                        g = g.sort_values(metric_col_real, ascending=False).reset_index(drop=True)
                        g.insert(0, '排名', range(1, len(g) + 1))
                        g[metric_display] = g[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                        table_df = g[['排名', 'Vendor', metric_display]].rename(columns={'Vendor': '厂商'})
                        ans_parts.append(f"🏆 按「{metric_display}」从高到低的厂商排名：")
                    elif '行业大类' in dfc.columns and len(dfc['行业大类'].dropna().unique()) > 1:
                        g = dfc.groupby('行业大类', dropna=False)[metric_col_real].sum(numeric_only=True).reset_index()
                        g = g.sort_values(metric_col_real, ascending=False).reset_index(drop=True)
                        g.insert(0, '排名', range(1, len(g) + 1))
                        g[metric_display] = g[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                        table_df = g[['排名', '行业大类', metric_display]]
                        ans_parts.append(f"🏆 按「{metric_display}」从高到低的行业大类排名：")
                    else:
                        s = year_agg.sort_values(metric_col_real, ascending=False).reset_index(drop=True)
                        s.insert(0, '排名', range(1, len(s) + 1))
                        s[metric_display] = s[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                        table_df = s[['排名', '年份', metric_display]]
                        ans_parts.append(f"🏆 按年份「{metric_display}」从高到低排名：")
                elif agg == 'group' or any(k in q for k in ['分布', '结构']):
                    # 分组结构：优先 行业大类 > Deployment > 厂商
                    group_col = None
                    if '行业大类' in dfc.columns and ('行业' in q or '结构' in q or len(dim_filter_values.get('行业大类', [])) > 0):
                        group_col = '行业大类'
                    elif 'Deployment' in dfc.columns and '部署' in q:
                        group_col = 'Deployment'
                    elif 'Vendor' in dfc.columns:
                        group_col = 'Vendor'
                    elif 'Technology' in dfc.columns:
                        group_col = 'Technology'
                    if group_col is None:
                        group_col = 'Technology' if 'Technology' in dfc.columns else dfc.columns[0]
                    g = dfc.groupby(group_col, dropna=False)[metric_col_real].sum(numeric_only=True).reset_index()
                    g = g.sort_values(metric_col_real, ascending=False).reset_index(drop=True)
                    total_v = g[metric_col_real].sum()
                    g['占比'] = g[metric_col_real].apply(lambda x: f"{x/total_v*100:.2f}%" if total_v and total_v != 0 else '—')
                    g[metric_display] = g[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                    table_df = g[[group_col, metric_display, '占比']]
                    ans_parts.append(f"🧩 按「{group_col}」分组的「{metric_display}」分布结构（合计 {_fmt(total_v, metric_display)}）：")
                elif agg == 'trend':
                    s = year_agg.sort_values('年份_raw').reset_index(drop=True)
                    if len(s) < 2:
                        ans_parts.append("⚠️ 有效年份少于 2 年，无法做趋势分析。")
                    else:
                        first_v = s.iloc[0][metric_col_real]; last_v = s.iloc[-1][metric_col_real]
                        first_y = s.iloc[0]['年份']; last_y = s.iloc[-1]['年份']
                        delta = last_v - first_v
                        pct = (delta / first_v * 100) if first_v and first_v != 0 else 0
                        drc = '📈 上升' if delta > 0 else ('📉 下降' if delta < 0 else '➡️ 持平')
                        ans_parts.append(
                            f"📊 「{metric_display}」趋势（{first_y} → {last_y}）：\n\n"
                            f"- 起始值（{first_y}）：**{_fmt(first_v, metric_display)}**\n"
                            f"- 最新值（{last_y}）：**{_fmt(last_v, metric_display)}**\n"
                            f"- 变化方向：**{drc}**，差值 **{_fmt(delta, metric_display)}**（{pct:+.2f}%）"
                        )
                        s_copy = s.copy()
                        yoy = []; pv = None
                        for _, r in s_copy.iterrows():
                            vv = r[metric_col_real]
                            if pv is None or pv == 0: yoy.append('—')
                            else: yoy.append(f"{(vv-pv)/pv*100:+.2f}%")
                            pv = vv
                        s_copy['同比'] = yoy
                        s_copy[metric_display] = s_copy[metric_col_real].apply(lambda vv: _fmt(vv, metric_display))
                        table_df = s_copy[['年份', metric_display, '同比']].reset_index(drop=True)
                elif agg == 'compare':
                    s = year_agg.sort_values('年份_raw').reset_index(drop=True)
                    if len(years) >= 2:
                        tdf = s[s['年份_raw'].astype(int).isin(years)]
                    else:
                        tdf = s.tail(2)
                    if len(tdf) < 2:
                        ans_parts.append("⚠️ 可对比数据少于 2 条，请明确对比年份。")
                    else:
                        vs = [(r['年份'], r[metric_col_real]) for _, r in tdf.iterrows()]
                        rows = [{'年份': y, metric_display: _fmt(v, metric_display)} for (y, v) in vs]
                        table_df = pd.DataFrame(rows)
                        if len(vs) >= 2:
                            diff = vs[-1][1] - vs[0][1]
                            pct = (diff / vs[0][1] * 100) if vs[0][1] and vs[0][1] != 0 else 0
                            ans_parts.append(
                                f"⚖️ 「{metric_display}」对比：{vs[0][0]} vs {vs[-1][0]}\n\n"
                                f"- {vs[0][0]}：**{_fmt(vs[0][1], metric_display)}**\n"
                                f"- {vs[-1][0]}：**{_fmt(vs[-1][1], metric_display)}**\n"
                                f"- 差异：**{_fmt(diff, metric_display)}**（{pct:+.2f}%）"
                            )
                else:  # value
                    if years:
                        s = year_agg
                        rows = []; found = False
                        for y in years:
                            mt = s[s['年份_raw'].astype(int) == y]
                            if len(mt) > 0:
                                found = True
                                for _, r in mt.iterrows():
                                    vv = r[metric_col_real]
                                    rows.append({'年份': r['年份'], metric_display: _fmt(vv, metric_display)})
                                    ans_parts.append(f"✅ {r['年份']}年「{metric_display}」= **{_fmt(vv, metric_display)}**")
                        if rows: table_df = pd.DataFrame(rows).reset_index(drop=True)
                        if not found:
                            ans_parts.append(f"⚠️ 年份 {years} 中未找到「{metric_display}」有效数据。")
                    else:
                        s = year_agg.sort_values('年份_raw').reset_index(drop=True)
                        r = s.iloc[-1]
                        vv = r[metric_col_real]
                        ans_parts.append(f"✅ 最新数据（{r['年份']}年）：「{metric_display}」 = **{_fmt(vv, metric_display)}**")
                        last3 = s.tail(3).copy()
                        last3[metric_display] = last3[metric_col_real].apply(lambda x: _fmt(x, metric_display))
                        table_df = last3[['年份', metric_display]].reset_index(drop=True)

                ans = "\n\n".join(ans_parts)
                # 统一翻译输出表格中的 Vendor/厂商 列为中文（英文名→中文名）
                if table_df is not None and len(table_df) > 0:
                    for vendor_col_name in ['Vendor', '厂商', '供应商']:
                        if vendor_col_name in table_df.columns:
                            table_df[vendor_col_name] = table_df[vendor_col_name].apply(lambda x: _vendor_translate(x))
                            if vendor_col_name == 'Vendor':
                                # 同时把列名也改中文
                                table_df = table_df.rename(columns={'Vendor': '厂商'})
                # 明细类表格过长时截断显示，最多 30 行
                if table_df is not None and len(table_df) > 30:
                    ans += f"\n\n> 📝 表格展示前 30 行，共 {len(table_df)} 条记录。"
                    table_df = table_df.head(30)
                return ans, table_df, did

        # ============================================================
        # UI：快捷问题 + 提问表单 + 对话历史
        # ============================================================
        st.markdown("### 🤖 快捷提问（点击可自动填充）")
        quick_questions = [
            # 分析汇总类
            "2025年锐捷份额是多少？",
            "哪一年锐捷DC收入最高？",
            "列出所有年份锐捷DC收入",
            "锐捷份额变化趋势？",
            "2024和2025年收入对比？",
            "竞争力指数排名？",
            "平均通信DC容量？",
            # 明细类
            "2026年Switch锐捷人民币收入？",
            "各厂商WLAN收入排名？",
            "Cisco Router 2025年收入？",
            "Switch各部署方式收入分布？",
            "2026年预测总出货量？",
        ]
        qcols = st.columns(4)
        for i, qq in enumerate(quick_questions):
            with qcols[i % 4]:
                if st.button(qq, key=f"qqv2_{i}", use_container_width=True):
                    st.session_state['qa_question_buffer'] = qq

        st.markdown("---")
        st.markdown("#### 💭 输入您的问题")
        buffer_val = st.session_state.get('qa_question_buffer', '')
        with st.form("qa_form_v2", clear_on_submit=False):
            fc1, fc2 = st.columns([8, 1])
            with fc1:
                question = st.text_input(
                    "问题",
                    value=buffer_val,
                    placeholder="例：2026年Switch锐捷收入？ / Cisco 2025 Router收入？ / 各厂商WLAN排名？ / VCC行业分布？",
                    label_visibility="collapsed"
                )
            with fc2:
                submitted = st.form_submit_button("🚀 提问", type="primary", use_container_width=True)

        if submitted and question:
            if 'chat_history' not in st.session_state:
                st.session_state.chat_history = []
            # 先用大模型把自然语言问题归一化成规则引擎可识别的规范问法（无 key / 失败则原样返回）
            question_for_engine = _llm_rewrite_question(question, datasets)
            ans, tbl, src_did = _build_answer_multi(question_for_engine, datasets)
            src_name = datasets[src_did]['name'] if src_did and src_did in datasets else '—'
            st.session_state.chat_history.append({
                'role': 'user',
                'content': question
            })
            st.session_state.chat_history.append({
                'role': 'assistant',
                'content': ans,
                'table': tbl,
                'dataset': src_name,
            })
            st.session_state['qa_question_buffer'] = ''
            st.rerun()

        # 清除对话按钮
        if 'chat_history' in st.session_state and len(st.session_state.chat_history) > 0:
            col_clear1, col_clear2 = st.columns([10, 1])
            with col_clear2:
                if st.button("🗑️ 清空对话"):
                    st.session_state.chat_history = []
                    st.rerun()

        # 对话历史（最新在上）
        if 'chat_history' in st.session_state and len(st.session_state.chat_history) > 0:
            st.markdown("#### 💬 对话记录（最新在上）")
            for msg in reversed(st.session_state.chat_history):
                if msg['role'] == 'user':
                    st.markdown(
                        f'<div style="background:#eef2ff;padding:10px 14px;border-radius:12px;margin:8px 0 4px 0;color:#3730a3">'
                        f'<b>👤 我：</b> {msg["content"]}</div>',
                        unsafe_allow_html=True
                    )
                else:
                    ds_tag = f'<span style="background:#ddd6fe;color:#4c1d95;padding:2px 8px;border-radius:6px;font-size:12px;margin-left:6px">📊 {msg.get("dataset","")}</span>' if msg.get('dataset') else ''
                    st.markdown(
                        f'<div style="background:#f0fdf4;padding:10px 14px;border-radius:12px;margin:8px 0 4px 40px;color:#14532d">'
                        f'<b>🤖 小助手：</b>{ds_tag}</div>',
                        unsafe_allow_html=True
                    )
                    st.markdown(msg['content'])
                    tbl = msg.get('table')
                    if tbl is not None and len(tbl) > 0:
                        st.dataframe(tbl, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ 还没有对话记录。点上方快捷问题或输入问题 → 点击「🚀 提问」开始吧！")

        # 数据来源与查询能力说明
        with st.expander("📖 数据集总览 + 可查询指标/维度说明", expanded=False):
            st.markdown("**所有加工结果表（{n}张）**：".format(n=len(datasets)))
            rows = []
            for did, ds in datasets.items():
                dfc = ds['df']
                rows.append({
                    'ID': did,
                    '分类': ds['category'],
                    '数据集名称': ds['name'],
                    '行数': f"{len(dfc):,}",
                    '列数': len(dfc.columns),
                    '年份列': ds['year_col'],
                    '实际/预测列': ds['data_type_col'],
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            st.markdown("**📊 分析汇总类指标（通信DC汇总）**：")
            am = pd.DataFrame([{'关键词': k, '对应列名': v} for k, v in ANALYSIS_COLUMN_ALIAS.items()])
            st.dataframe(am, use_container_width=True, hide_index=True)

            st.markdown("**📦 明细类指标/维度（全产品+单产品加工表）**：")
            dm = pd.DataFrame([{'关键词': k, '对应列名': v, '类型': '数值指标' if v in ['Vendor Revenue (CNY M)', 'Vendor Revenue (USD M)', 'Units'] else '过滤维度'} for k, v in DETAIL_COLUMN_ALIAS.items()])
            st.dataframe(dm, use_container_width=True, hide_index=True)

            st.markdown("**🧭 数据集路由规则**（按优先级，可通过问题中产品关键词路由到对应表）：\n")
            st.caption("产品关键词（交换机/Switch/WLAN/Router/VCC/锐捷/通信DC）→ 命中单产品加工表优先；否则使用全产品明细；纯DC汇总类问题自动路由到分析结果表。")
