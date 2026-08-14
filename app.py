import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="LTV版本对比工具", layout="wide")

st.title("🎮 LTV版本对比分析工具")
st.markdown("上传两个版本的Global LTV数据，自动分析收益前5国家的LTV表现")

# ==================== 侧边栏配置 ====================
with st.sidebar:
    st.header("⚙️ 配置")
    
    ltv_options = {
        'LTV1': 'ltv01',
        'LTV7': 'ltv07',
        'LTV14': 'ltv14',
        'LTV30': 'ltv30' if 'ltv30' in st.session_state else 'ltv30'
    }
    
    selected_ltvs = st.multiselect(
        "选择要分析的LTV指标",
        options=['LTV1', 'LTV7', 'LTV14', 'LTV30'],
        default=['LTV1', 'LTV7', 'LTV14']
    )
    
    top_n = st.slider("显示Top N国家", min_value=3, max_value=10, value=5)
    
    st.markdown("---")
    st.markdown("""
    **📌 数据格式要求：**
    - CSV文件，UTF-8编码
    - 必须包含列：`weidu`(国家), `new_user`(新增用户)
    - LTV列：`ltv01`, `ltv07`, `ltv14` 等
    - 日期列：`first_open_date_day`
    """)

# ==================== 核心函数 ====================
def load_and_clean(file):
    """加载并清洗CSV数据"""
    try:
        df = pd.read_csv(file)
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(file, encoding='utf-16', sep='\t')
        except:
            df = pd.read_csv(file, encoding='latin-1')
    
    df = df.replace('', np.nan)
    df = df.replace(' ', np.nan)
    
    # 转换数值列
    numeric_cols = ['new_user', 'rev00', 'dav00', 'ltv00', 'rev01', 'dav01', 'ltv01', 
                    'rev03', 'dav03', 'ltv03', 'rev07', 'dav07', 'ltv07', 'rev14', 'dav14', 'ltv14',
                    'rev30', 'dav30', 'ltv30']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def get_top_countries(df, n=5):
    """计算收益前N国家"""
    if 'rev00' not in df.columns and 'ltv00' not in df.columns:
        # 如果没有收益列，用LTV1 * new_user 估算
        if 'ltv01' in df.columns and 'new_user' in df.columns:
            df['_revenue_est'] = df['ltv01'] * df['new_user']
            country_earnings = df.groupby('weidu')['_revenue_est'].sum().sort_values(ascending=False)
        else:
            st.error("❌ 找不到收益列，请确认数据包含 'rev00' 或 'ltv01' 列")
            return []
    else:
        # 使用 rev00 作为收益
        country_earnings = df.groupby('weidu')['rev00'].sum().sort_values(ascending=False)
    
    return country_earnings.head(n).index.tolist()

def calculate_country_metrics(df, country):
    """计算单个国家的LTV指标"""
    country_df = df[df['weidu'] == country]
    if len(country_df) == 0:
        return None
    
    metrics = {
        '国家': country,
        '新增用户': country_df['new_user'].sum() if 'new_user' in country_df.columns else 0,
    }
    
    # 计算各LTV平均值
    ltv_cols = [col for col in ['ltv01', 'ltv07', 'ltv14', 'ltv30'] if col in df.columns]
    for col in ltv_cols:
        metrics[col] = country_df[col].mean()
    
    return metrics

def create_ltv_comparison(df_a, df_b, countries, version_a, version_b, ltv_cols):
    """创建LTV对比数据"""
    results = []
    
    for country in countries:
        metrics_a = calculate_country_metrics(df_a, country)
        metrics_b = calculate_country_metrics(df_b, country)
        
        if metrics_a is None or metrics_b is None:
            continue
        
        row = {
            '国家': country,
            f'{version_a}_用户': metrics_a['新增用户'],
            f'{version_b}_用户': metrics_b['新增用户'],
            '用户变化%': ((metrics_b['新增用户'] / metrics_a['新增用户'] - 1) * 100) if metrics_a['新增用户'] > 0 else 0,
        }
        
        for ltv_col in ltv_cols:
            val_a = metrics_a.get(ltv_col, 0)
            val_b = metrics_b.get(ltv_col, 0)
            row[f'{version_a}_{ltv_col}'] = val_a
            row[f'{version_b}_{ltv_col}'] = val_b
            row[f'{ltv_col}_变化%'] = ((val_b / val_a - 1) * 100) if val_a > 0 else 0
        
        results.append(row)
    
    return pd.DataFrame(results)

def create_global_summary(df_a, df_b, version_a, version_b, ltv_cols):
    """创建全球汇总"""
    summary = {'维度': ['Global']}
    
    for df, ver in [(df_a, version_a), (df_b, version_b)]:
        summary[f'{ver}_用户'] = df['new_user'].sum() if 'new_user' in df.columns else 0
        for col in ltv_cols:
            if col in df.columns:
                summary[f'{ver}_{col}'] = df[col].mean()
    
    # 计算变化
    for col in ltv_cols:
        val_a = summary.get(f'{version_a}_{col}', 0)
        val_b = summary.get(f'{version_b}_{col}', 0)
        summary[f'{col}_变化%'] = ((val_b / val_a - 1) * 100) if val_a > 0 else 0
    
    return pd.DataFrame([summary])

def format_excel_file(filepath):
    """美化Excel"""
    wb = load_workbook(filepath)
    
    header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='2E86AB', end_color='2E86AB', fill_type='solid')
    positive_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
    negative_fill = PatternFill(start_color='FADBD8', end_color='FADBD8', fill_type='solid')
    positive_font = Font(color='1A7A3A')
    negative_font = Font(color='C0392B')
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=1), 1):
            for cell in row:
                if cell.value is not None:
                    if row_idx == 1:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.border = border
                        col_name = ws.cell(row=1, column=cell.column).value
                        if isinstance(cell.value, (int, float)):
                            if col_name and '变化%' in str(col_name):
                                if cell.value > 0:
                                    cell.font = positive_font
                                    cell.fill = positive_fill
                                elif cell.value < 0:
                                    cell.font = negative_font
                                    cell.fill = negative_fill
                            if isinstance(cell.value, float):
                                cell.value = round(cell.value, 4)
        
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
            ws.column_dimensions[col_letter].width = min(max(max_length + 3, 12), 30)
    
    wb.save(filepath)

def generate_plot(df_compare, version_a, version_b, ltv_cols, title):
    """生成LTV对比图"""
    countries = df_compare['国家'].tolist()
    n_countries = len(countries)
    
    fig, axes = plt.subplots(1, len(ltv_cols), figsize=(6 * len(ltv_cols), 6))
    if len(ltv_cols) == 1:
        axes = [axes]
    
    x = np.arange(n_countries)
    width = 0.35
    
    for idx, ltv_col in enumerate(ltv_cols):
        ax = axes[idx]
        
        val_a = df_compare[f'{version_a}_{ltv_col}'].values
        val_b = df_compare[f'{version_b}_{ltv_col}'].values
        
        bars1 = ax.bar(x - width/2, val_a, width, label=version_a, color='#2E86AB')
        bars2 = ax.bar(x + width/2, val_b, width, label=version_b, color='#E84855')
        
        # 添加数值标签
        for bar in bars1:
            height = bar.get_height()
            ax.annotate(f'{height:.3f}', xy=(bar.get_x() + bar.get_width()/2, height),
                       xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=9)
        for bar in bars2:
            height = bar.get_height()
            ax.annotate(f'{height:.3f}', xy=(bar.get_x() + bar.get_width()/2, height),
                       xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=9)
        
        ax.set_xlabel('Country', fontsize=12)
        ax.set_ylabel(ltv_col.upper(), fontsize=12)
        ax.set_title(f'{title} - {ltv_col.upper()} Comparison', fontsize=14, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(countries, rotation=45, ha='right')
        ax.legend(fontsize=11)
        ax.grid(True, alpha=0.3, axis='y')
    
    plt.tight_layout()
    return fig

# ==================== 主界面 ====================
# 文件上传
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 版本A (旧版本)")
    file_a = st.file_uploader("上传版本A的CSV文件", type=['csv', 'txt'], key='file_a')
    version_a = st.text_input("版本A名称", value="v1.0.0", key='ver_a')

with col2:
    st.subheader("📁 版本B (新版本)")
    file_b = st.file_uploader("上传版本B的CSV文件", type=['csv', 'txt'], key='file_b')
    version_b = st.text_input("版本B名称", value="v2.0.0", key='ver_b')

# 分析按钮
if st.button("🚀 开始分析", type="primary", use_container_width=True):
    if file_a is None or file_b is None:
        st.error("❌ 请上传两个版本的CSV文件")
        st.stop()
    
    with st.spinner("🔄 正在分析数据..."):
        try:
            # 加载数据
            df_a = load_and_clean(file_a)
            df_b = load_and_clean(file_b)
            
            # 获取LTV列
            ltv_cols = [col for col in ['ltv01', 'ltv07', 'ltv14', 'ltv30'] 
                       if col in df_a.columns and col in df_b.columns]
            
            if not ltv_cols:
                st.error("❌ 数据中找不到LTV列 (ltv01, ltv07, ltv14, ltv30)")
                st.stop()
            
            # 自动检测游戏名
            game_name = "Game"
            if 'weidu' in df_a.columns:
                # 从国家列推断
                pass
            
            # 获取Top N国家
            countries = get_top_countries(df_a, top_n)
            if not countries:
                st.error("❌ 无法计算收益排名，请确认数据包含 'rev00' 或 'ltv01' 列")
                st.stop()
            
            # 创建对比数据
            df_compare = create_ltv_comparison(df_a, df_b, countries, version_a, version_b, ltv_cols)
            df_global = create_global_summary(df_a, df_b, version_a, version_b, ltv_cols)
            
            # ===== 显示结果 =====
            st.success(f"✅ 分析完成！Top {top_n} 国家: {', '.join(countries)}")
            
            # Tab显示
            tab1, tab2, tab3, tab4 = st.tabs(["📊 国家LTV对比", "🌍 全球汇总", "📈 趋势图表", "📥 下载报告"])
            
            with tab1:
                st.subheader(f"📊 Top {top_n} 国家 LTV 对比")
                
                # 格式化显示
                display_df = df_compare.copy()
                for col in display_df.columns:
                    if '变化%' in col:
                        display_df[col] = display_df[col].apply(lambda x: f"{x:+.2f}%")
                    elif '用户' in col and not '变化' in col:
                        display_df[col] = display_df[col].apply(lambda x: f"{int(x):,}")
                    elif 'ltv' in col.lower():
                        display_df[col] = display_df[col].apply(lambda x: f"{x:.4f}")
                
                st.dataframe(display_df, use_container_width=True)
                
                # 添加条件格式说明
                st.markdown("""
                **📌 颜色说明：**
                - 🟢 绿色：变化率为正（表现提升）
                - 🔴 红色：变化率为负（表现下降）
                """)
            
            with tab2:
                st.subheader("🌍 全球汇总")
                display_global = df_global.copy()
                for col in display_global.columns:
                    if '变化%' in col:
                        display_global[col] = display_global[col].apply(lambda x: f"{x:+.2f}%")
                    elif 'ltv' in col.lower():
                        display_global[col] = display_global[col].apply(lambda x: f"{x:.4f}")
                    elif '用户' in col and '变化' not in col:
                        display_global[col] = display_global[col].apply(lambda x: f"{int(x):,}")
                st.dataframe(display_global, use_container_width=True)
            
            with tab3:
                st.subheader("📈 LTV趋势对比图")
                
                # 国家LTV对比图
                fig = generate_plot(df_compare, version_a, version_b, ltv_cols, "Top Countries")
                st.pyplot(fig)
                plt.close(fig)
                
                # 显示各国每日趋势（可选）
                st.subheader("📈 各国每日LTV趋势")
                country_select = st.selectbox("选择国家查看趋势", countries)
                
                if country_select:
                    fig2, axes2 = plt.subplots(1, len(ltv_cols), figsize=(6 * len(ltv_cols), 4))
                    if len(ltv_cols) == 1:
                        axes2 = [axes2]
                    
                    country_df_a = df_a[df_a['weidu'] == country_select]
                    country_df_b = df_b[df_b['weidu'] == country_select]
                    
                    for idx, ltv_col in enumerate(ltv_cols):
                        ax = axes2[idx]
                        if 'first_open_date_day' in country_df_a.columns:
                            daily_a = country_df_a.groupby('first_open_date_day')[ltv_col].mean()
                            ax.plot(daily_a.index, daily_a.values, 'o-', label=version_a, color='#2E86AB')
                        if 'first_open_date_day' in country_df_b.columns:
                            daily_b = country_df_b.groupby('first_open_date_day')[ltv_col].mean()
                            ax.plot(daily_b.index, daily_b.values, 's-', label=version_b, color='#E84855')
                        
                        ax.set_xlabel('Date')
                        ax.set_ylabel(ltv_col.upper())
                        ax.set_title(f'{country_select} - {ltv_col.upper()}')
                        ax.legend()
                        ax.grid(True, alpha=0.3)
                        ax.tick_params(axis='x', rotation=45)
                    
                    plt.tight_layout()
                    st.pyplot(fig2)
                    plt.close(fig2)
            
            with tab4:
                st.subheader("📥 下载Excel报告")
                
                # 生成Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_global.to_excel(writer, sheet_name='全球汇总', index=False)
                    df_compare.to_excel(writer, sheet_name='国家LTV对比', index=False)
                    
                    # 添加各国每日趋势
                    for country in countries:
                        country_df_a = df_a[df_a['weidu'] == country]
                        country_df_b = df_b[df_b['weidu'] == country]
                        
                        if len(country_df_a) > 0 or len(country_df_b) > 0:
                            daily_a = country_df_a.groupby('first_open_date_day').agg({
                                **{col: 'mean' for col in ltv_cols if col in df_a.columns},
                                'new_user': 'sum' if 'new_user' in df_a.columns else None
                            }).reset_index() if len(country_df_a) > 0 else pd.DataFrame()
                            
                            daily_b = country_df_b.groupby('first_open_date_day').agg({
                                **{col: 'mean' for col in ltv_cols if col in df_b.columns},
                                'new_user': 'sum' if 'new_user' in df_b.columns else None
                            }).reset_index() if len(country_df_b) > 0 else pd.DataFrame()
                            
                            if not daily_a.empty or not daily_b.empty:
                                daily_a['版本'] = version_a if not daily_a.empty else None
                                daily_b['版本'] = version_b if not daily_b.empty else None
                                
                                daily_combined = pd.concat([daily_a, daily_b], ignore_index=True)
                                sheet_name = country.replace('/', '_').replace('\\', '_')[:31]
                                daily_combined.to_excel(writer, sheet_name=sheet_name, index=False)
                
                excel_buffer.seek(0)
                
                st.download_button(
                    label="📥 下载Excel报告",
                    data=excel_buffer,
                    file_name=f"{game_name}_LTV_Comparison.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # 下载图表
                fig_buffer = io.BytesIO()
                fig = generate_plot(df_compare, version_a, version_b, ltv_cols, "Top Countries")
                fig.savefig(fig_buffer, dpi=300, bbox_inches='tight')
                fig_buffer.seek(0)
                plt.close(fig)
                
                st.download_button(
                    label="📊 下载图表PNG",
                    data=fig_buffer,
                    file_name=f"{game_name}_LTV_Charts.png",
                    mime="image/png",
                    use_container_width=True
                )
        
        except Exception as e:
            st.error(f"❌ 分析失败: {str(e)}")
            st.exception(e)

# ==================== 底部 ====================
st.markdown("---")
st.caption("💡 上传两个版本的Global LTV数据，系统自动识别Top国家并进行对比分析")