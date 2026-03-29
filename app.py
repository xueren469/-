import streamlit as st
import pandas as pd
import io, os, sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from analysis.loader import load_and_validate
from analysis.engine import run_all_chapters
from report.generator import build_html_report

# ── 页面配置 ──
st.set_page_config(
    page_title="财务分析报告生成系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 全局样式 ──
st.markdown("""
<style>
body, .stApp { background-color: #0d1117; color: #e2e8f0; }
.stButton>button {
    background: #2e75b6; color: white; border: none;
    border-radius: 6px; padding: 10px 24px; font-size: 15px;
    font-weight: 600; width: 100%; cursor: pointer;
}
.stButton>button:hover { background: #1d4ed8; }
.upload-box {
    background: #161b27; border: 1px dashed #334155;
    border-radius: 8px; padding: 12px; margin-bottom: 8px;
}
.section-title {
    font-size: 13px; font-weight: 600; color: #94a3b8;
    padding-bottom: 6px; border-bottom: 1px solid #1e293b;
    margin-bottom: 12px;
}
.tag-must { background: #7f1d1d; color: #fca5a5; padding: 1px 7px;
    border-radius: 3px; font-size: 10px; font-weight: 700; }
.tag-rec { background: #1e3a5f; color: #93c5fd; padding: 1px 7px;
    border-radius: 3px; font-size: 10px; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# ── 侧边栏配置 ──
with st.sidebar:
    st.markdown("## ⚙️ 报告配置")
    st.markdown("---")

    st.markdown("**报告对象**")
    report_target = st.selectbox(
        "", ["董事会", "CFO", "区域总"], label_visibility="collapsed"
    )

    st.markdown("**分析年份**")
    report_year = st.selectbox(
        "", ["2025年", "2024年", "2023年"], label_visibility="collapsed"
    )

    st.markdown("**收入增速战略目标（%）**")
    revenue_target = st.number_input(
        "", min_value=0.0, max_value=100.0, value=25.0,
        step=0.5, label_visibility="collapsed"
    )

    st.markdown("**关联方项目清单**")
    related_party_input = st.text_area(
        "",
        placeholder="每行输入一个关联方项目的标准名称，例如：\n鲁山教育产业园项目餐饮运营部\n教育产业园餐饮总、超市及文印",
        height=120,
        label_visibility="collapsed"
    )
    related_parties = [
        x.strip() for x in related_party_input.strip().splitlines() if x.strip()
    ]

    st.markdown("**关联方在AR表中的标识字段值**")
    rp_ar_tag = st.text_input(
        "", value="兴源高中", label_visibility="collapsed",
        help="应收账款表中'内外部'字段用于标记关联方的特殊值"
    )

    st.markdown("---")
    st.markdown("""
    <div style='font-size:11px;color:#475569;line-height:1.7'>
    💡 <strong style='color:#94a3b8'>使用方式</strong><br>
    1. 在右侧上传所需底表<br>
    2. 在此处填写配置信息<br>
    3. 点击"生成报告"按钮<br>
    4. 下载HTML报告文件
    </div>
    """, unsafe_allow_html=True)

# ── 主界面 ──
st.markdown("# 📊 财务分析报告生成系统")
st.markdown(
    '<p style="color:#475569;font-size:13px;margin-bottom:20px">'
    '上传数据底表，自动生成五章完整财务分析报告（深色主题HTML）</p>',
    unsafe_allow_html=True
)

# ── 文件上传区域 ──
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="section-title">📁 必须上传 <span class="tag-must">REQUIRED</span></div>',
                unsafe_allow_html=True)

    file_ic = st.file_uploader(
        "① 收入成本明细表（Sheet1）",
        type=["xlsx", "xls"],
        key="income_cost",
        help="字段：摘要/科目/本期借方/类型/合作自营/区域/标准名称/分析是否剔除/入场时间/撤场时间"
    )
    file_fs = st.file_uploader(
        "② 三大报表（利润表+资产负债表+现金流量表）",
        type=["xlsx", "xls"],
        key="financial_stmt",
        help="包含利润表、资产负债表、现金流量表三张合并报表"
    )
    file_ar = st.file_uploader(
        "③ 应收账款表",
        type=["xlsx", "xls"],
        key="ar",
        help="字段：摘要(含过去3-5年)/期末余额/标准项目/业态/内外部/合作自有/区域/分析是否剔除"
    )
    file_ap = st.file_uploader(
        "④ 应付账款表",
        type=["xlsx", "xls"],
        key="ap",
        help="字段：摘要/期末余额/标准项目/内外部/类型/合作自营/公司（区域）"
    )
    file_adv = st.file_uploader(
        "⑤ 预收账款表",
        type=["xlsx", "xls"],
        key="adv",
        help="字段：摘要/期末余额/标准项目/内外部/类型/合作自营/公司（区域）"
    )

with col2:
    st.markdown('<div class="section-title">📁 建议上传 <span class="tag-rec">RECOMMENDED</span></div>',
                unsafe_allow_html=True)

    file_cf = st.file_uploader(
        "⑥ 项目级经营现金流表（经营性现金流Sheet）",
        type=["xlsx", "xls"],
        key="cf",
        help="字段：摘要/现金流量表项/值/收支/标准名称/区域/合作自营/分析是否剔除。缺少此表则第四章仅做整体分析，无法做区域拆解。"
    )

    st.markdown("---")
    st.markdown("""
    <div style='background:#161b27;border-radius:8px;padding:14px;font-size:11.5px;color:#64748b;line-height:1.8'>
    <strong style='color:#94a3b8'>📋 数据格式说明</strong><br>
    • 支持 .xlsx 和 .xls 格式<br>
    • 收入成本表取 Sheet1<br>
    • 应收/应付/预收取对应Sheet<br>
    • 现金流表取"经营性现金流"Sheet<br>
    • 字段名称需与底表一致<br>
    • "分析是否剔除"=1的行将被过滤<br><br>
    <strong style='color:#94a3b8'>⚠ 关联方处理</strong><br>
    关联方项目将从主体分析中剥离，<br>
    在第三章单独列示，并在各章说明影响。
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── 上传状态检查 ──
    required_files = {
        "收入成本表": file_ic,
        "三大报表": file_fs,
        "应收账款表": file_ar,
        "应付账款表": file_ap,
        "预收账款表": file_adv
    }
    missing = [k for k, v in required_files.items() if v is None]
    uploaded_count = 5 - len(missing) + (1 if file_cf else 0)

    if not missing:
        st.success(f"✅ 必须文件已全部上传（{uploaded_count}/6）")
    else:
        st.warning(f"⚠ 还缺少：{'、'.join(missing)}")

# ── 生成按钮 ──
st.markdown("---")
gen_col1, gen_col2, gen_col3 = st.columns([1, 2, 1])

with gen_col2:
    generate_btn = st.button(
        "🚀 生成财务分析报告",
        disabled=bool(missing),
        use_container_width=True
    )

# ── 生成逻辑 ──
if generate_btn:
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        config = {
            "report_target": report_target,
            "report_year": report_year,
            "base_year": str(int(report_year[:4]) - 1) + "年",
            "revenue_target": revenue_target,
            "related_parties": related_parties,
            "rp_ar_tag": rp_ar_tag,
            "report_date": datetime.now().strftime("%Y年%m月%d日"),
            "company_name": "北京兴业源科技服务集团股份有限公司"
        }

        status_text.markdown("⏳ 正在加载和验证数据...")
        progress_bar.progress(10)

        data = load_and_validate(
            file_ic=file_ic, file_fs=file_fs, file_ar=file_ar,
            file_ap=file_ap, file_adv=file_adv, file_cf=file_cf,
            config=config
        )

        status_text.markdown("⏳ 正在分析第一章：整体经营健康度...")
        progress_bar.progress(25)

        status_text.markdown("⏳ 正在分析第二章：利润结构解剖...")
        progress_bar.progress(45)

        status_text.markdown("⏳ 正在分析第三章：应收账款...")
        progress_bar.progress(60)

        status_text.markdown("⏳ 正在分析第四章：现金流与偿债...")
        progress_bar.progress(75)

        status_text.markdown("⏳ 正在生成第五章：行动建议...")
        progress_bar.progress(85)

        results = run_all_chapters(data, config)

        status_text.markdown("⏳ 正在渲染HTML报告...")
        progress_bar.progress(95)

        html_content = build_html_report(results, config)

        progress_bar.progress(100)
        status_text.markdown("✅ 报告生成完成！")

        # ── 下载按钮 ──
        filename = f"财务分析报告_{config['company_name']}_{report_year}.html"
        dl_col1, dl_col2, dl_col3 = st.columns([1, 2, 1])
        with dl_col2:
            st.download_button(
                label="⬇️ 下载HTML报告",
                data=html_content.encode("utf-8"),
                file_name=filename,
                mime="text/html",
                use_container_width=True
            )

        # ── 关键指标预览 ──
        st.markdown("### 📋 关键指标速览")
        kpi = results.get("kpi_summary", {})
        if kpi:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("营业收入", kpi.get("revenue_25", "—"), kpi.get("revenue_chg", ""))
            k2.metric("综合毛利率", kpi.get("gm_rate_25", "—"), kpi.get("gm_rate_chg", ""))
            k3.metric("资产负债率", kpi.get("debt_ratio_25", "—"), kpi.get("debt_ratio_chg", ""))
            k4.metric("FCF自由现金流", kpi.get("fcf", "—"), "")

    except Exception as e:
        progress_bar.progress(0)
        st.error(f"❌ 生成报告时出错：{str(e)}")
        st.markdown("**常见问题排查：**")
        st.markdown("- 请检查Excel文件的Sheet名称是否正确")
        st.markdown("- 请检查字段名称是否与标准大纲一致")
        st.markdown("- 请确认'分析是否剔除'字段为数字（0或1）")
        with st.expander("查看详细错误信息"):
            import traceback
            st.code(traceback.format_exc())
