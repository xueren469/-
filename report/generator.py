"""
generator.py — HTML报告生成器
接收五章分析结果，输出完整的深色主题HTML报告字符串
"""

CSS = """
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0d1117;color:#e2e8f0;font-family:'PingFang SC','Microsoft YaHei',Arial,sans-serif;font-size:13px;line-height:1.6}
.nav{position:sticky;top:0;background:#0d1117;border-bottom:1px solid #1e293b;padding:8px 20px;display:flex;gap:6px;flex-wrap:wrap;z-index:100}
.nav a{font-size:11px;padding:3px 10px;border-radius:4px;background:#161b27;color:#94a3b8;text-decoration:none;border:1px solid #1e293b}
.nav a:hover{background:#1e293b;color:#e2e8f0}
.cover{min-height:100vh;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;padding:40px;background:linear-gradient(135deg,#0d1117 0%,#161b27 100%)}
.cover-co{font-size:14px;color:#64748b;margin-bottom:14px}
.cover-title{font-size:38px;font-weight:700;color:#e2e8f0;line-height:1.2;margin-bottom:10px}
.cover-sub{font-size:16px;color:#60a5fa;margin-bottom:30px;font-style:italic}
.cover-div{width:120px;height:2px;background:linear-gradient(90deg,#2e75b6,#60a5fa);margin:0 auto 22px}
.cover-meta{font-size:13px;color:#64748b;line-height:2.2}
.cover-warn{margin-top:30px;font-size:11px;color:#334155;font-style:italic}
.section{padding:22px 24px 30px;border-bottom:2px solid #1e293b}
.ch-tag{font-size:11px;color:#475569;margin-bottom:3px}
.ch-title{font-size:22px;font-weight:700;margin-bottom:3px}
.ch-sub{font-size:12px;color:#475569;margin-bottom:18px}
.panel{background:#161b27;border-radius:10px;padding:15px;margin-bottom:12px}
.st{font-size:12.5px;font-weight:600;color:#94a3b8;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #1e293b}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:12px}
.g4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:10px;margin-bottom:12px}
.box{background:#0d1117;border-radius:7px;padding:10px 12px}
.kpi{border-radius:8px;padding:13px;border-left:4px solid}
.kpi-l{font-size:10.5px;color:#64748b;margin-bottom:4px}
.kpi-v{font-size:22px;font-weight:700;line-height:1}
.kpi-s{font-size:10px;margin-top:3px;color:#64748b}
.k-blue{background:#0a0f1a;border-color:#3b82f6}.k-blue .kpi-v{color:#60a5fa}
.k-green{background:#0a1a0a;border-color:#22c55e}.k-green .kpi-v{color:#4ade80}
.k-red{background:#1a0505;border-color:#ef4444}.k-red .kpi-v{color:#f87171}
.k-orange{background:#1a0c05;border-color:#f97316}.k-orange .kpi-v{color:#fb923c}
.k-yellow{background:#1a1200;border-color:#f59e0b}.k-yellow .kpi-v{color:#fbbf24}
.k-purple{background:#0d0820;border-color:#a855f7}.k-purple .kpi-v{color:#c084fc}
.al{border-radius:8px;padding:11px 13px;margin-bottom:9px;display:flex;gap:9px}
.al-r{background:#1a0505;border:1px solid #7f1d1d}
.al-y{background:#1a1200;border:1px solid #78350f}
.al-g{background:#051a0a;border:1px solid #14532d}
.al-b{background:#0a0f1a;border:1px solid #1e3a5f}
.al-p{background:#0d0820;border:1px solid #4c1d95}
.al-icon{font-size:15px;flex-shrink:0;line-height:1.5}
.al-title{font-size:12px;font-weight:700;margin-bottom:3px}
.al-r .al-title{color:#f87171}.al-y .al-title{color:#fbbf24}
.al-g .al-title{color:#4ade80}.al-b .al-title{color:#60a5fa}.al-p .al-title{color:#c084fc}
.al-body{font-size:11.5px;color:#94a3b8;line-height:1.6}
.al-body em{font-style:normal;font-weight:700;color:#e2e8f0}
.tbl{width:100%;border-collapse:collapse;font-size:11.5px;margin-bottom:6px}
.tbl th{background:#1e293b;color:#94a3b8;font-weight:600;padding:7px 9px;text-align:left}
.tbl th.r,.tbl td.r{text-align:right}
.tbl td{padding:5px 9px;border-bottom:1px solid #0d1117;color:#cbd5e1}
.tbl tr:hover td{background:#1e293b22}
.row-r td{background:#1a050522}.row-y td{background:#1a120022}
.row-g td{background:#051a0a22}.row-b td{background:#0a0f1a22}
.row-p td{background:#0d082022}
.sub-hd td{background:#1e293b;color:#64748b;font-size:10.5px;font-style:italic;padding:4px 9px}
.t-red{color:#f87171;font-weight:600}.t-orange{color:#fb923c;font-weight:600}
.t-green{color:#4ade80;font-weight:600}.t-blue{color:#60a5fa;font-weight:600}
.t-yellow{color:#fbbf24;font-weight:600}.t-gray{color:#64748b}
.t-purple{color:#c084fc;font-weight:600}
.bar-row{display:flex;align-items:center;gap:8px;margin-bottom:6px}
.bar-name{font-size:11px;color:#94a3b8;width:140px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar-track{flex:1;background:#0d1117;border-radius:3px;height:8px;overflow:hidden}
.bar-fill{height:100%;border-radius:3px}
.bf-red{background:#ef4444}.bf-orange{background:#f97316}
.bf-yellow{background:#f59e0b}.bf-green{background:#22c55e}
.bf-blue{background:#3b82f6}.bf-purple{background:#a855f7}
.bar-val{font-size:11px;font-weight:700;width:70px;text-align:right;flex-shrink:0}
.score-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px}
.score-card{background:#0d1117;border-radius:8px;padding:11px;text-align:center}
.score-val{font-size:26px;font-weight:700}
.score-lbl{font-size:10px;color:#64748b;margin-top:3px;line-height:1.4}
.tag{display:inline-block;padding:1px 7px;border-radius:3px;font-size:10px;font-weight:700;margin-left:3px}
.tag-r{background:#7f1d1d;color:#fca5a5}
.tag-g{background:#14532d;color:#86efac}
.tag-y{background:#78350f;color:#fcd34d}
.tag-b{background:#1e3a5f;color:#93c5fd}
.tag-p{background:#4c1d95;color:#c4b5fd}
.summary-box{background:#1f3864;border-radius:10px;padding:20px;text-align:center;margin:14px 0}
.s-title{font-size:14px;font-weight:700;color:#ffd700;margin-bottom:8px}
.s-body{font-size:11.5px;color:#bdd7ee;line-height:1.8}
.note{font-size:10.5px;color:#475569;margin-top:6px;line-height:1.6}
.divider{height:1px;background:#1e293b;margin:12px 0}
.sp{margin:8px 0}
.h3{font-size:13px;font-weight:600;color:#e2e8f0;margin:12px 0 7px}
</style>
"""


def _fmt_wan(v, decimals=0):
    """格式化万元"""
    if abs(v) >= 1e8:
        return f"{v/1e8:.{decimals}f}亿"
    return f"{v/1e4:.{decimals}f}万"


def _pct(v):
    return f"{v*100:.1f}%"


def _ppt(v):
    s = "+" if v >= 0 else ""
    return f"{s}{v*100:.2f}ppt"


def _bar(name, val, max_val, color="bf-green", val_fmt=None, color_class=None):
    w = min(int(abs(val) / max(abs(max_val), 1) * 100), 100)
    v_str = val_fmt or f"{val/1e4:.0f}万"
    c = color_class or ("t-green" if val >= 0 else "t-red")
    return (
        f'<div class="bar-row">'
        f'<div class="bar-name">{name}</div>'
        f'<div class="bar-track"><div class="bar-fill {color}" style="width:{w}%"></div></div>'
        f'<div class="bar-val {c}">{v_str}</div>'
        f'</div>'
    )


def build_html_report(results, config):
    ch1 = results["ch1"]
    ch2 = results["ch2"]
    ch3 = results["ch3"]
    ch4 = results["ch4"]
    ch5 = results["ch5"]
    company = config.get("company_name", "公司")
    year = config["report_year"]
    base = config["base_year"]
    report_date = config["report_date"]
    target_pct = config["revenue_target"]

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="UTF-8"><title>{company}·{year}财务分析报告</title>{CSS}</head>
<body>
<div class="nav">
  <a href="#cover">封面</a>
  <a href="#exec">执行摘要</a>
  <a href="#ch1">第一章</a>
  <a href="#ch2">第二章</a>
  <a href="#ch3">第三章</a>
  <a href="#ch4">第四章</a>
  <a href="#ch5">第五章</a>
</div>
"""

    # ── 封面 ──
    html += f"""
<div class="cover" id="cover">
  <div class="cover-co">{company}</div>
  <div class="cover-title">{year}度综合财务分析报告</div>
  <div class="cover-sub">供{config.get('report_target','董事会')}决策参考</div>
  <div class="cover-div"></div>
  <div class="cover-meta">报告日期：{report_date}<br>分析范围：{base}–{year}</div>
  <div class="cover-warn">本报告含敏感经营信息，仅限内部使用，禁止对外传阅</div>
</div>
"""

    # ── 执行摘要 ──
    rev = ch1.get("rev_25", 0)
    rev_g = ch1.get("rev_growth", 0)
    gm_chg = ch1.get("gm_rate_25", 0) - ch1.get("gm_rate_24", 0)
    net = ch1.get("net_25", 0)
    adj_net = ch1.get("adj_net_25", 0)
    adj_g = (adj_net - ch1.get("net_24", 0)) / ch1.get("net_24", 1) if ch1.get("net_24", 0) > 0 else 0
    ar_total = ch3.get("ar_total_25", 0)
    da = ch1.get("da_ratio_25", 0)

    scores = ch5.get("scores", {})
    score_colors = {"A": "#4ade80", "B+": "#4ade80", "B": "#86efac",
                    "B-": "#fbbf24", "C+": "#fbbf24", "C": "#fb923c",
                    "D+": "#f87171", "D": "#ef4444"}

    score_html = ""
    for dim, sc in scores.items():
        c = score_colors.get(sc, "#94a3b8")
        score_html += f"""
        <div class="score-card">
          <div class="score-val" style="color:{c}">{sc}</div>
          <div class="score-lbl">{dim}</div>
        </div>"""

    html += f"""
<div class="section" id="exec">
  <div class="ch-tag">执行摘要</div>
  <div class="ch-title">核心发现速览</div>
  <div class="ch-sub">回答三个核心问题：钱是怎么赚的 · 赚得是否健康 · 能否持续赚</div>
  <div class="g2">
    <div>
      <div class="h3">五个核心发现</div>
      <div class="al al-r"><div class="al-icon">🔴</div><div>
        <div class="al-title">发现①：净利润增长存在水分</div>
        <div class="al-body">账面净利润<em>{_fmt_wan(net)}</em>，但含一次性收益<em>{_fmt_wan(ch1.get('onetime_gain',0))}</em>。剔除后实质净利润<em>{_fmt_wan(adj_net)}</em>，实质增幅仅<em>{adj_g*100:.1f}%</em>。</div>
      </div></div>
      <div class="al al-r"><div class="al-icon">🔴</div><div>
        <div class="al-title">发现②：增长依赖新项目输血</div>
        <div class="al-body">收入增速{_pct(rev_g)}（战略目标{target_pct}%），存续老项目和撤场项目合计流失大量收入，新项目在填补缺口。</div>
      </div></div>
      <div class="al al-r"><div class="al-icon">🔴</div><div>
        <div class="al-title">发现③：应收账款持续积累</div>
        <div class="al-body">期末应收账款<em>{_fmt_wan(ar_total)}</em>，净资金敞口持续扩大，是制约并购扩张的首要瓶颈。</div>
      </div></div>
      <div class="al al-g"><div class="al-icon">🟢</div><div>
        <div class="al-title">发现④：财务结构改善</div>
        <div class="al-body">资产负债率<em>{_pct(da)}</em>，净有息负债<em>{_fmt_wan(ch1.get('net_debt',0))}</em>，偿债能力健康。</div>
      </div></div>
      <div class="al al-{"r" if gm_chg < -0.005 else "y"}"><div class="al-icon">{"🔴" if gm_chg < -0.005 else "🟡"}</div><div>
        <div class="al-title">发现⑤：毛利率{"下滑" if gm_chg < 0 else "基本稳定"}</div>
        <div class="al-body">综合毛利率{_pct(ch1.get('gm_rate_25',0))}，较上年变动<em>{_ppt(gm_chg)}</em>。{"毛利额不增反降，增收不增利。" if gm_chg < 0 else "毛利管控较好。"}</div>
      </div></div>
    </div>
    <div>
      <div class="h3">{year}经营健康度评分</div>
      <div class="score-grid">{score_html}</div>
      <div class="summary-box">
        <div class="s-title">{year}核心战略指令</div>
        <div class="s-body">财务结构{"已安全" if da < 0.6 else "仍有压力"}，{"但经营质量需要关注。" if gm_chg < 0 else "经营质量保持稳定。"}<br>
        <strong style="color:#ffd700;font-size:13px">首要任务：{"做实比做大更重要" if ch3.get('ar_total_25',0)/max(rev,1)>0.3 else "保持增长同时管控应收"}：</strong><br>
        压缩应收账款 → 释放现金 → 修复问题区域盈利能力 → 再以更健康的基本面推进扩张</div>
      </div>
    </div>
  </div>
</div>
"""

    # ── 第一章 ──
    sell_24 = ch1.get("sell_24", 0)
    sell_25 = ch1.get("sell_25", 0)
    mgmt_24 = ch1.get("mgmt_24", 0)
    mgmt_25 = ch1.get("mgmt_25", 0)
    fin_24 = ch1.get("fin_24", 0)
    fin_25 = ch1.get("fin_25", 0)
    rev_24 = ch1.get("rev_24", 1)
    rev_25 = ch1.get("rev_25", 1)

    def rate_str(v24, v25, rev24, rev25):
        r24 = v24/rev24 if rev24 > 0 else 0
        r25 = v25/rev25 if rev25 > 0 else 0
        chg = r25 - r24
        cc = "t-green" if chg < 0 else "t-orange"
        return f'{_pct(r25)}&nbsp;<span class="{cc}">{_ppt(chg)}</span>'

    html += f"""
<div class="section" id="ch1">
  <div class="ch-tag">第一章</div>
  <div class="ch-title">公司整体经营健康度</div>
  <div class="ch-sub">核心问题：公司整体是在变好还是变差？账面数字和真实经营能力之间是否存在差距？</div>

  <div class="panel">
    <div class="st">① 规模与增长</div>
    <div class="g4">
      <div class="kpi k-blue"><div class="kpi-l">营业收入（{year}）</div>
        <div class="kpi-v">{_fmt_wan(rev_25)}</div>
        <div class="kpi-s">较{base}增长{_pct(rev_g)}（目标{target_pct}%）</div></div>
      <div class="kpi {"k-red" if gm_chg < 0 else "k-green"}">
        <div class="kpi-l">综合毛利率</div>
        <div class="kpi-v">{_pct(ch1.get("gm_rate_25",0))}</div>
        <div class="kpi-s">变动 {_ppt(gm_chg)}</div></div>
      <div class="kpi {"k-green" if da < 0.55 else "k-yellow"}">
        <div class="kpi-l">资产负债率</div>
        <div class="kpi-v">{_pct(da)}</div>
        <div class="kpi-s">净有息负债 {_fmt_wan(ch1.get("net_debt",0))}</div></div>
      <div class="kpi k-{"green" if ch1.get("ocf_quality",0)>1 else "yellow"}">
        <div class="kpi-l">OCF/实质净利润</div>
        <div class="kpi-v">{_pct(ch1.get("ocf_quality",0))}</div>
        <div class="kpi-s">FCF {_fmt_wan(ch1.get("fcf",0))}</div></div>
    </div>
  </div>

  <div class="panel">
    <div class="st">② 完整利润表与三项费用率</div>
    <table class="tbl">
      <thead><tr><th>利润层级</th>
        <th class="r">{base}(万)</th><th class="r">{base}占收入%</th>
        <th class="r">{year}(万)</th><th class="r">{year}占收入% / 变动</th></tr></thead>
      <tbody>
        <tr class="row-b"><td><strong>营业收入</strong></td>
          <td class="r"><strong>{ch1.get("rev_24",0)/1e4:.0f}</strong></td><td class="r">100%</td>
          <td class="r"><strong>{rev_25/1e4:.0f}</strong></td><td class="r">100%</td></tr>
        <tr><td>减：营业成本</td>
          <td class="r">{ch1.get("cst_24",0)/1e4:.0f}</td>
          <td class="r">{_pct(ch1.get("cst_24",0)/max(rev_24,1))}</td>
          <td class="r">{ch1.get("cst_25",0)/1e4:.0f}</td>
          <td class="r {"t-red" if ch1.get("cst_growth",0)>ch1.get("rev_growth",0) else ""}">{_pct(ch1.get("cst_25",0)/max(rev_25,1))}&nbsp;{_ppt(ch1.get("cst_25",0)/max(rev_25,1)-ch1.get("cst_24",0)/max(rev_24,1))}</td></tr>
        <tr class="row-g"><td><strong>毛利额</strong></td>
          <td class="r t-green"><strong>{ch1.get("gm_24",0)/1e4:.0f}</strong></td>
          <td class="r t-green"><strong>{_pct(ch1.get("gm_rate_24",0))}</strong></td>
          <td class="r {"t-orange" if gm_chg<0 else "t-green"}"><strong>{ch1.get("gm_25",0)/1e4:.0f}</strong></td>
          <td class="r {"t-red" if gm_chg<0 else "t-green"}"><strong>{_pct(ch1.get("gm_rate_25",0))}&nbsp;{_ppt(gm_chg)}</strong></td></tr>
        <tr><td>减：销售费用</td>
          <td class="r">{sell_24/1e4:.0f}</td><td class="r">{_pct(sell_24/max(rev_24,1))}</td>
          <td class="r">{sell_25/1e4:.0f}</td><td class="r">{rate_str(sell_24,sell_25,rev_24,rev_25)}</td></tr>
        <tr class="{"row-g" if mgmt_25<mgmt_24 else ""}"><td>减：管理费用{'&nbsp;★节省'+_fmt_wan(mgmt_24-mgmt_25) if mgmt_25<mgmt_24 else ''}</td>
          <td class="r">{mgmt_24/1e4:.0f}</td><td class="r">{_pct(mgmt_24/max(rev_24,1))}</td>
          <td class="r {"t-green" if mgmt_25<mgmt_24 else ""}">{mgmt_25/1e4:.0f}</td>
          <td class="r {"t-green" if mgmt_25<mgmt_24 else ""}">{rate_str(mgmt_24,mgmt_25,rev_24,rev_25)}</td></tr>
        <tr class="{"row-g" if fin_25<fin_24 else ""}"><td>减：财务费用{'&nbsp;★节省'+_fmt_wan(fin_24-fin_25) if fin_25<fin_24 else ''}</td>
          <td class="r">{fin_24/1e4:.0f}</td><td class="r">{_pct(fin_24/max(rev_24,1))}</td>
          <td class="r {"t-green" if fin_25<fin_24 else ""}">{fin_25/1e4:.0f}</td>
          <td class="r {"t-green" if fin_25<fin_24 else ""}">{rate_str(fin_24,fin_25,rev_24,rev_25)}</td></tr>
        <tr class="row-b"><td><strong>三费合计</strong></td>
          <td class="r"><strong>{(sell_24+mgmt_24+fin_24)/1e4:.0f}</strong></td>
          <td class="r"><strong>{_pct((sell_24+mgmt_24+fin_24)/max(rev_24,1))}</strong></td>
          <td class="r t-green"><strong>{(sell_25+mgmt_25+fin_25)/1e4:.0f}</strong></td>
          <td class="r t-green"><strong>{rate_str(sell_24+mgmt_24+fin_24,sell_25+mgmt_25+fin_25,rev_24,rev_25)}</strong></td></tr>
        <tr class="row-y"><td>加：其他收益（含一次性处置收益）</td>
          <td class="r t-gray">—</td><td class="r">—</td>
          <td class="r t-orange">{ch1.get("other_income_25",0)/1e4:.0f}</td>
          <td class="r t-orange">含一次性{_fmt_wan(ch1.get("onetime_gain",0))}</td></tr>
        <tr class="row-b"><td><strong>净利润（账面）</strong></td>
          <td class="r"><strong>{ch1.get("net_24",0)/1e4:.0f}</strong></td><td class="r t-gray">—</td>
          <td class="r"><strong>{net/1e4:.0f}</strong></td><td class="r t-blue"><strong>↑{(net-ch1.get("net_24",0))/max(ch1.get("net_24",1),1)*100:.1f}%</strong></td></tr>
        <tr class="row-y"><td><strong>净利润（剔除一次性收益后）</strong></td>
          <td class="r"><strong>{ch1.get("net_24",0)/1e4:.0f}</strong></td><td class="r t-gray">—</td>
          <td class="r t-orange"><strong>{adj_net/1e4:.0f}</strong></td>
          <td class="r t-orange"><strong>实质增幅{adj_g*100:.1f}%&nbsp;⚠</strong></td></tr>
      </tbody>
    </table>
    <p class="note">★ 标注项目为本年营业利润增量的关键驱动因素</p>
  </div>

  <div class="panel">
    <div class="st">③ 财务安全性</div>
    <div class="g4">
      <div class="kpi {"k-green" if da<0.55 else "k-yellow"}">
        <div class="kpi-l">资产负债率</div>
        <div class="kpi-v">{_pct(da)}</div>
        <div class="kpi-s">{"✓ 财务目标达成" if da<0.55 else "⚠ 仍有压力"}</div></div>
      <div class="kpi k-green">
        <div class="kpi-l">有息负债/EBITDA</div>
        <div class="kpi-v">{ch1.get("debt_ebitda",0):.1f}x</div>
        <div class="kpi-s">安全线&lt;3x {"✓" if ch1.get("debt_ebitda",0)<3 else "⚠"}</div></div>
      <div class="kpi k-green">
        <div class="kpi-l">利息保障倍数</div>
        <div class="kpi-v">{min(ch1.get("interest_cover",99),99):.0f}x</div>
        <div class="kpi-s">警戒线&gt;3x {"✓" if ch1.get("interest_cover",0)>3 else "⚠"}</div></div>
      <div class="kpi k-{"orange" if ch1.get("goodwill_25",0)/max(ch1.get("total_asset_25",1),1)>0.15 else "blue"}">
        <div class="kpi-l">商誉</div>
        <div class="kpi-v">{_fmt_wan(ch1.get("goodwill_25",0))}</div>
        <div class="kpi-s">占总资产{_pct(ch1.get("goodwill_25",0)/max(ch1.get("total_asset_25",1),1))}&nbsp;{"⚠" if ch1.get("goodwill_25",0)/max(ch1.get("total_asset_25",1),1)>0.15 else ""}</div></div>
    </div>
  </div>
</div>
"""

    # ── 第二章 ──
    quads = ch2.get("quadrants", {})
    quad_configs = [
        ("住宅·自有", "1a0808", "7f1d1d", "f87171", "⚠ 核心问题来源"),
        ("后勤·自有", "0e1a10", "166534", "4ade80", "✓ 基本稳定"),
        ("住宅·合作", "0e1220", "1e3a5f", "60a5fa", "↑ 改善中"),
        ("后勤·合作", "161200", "78350f", "fbbf24", "⚠ 微利边缘"),
    ]

    quads_html = ""
    for name, bg, border, color, status in quad_configs:
        q = quads.get(name, {})
        if not q:
            continue
        gm_chg_q = q.get("gm_rate_25", 0) - q.get("gm_rate_24", 0)
        quads_html += f"""
        <div style="background:#{bg};border:1px solid #{border};border-radius:8px;padding:12px">
          <div style="font-size:10.5px;color:#64748b;margin-bottom:4px">{status}</div>
          <div style="font-size:13px;font-weight:700;color:#{color};margin-bottom:7px">{name}</div>
          <div style="font-size:11px;color:#94a3b8;line-height:1.8">
            {year}收入：<strong style="color:#e2e8f0">{_fmt_wan(q.get("rev_25",0))}</strong><br>
            收入增速：<span style="color:{"#4ade80" if q.get("rev_chg",0)>=0 else "#f87171"}">{q.get("rev_chg",0)*100:+.1f}%</span><br>
            毛利率：{_pct(q.get("gm_rate_24",0))} → <strong style="color:{"#4ade80" if gm_chg_q>=0 else "#f87171"}">{_pct(q.get("gm_rate_25",0))}（{_ppt(gm_chg_q)}）</strong><br>
            毛利额：<span style="color:{"#4ade80" if q.get("gm_25",0)>=q.get("gm_24",0) else "#f87171"}">{_fmt_wan(q.get("gm_25",0))}（{(q.get("gm_25",0)-q.get("gm_24",0))/1e4:+.0f}万）</span>
          </div>
        </div>"""

    # 区域毛利排行
    regions = ch2.get("region_analysis", {})
    regions_sorted = sorted(regions.items(), key=lambda x: -x[1].get("gm_rate_25", 0))
    region_rows = ""
    for rgn, v in regions_sorted:
        gm25 = v.get("gm_rate_25", 0)
        ppt_v = v.get("gm_ppt", 0)
        row_cls = "row-g" if gm25 > 0.20 else ("row-r" if gm25 < 0.05 else "row-y" if ppt_v < -0.03 else "")
        ppt_cls = "t-green" if ppt_v >= 0 else ("t-red" if ppt_v < -0.05 else "t-orange")
        region_rows += f"""
        <tr class="{row_cls}">
          <td>{rgn}</td>
          <td class="r">{_fmt_wan(v.get("rev_25",0))}</td>
          <td class="r">{_pct(v.get("gm_rate_24",0))}</td>
          <td class="r"><strong>{_pct(gm25)}</strong></td>
          <td class="r {ppt_cls}"><strong>{_ppt(ppt_v)}</strong></td>
          <td class="r">{_fmt_wan(v.get("gm_chg",0),"" )}</td>
        </tr>"""

    # 项目分布
    proj_dist = ch2.get("proj_dist", [])
    dist_html = ""
    dist_colors = {"亏损(<0%)": ("#7f1d1d","#fca5a5"), "微利(0-5%)": ("#78350f","#fcd34d"),
                   "正常(5-10%)": ("#1e293b","#94a3b8"), "良好(10-20%)": ("#14532d","#86efac"),
                   "优质(>20%)": ("#166534","#4ade80")}
    total_proj = sum(d.get("count", 0) for d in proj_dist)
    for d in proj_dist:
        bg, fg = dist_colors.get(str(d.get("band","")), ("#1e293b","#94a3b8"))
        cnt = d.get("count", 0)
        dist_html += f"""
        <div class="box" style="text-align:center">
          <div style="font-size:22px;font-weight:700;color:{fg}">{cnt}</div>
          <div style="font-size:9.5px;color:#64748b;margin-top:2px">占比 {cnt/max(total_proj,1)*100:.0f}%</div>
          <div style="margin-top:5px;padding:2px 0;border-radius:3px;font-size:9px;font-weight:700;background:#{bg};color:{fg}">{d.get("band","")}</div>
          <div style="font-size:9.5px;color:#475569;margin-top:3px">收入 {d.get("rev_sum",0)/1e4:.0f}万</div>
        </div>"""

    html += f"""
<div class="section" id="ch2">
  <div class="ch-tag">第二章</div>
  <div class="ch-title">利润结构解剖：谁在赚钱，谁在流血</div>
  <div class="ch-sub">核心问题：总利润由哪些业务贡献？哪些板块是拖累？毛利下滑根本原因是什么？</div>

  <div class="panel">
    <div class="st">① 业态×运营模式四象限</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">{quads_html}</div>
  </div>

  <div class="panel">
    <div class="st">② 各区域综合毛利率排行（{year}）</div>
    <table class="tbl">
      <thead><tr><th>区域</th><th class="r">{year}收入(万)</th>
        <th class="r">{base}毛利率</th><th class="r">{year}毛利率</th>
        <th class="r">变动(ppt)</th><th class="r">毛利额变动(万)</th></tr></thead>
      <tbody>{region_rows}</tbody>
    </table>
  </div>

  <div class="panel">
    <div class="st">③ 全部项目毛利率区间分布（{year}）</div>
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin-bottom:10px">{dist_html}</div>
  </div>
</div>
"""

    # ── 第三章 ──
    ar_trend = ch3.get("ar_trend", {})
    trend_vals = sorted(ar_trend.items())
    max_ar = max(v for _, v in trend_vals) if trend_vals else 1
    trend_bars = ""
    for yr, v in trend_vals:
        h = int(v / max_ar * 100)
        c = "#ef4444" if yr == max(k for k, _ in trend_vals) else "#f59e0b"
        trend_bars += f"""
        <div style="flex:1;display:flex;flex-direction:column;align-items:center">
          <div style="font-size:9.5px;font-weight:700;color:{c};margin-bottom:2px">{v/1e4:.0f}万</div>
          <div style="background:{c};border-radius:2px 2px 0 0;width:100%;height:{h}px"></div>
          <div style="font-size:9px;color:#64748b;margin-top:3px">{yr}</div>
        </div>"""

    region_df = ch3.get("region_df", {})
    max_gap = max((abs(v.get("gap", 0)) for v in region_df.values()), default=1)
    region_rows3 = ""
    for rgn, v in region_df.items():
        gap = v.get("gap", 0)
        days = v.get("ar_days", 0)
        gap_cls = "t-red" if gap > 1e7 else ("t-orange" if gap > 5e6 else ("t-green" if gap < 0 else ""))
        days_cls = "t-red" if days > 365 else ("t-orange" if days > 180 else ("t-green" if days < 90 else ""))
        row_cls = "row-r" if gap > 1e7 else ("row-y" if gap > 5e6 else ("row-g" if gap < 0 else ""))
        region_rows3 += f"""
        <tr class="{row_cls}">
          <td>{rgn}</td>
          <td class="r">{v.get("ar_25",0)/1e4:.0f}</td>
          <td class="r">{v.get("ap_25",0)/1e4:.0f}</td>
          <td class="r">{v.get("adv_25",0)/1e4:.0f}</td>
          <td class="r {gap_cls}"><strong>{gap/1e4:.0f}</strong></td>
          <td class="r {days_cls}"><strong>{days:.0f}天</strong></td>
        </tr>"""

    html += f"""
<div class="section" id="ch3">
  <div class="ch-tag">第三章</div>
  <div class="ch-title">应收·应付·预收：钱在哪里，谁在垫资</div>
  <div class="ch-sub">净资金敞口 = 期末应收 − 期末应付 − 期末预收 | 正值=公司在垫资；负值=公司用别人的钱运转</div>

  <div class="g2">
    <div class="panel">
      <div class="st">① 应收账款历年趋势（万元）</div>
      <div style="display:flex;align-items:flex-end;gap:5px;height:110px;margin:10px 0 6px">
        {trend_bars}
      </div>
      <div class="al al-r"><div class="al-icon">🔴</div><div>
        <div class="al-title">5年从未下降，是结构性危机</div>
        <div class="al-body">应收账款每年持续增长，新增欠款远超回收金额，是商业模式中长期无法向甲方完全收款的结构性缺陷。全公司净资金敞口持续扩大，制约并购扩张。</div>
      </div></div>
    </div>
    <div class="panel">
      <div class="st">② 分业态应收账款周转天数</div>
      <table class="tbl">
        <thead><tr><th>业务口径</th><th class="r">平均AR(万)</th><th class="r">周转天数</th><th>参考标准</th></tr></thead>
        <tbody>
          <tr class="row-b"><td><strong>全公司整体</strong></td>
            <td class="r"><strong>—</strong></td>
            <td class="r"><strong style="color:{"#fb923c" if ch3.get("overall_days",0)>150 else "#4ade80"}">{ch3.get("overall_days",0):.0f}天</strong></td>
            <td>—</td></tr>
          <tr class="{"row-r" if ch3.get("residential_days",0)>180 else "row-y"}">
            <td>居民生活服务（住宅）</td><td class="r">—</td>
            <td class="r"><strong class="{"t-red" if ch3.get("residential_days",0)>180 else "t-orange"}">{ch3.get("residential_days",0):.0f}天</strong></td>
            <td>&lt;180天正常</td></tr>
          <tr class="{"row-g" if ch3.get("commercial_days",0)<90 else ""}">
            <td>企事业后勤综合服务</td><td class="r">—</td>
            <td class="r"><strong class="{"t-green" if ch3.get("commercial_days",0)<90 else "t-orange"}">{ch3.get("commercial_days",0):.0f}天</strong></td>
            <td>&lt;90天优秀</td></tr>
          <tr class="row-p"><td>关联方（单独列示）</td>
            <td class="r t-purple">{ch3.get("rp_ar_25",0)/1e4:.0f}</td>
            <td class="r t-red">—</td><td>单独管理</td></tr>
        </tbody>
      </table>
      <p class="note">关联方AR：{year}末{ch3.get("rp_ar_25",0)/1e4:.0f}万（{base}末{ch3.get("rp_ar_24",0)/1e4:.0f}万），未纳入主体分析，需董事会单独关注。</p>
    </div>
  </div>

  <div class="panel">
    <div class="st">③ 各区域净资金敞口与AR周转天数（{year}末）</div>
    <table class="tbl">
      <thead><tr><th>区域</th><th class="r">AR余额(万)</th><th class="r">AP余额(万)</th>
        <th class="r">预收(万)</th><th class="r">净敞口(万)</th><th class="r">AR周转天</th></tr></thead>
      <tbody>{region_rows3}</tbody>
    </table>
    <p class="note">颜色：红=净敞口&gt;1,000万（极危）/ 黄=净敞口200-1,000万（警示）/ 绿=净敞口&lt;0（优质）</p>
  </div>
</div>
"""

    # ── 第四章 ──
    region_cf = ch4.get("region_cf", {})
    max_cf = max((abs(v.get("net", 0)) for v in region_cf.values()), default=1) if region_cf else 1

    drain_html, feed_html = "", ""
    drain_items = sorted([(k, v) for k, v in region_cf.items() if v.get("net", 0) < 0], key=lambda x: x[1]["net"])
    feed_items = sorted([(k, v) for k, v in region_cf.items() if v.get("net", 0) >= 0], key=lambda x: -x[1]["net"])

    for nm, v in drain_items:
        net_v = v.get("net", 0)
        drain_html += _bar(nm, abs(net_v), max_cf, "bf-red", f"{net_v/1e4:.0f}万", "t-red")

    for nm, v in feed_items[:8]:
        net_v = v.get("net", 0)
        feed_html += _bar(nm, net_v, max_cf, "bf-green", f"+{net_v/1e4:.0f}万", "t-green")

    cf_section = ""
    if ch4.get("has_cf") and (drain_html or feed_html):
        cf_section = f"""
  <div class="panel">
    <div class="st">③ 各区域经营现金净额（纯市场化口径，{year}）</div>
    <div class="g2">
      <div>
        <div class="h3" style="color:#f87171">▼ 现金净消耗区域</div>
        {drain_html or '<p class="note">无净消耗区域</p>'}
      </div>
      <div>
        <div class="h3" style="color:#4ade80">▲ 现金净贡献区域（TOP8）</div>
        {feed_html}
      </div>
    </div>
  </div>"""

    html += f"""
<div class="section" id="ch4">
  <div class="ch-tag">第四章</div>
  <div class="ch-title">现金流与偿债能力：公司还撑得住吗</div>
  <div class="ch-sub">核心问题：利润是真实的吗？有没有资金链风险？各区域谁在造血、谁在消耗？</div>

  <div class="panel">
    <div class="st">① 偿债能力全面健康</div>
    <div class="g4">
      <div class="kpi {"k-green" if ch4.get("debt_ebitda",99)<3 else "k-red"}">
        <div class="kpi-l">有息负债/EBITDA</div>
        <div class="kpi-v">{ch4.get("debt_ebitda",0):.2f}x</div>
        <div class="kpi-s">安全线&lt;3x {"✓" if ch4.get("debt_ebitda",99)<3 else "⚠"}</div></div>
      <div class="kpi {"k-green" if ch4.get("interest_cover",0)>3 else "k-red"}">
        <div class="kpi-l">利息保障倍数</div>
        <div class="kpi-v">{min(ch4.get("interest_cover",0),99):.0f}x</div>
        <div class="kpi-s">警戒线&gt;3x {"✓" if ch4.get("interest_cover",0)>3 else "⚠"}</div></div>
      <div class="kpi {"k-green" if ch4.get("cash_cover",0)>1.5 else "k-yellow"}">
        <div class="kpi-l">现金覆盖短期借款</div>
        <div class="kpi-v">{ch4.get("cash_cover",0):.1f}x</div>
        <div class="kpi-s">现金{_fmt_wan(ch4.get("cash_25",0))} vs 短贷{_fmt_wan(ch4.get("st_loan_25",0))}</div></div>
      <div class="kpi {"k-green" if ch4.get("net_debt",0)<0 else "k-orange"}">
        <div class="kpi-l">净有息负债</div>
        <div class="kpi-v">{_fmt_wan(ch4.get("net_debt",0))}</div>
        <div class="kpi-s">{"✓ 净现金状态" if ch4.get("net_debt",0)<0 else "⚠ 净债务状态"}</div></div>
    </div>
  </div>

  <div class="panel">
    <div class="st">② 利润含金量分析</div>
    <div class="g2">
      <div>
        <table class="tbl">
          <thead><tr><th>指标</th><th class="r">数值</th><th>解读</th></tr></thead>
          <tbody>
            <tr><td>账面净利润</td><td class="r">{_fmt_wan(ch4.get("net_25",0))}</td><td>含一次性项目</td></tr>
            <tr class="row-r"><td>减：一次性收益（估算·税后）</td>
              <td class="r t-red">-{_fmt_wan(ch4.get("onetime_gain",0))}</td><td class="t-red">不可持续</td></tr>
            <tr class="row-y"><td><strong>实质净利润</strong></td>
              <td class="r t-orange"><strong>{_fmt_wan(ch4.get("adj_net_25",0))}</strong></td><td class="t-orange">实际增速更低</td></tr>
            <tr class="row-g"><td><strong>OCF / 实质净利润</strong></td>
              <td class="r t-green"><strong>{_pct(ch4.get("ocf_quality",0))}</strong></td>
              <td class="t-green">{"✓ 利润质量高" if ch4.get("ocf_quality",0)>1 else "⚠ 利润质量偏低"}</td></tr>
            <tr><td>销售收现率</td>
              <td class="r {"t-green" if ch4.get("sales_collection_rate",0)>1 else "t-orange"}">{_pct(ch4.get("sales_collection_rate",0))}</td>
              <td>{"✓ 在回收历史欠款" if ch4.get("sales_collection_rate",0)>1 else "⚠ 当年款未全部收回"}</td></tr>
            <tr class="row-y"><td><strong>自由现金流（FCF）</strong></td>
              <td class="r t-orange"><strong>{_fmt_wan(ch4.get("fcf",0))}</strong></td>
              <td class="t-orange">{"⚠ 难支撑大型并购" if ch4.get("fcf",0)<5e7 else "✓ 充足"}</td></tr>
          </tbody>
        </table>
      </div>
      <div>
        <div class="al al-g"><div class="al-icon">🟢</div><div>
          <div class="al-title">核心业务造血能力健康</div>
          <div class="al-body">OCF/实质净利润={_pct(ch4.get("ocf_quality",0))}，说明账面利润有真实现金支撑。偿债能力指标全面健康，短期无资金链风险。</div>
        </div></div>
        <div class="al al-y"><div class="al-icon">⚠</div><div>
          <div class="al-title">并购子弹不足</div>
          <div class="al-body">FCF仅{_fmt_wan(ch4.get("fcf",0))}，而大量资金被锁在应收账款中。必须先压缩应收账款释放存量资金，才能支撑并购扩张计划。</div>
        </div></div>
      </div>
    </div>
  </div>
  {cf_section}
</div>
"""

    # ── 第五章 ──
    issues = ch5.get("issues", [])
    issue_html = ""
    for iss in issues:
        level = iss.get("level", "yellow")
        cls = f"al-{level[0]}"
        icon = "🔴" if level == "red" else ("🟡" if level == "yellow" else "🟢")
        issue_html += f"""
    <div class="al {cls}">
      <div class="al-icon">{icon}&nbsp;<strong style="font-size:12px">{iss.get("priority","")}</strong></div>
      <div>
        <div class="al-title">{iss.get("title","")}</div>
        <div class="al-body">{iss.get("detail","")}<span style="font-size:10px;color:#475569">&nbsp;来源：{iss.get("chapter","")}</span></div>
      </div>
    </div>"""

    html += f"""
<div class="section" id="ch5">
  <div class="ch-tag">第五章</div>
  <div class="ch-title">核心问题与优先级行动建议</div>
  <div class="ch-sub">基于四章数据分析的综合研判，供{config.get("report_target","董事会")}直接决策</div>

  <div class="panel">
    <div class="st">① 经营健康度综合评分</div>
    <div class="score-grid">
      {"".join(f'<div class="score-card"><div class="score-val" style="color:{score_colors.get(sc,"#94a3b8")}">{sc}</div><div class="score-lbl">{dim}</div></div>' for dim, sc in scores.items())}
    </div>
  </div>

  <div class="panel">
    <div class="st">② 自动识别的优先级行动清单</div>
    <p class="note" style="margin-bottom:10px">基于数据自动识别的问题，按严重程度排序。红色=立即行动，黄色=年内落地，绿色=战略规划。</p>
    {issue_html if issue_html else '<div class="al al-g"><div class="al-icon">🟢</div><div><div class="al-title">各项指标基本健康</div><div class="al-body">当前分析周期内未发现严重异常指标，请结合业务实际判断。</div></div></div>'}
  </div>

  <div class="summary-box">
    <div class="s-title" style="font-size:16px">给{config.get("report_target","董事会")}的一句话总结</div>
    <div class="s-body" style="font-size:12pt">
      公司财务结构{"已安全（负债率"+_pct(da)+"）" if da<0.6 else "仍有压力"}，{"但经营质量需要提升（毛利率"+_pct(gm_chg)+"变动）。" if gm_chg<0 else "经营质量基本稳定。"}<br>
      <strong style="color:#ffd700;font-size:14pt">{year}的首要任务：</strong><br>
      压缩应收账款 → 释放现金 → 修复问题区域盈利能力 → 再以更健康的基本面推进扩张
    </div>
  </div>

  <div style="margin-top:24px;padding-top:10px;border-top:1px solid #1e293b;font-size:11px;color:#334155;line-height:1.8">
    报告生成日期：{report_date}&nbsp;|&nbsp;分析范围：{base}–{year}&nbsp;|&nbsp;由财务分析报告生成系统自动生成<br>
    本报告含敏感经营信息，仅限内部使用，禁止对外传阅
  </div>
</div>
</body>
</html>"""

    return html
