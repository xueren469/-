# 财务分析报告生成系统

上传Excel底表，自动生成五章完整财务分析报告（深色主题HTML）。

---

## 📁 项目结构

```
fin_app/
├── app.py                  # 主程序（Streamlit界面）
├── requirements.txt        # 依赖包
├── analysis/
│   ├── loader.py           # 数据加载与验证
│   └── engine.py           # 五章分析引擎
└── report/
    └── generator.py        # HTML报告生成器
```

---

## 🚀 本地运行（最快，5分钟搞定）

### 第一步：安装Python环境
推荐使用 Python 3.10 或 3.11。

### 第二步：安装依赖
```bash
pip install -r requirements.txt
```

### 第三步：启动应用
```bash
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`，就可以使用了。

---

## ☁️ 部署到Streamlit Cloud（免费，有网址，随时访问）

### 第一步：注册账号
1. 注册 GitHub 账号：https://github.com
2. 注册 Streamlit Cloud 账号：https://share.streamlit.io（用GitHub登录）

### 第二步：上传代码到GitHub
1. 在 GitHub 创建一个新的**私有仓库**（Private Repository，保护数据安全）
2. 把 `fin_app` 文件夹下的所有文件上传到仓库根目录

   方法一：直接在GitHub网页上传文件（适合不懂Git的用户）
   - 进入仓库页面 → Add file → Upload files
   - 把所有文件拖进去上传

   方法二：用Git命令
   ```bash
   git init
   git add .
   git commit -m "初始提交"
   git remote add origin https://github.com/你的用户名/你的仓库名.git
   git push -u origin main
   ```

### 第三步：在Streamlit Cloud部署
1. 登录 https://share.streamlit.io
2. 点击 "New app"
3. 选择你的GitHub仓库
4. Main file path 填写：`app.py`
5. 点击 "Deploy"
6. 等待1-2分钟，会生成一个 `https://xxx.streamlit.app` 的网址

**完成！** 把这个网址发给同事，大家都可以直接用。

---

## 📋 所需底表格式

| 底表 | 必须/建议 | 关键字段 |
|------|---------|---------|
| 收入成本明细表 | **必须** | 摘要/科目/本期借方/类型/合作自营/区域/标准名称/分析是否剔除/入场时间/撤场时间 |
| 三大报表 | **必须** | 利润表+资产负债表+现金流量表（三个Sheet或合并一个文件） |
| 应收账款表 | **必须** | 摘要(含历史年份)/期末余额/标准项目/业态/内外部/合作自有/区域/分析是否剔除 |
| 应付账款表 | **必须** | 摘要/期末余额/标准项目/内外部/类型/合作自营/公司 |
| 预收账款表 | **必须** | 摘要/期末余额/标准项目/内外部/类型/合作自营/公司 |
| 项目级经营现金流表 | **建议** | 摘要/现金流量表项/值/收支/标准名称/区域/合作自营/分析是否剔除 |

---

## ❓ 常见问题

**Q: 报错"找不到Sheet"**
A: 检查Excel文件的Sheet名称。收入成本表取Sheet1；应收/应付/预收取对应名称的Sheet；现金流取"经营性现金流"Sheet。

**Q: 报错"字段不存在"**
A: 字段名称需与底表一致。常见问题：`合作/自有` vs `合作/自营`（系统会自动处理这个差异）。

**Q: 三大报表读取失败**
A: 确认三大报表文件中有三个Sheet（利润表、资产负债表、现金流量表），或者文件只有一个Sheet时按顺序读取。

**Q: 关联方数据混入了主体分析**
A: 在左侧配置栏的"关联方项目清单"中填入关联方的标准名称，每行一个。同时确认关联方在应收账款表中的"内外部"字段有特殊标记（如"兴源高中"），填入"关联方AR标识字段"中。

**Q: 部署后报告内容显示不完整**
A: 这通常是三大报表的格式问题。系统会用关键字匹配科目名称，如果匹配失败则显示0。检查报表中"营业利润"、"净利润"、"货币资金"等科目的名称是否包含这些关键词。

---

## 🔧 定制化修改

如果需要调整分析逻辑：
- **修改取数规则**：编辑 `analysis/loader.py`
- **修改分析计算**：编辑 `analysis/engine.py`
- **修改报告样式/内容**：编辑 `report/generator.py`
- **修改上传界面**：编辑 `app.py`

每次修改代码后，如果是本地运行，Streamlit会自动检测变化并提示刷新。如果是Streamlit Cloud，把代码推送到GitHub后会自动重新部署。

---

版本：v1.0 | 基于北京兴业源科技服务集团2025年度分析实践
