"""
word_generator.py — Word (.docx) 报告生成器
使用 Node.js docx 库，通过子进程生成 .docx 文件
"""
import subprocess, tempfile, os, json


def _fmt(v, unit="万", dec=0):
    v = float(v or 0)
    if unit == "万":
        return f"{v/1e4:.{dec}f}万"
    elif unit == "亿":
        return f"{v/1e8:.{dec}f}亿" if abs(v) >= 1e8 else f"{v/1e4:.{dec}f}万"
    return f"{v:.{dec}f}"

def _pct(v, dec=1):
    return f"{float(v or 0)*100:.{dec}f}%"

def _ppt(v, dec=2):
    v = float(v or 0)
    s = "+" if v >= 0 else ""
    return f"{s}{v*100:.{dec}f}ppt"


def build_word_report(results, config, out_path):
    """
    生成 Word 报告。
    results: run_all_chapters() 的返回值
    config:  报告配置
    out_path: 输出 .docx 文件路径
    """
    ch1 = results["ch1"]
    ch2 = results["ch2"]
    ch3 = results["ch3"]
    ch4 = results["ch4"]
    ch5 = results["ch5"]
    company    = config.get("company_name", "公司")
    year       = config["report_year"]
    base       = config["base_year"]
    rpt_date   = config["report_date"]
    rpt_target = config.get("report_target", "董事会")

    # ── 把分析结果序列化为 JSON 传给 Node 脚本 ──
    payload = {
        "outPath": out_path,
        "company": company,
        "year": year,
        "base": base,
        "rptDate": rpt_date,
        "rptTarget": rpt_target,
        "ch1": _serialize(ch1),
        "ch2": _serialize(ch2),
        "ch3": _serialize(ch3),
        "ch4": _serialize(ch4),
        "ch5": _serialize(ch5),
    }

    # 把 payload 写到临时 JSON 文件
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json',
                                     delete=False, encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False)
        json_path = f.name

    # 调用 Node.js 脚本生成 docx
    script_dir = os.path.dirname(__file__)
    node_script = os.path.join(script_dir, "word_builder.js")

    try:
        result = subprocess.run(
            ["node", node_script, json_path],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(f"Word生成失败：{result.stderr}")
    finally:
        os.unlink(json_path)

    return out_path


def _serialize(obj):
    """把含有 numpy/pandas 类型的对象序列化为可 JSON 化的格式"""
    import numpy as np
    if isinstance(obj, dict):
        return {str(k): _serialize(v) for k, v in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [_serialize(i) for i in obj]
    elif isinstance(obj, float) and (obj != obj):  # NaN
        return 0
    elif hasattr(obj, 'item'):  # numpy scalar
        return obj.item()
    elif hasattr(obj, 'to_dict'):  # DataFrame
        return _serialize(obj.to_dict())
    else:
        try:
            json.dumps(obj)
            return obj
        except Exception:
            return str(obj)
