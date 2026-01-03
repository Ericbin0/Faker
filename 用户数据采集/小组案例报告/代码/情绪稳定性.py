import pandas as pd
from snownlp import SnowNLP
from docx import Document
import re
import os
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt
import platform

# ==========================================
# 1. 配置与字体设置 (解决中文乱码)
# ==========================================
system_name = platform.system()
if system_name == "Windows":
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
elif system_name == "Darwin":  # MacOS
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
else:
    plt.rcParams['font.sans-serif'] = ['sans-serif']

plt.rcParams['axes.unicode_minus'] = False


def extract_text_from_docx(file_path):
    """从 Word 文档中提取所有文本"""
    if not os.path.exists(file_path):
        return ""
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())
        return "\n".join(full_text)
    except Exception as e:
        print(f"读取错误: {e}")
        return ""


def analyze_sentiment(text, period_label, source_label):
    """
    对文本进行分句清洗和情感打分
    period_label: '前期' 或 '后期'
    source_label: '采访' 或 '纪录片' (文件名)
    """
    # 1. 简单清洗：去除多余空白和常见的时间戳格式 (如 [12:30])
    text = re.sub(r'\[\d{2}:\d{2}.*?\]', '', text)
    text = text.replace('\n', ' ').replace('\r', ' ')

    # 2. 分句：按中文标点切分
    sentences = re.split(r'[。！？!?]', text)

    data = []
    for sent in sentences:
        sent = sent.strip()
        # 过滤掉太短的句子
        if len(sent) < 4:
            continue

        # 3. 情感打分 (SnowNLP)
        try:
            s = SnowNLP(sent)
            # 映射到 -1 到 1
            score = (s.sentiments - 0.5) * 2
        except:
            score = 0.0

        data.append({
            'Period': period_label,
            'Source': source_label,
            'Sentence': sent,
            'Sentiment_Score': round(score, 4)
        })

    return data


def select_files(title):
    """弹出文件选择框"""
    print(f"\n请选择【{title}】的 Word 文档 (支持多选)...")
    file_paths = filedialog.askopenfilenames(title=f"选择{title}文档", filetypes=[("Word", "*.docx")])
    return file_paths


def plot_variance_comparison(stats_df):
    """
    绘制情感方差对比图
    stats_df: 包含 'Period' 和 'var' 列的 DataFrame
    """
    if stats_df.empty or 'var' not in stats_df.columns:
        print("无有效统计数据，无法绘图。")
        return

    # 准备数据
    periods = stats_df.index.tolist()
    variances = stats_df['var'].fillna(0).tolist()

    # 设置颜色：前期红色(波动大)，后期绿色(平稳)
    colors = ['#d62728' if '前期' in str(p) else '#2ca02c' for p in periods]

    # 创建画布
    plt.figure(figsize=(10, 6), dpi=120)

    # 绘制柱状图
    bars = plt.bar(periods, variances, color=colors, alpha=0.8, width=0.5)

    # 添加数值标签
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2., height,
                 f'{height:.4f}',
                 ha='center', va='bottom', fontsize=12, fontweight='bold')

    # 装饰图表
    plt.title('Faker 职业生涯情感稳定性对比 (方差越小越稳定)', fontsize=16, pad=20)
    plt.ylabel('情感得分方差 (Variance)', fontsize=12)
    plt.xlabel('职业阶段', fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.3)

    # 添加解读文本
    if len(variances) >= 2:
        diff = variances[0] - variances[-1]
        if diff > 0:
            note = f"📉 方差下降 {diff:.3f}\n(情绪控制力显著提升)"
            plt.annotate(note,
                         xy=(1, variances[-1]),
                         xytext=(0.5, max(variances) * 0.8),
                         arrowprops=dict(facecolor='gray', shrink=0.05, linestyle='--'),
                         fontsize=11, bbox=dict(boxstyle="round", fc="white", ec="gray", alpha=0.9))

    plt.tight_layout()

    # 保存并显示
    save_path = 'faker_variance_comparison.png'
    plt.savefig(save_path)
    print(f"\n[可视化完成] 图表已保存为: {save_path}")
    plt.show()


def main():
    print("=== Faker 文本情感量化工具 (通用版 + 可视化) ===")

    root = tk.Tk()
    root.withdraw()

    all_data = []

    # 1. 选择前期文件
    early_files = select_files("前期 (Early Career)")
    for f in early_files:
        print(f"正在处理前期文档: {os.path.basename(f)}...")
        text = extract_text_from_docx(f)
        if text:
            all_data.extend(analyze_sentiment(text, "前期", os.path.basename(f)))

    # 2. 选择后期文件
    late_files = select_files("后期 (Late Career)")
    for f in late_files:
        print(f"正在处理后期文档: {os.path.basename(f)}...")
        text = extract_text_from_docx(f)
        if text:
            all_data.extend(analyze_sentiment(text, "后期", os.path.basename(f)))

    # 3. 导出与分析
    if all_data:
        df = pd.DataFrame(all_data)
        output_file = "faker_sentiment_analysis_final.xlsx"
        df.to_excel(output_file, index=False)

        print("\n" + "=" * 30)
        print(f"处理完成！数据已保存为: {output_file}")
        print(f"共提取句子: {len(df)} 条")
        print("=" * 30)

        # 自动计算方差
        print("\n[关键指标预览: 情绪稳定性分析]")
        # 聚合计算均值和方差
        stats = df.groupby('Period')['Sentiment_Score'].agg(['count', 'mean', 'var'])
        print(stats)

        # === 新增：调用可视化函数 ===
        plot_variance_comparison(stats)

    else:
        print("未选择任何文件或提取失败。")


if __name__ == "__main__":
    main()