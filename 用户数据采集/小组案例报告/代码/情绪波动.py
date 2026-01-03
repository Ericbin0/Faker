import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import platform
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import os

# ==========================================
# 1. 配置与中文字体
# ==========================================
system_name = platform.system()
if system_name == "Windows":
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
elif system_name == "Darwin":  # MacOS
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
else:
    plt.rcParams['font.sans-serif'] = ['sans-serif']

plt.rcParams['axes.unicode_minus'] = False


# ==========================================
# 2. 核心分析逻辑
# ==========================================

def load_data(filepath):
    """读取 CSV 或 Excel 文件"""
    if not filepath:
        return None

    try:
        if filepath.endswith('.csv'):
            # 尝试不同的编码读取 CSV
            try:
                df = pd.read_csv(filepath, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(filepath, encoding='gbk')
        elif filepath.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(filepath)
        else:
            messagebox.showerror("错误", "不支持的文件格式。请选择 CSV 或 Excel 文件。")
            return None
        return df
    except Exception as e:
        messagebox.showerror("读取失败", f"无法读取文件：{e}")
        return None


def analyze_file_volatility(df):
    """
    分析数据框中的情感波动性
    """
    # 1. 寻找分数列
    score_col = None
    possible_cols = ['Sentiment_Score', 'Score', 'Sentiment', '得分', '情感得分', '分数']

    for col in possible_cols:
        if col in df.columns:
            score_col = col
            break

    if not score_col:
        # 如果没找到，让用户输入
        score_col = simpledialog.askstring("列名确认",
                                           f"未找到默认得分列。\n现有列名：{list(df.columns)}\n请输入包含情感得分的列名：")
        if not score_col or score_col not in df.columns:
            messagebox.showerror("错误", "无效的列名，无法分析。")
            return None, None, None

    # 2. 寻找分组列 (例如年份或时期)
    group_col = None
    possible_group_cols = ['Year', 'Period', 'Stage', 'Event', '时期', '年份', '阶段']

    for col in possible_group_cols:
        if col in df.columns:
            group_col = col
            break

    if not group_col:
        group_col = simpledialog.askstring("列名确认",
                                           f"未找到默认分组列(如Year/Period)。\n现有列名：{list(df.columns)}\n请输入用于分组(前期/中期/后期)的列名：")
        if not group_col or group_col not in df.columns:
            # 如果用户不输入分组，就当做整体分析
            print("未指定分组，将视为单组数据分析。")
            group_col = None

    # 3. 开始分析
    results = {}
    raw_scores = {}

    if group_col:
        # 按组分析
        groups = df[group_col].unique()
        # 尝试排序 (如果组名包含年份)
        try:
            groups = sorted(groups)
        except:
            pass

        for group in groups:
            group_data = df[df[group_col] == group][score_col].dropna()
            if len(group_data) > 0:
                raw_scores[str(group)] = group_data.values
                results[str(group)] = np.std(group_data.values)
    else:
        # 整体分析
        data = df[score_col].dropna()
        if len(data) > 0:
            raw_scores["All Data"] = data.values
            results["All Data"] = np.std(data.values)

    return results, raw_scores, score_col


def main():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    print("=== Faker 情感波动性分析工具 (自定义文件版) ===")
    print("请选择包含情感得分的 CSV 或 Excel 文件...")

    # 1. 选择文件
    file_path = filedialog.askopenfilename(
        title="选择情感分析结果文件",
        filetypes=[("Data Files", "*.csv *.xlsx *.xls")]
    )

    if not file_path:
        print("未选择文件，程序退出。")
        return

    print(f"正在读取: {os.path.basename(file_path)}...")
    df = load_data(file_path)

    if df is None:
        return

    # 2. 分析数据
    volatilities, all_scores_dict, score_col_name = analyze_file_volatility(df)

    if not volatilities:
        print("分析失败，没有有效数据。")
        return

    # 准备绘图数据
    labels = list(volatilities.keys())
    vol_values = list(volatilities.values())
    score_distributions = list(all_scores_dict.values())

    print(f"\n📊 分析结果 (基于列: {score_col_name}):")
    print(f"{'分组':<15} | {'波动性 (标准差)':<15} | {'心理状态评价'}")
    print("-" * 60)

    for label, vol in volatilities.items():
        if vol > 0.5:
            status = "极度不稳定 (High)"  # 假设得分是 -1到1 或类似的小数
        elif vol > 15:
            status = "极度不稳定 (High)"  # 假设得分是 0-100
        elif vol > 10:
            status = "波动较大 (Moderate)"
        else:
            status = "相对稳定 (Stable)"
        print(f"{label:<15} | {vol:<15.4f} | {status}")

    # ==========================================
    # 3. 可视化
    # ==========================================
    plt.figure(figsize=(12, 8))

    # --- 子图 1: 箱线图 ---
    plt.subplot(2, 1, 1)
    box = plt.boxplot(score_distributions, labels=labels, patch_artist=True, vert=False)

    # 自动生成颜色
    colors = plt.cm.Set3(np.linspace(0, 1, len(labels)))
    for patch, color in zip(box['boxes'], colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.7)

    plt.title(f'各时期情感得分分布 (列: {score_col_name})', fontsize=14)
    plt.xlabel('情感得分 (Score)', fontsize=12)
    plt.grid(axis='x', linestyle='--', alpha=0.3)

    # --- 子图 2: 波动性趋势 ---
    plt.subplot(2, 1, 2)
    x = np.arange(len(labels))
    plt.plot(x, vol_values, marker='o', markersize=10, linewidth=3, color='#FF5733', linestyle='-')
    plt.fill_between(x, vol_values, color='#FF5733', alpha=0.1)

    plt.title('心理波动性 (标准差) 演变趋势', fontsize=14)
    plt.ylabel('标准差 (Standard Deviation)', fontsize=12)
    plt.xticks(x, labels, fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.3)

    # 尝试自动标注最大最小值
    max_idx = np.argmax(vol_values)
    min_idx = np.argmin(vol_values)

    plt.annotate('波动最大', xy=(max_idx, vol_values[max_idx]),
                 xytext=(max_idx, vol_values[max_idx] * 1.1),
                 ha='center', color='#d62728', fontweight='bold',
                 arrowprops=dict(arrowstyle='->', color='#d62728'))

    plt.annotate('最稳定', xy=(min_idx, vol_values[min_idx]),
                 xytext=(min_idx, vol_values[min_idx] * 1.1),
                 ha='center', color='#2ca02c', fontweight='bold',
                 arrowprops=dict(arrowstyle='->', color='#2ca02c'))

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    main()