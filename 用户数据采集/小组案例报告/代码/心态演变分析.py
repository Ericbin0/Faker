import matplotlib.pyplot as plt
import numpy as np
import jieba
import docx
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import platform

# ==========================================
# 1. 配置与中文字体
# ==========================================
# 解决 Matplotlib 中文乱码问题
system_name = platform.system()
if system_name == "Windows":
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
elif system_name == "Darwin":  # MacOS
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
else:
    plt.rcParams['font.sans-serif'] = ['sans-serif']

plt.rcParams['axes.unicode_minus'] = False

# ==========================================
# 2. 定义情感维度词典 (核心评分标准)
# ==========================================

# 🔴 维度 A: 锋芒/攻击性 (Aggression/Ego)
# 代表：早期“大魔王”时期的自信、征服欲和自我中心
aggression_keywords = {
    "杀", "击杀", "单杀", "碾压", "摧毁", "打爆",
    "证明", "最强", "第一", "冠军", "无敌", "神",
    "我", "我的", "自己", "统治", "愤怒", "垃圾",
    "处刑", "傲慢", "野心", "王座", "必须赢", "赢"
}

# 🟢 维度 B: 沉稳/谦虚 (Maturity/Humility)
# 代表：后期“求道者”时期的感恩、团队、客观和哲学思考
maturity_keywords = {
    "感谢", "感激", "谢谢", "运气", "多亏", "抱歉",
    "队友", "团队", "我们", "大家", "配合", "失误",
    "学习", "过程", "客观", "健康", "心态", "读书",
    "冥想", "平静", "享受", "准备", "不足", "改进",
    "粉丝", "责任", "沉稳", "接受", "下滑"
}


# ==========================================
# 3. 工具函数
# ==========================================

def read_word_file(filepath):
    """读取 .docx 文件中的所有文本"""
    if not filepath:
        return ""
    try:
        doc = docx.Document(filepath)
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())
        return "\n".join(full_text)
    except Exception as e:
        print(f"❌ 读取文件失败: {filepath}\n错误: {e}")
        return ""


def calculate_density(text):
    """
    计算文本中两类关键词的密度
    返回: (锋芒得分, 沉稳得分)
    """
    if not text:
        return 0, 0

    words = list(jieba.cut(text))
    total_words = len(words)

    if total_words == 0:
        return 0, 0

    agg_count = sum(1 for w in words if w in aggression_keywords)
    mat_count = sum(1 for w in words if w in maturity_keywords)

    # 计算密度系数 (为了图表显示效果，乘以 100)
    agg_score = (agg_count / total_words) * 100
    mat_score = (mat_count / total_words) * 100

    return agg_score, mat_score


# ==========================================
# 4. 主程序
# ==========================================

def main():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    print("=== Faker 心态演变分析工具启动 ===")

    # 存储三个阶段的数据
    stages = ["前期 (2013-2017)", "中期 (2018-2021)", "后期 (2022-至今)"]
    file_paths = []

    # --- 1. 依次选择文件 ---
    messagebox.showinfo("步骤说明", "请依次选择三个时期的 Word 文档：\n1. 前期\n2. 中期\n3. 后期")

    for stage in stages:
        print(f"📂 请选择 [{stage}] 的文档...")
        path = filedialog.askopenfilename(
            title=f"选择 {stage} 的采访文档",
            filetypes=[("Word Documents", "*.docx")]
        )
        if not path:
            print(f"⚠️ 跳过或未选择 {stage}，程序将使用模拟数据演示该阶段。")
            file_paths.append(None)
        else:
            file_paths.append(path)

    # --- 2. 计算得分 ---
    agg_scores = []
    mat_scores = []

    # 默认模拟数据 (以防用户未选择文件)
    mock_data = [
        ("我要杀光他们证明我是最强", 8.0, 1.0),  # 前期: 高锋芒
        ("输了很难过但我必须承担责任", 4.0, 3.5),  # 中期: 纠结
        ("感谢队友和粉丝让我享受过程", 1.5, 7.0)  # 后期: 高沉稳
    ]

    print("\n📊 分析结果:")
    print("-" * 50)
    print(f"{'阶段':<15} | {'锋芒指数 (Agg)':<15} | {'沉稳指数 (Mat)':<15}")
    print("-" * 50)

    for i, path in enumerate(file_paths):
        if path:
            text = read_word_file(path)
            a_score, m_score = calculate_density(text)
            # 简单的归一化/放大处理，确保图表好看
            # 如果文本很长，密度可能会很小，这里做个动态调整
            scale_factor = 2.0
            a_score *= scale_factor
            m_score *= scale_factor
        else:
            # 使用模拟数据
            a_score, m_score = mock_data[i][1], mock_data[i][2]

        agg_scores.append(a_score)
        mat_scores.append(m_score)
        print(f"{stages[i]:<15} | {a_score:<15.2f} | {m_score:<15.2f}")

    # ==========================================
    # 5. 可视化生成
    # ==========================================

    # --- 图表 A: 演变折线图 ---
    plt.figure(figsize=(14, 6))

    x_axis = np.arange(len(stages))

    # 绘制曲线
    plt.plot(x_axis, agg_scores, marker='o', linestyle='-', linewidth=3, color='#d62728',
             label='锋芒/攻击性 (Aggression)')
    plt.plot(x_axis, mat_scores, marker='s', linestyle='-', linewidth=3, color='#2ca02c', label='沉稳/谦逊 (Humility)')

    # 填充交叉区域
    # 为了 fill_between 正常工作，需要插值让曲线平滑 (这里简化处理，直接连线)
    plt.fill_between(x_axis, agg_scores, mat_scores, where=(np.array(agg_scores) > np.array(mat_scores)),
                     interpolate=True, color='#d62728', alpha=0.1)
    plt.fill_between(x_axis, agg_scores, mat_scores, where=(np.array(agg_scores) <= np.array(mat_scores)),
                     interpolate=True, color='#2ca02c', alpha=0.1)

    # 装饰图表
    plt.title('Faker 职业生涯心态演变轨迹 (基于词频占比分析)', fontsize=16, pad=20)
    plt.ylabel('关键词密度指数', fontsize=12)
    plt.xticks(x_axis, stages, fontsize=12)
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.legend(fontsize=12)

    # 标注关键节点
    for i, txt in enumerate(agg_scores):
        plt.annotate(f"{txt:.1f}", (x_axis[i], agg_scores[i]), textcoords="offset points", xytext=(0, 5), ha='center',
                     color='#d62728')
    for i, txt in enumerate(mat_scores):
        plt.annotate(f"{txt:.1f}", (x_axis[i], mat_scores[i]), textcoords="offset points", xytext=(0, -15), ha='center',
                     color='#2ca02c')

    plt.tight_layout()
    plt.show()

    # --- 图表 B: 前后期雷达对比图 ---
    # 这是一个简化的五维推断，基于我们的两个核心得分进行映射
    # 逻辑：
    # 攻击欲 ≈ 锋芒指数
    # 自我中心 ≈ 锋芒指数 * 0.8
    # 团队意识 ≈ 沉稳指数 * 1.2
    # 抗压能力 ≈ (沉稳指数 + 锋芒指数) / 2 (中期通常最低)
    # 哲学/感恩 ≈ 沉稳指数

    def get_radar_data(agg, mat):
        # 限制在 0-10 分之间
        def limit(x): return min(max(x, 1), 10)

        return [
            limit(agg * 1.2),  # 攻击欲
            limit(agg * 1.0),  # 自我中心
            limit(mat * 1.5),  # 团队意识
            limit((agg + mat) / 1.5),  # 抗压/心态管理
            limit(mat * 1.2)  # 哲学/感恩
        ]

    labels = np.array(['攻击欲', '自我中心', '团队意识', '抗压/心态', '哲学/感恩'])
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    angles += angles[:1]  # 闭合

    # 获取前期和后期的数据
    data_early = get_radar_data(agg_scores[0], mat_scores[0])
    data_late = get_radar_data(agg_scores[2], mat_scores[2])

    data_early += data_early[:1]
    data_late += data_late[:1]

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))

    # 绘图
    ax.plot(angles, data_early, color='#d62728', linewidth=2, label='前期 (Early)')
    ax.fill(angles, data_early, color='#d62728', alpha=0.25)

    ax.plot(angles, data_late, color='#2ca02c', linewidth=2, label='后期 (Late)')
    ax.fill(angles, data_late, color='#2ca02c', alpha=0.25)

    ax.set_yticklabels([])
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=12)
    plt.title('Faker 心态模型重构对比 (前期 vs 后期)', fontsize=16, pad=20)
    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    main()