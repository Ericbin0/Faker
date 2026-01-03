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
# 自动检测系统并设置中文字体，防止乱码
system_name = platform.system()
if system_name == "Windows":
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
elif system_name == "Darwin":  # MacOS
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
else:
    plt.rcParams['font.sans-serif'] = ['sans-serif']

plt.rcParams['axes.unicode_minus'] = False

# ==========================================
# 2. 定义“身份认同”词典 (Identity Dictionaries)
# ==========================================

# 🔴 个人/自我 (Personal/Self)
# 核心逻辑：强调“我”的主体性，关注个人表现、荣誉与责任
personal_keywords = {
    # 第一人称代词
    "我", "我的", "自己", "个人", "私心",
    # 强调个人成就/行为的词
    "单杀", "证明", "最强", "第一", "无敌", "统治",
    "必须", "赢", "夺冠", "表现", "当饭吃", "焦点",
    "责任", "误判", "不足", "评价", "反思", "压力", "方向", "出路", "信心",
    "碾压", "愤怒", "击败", "完美"
}

# 🔵 团队/集体 (Team/Collective)
# 核心逻辑：强调“我们”的共同体，关注连接、协作与他者
team_keywords = {
    # 复数代词与集体名词
    "我们", "我们的", "团队", "队伍", "队友", "大家", "SKT", "T1", "兄弟们",
    # 强调连接/互动的词
    "配合", "合作", "帮助", "感谢", "感激", "谢谢",
    "粉丝", "支持", "夸赞", "对手", "一起", "享受", "幸福", "热情", "感恩", "过程",
    "快乐", "启发", "信任"
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


def calculate_identity_density(text):
    """计算文本中两类关键词的密度"""
    if not text: return 0, 0
    words = list(jieba.cut(text))
    total_words = len(words)
    if total_words == 0: return 0, 0

    p_count = sum(1 for w in words if w in personal_keywords)
    t_count = sum(1 for w in words if w in team_keywords)

    # 计算密度 (x100 转为百分比，再乘系数放大视觉差异)
    p_score = (p_count / total_words) * 100 * 2.5
    t_score = (t_count / total_words) * 100 * 2.5

    return p_score, t_score


# ==========================================
# 4. 主程序逻辑
# ==========================================

def main():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    print("=== Faker 身份认同转变分析工具 (本地文件版) ===")
    print("请按照提示依次选择三个时期的 Word 文档 (.docx)")

    stages = ["前期 (Early)", "中期 (Middle)", "后期 (Late)"]
    p_scores = []
    t_scores = []
    file_names = []

    # 依次选择文件
    for stage in stages:
        messagebox.showinfo("选择文件", f"请选择【{stage}】的采访文档 (.docx)")
        path = filedialog.askopenfilename(
            title=f"选择 {stage} 文档",
            filetypes=[("Word Documents", "*.docx")]
        )

        if path:
            print(f"正在分析: {os.path.basename(path)}...")
            text = read_word_file(path)
            p, t = calculate_identity_density(text)
            p_scores.append(p)
            t_scores.append(t)
            file_names.append(os.path.basename(path))
        else:
            print(f"⚠️ 跳过 {stage} (未选择文件)，数值记为 0")
            p_scores.append(0)
            t_scores.append(0)
            file_names.append("未选择")

    print("\n📊 分析结果:")
    print(f"{'阶段':<15} | {'个人词频 (I/Me)':<18} | {'团队词频 (We/Us)':<18}")
    print("-" * 55)
    for i in range(3):
        print(f"{stages[i]:<15} | {p_scores[i]:<18.2f} | {t_scores[i]:<18.2f}")

    # ==========================================
    # 5. 可视化绘制
    # ==========================================

    plt.figure(figsize=(12, 7))
    x = np.arange(len(stages))
    width = 0.35

    # --- 绘制双柱状图 ---
    bars1 = plt.bar(x - width / 2, p_scores, width, label='个人/自我 (I/Me)', color='#d62728', alpha=0.85)
    bars2 = plt.bar(x + width / 2, t_scores, width, label='团队/集体 (We/Team)', color='#1f77b4', alpha=0.85)

    # --- 绘制趋势线 ---
    plt.plot(x - width / 2, p_scores, color='#d62728', marker='o', linewidth=2, linestyle='--', alpha=0.4)
    plt.plot(x + width / 2, t_scores, color='#1f77b4', marker='s', linewidth=2, linestyle='--', alpha=0.4)

    # 装饰图表
    plt.title('从“我”到“我们”：Faker 职业生涯身份认同转变分析', fontsize=16, pad=20)
    plt.ylabel('词汇密度指数 (Density Index)', fontsize=12)
    plt.xticks(x, stages, fontsize=12)
    plt.legend(fontsize=11)
    plt.grid(axis='y', linestyle='--', alpha=0.3)

    # 显示数值
    def add_labels(bars, color):
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2., height + 0.1,
                     f'{height:.1f}', ha='center', va='bottom', color=color, fontweight='bold')

    add_labels(bars1, '#d62728')
    add_labels(bars2, '#1f77b4')

    # 添加解读标签 (根据数值大小动态调整位置)
    try:
        if p_scores[0] > t_scores[0]:
            plt.annotate('孤胆英雄', xy=(0 - width / 2, p_scores[0]), xytext=(0 - width / 2, p_scores[0] + 2),
                         ha='center', color='#d62728', fontweight='bold')

        if t_scores[2] > p_scores[2]:
            plt.annotate('精神领袖', xy=(2 + width / 2, t_scores[2]), xytext=(2 + width / 2, t_scores[2] + 2),
                         ha='center', color='#1f77b4', fontweight='bold')
    except:
        pass

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    main()