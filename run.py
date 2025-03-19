import pandas as pd
import shutil
import os
import sys

# 在当前代码运行前执行“源文件整理1-统一格式.py”
source_script = "源文件整理1-统一格式.py"
try:
    # 使用 exec() 执行外部 Python 脚本
    with open(source_script, 'r', encoding='utf-8') as file:
        exec(file.read())
    print(f"已成功运行 {source_script}")
except FileNotFoundError:
    print(f"错误：找不到 {source_script} 文件")
    sys.exit(1)
except Exception as e:
    print(f"运行 {source_script} 时出错：{str(e)}")
    sys.exit(1)

# 定义文件路径
original_file = "处理后的文件/整理源文件格式/英语AB级题目及知识点对照表.xlsx"
temp_file = "处理后的文件/整理源文件格式/英语AB级题目及知识点对照表_临时副本.xlsx"
knowledge_file = "处理后的文件/整理源文件格式/英语AB级知识点总表_已处理.xlsx"
types_file = "处理后的文件/整理源文件格式/英语AB级题目类型总表.xlsx"
vocab_file = "处理后的文件/整理源文件格式/英语AB级词汇总表_已处理.xlsx"

# 创建临时副本
shutil.copyfile(original_file, temp_file)

# 读取所有需要的 Excel 文件
df_questions = pd.read_excel(temp_file)
df_knowledge = pd.read_excel(knowledge_file)
df_types = pd.read_excel(types_file)
df_vocab = pd.read_excel(vocab_file)

# 获取需要排除的列
exclude_columns = ["编号", "试卷类型", "年月", "详解"]
# 获取需要检查的列
check_columns = [col for col in df_questions.columns if col not in exclude_columns]

# 获取三个参考表的编号列
knowledge_ids = set(df_knowledge["编号"].astype(str).tolist())
type_ids = set(df_types["编号"].astype(str).tolist())
vocab_ids = set(df_vocab["编号"].astype(str).tolist())

# 存储未匹配的数据
unmatched_data = []

# 遍历每一行每一列的数据
for index, row in df_questions.iterrows():
    for col in check_columns:
        cell_value = str(row[col])  # 转换为字符串以确保匹配一致性
        # 忽略空值和"nan"
        if (pd.notna(cell_value) and 
            cell_value.strip() and 
            cell_value.lower() != "nan"):
            # 检查是否在任一参考表中找到匹配
            if (cell_value not in knowledge_ids and 
                cell_value not in type_ids and 
                cell_value not in vocab_ids):
                unmatched_data.append(cell_value)

# 删除临时副本
os.remove(temp_file)

# 在程序窗口输出结果
if unmatched_data:
    print("未匹配的数据如下：")
    for value in unmatched_data:
        print(value)
    print(f"\n总共找到 {len(unmatched_data)} 个未匹配的单元格")
else:
    print("所有数据都在参考表中找到匹配")
