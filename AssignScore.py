import pandas as pd
import numpy as np
import sys

def print_help():
    print("用法: python rank.py <输入数据>")
    print("输入数据: 无表头的学生名、原始分数的两列excel文件")
    print("输出数据: '赋分后结果.xlsx'")
    print("请确保'等级折算赋分方案.xlsx'已在当前目录下")
    sys.exit(1)

if len(sys.argv) != 2:
    print("请输入学生名、原始分数的两列excel文件")
    print_help()

# 读取等级折算赋分方案，并重新计算总占比
try:
    rank_df = pd.read_excel("./等级折算赋分方案.xlsx")
except FileNotFoundError:
    print("请将'等级折算赋分方案.xlsx'放在当前目录下")
    print_help()
except Exception as e:
    print(e)
    print("请检查'等级折算赋分方案.xlsx'格式是否正确")
    print("请确保已安装pandas库以及openpyxl库")
rank_percent = rank_df["占比"]
sum_percent = [sum(rank_percent[: i + 1]) for i in range(len(rank_percent))]
score = rank_df["赋分"]
ranking = rank_df["等级"]
rank_df = pd.DataFrame(
    {"ranking": ranking, "score": score, "top_percent": np.array(sum_percent) / 100}
)

# 输入无表头的学生名、原始分数的两列excel文件
data = sys.argv[1]
data = pd.read_excel(data, header=None)
# 删除空值
data = data.dropna()
# 重命名列名
clean_data = pd.DataFrame({"name": data[0], "score": data[1]})

# 对原始分数进行排序，并计算排名与排名占比
clean_data.sort_values(by="score", ascending=False, inplace=True)
clean_data["rank"] = clean_data["score"].rank(ascending=False, method="min")

total_people = len(clean_data)
print("总人数：", total_people)

clean_data["rank_percentage"] = clean_data["rank"] / total_people

# 将clean_data的rank_percentage按照rank_df的top_percent进行匹配分组
clean_data["ranking"] = clean_data["rank_percentage"].apply(
    lambda x: rank_df[rank_df["top_percent"] >= x].iloc[0]["ranking"]
)
clean_data["rank_score"] = clean_data["ranking"].apply(
    lambda x: rank_df[rank_df["ranking"] == x].iloc[0]["score"]
)
# 输出中文结果
clean_data.rename(
    columns={
        "name": "姓名",
        "score": "原始分数",
        "rank": "排名",
        "rank_percentage": "排名占比",
        "ranking": "等级",
        "rank_score": "赋分",
    },
    inplace=True,
)

clean_data.to_excel("赋分后结果.xlsx", index=False)
