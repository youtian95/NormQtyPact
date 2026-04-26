"""
NormQtyPact 基本使用示例。

演示如何将 NormQtyPact 作为 Python API 使用，
为一栋 3 层公寓楼生成非结构构件数量文件。

运行要求：
    - Windows 操作系统，并已安装 Microsoft Excel
    - 已安装 normqtypact 包（pip install normqtypact）

输出文件（写入当前工作目录）：
    - PactComponentDirectory.csv
    - PelucunComponentDirectory.csv
    - PelucunComponentDirectory.json
"""

from normqtypact import NormQtyPact

# --- 建筑参数 ---
NUM_STORIES = 3
FLOOR_AREA = [1000.0, 1000.0, 1000.0]          # 各层楼面面积（平方英尺）
OCCUPANCY_1 = ["APARTMENT", "APARTMENT", "APARTMENT"]
OCCUPANCY_2 = ["none", "none", "none"]
OCCUPANCY_3 = ["none", "none", "none"]
AREA_FRAC_1 = [1.0, 1.0, 1.0]                  # 各层主要使用类型面积占比
AREA_FRAC_2 = [0.0, 0.0, 0.0]
AREA_FRAC_3 = [0.0, 0.0, 0.0]

# --- 创建 NormQtyPact 实例 ---
obj = NormQtyPact(
    NumOfStories=NUM_STORIES,
    FloorAreaList=FLOOR_AREA,
    Occupancy1Type=OCCUPANCY_1,
    Occupancy2Type=OCCUPANCY_2,
    Occupancy3Type=OCCUPANCY_3,
    Occupancy1Area=AREA_FRAC_1,
    Occupancy2Area=AREA_FRAC_2,
    Occupancy3Area=AREA_FRAC_3,
)

# --- 生成 PACT 格式输出 ---
obj.Output_PactComponentDirectory("PactComponentDirectory.csv")
print("PACT 构件目录已写入：PactComponentDirectory.csv")

# --- 生成 Pelicun 格式输出 ---
obj.Output_PelicunComponentDirectory(
    json_path="PelucunComponentDirectory.json",
    csv_path="PelucunComponentDirectory.csv",
)
print("Pelicun 构件目录已写入：")
print("  PelucunComponentDirectory.json")
print("  PelucunComponentDirectory.csv")
