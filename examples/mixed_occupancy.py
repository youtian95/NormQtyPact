"""
NormQtyPact 混合使用类型建筑示例。

演示一栋 5 层建筑：首层为零售 + 办公混合使用，其余楼层为纯办公使用。
"""

from normqtypact import NormQtyPact

obj = NormQtyPact(
    NumOfStories=5,
    FloorAreaList=[2000.0, 1500.0, 1500.0, 1500.0, 1500.0],
    # 首层：60% 零售 + 40% 办公；其余楼层：100% 办公
    Occupancy1Type=["RETAIL", "OFFICE", "OFFICE", "OFFICE", "OFFICE"],
    Occupancy2Type=["OFFICE", "none",   "none",   "none",   "none"],
    Occupancy3Type=["none",   "none",   "none",   "none",   "none"],
    Occupancy1Area=[0.6,  1.0, 1.0, 1.0, 1.0],
    Occupancy2Area=[0.4,  0.0, 0.0, 0.0, 0.0],
    Occupancy3Area=[0.0,  0.0, 0.0, 0.0, 0.0],
)

obj.Output_PactComponentDirectory("PactComponentDirectory_mixed.csv")
obj.Output_PelicunComponentDirectory(
    json_path="PelucunComponentDirectory_mixed.json",
    csv_path="PelucunComponentDirectory_mixed.csv",
)

print("混合使用类型示例输出文件已写入。")
