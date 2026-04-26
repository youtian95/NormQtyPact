import argparse
import sys

from .core import NormQtyPact


def main(args=None):
    """NormQtyPact 命令行入口。"""
    parser = argparse.ArgumentParser(
        prog="normqtypact",
        description="使用 FEMA P-58 标准数量估算工具生成非结构构件数量文件。",
    )
    parser.add_argument("--NumOfStories", type=int, default=3,
                        help="楼层数（默认：3）")
    parser.add_argument("--FloorAreaList", type=str, default="1,1,1",
                        help="各层楼面面积（平方英尺），逗号分隔（默认：'1,1,1'）")
    parser.add_argument("--Occupancy1Type", type=str, default="APARTMENT,APARTMENT,APARTMENT",
                        help="各层主要使用类型，逗号分隔")
    parser.add_argument("--Occupancy2Type", type=str, default="none,none,none",
                        help="各层次要使用类型，逗号分隔")
    parser.add_argument("--Occupancy3Type", type=str, default="none,none,none",
                        help="各层第三使用类型，逗号分隔")
    parser.add_argument("--Occupancy1Area", type=str, default="1,1,1",
                        help="各层主要使用类型面积占比，逗号分隔")
    parser.add_argument("--Occupancy2Area", type=str, default="0,0,0",
                        help="各层次要使用类型面积占比，逗号分隔")
    parser.add_argument("--Occupancy3Area", type=str, default="0,0,0",
                        help="各层第三使用类型面积占比，逗号分隔")

    parsed = parser.parse_args(args)

    obj = NormQtyPact(
        NumOfStories=parsed.NumOfStories,
        FloorAreaList=list(map(float, parsed.FloorAreaList.split(","))),
        Occupancy1Type=parsed.Occupancy1Type.split(","),
        Occupancy2Type=parsed.Occupancy2Type.split(","),
        Occupancy3Type=parsed.Occupancy3Type.split(","),
        Occupancy1Area=list(map(float, parsed.Occupancy1Area.split(","))),
        Occupancy2Area=list(map(float, parsed.Occupancy2Area.split(","))),
        Occupancy3Area=list(map(float, parsed.Occupancy3Area.split(","))),
    )
    obj.Output_PactComponentDirectory()
    obj.Output_PelicunComponentDirectory()


if __name__ == "__main__":
    main(sys.argv[1:])
