########################################################
# 使用 Python 脚本调用 FEMA P-58 标准数量估算工具（Excel 文件）。
########################################################

import os
import csv
from win32com.client import Dispatch, GetActiveObject
import json


class NormQtyPact:
    """FEMA P-58 标准数量估算工具（Excel）的 Python 接口。

    参数
    ----
    NumOfStories : int
        楼层数。
    FloorAreaList : list of float
        各层楼面面积（平方英尺）。
    Occupancy1Type : list of str
        各层主要使用类型。
    Occupancy2Type : list of str
        各层次要使用类型。
    Occupancy3Type : list of str
        各层第三使用类型。
    Occupancy1Area : list of float
        各层主要使用类型面积占比。
    Occupancy2Area : list of float
        各层次要使用类型面积占比。
    Occupancy3Area : list of float
        各层第三使用类型面积占比。
    """

    # 定位与本文件同级的 resources 目录中打包的 Excel 宏文件
    _RESOURCE_DIR = os.path.join(os.path.dirname(__file__), "resources")
    pathXls = os.path.abspath(
        os.path.join(_RESOURCE_DIR,
                     "FEMA P-58_NormativeQuantityEstimationTool_031818.xlsm")
    )

    NumOfStories = 3
    FloorAreaList = [1, 1, 1]
    Occupancy1Type = ["APARTMENT", "APARTMENT", "APARTMENT"]
    Occupancy2Type = ["none", "none", "none"]
    Occupancy3Type = ["none", "none", "none"]
    Occupancy1Area = [1, 1, 1]
    Occupancy2Area = [0, 0, 0]
    Occupancy3Area = [0, 0, 0]

    def __init__(
        self,
        NumOfStories=3,
        FloorAreaList=None,
        Occupancy1Type=None,
        Occupancy2Type=None,
        Occupancy3Type=None,
        Occupancy1Area=None,
        Occupancy2Area=None,
        Occupancy3Area=None,
    ):
        self.NumOfStories = NumOfStories
        self.FloorAreaList = FloorAreaList if FloorAreaList is not None else [1, 1, 1]
        self.Occupancy1Type = Occupancy1Type if Occupancy1Type is not None else ["APARTMENT"] * NumOfStories
        self.Occupancy2Type = Occupancy2Type if Occupancy2Type is not None else ["none"] * NumOfStories
        self.Occupancy3Type = Occupancy3Type if Occupancy3Type is not None else ["none"] * NumOfStories
        self.Occupancy1Area = Occupancy1Area if Occupancy1Area is not None else [1] * NumOfStories
        self.Occupancy2Area = Occupancy2Area if Occupancy2Area is not None else [0] * NumOfStories
        self.Occupancy3Area = Occupancy3Area if Occupancy3Area is not None else [0] * NumOfStories

    # ------------------------------------------------------------------
    # 内部方法
    # ------------------------------------------------------------------

    def ExecutePactXlsmNormTool(self):
        """运行 Excel 宏，填充标准数量估算工作表。"""
        try:
            docApp = GetActiveObject("Excel.Application")
            excel_was_running = True
        except Exception:
            docApp = Dispatch("Excel.Application")
            excel_was_running = False
        try:
            doc = docApp.Workbooks.Open(self.pathXls)
            ws = doc.Worksheets("Normative Quantity Estimate")

            ws.Range(ws.Cells(10, 2), ws.Cells(100, 10)).ClearContents()
            for i in range(self.NumOfStories):
                ws.Range("B" + str(10 + self.NumOfStories - i)).value = i + 1
                ws.Range("C" + str(10 + self.NumOfStories - i)).value = i + 1
                ws.Range("D" + str(10 + self.NumOfStories - i)).value = self.FloorAreaList[i]
                ws.Range("E" + str(10 + self.NumOfStories - i)).value = self.Occupancy1Type[i]
                ws.Range("F" + str(10 + self.NumOfStories - i)).value = self.Occupancy1Area[i]
                ws.Range("G" + str(10 + self.NumOfStories - i)).value = self.Occupancy2Type[i]
                ws.Range("H" + str(10 + self.NumOfStories - i)).value = self.Occupancy2Area[i]
                ws.Range("I" + str(10 + self.NumOfStories - i)).value = self.Occupancy3Type[i]
                ws.Range("J" + str(10 + self.NumOfStories - i)).value = self.Occupancy3Area[i]

            ws.Range("B10").value = "Roof"
            ws.Range("C10").value = self.NumOfStories + 1
            ws.Range("D10").value = 0
            ws.Range("E10").value = "none"
            ws.Range("F10").value = 1
            ws.Range("G10").value = "none"
            ws.Range("H10").value = 0
            ws.Range("I10").value = "none"
            ws.Range("J10").value = 0

            doc.Application.Run("Sheet7.compile_fragility")
            doc.Save()
            doc.Close(SaveChanges=False)
            print("Excel macro runs successfully...")
        except Exception as e:
            print(e)
        finally:
            if not excel_was_running:
                docApp.DisplayAlerts = False
                docApp.Quit()

    # ------------------------------------------------------------------
    # 公共接口
    # ------------------------------------------------------------------

    def Output_PactComponentDirectory(self, output_path="PactComponentDirectory.csv"):
        """生成 PACT 格式的构件目录 CSV 文件。

        参数
        ----
        output_path : str
            输出 CSV 文件路径。
        """
        self.ExecutePactXlsmNormTool()

        headers = [
            "No.", "Component Type", "Performance Group Quantity",
            "Quantity Dispersion", "Fragility Correlated", "Population Model",
            "Demand Parameters", "Dir", "Floor Num",
        ]

        try:
            docApp = GetActiveObject("Excel.Application")
            excel_was_running = True
        except Exception:
            docApp = Dispatch("Excel.Application")
            excel_was_running = False
        rows = []
        try:
            doc = docApp.Workbooks.Open(self.pathXls, ReadOnly=True)
            ws = doc.Worksheets("Normative Quantity Estimate")
            i_row = 11
            while True:
                i_floor = ws.Cells(i_row, 15).value
                if i_floor is None:
                    postag = ws.Cells(i_row, 13).value
                    if postag != "END OF BUILDING SUM INPUT":
                        i_row += 1
                        continue
                    else:
                        break
                PactNo = ws.Cells(i_row, 17).value
                ComponentType = ""
                PerformanceGroupQuantity = ws.Cells(i_row, 20).value
                Direction = [1, 2]
                if isinstance(PerformanceGroupQuantity, str) and PerformanceGroupQuantity.strip("-") == "":  # 无方向性构件
                    PerformanceGroupQuantity = ws.Cells(i_row, 21).value
                    Direction = [3]
                QuantityDispersion = ws.Cells(i_row, 26).value
                FragilityCorrelated = "FALSE"
                PopulationModel = ""
                DemandParameters = ""
                for i_Dir in Direction:
                    rows.append([
                        PactNo, ComponentType, PerformanceGroupQuantity,
                        QuantityDispersion, FragilityCorrelated,
                        PopulationModel, DemandParameters, i_Dir, i_floor,
                    ])
                i_row += 1
            doc.Close(SaveChanges=False)
            print("Output successfully...")
        except Exception as e:
            print(e)
        finally:
            if not excel_was_running:
                docApp.DisplayAlerts = False
                docApp.Quit()

        with open(output_path, "w", encoding="utf8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)

    def Output_PelicunComponentDirectory(
        self,
        json_path="PelucunComponentDirectory.json",
        csv_path="PelucunComponentDirectory.csv",
    ):
        """生成 Pelicun 格式的构件目录 JSON 和 CSV 文件。

        参数
        ----
        json_path : str
            输出 JSON 文件路径。
        csv_path : str
            输出 CSV 文件路径。
        """
        self.ExecutePactXlsmNormTool()

        try:
            docApp = GetActiveObject("Excel.Application")
            excel_was_running = True
        except Exception:
            docApp = Dispatch("Excel.Application")
            excel_was_running = False
        jsondata = {}
        try:
            doc = docApp.Workbooks.Open(self.pathXls, ReadOnly=True)
            ws = doc.Worksheets("Normative Quantity Estimate")
            i_row = 11
            while True:
                i_floor = ws.Cells(i_row, 15).value
                if i_floor is None:
                    postag = ws.Cells(i_row, 13).value
                    if postag != "END OF BUILDING SUM INPUT":
                        i_row += 1
                        continue
                    else:
                        break
                elif isinstance(i_floor, float):
                    i_floor = int(i_floor)
                elif i_floor == "ALL":
                    i_floor = "1"

                PactNo = ws.Cells(i_row, 17).value
                PerformanceGroupQuantity = ws.Cells(i_row, 20).value
                Direction = "1, 2"
                if isinstance(PerformanceGroupQuantity, str) and PerformanceGroupQuantity.strip("-") == "":  # 无方向性构件
                    PerformanceGroupQuantity = ws.Cells(i_row, 21).value
                    Direction = "3"
                QuantityDispersion = ws.Cells(i_row, 26).value
                if QuantityDispersion == "p90 low":
                    QuantityDispersion = 0.0001
                elif isinstance(QuantityDispersion, str):
                    pass
                else:
                    if QuantityDispersion < 0.0001:
                        QuantityDispersion = 0.0001

                UnitType = ws.Cells(i_row, 24).value
                UnitType = UnitType.lower()
                # 统一单位名称为 Pelicun 格式
                if UnitType == "lf":
                    UnitType = "ft"
                elif UnitType == "sf":
                    UnitType = "ft2"

                PactUnit = ws.Cells(i_row, 19).value
                PactUnit_float = float("".join(filter(str.isdigit, PactUnit)))
                # 将 PACT 中的组数换算为物理数量
                if UnitType in ["ft", "ea", "ft2"]:
                    MedianQuantity = PerformanceGroupQuantity * PactUnit_float
                else:
                    MedianQuantity = PerformanceGroupQuantity
                    UnitType = "ea"

                ComponentContents = {
                    "location": str(i_floor),
                    "direction": Direction,
                    "Theta_0": str(MedianQuantity),
                    "distribution": "lognormal",
                    "cov": str(QuantityDispersion),
                    "unit": UnitType.lower(),
                    "Blocks": str(MedianQuantity / PactUnit_float),
                }
                if jsondata.get(PactNo) is None:
                    jsondata[PactNo] = [ComponentContents]
                else:
                    jsondata[PactNo].append(ComponentContents)
                i_row += 1
            doc.Close(SaveChanges=False)
            print("Output successfully...")
        except Exception as e:
            print(e)
        finally:
            if not excel_was_running:
                docApp.DisplayAlerts = False
                docApp.Quit()

        with open(json_path, "w") as fp:
            json.dump(jsondata, fp, indent=4)

        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([
                "ID", "Units", "Location", "Direction",
                "Theta_0", "Blocks", "Family", "Theta_1", "Comment",
            ])
            for key, items in jsondata.items():
                keyName = ".".join([key[0], key[1:3], key[3:]])
                for item in items:
                    writer.writerow([
                        keyName,
                        item["unit"],
                        item["location"],
                        item["direction"] if item["direction"] != "3" else "0",
                        item["Theta_0"],
                        item["Blocks"],
                    ])
