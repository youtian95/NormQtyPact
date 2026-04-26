# NormQtyPact

通过调用 FEMA P-58 附带的非结构构件标准数量 Excel 文件，根据建筑各层的平面面积输出非结构构件数量文件。生成的文件可以直接导入 PACT 或 Pelicun 中使用。

> **注意：** 本包依赖 Windows COM 接口调用 Excel，仅支持 Windows 平台且需要安装 Microsoft Excel。

## 输出文件

- `PactComponentDirectory.csv`：可直接导入 [PACT](https://femap58.atcouncil.org/pact) 的非结构构件数量文件。
- `PelucunComponentDirectory.csv` / `PelucunComponentDirectory.json`：可作为 [Pelicun](https://github.com/NHERI-SimCenter/pelicun) 的非结构构件数量输入文件。

## 安装

### 从 PyPI 安装（发布后）

```bash
pip install normqtypact
```

## 依赖

- Python >= 3.7
- Windows 操作系统 + Microsoft Excel
- [pypiwin32](https://pypi.org/project/pypiwin32/)
- 构建工具：[hatch](https://hatch.pypa.io/)（仅开发/打包时需要）

## 使用

### 命令行

安装后可直接使用 `normqtypact` 命令：

```bash
normqtypact --NumOfStories 3 --FloorAreaList "1,1,1" --Occupancy1Type "APARTMENT,APARTMENT,APARTMENT" --Occupancy2Type "none,none,none" --Occupancy3Type "none,none,none" --Occupancy1Area "1,1,1" --Occupancy2Area "0,0,0" --Occupancy3Area "0,0,0"
```

参数说明：

| 参数 | 类型 | 说明 | 示例 |
|------|------|------|------|
| `--NumOfStories` | int | 楼层数 | `3` |
| `--FloorAreaList` | str | 各层楼面面积（平方英尺），逗号分隔 | `"1000,1000,1000"` |
| `--Occupancy1Type` | str | 各层主要使用类型，逗号分隔 | `"APARTMENT,APARTMENT,APARTMENT"` |
| `--Occupancy2Type` | str | 各层次要使用类型，逗号分隔 | `"none,none,none"` |
| `--Occupancy3Type` | str | 各层第三使用类型，逗号分隔 | `"none,none,none"` |
| `--Occupancy1Area` | str | 主要使用类型面积占比，逗号分隔 | `"1,1,1"` |
| `--Occupancy2Area` | str | 次要使用类型面积占比，逗号分隔 | `"0,0,0"` |
| `--Occupancy3Area` | str | 第三使用类型面积占比，逗号分隔 | `"0,0,0"` |

### Python API

```python
from normqtypact import NormQtyPact

obj = NormQtyPact(
    NumOfStories=3,
    FloorAreaList=[1000, 1000, 1000],
    Occupancy1Type=["APARTMENT", "APARTMENT", "APARTMENT"],
    Occupancy2Type=["none", "none", "none"],
    Occupancy3Type=["none", "none", "none"],
    Occupancy1Area=[1, 1, 1],
    Occupancy2Area=[0, 0, 0],
    Occupancy3Area=[0, 0, 0],
)

# 生成 PACT 格式输出
obj.Output_PactComponentDirectory("PactComponentDirectory.csv")

# 生成 Pelicun 格式输出
obj.Output_PelicunComponentDirectory(
    json_path="PelucunComponentDirectory.json",
    csv_path="PelucunComponentDirectory.csv",
)
```

## 示例

`examples/` 目录包含可直接运行的示例脚本：

| 文件 | 说明 |
|------|------|
| `basic_usage.py` | 3 层公寓建筑（单一使用类型） |
| `mixed_occupancy.py` | 5 层混合使用建筑（零售 + 办公） |
| `PactComponentDirectory.csv` | PACT 格式示例输出 |
| `PelucunComponentDirectory.csv` | Pelicun CSV 格式示例输出 |
| `PelucunComponentDirectory.json` | Pelicun JSON 格式示例输出 |

运行示例：

```bash
cd examples
python basic_usage.py
```
