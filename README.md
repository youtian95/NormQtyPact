# NormQtyPact

通过调用FEMA P-58附带的非结构构件标准数量Excel文件，根据建筑各层的平面面积输出非结构构件数量文件。生成的文件可以复制进PACT或者Pelicun中。

## 输出文件
- PactComponentDirectory.csv：可以直接复制进[PACT](https://femap58.atcouncil.org/pact)中提供非结构构件数量信息。
- PelucunComponentDirectory.csv：可以作为[Pelicun](https://github.com/NHERI-SimCenter/pelicun)的非结构构件数量文件。

## 依赖

需要python安装以下包：
 - pypiwin32
 - json

## 使用

直接在命令行中运行NormQtyPact.py文件，例子：

```
python NormQtyPact.py --NumOfStories 3 --FloorAreaList '1,1,1' --Occupancy1Type 'APARTMENT,APARTMENT,APARTMENT' --Occupancy2Type 'none,none,none' --Occupancy3Type 'none,none,none' --Occupancy1Area '1,1,1' --Occupancy2Area '0,0,0' --Occupancy3Area '0,0,0'
```

参数:
```
NumOfStories (int): Number of stories. e.g. 3
FloorAreaList (str): Floor area of each story. e.g. '1,1,1'
Occupancy1Type (str): 1st Occupancy type of each story. e.g. 'APARTMENT,APARTMENT,APARTMENT'. 
Occupancy2Type (str): 2nd Occupancy type of each story. e.g. 'none,none,none'
Occupancy3Type (str): 3rd Occupancy type of each story. e.g. 'none,none,none'
Occupancy1Area (str): area percentage of 1st Occupancy of each story. e.g. '1,1,1'
Occupancy2Area (str): area percentage of 2nd Occupancy of each story. e.g. '0,0,0'
Occupancy3Area (str): area percentage of 3rd Occupancy of each story. e.g. '0,0,0'
```
