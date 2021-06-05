 
# 配置表格定义
## Excel表格文件结构
- Excel2007之前或者之后的版本均支持（.xls .xlsx）
- Excel文件名=导出数据的类名（或者说表名）
- Excel文件中的第一张格为数据定义区域（即sheet1），
- Excel文件中sheet1通过修改表名指定**数据结构**，如没有指定结构，采用默认值list
- Excel文件名命名字符只能包含 英文字母 和 "+"(加号) "-"(减号) "_"(下划线)
  - \+ 表格合并标记
  - \- 可以被外链，不可被导出标记

## 数据结构 定义

| list        | 列表     | |
| ----------- | -------- | --------------------- |
| dict        | 字典     | 第一个字段用于构造字典key |
| object      | 键值     | |
| matrix      | 矩阵     | |
| matrix(csr) | 稀疏矩阵  | |


## 表头 定义
- 第一行(列)**注释**，可以为空
- 第二行(列)**字段名称**。只能包含英文字母、数字和下划线，且不能为空
- 第三行(列)**字段类型**。不能为空
  - int
  - float
  - double
  - bool
  - string
  - 外链查询定义（Item[]|id,略微复杂详情见下）
- 第四行(列) 如果出现匹配匹配标记则认为是匹配行，否则视为数据行(列)

- 字段不导出的情形
  - 字段字符非法，该列不导出
  - 类型定义非法，该列不导出



# 示例
## list
Item.xls 

| ID  | 值    |
| --- | ----- |
| id  | value |
| int | int   |
| 1   | 100   |
| 2   | 200   |

```JSON
{
 "Item":
  [
   {"id":1, "value":100},
   {"id":2, "value":200}
  ]
}
```
---

## dict 
Item.xls

| ID  | 值    |
| --- | ----- |
| id  | value |
| int | int   |
| 1   | 100   |
| 2   | 200   |

```JSON
{
 "Item":
  {
   "1":{"id":1, "value":100},
   "2":{"id":2, "value":200}
  }
}
```
---

## object
GameConsts.xls

| 注释 | cd  | int | 86400 |
| --- | --- | --- | ----- |
| 注释 | hp | int  | 100   |
| 注释 | mp | int  | 200   |

```JSON
{
 "GameConsts":
 {
  "cd":86400,
  "hp":100,
  "mp":200
 }
}
```

---

## matrix   matrix(csr)
Buff.xls

|     |     | 注释 | 注释 | 注释 | 注释 |
| --- | --- | --- | ---- | --- | --- |
|     | int | C1  | C2   | C3  | C4  |
| 注释 | R1  | 0   | a12  | a13 | 0   |
| 注释 | R2  | 0   | 0    | 0   | 0   |
| 注释 | R3  | 0   | a32  | 0   | 0   |
| 注释 | R4  | 0   | 0    | 0   | 0   |

 
```JSON
----------普通导出----------
{
 "Buff":{
  "type":"int",
  "col_head":[C1, C2, C3, C4],
  "row_head":[R1, C2, R3, R4],
  "matrix":[
    0, a12, a13, 0,
    0, 0,   0,   0,
    0, a32, 0,   0,
    0, 0,   0,   0
  ]
 }
}

----------CSR导出----------
Matrix CSR Storage in JSON
        not empty row count
        |  row indexe
        |  |     row index
        |  |     |  ___________________
        |  |  ___|_|___                |
        |  | |   | |   |               |
        |  | |   | |  col val col val col val
        |  | |   | |   |  |    |  |    |  |
        |  | |   | |   v  |    |  |    v  |
array   2, 0,5 , 3,9 , 1,a12 , 2,a13 , 1,a32
index   0  1 2   3 4   5 6     7 8     9 10
              
{
"Buff":{
  "type":"matrix(csr)",
  "row_head":[C1,C2,C3,C4],
  "col_head":[R1,R2,R3,R4],
  "matrix":[
          2,        //非空行数量
          0,5, 3,9, //非空行号,非空行开始下标
          1,1, 2,2, //行标，值
          1,1       //行标，值
          ]
 }
}

代码实现查找时，可用二分法减少访问行标的时间复杂度
复杂度
```

---

## 纯数表, 无字段 (int[])
假如想导出这样的结构应该如何定义表格
```JSON
{ "Numbers":[1,2,3,4]}
```

| 注释  |
| ----- |
| _     |
| 100   |
| 500   |
| 1000  |
| 20000 |

---

## 字段中嵌套简单数组
NestArray.xls

| 注释    | 注释  | 注释   | 注释  | 注释   |
| ------ | ----- | ----- | ----- | ----- |
| name   | ids_0 | ids_1 | ids_2 | ids_3 |
| string | int   | int   | int   | int   |
| abc    | 0     | 1     | 2     | 3     |

<!-- ## 嵌套定义复杂列表 这种形式定义和解析太复杂，不推荐
Bundle.xls
| ID   | 解包类型  | 选中项数 | 嵌套标记  | 权重    | 概率            | 物品id   | 物品数量   |
| ---- | -------- | ------- | ------- | ------- | -------------- | -------- | --------- |
| id   | openType | count   | .items  | .weight | .probability   | .item_id | .item_num |
| int  | int      | int     | []      | int     | float          | int      | int       |
| 101  | 1        | 2       |         | 100     |                | 1001     | 1000      |
| 102  | 2        | 1       |         | 400     |                | 1004     | 1         |
| 102  | 2        | 1       |         |         | 0.5            | 2001     | 1         | -->

## 外键查询, 嵌套复杂对象
Bundle.xls (主表 dict)

| ID   | 解包类型  | 选中项数 | 物品列表          |
| ---- | -------- | ------- | ---------------- |
| id   | openType | count   | items            |
| int  | int      | int     | BundleItem[]\|id |
| 101  | 1        | 2       | 101              |
| 102  | 2        | 1       | 102              |

BundleItem.xls (附表 list)

| ID   | 权重   | 概率           | 物品id   | 物品数量  |
| ---- | ------ | ------------- | ------- | -------- |
| id   | weight | probability   | item_id | item_num |
| int  | int    | float         | int     | int      |
| 101  | 100    |               | 1001    | 1000     |
| 101  | 200    |               | 1002    | 100      |
| 101  | 300    |               | 1003    | 10       |
| 101  | 400    |               | 1004    | 1        |
| 102  |        | 0.5           | 2001    | 1        |
| 102  |        | 0.2           | 2002    | 1        |
| 102  |        | 0.1           | 2003    | 1        |
| 102  |        | 0.05          | 2004    | 1        |

```JSON
{
  "Bundle":{
    "101":{
      "id":101, 
      "openType":1, 
      "count":2,
      "items":[
        {"id":101, "weight":100, "probability":0, "itemId":1001, "itemNum":1000},
        {"id":101, "weight":100, "probability":0, "itemId":1002, "itemNum":100},
        {"id":101, "weight":100, "probability":0, "itemId":1003, "itemNum":10}
      ]
    },
    "102":{
      "id":102, 
      "openType":2, 
      "count":1,
      "items":[
        {"id":102, "weight":0, "probability":0.5, "itemId":2001, "itemNum":1},
        {"id":102, "weight":0, "probability":0.2, "itemId":2002, "itemNum":1},
        {"id":102, "weight":0, "probability":0.1, "itemId":2003, "itemNum":1},
        {"id":102, "weight":0, "probability":0.05, "itemId":2004, "itemNum":1}
      ]
    }
  }
}
```


```JSON
{
"NestArray":{
    "name":"abc",
    "ids":[0,1,2,3]
 }
}
```

# 导出文件
- 数据表
- 类型信息表

# 脚本参数
### --excel 表格目录路径, 默认当前工作目录

### --match 匹配字符 
- \+s 空和带s匹配的字段(s为任意字符串)
- \-s 不带s匹配的字段(s为任意字符串)
- \s  所有s匹配的字段(s为任意字符串)
- \-* 所有无匹配的字段
- \*  所有有匹配的字段

### --file JSON文件名，默认值Config.json
### --start 表头偏移量 (默认值=1 第一行/列保留给注释用)
### --excludes 排除表格
### --workspace 工作目录