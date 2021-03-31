# 
在过去的工作中，我接触了各种各样的游戏数据配置表格，有的复杂，有的简单，或多或少遇到一些不满意。这里就在简单与复杂中，尝试一个折中方案。用任何语言都可以实现，这里采用python仅仅是为了简单实现原型。

# 表格解析与转换
        Excel
        /   \
      Meta  JSON
        |    / \
        |   /   简单文本转换(lua/js...)
    复杂序列化(protobuf/navtive serialization object)

Meta 表格类型描述
- 集合类型
- 字段
- 字段类型

JSON
- 结构简单, 生成容易, 易读写
- 各语言支持良好

作为一个通用的表格解析器，推荐 把Meta 和JSON 作为解析的中间件导出(或者说首要导出)，其他诸如lua/js/protobuf/.... 当作子程序或者其他程序在基于 Meta/JSON 二次导出

因为在不知道具体编程环境，很难导出完美契合项目的代码, 所以这最后的一步，交给最终编程环境去实现会非常简单且完美契合，其实直接使用JSON也可以了。

# 的表格形式
- 列表(dict/list)
- 常数枚举
- 矩阵表格

## 列表(list/dict)

| id  | lvl | exp | name   | parse_list | query_one           | query_many               |
| --- | --- | --- | ------ | ---------- | ------------------- | ------------------------ |
| id  | lvl | exp | name   | parse_list | query_one           | query_many               |
| int | int | int | string | int[],     | sheet_name\|id,name | sheet_name\[\]\|name,age |
| 1   | 100 | 100 | lucy   | 1,2,3      | 100,aaa             | aaa, 100                 |
| 2   | 101 | 200 | mark   | 4,5,6      | 200,bbb             | bbb, 100                 |


row(0+start)    定义字段

row(1+start)    定义类型

row(2+start~n)  定义数值

这种类型的配置表格占绝大多数

dict/list 区分仅仅是在存储/访问对象上提供一点点方便，解析定义表格方面二者结构上没有差异

start 起始索引位置。表头定义一些人类可读的显示名称也是很常见的操作，如下

| ID  | 等级 | 经验 | 名称   | 数组解析式  | 一对一查询            | 一对多查询                |
| --- | --- | --- | ------ | ---------- | ------------------- | ------------------------ |
| id  | lvl | exp | name   | parse_list | query_one           | query_many               |
| int | int | int | string | int[],     | sheet_name\|id,name | sheet_name\[\]\|name,age |
| 1   | 100 | 100 | lucy   | 1,2,3      | 100,aaa             | aaa, 100                 |
| 2   | 101 | 200 | mark   | 4,5,6      | 200,bbb             | bbb, 100                 |


## 稀疏矩阵
- type 存储类型
- row_head 列表头
- col_head 行表头
- matrix CSR方式存取:非空行数量,{非空行号,非空行开始下标,...}, {行标,值,行标,值,...}

**非稀疏矩阵采用数组全部元素存储**


例:buff之间关系矩阵

| int | C1  | C2  | C3  | C4  |
| --- | --- | --- | --- | --- |
| R1  | 0   | a12 | a13 | 0   |
| R2  | 0   | 0   | 0   | 0   |
| R3  | 0   | a32 | 0   | 0   |
| R4  | 0   | 0   | 0   | 0   |

```JSON
Matrix CSR Storage in JSON
        not empty row count
        |  row index
        |  |     row index
        |  |     |  ___________________
        |  |  ___|_|___                |
        |  | |   | |   |               |
        |  | |   | |  col val col val col val
        |  | |   | |   |  |    |  |    |  |
        |  | |   | |   v  |    |  |    v  |
array   2, 0,5 , 3,9 , 1,a12 , 2,a13 , 1,a32
index   0  1 2   3 4   5 6     7 8     9 10
              
json string
{
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
```

代码实现查找时，可用二分法减少访问行标的时间复杂度


## 常量表
| name          | type | value |
| ------------- | ---- | ----- |
| max_acc_level | int  | 10    |
| dungeon_cd    | int  | 1000  |

常量表主要是枚举一些特殊的值，相互可能无关，数量也可能较多，竖向更利于维护



# 导出类型
sheet_name
- list
- dict
- object
- matrix
- matrix(csr)

表格sheet名称就是导出数据类型。默认list

# 字段类型定义
## 数据类型
- int
- float
- doubble
- bool
- string

## 数据列表解析类型
**不再推荐使用，请使用字段合并，更优雅**

T[]S  (T:基础数据类型,S:分割字符)

## 数组字段合并
字段以 下划线"_"和数字结尾


| name | ids_0 | ids_1 | ids_2 | ids_3 |
| ---- | ----- | ----- | ----- | ----- |
| abc  | 0     | 1     | 2     | 3     |

```JSON
{
    "name":"abc",
    "ids":[0,1,2,3]
}
```



## 查询表达式类型
- 查询一个 Table|field1,... (Table:数据表格名,filed:字段名)
- 查询所有 Table[]|field1,...

# excel表格命名连接词
连接词在导出文件时会被忽略
+ 分表连接
User.xls
User+Attr.xls
User+Body.xls
上面三个表最终合并成User一个表格导出。其作用时拆分巨大无比的表格，让编辑维护更简单

- 禁止导出
-UserAni.xls
这样表示这个表格不会被单独导出，但是本身是可以被查询到的

-- 忽略表格
--UserAni.xls
这样表示跳过解析这个表格，当其不存在

# 导出格式
json

# 字段匹配符
|  id | lvl |
| --- | --- |
| int | int |
|     | *s  |
| 1   | 1   |
| 2   | 1   |

第三行第二列  *s 这是一个匹配符
用于匹配传入的匹配字符，确定是否过滤或者匹配该字段
匹配行是可选的

命令行 --match
--match=+s 空和带s匹配的字段(s可以是任意字符)<br>
--match=-s 不带s匹配的字段(s可以是任意字符)<br>
--match=s  所有s匹配的字段(s可以是任意字符)<br>
--match=-* 所有无匹配的字段<br>
--match=*  所有有匹配的字段<br>

# 字段导出前提=>命名合法且有类型标记
非法命名字段忽略导出
没有类型标记的字段忽略导出

字段命名合法字符集 {小写字母, 大小字母, 数字, _}

# )数据导出中间件
默认导出为json
但是为支持json到其他数据的转换如protobuf/c#/java 等数据的转换，导出json的同时还会生成一份所有数据表格的类型定义

# )排除文件 --excludes
--excludes=UserItem.xls,UserSkill.xls