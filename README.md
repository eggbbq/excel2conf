# 1)表格

| int | b1 | b2 | b3 |
| --- | -- | -- | -- |
| b1  | 0  | 1  | 2  |
| b2  | 1  | 2  | 3  |
| b3  | 1  | 2  | 3  |

矩阵数据类型定义在左上角
```
{
    "head":[b1,b2,b3,b4]
    "array":[ 0,1,2
              1,2,3
              1,2,3
            ]
}
```

# 2)常量表
| name          | type | value |
| ------------- | ---- | ----- |
| max_acc_level | int  | 10    |
| dungeon_cd    | int  | 1000  |

# 3)list/dict
| id  | lvl | exp | name   | parse_list | query_one           | query_many            |
| --- | --- | --- | ------ | ---------- | ------------------- | --------------------- |
| int | int | int | string | int[],     | sheet_name|id,name  | sheet_name[]|name,age |
| 1   | 100 | 100 | lucy   | 1,2,3      | 100,aaa             | aaa, 100              |
| 2   | 101 | 200 | mark   | 4,5,6      | 200,bbb             | bbb, 100              |

# 4)导出类型
sheet_name
- list
- dict
- object
- table

表格sheet名称就是导出数据类型。默认list

# 5)字段类型定义
基础数据类型
int
float
doubble
bool
string

基础数据列表解析类型
T[]S  (T:基础数据类型,S:分割字符)

查询表达式类型
查询一个 Table|field1,... (Table:数据表格名,filed:字段名)
查询所有 Table[]|field1,...

# 6)excel表格命名连接词
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

# 7)导出格式
json

# 8)字段匹配符
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

# 9)字段导出前提=>命名合法且有类型标记
非法命名字段忽略导出
没有类型标记的字段忽略导出

字段命名合法字符集 {小写字母, 大小字母, 数字, _}

# 10)数据导出中间件
默认导出为json
但是为支持json到其他数据的转换如protobuf/c#/java 等数据的转换，导出json的同时还会生成一份所有数据表格的类型定义

# 11)排除文件 --excludes
--excludes=UserItem.xls,UserSkill.xls