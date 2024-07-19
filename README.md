# ExcelReader

Excel导出cs文件工具，会依照数据分别生成类和数据。
表的存储类型分为了List和Dict。
List适宜全部加载，需求遍历全部数据
Dict按需动态加载，适宜存储大量数据
原理参照了[ConfigExcel](https://github.com/yukuyoulei/ConfigExcel.git)。

# Excel导出细则

**NameSpace** 默认Config

**Excel名** 默认类名

**Sheet名** 默认字段名
    1.#开头为注释sheet，导出时自动忽略
    2.@开头为常量sheet，导出依照常量表导出规则

**Sheet批注** 表内第一行第一列的批注为特性属性列表
    1.@常量表不需要特性批注
    2.特性批注现在有2个：Name和Type。批注示例:
        Name:xxx
        Type:List

**Sheet表**
    1.#开头的行或者列都会自动忽略
    2.常量表：特殊表，每行只1个字段，默认前三列分别为：字段名、字段类型、备注
    3.普通表：需要特性批注的表，默认前三行分别为：字段名、字段类型、备注

