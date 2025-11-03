# Excel Compare

## Project Overview

针对 Excel 文件内容对比的工具：

- 可以在配置文件中指定多列作为联合唯一业务标识，以此关联两表的行
- 可以在配置文件中指定多列进行数据对比
- 单独对比每一列的值，并标识出差异（黄色背景及批注）
- 对比结果中会新增两列：联合唯一业务标识，整行的对比结果
- 对比结果同时包含统计信息：总行数，一致数、不一致数、缺失数及主键重复数

- 可以指定运行时使用的配置文件

## Requirements
- Python >= 3.7

## Installation
``` 
pip install pandas openpyxl
```

## Config

- `main_compare_file_path`

  主文件路径（生成的对比结果文件会以此文件为基础）

- `sub_compare_file_path`

  副文件路径

- `unique_columns`

  作为联合唯一业务标识的列

- `compare_columns`

  数据对比的列

- `output_path`

  **默认值**：""

  对比结果文件的生成路径（无需指定文件名）(未配置时会在当前命令执行目录生成)

- `skip_lines`

  **默认值**：0

  对比时跳过的行数（第一行默认为表头默认跳过），例如需要跳过第二行时，配置此参数为 1

- `skip_both`

  **默认值**：True

  在跳过时，副文件是否同时跳过

```json
{
  "main_compare_file_path": "your path/main.xlsx",
  "sub_compare_file_path": "your path/sub.xlsx",
  "output_path": "your path/",
  "skip_lines": 3,
  "skip_both": true,
  "unique_columns": [
    "部门",
    "系统名称",
    "系统名称2"
  ],
  "compare_columns": [
    "阶段1",
    "阶段2",
    "阶段3"
  ]
}
```

## Usage

Run the project with the following command:
>[!NOTE]
> `配置文件路径` 非必须，未指定时会自动读取当前命令执行目录下的 `config.json` 文件

```bash
py .\main.py <配置文件路径>
```
