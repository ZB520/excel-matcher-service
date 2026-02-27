## Excel 旧表/新表 图书匹配工具

### 功能概述

- 从“旧表”和“新表”两个 Excel 中，按书名和数量总和进行匹配。
- 对成功匹配的书目：
  - 生成 `matched_details.xlsx`：保留旧表的所有行和结构，仅将 `书号`、`书名`、`出版社`、`单价` 四列替换为新表信息。
- 对未匹配（或数量不一致）的书目：
  - 生成 `unmatched_items.xlsx`：从新表中列出对应行，便于后续人工核查。

### 依赖

使用前请先安装依赖（建议在虚拟环境中执行）：

```bash
pip install -r excel_book_matching_requirements.txt
```

或单独安装：

```bash
pip install pandas openpyxl
```

### 使用步骤

1. 准备两个 Excel 文件：
   - 旧表（列名至少包含）：`学校`, `序号`, `班级人数`, `书号`, `书名`, `出版社`, `单价`, `学生数量`
   - 新表（列名至少包含）：`书号`, `书名`, `出版社`, `单价`, `数量`

2. 打开 `excel_book_matcher.py`，在文件顶部修改路径常量：
   - 将 `OLD_FILE_PATH` 设置为你的旧表路径，例如：
     - `OLD_FILE_PATH = r"C:\\Users\\ZB\\Desktop\\old.xlsx"`
   - 将 `NEW_FILE_PATH` 设置为你的新表路径，例如：
     - `NEW_FILE_PATH = r"C:\\Users\\ZB\\Desktop\\new.xlsx"`

3. 运行脚本（在命令行/终端中）：

```bash
python excel_book_matcher.py
```

4. 脚本成功执行后，会在当前工作目录生成三个文件：
   - `matched_details.xlsx`：匹配成功的明细表（结构同旧表，但 4 个字段已更新为新表信息）。
   - `unmatched_items.xlsx`：新表中未能成功匹配的书目记录。
   - `match_log.txt`：匹配统计与日志信息。

### 匹配规则说明（摘要）

- 先对书名做“强标准化”：
  - 转小写、去前后空格。
  - 合并多余空白。
  - 去掉所有空格和常见标点（全角/半角）。
- 分别在旧表、新表中按标准化后的书名分组求和：
  - 旧表：`学生数量` 总和。
  - 新表：`数量` 总和。
- 仅当 **标准化书名相同 且 数量总和相等** 时，视为匹配成功：
  - 旧表所有对应行进入 `matched_details.xlsx`，并用新表信息更新指定 4 列。
  - 若书名存在于新旧两表，但数量不一致，则视为“冲突”，对应新表记录会进入 `unmatched_items.xlsx`。
  - 仅出现在新表、不在旧表的书名，同样进入 `unmatched_items.xlsx`。

