# DataProcBase



# 智汇库存分析助手使用说明

## 软件简介

智汇库存分析助手是一款专业的库存数据分析工具，能够帮助用户快速处理和分析库存相关数据，支持自定义计算公式，让数据分析更加灵活高效。

## 主要功能

1. Excel文件数据导入
2. 自定义列选择
3. 强大的计算列功能
4. 数据预览功能
5. 结果导出功能

## 使用指南

### 1. 数据导入

1. 启动软件后，点击界面上的"选择文件"按钮
2. 在弹出的文件选择对话框中选择需要分析的Excel文件
3. 软件会自动读取文件内容并显示

### 2. 列选择和数据预览

1. 在列选择界面，您可以：
   - 使用"全选"或"全不选"按钮快速选择
   - 手动勾选需要的列
   - 实时预览选中的数据（右侧预览窗口显示前5行数据）

### 3. 添加计算列

1. 勾选"添加计算列"选项
2. 输入新列名称
3. 选择预设公式或输入自定义公式
   - 预设公式包括：
     * 库存天数(不含在途) = 运营云仓可用数/30天发货量 * 30
     * 库存天数(含在途) = (运营云仓可用数 + 采购在途)/30天发货量 * 30
   - 自定义公式说明：
     * 使用英文方括号[列名]引用列
     * 支持基本运算符：+, -, *, /
     * 支持使用()设置运算优先级
     * 示例：[销售额] - [成本]
4. 点击"添加计算列"按钮生成新列

### 4. 数据导出

1. 确认选择的列和计算列都正确后
2. 点击"确认"按钮
3. 选择保存位置，输入文件名
4. 系统会自动导出Excel文件

## 使用技巧

1. 在输入公式时，可以双击"可用列"列表中的列名，系统会自动将其添加到公式中
2. 使用预览功能及时查看数据，确保结果符合预期
3. 添加计算列时注意检查公式格式是否正确

## 常见问题解答

### Q1: 为什么我的公式计算结果显示错误？

A1: 请检查：

- 列名是否使用英文方括号[]正确引用
- 运算符是否为英文字符
- 括号是否配对
- 引用的列名是否存在

### Q2: 如何批量处理多个文件？

A2: 目前需要逐个文件处理，建议将数据合并到一个Excel文件中再进行处理

### Q3: 计算结果显示"除以零"错误怎么办？

A3: 这是因为公式中除数出现了0，请检查数据中是否存在异常值，必要时可以添加条件处理

## 注意事项

1. 使用前请确保Excel文件格式正确
2. 大文件处理可能需要较长时间，请耐心等待
3. 建议定期保存计算结果
4. 使用自定义公式时注意数据类型的匹配

## 技术支持

如果您在使用过程中遇到问题，请检查以上文档是否有相关解答。如果问题仍未解决，请联系技术支持。

祝您使用愉快！
