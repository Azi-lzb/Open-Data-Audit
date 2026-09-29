# Excel、Direct OOXML、openpyxl 与 Office 边界

## 1. 决策顺序

处理 Excel 文件时，默认按以下顺序选择：

```text
能否用纯 Python 完成？
├─ 固定 OOXML + 性能敏感 + 结构可控 → Direct OOXML
├─ 通用静态 Excel 读写              → openpyxl
└─ 需要真实工作簿计算                → Office
```

## 2. Direct OOXML

`xlsx/xlsm = ZIP + OOXML`。

适合：
- 固定模板；
- 高速批量渲染；
- 只修改明确 XML；
- 需要保留其他部件；
- 性能瓶颈明确。

要求：
- XML 使用 namespace-aware parser；
- 禁止依赖 `ns0:` 等具体前缀的正则解析结构；
- 正确维护 Relationship、Content Types、Sheet graph；
- 删除 Sheet 后校验不存在悬空引用；
- 输出必须能重新用 openpyxl/Office 打开。

不适合：
- 工作簿真实计算；
- 未知任意 Excel 的万能编辑器；
- 不能证明结构安全的复杂对象修改。

## 3. openpyxl

作为通用静态 OOXML 读写层。

适合：
- 单元格；
- 工作表；
- 公式文本；
- 样式；
- 命名区域；
- 静态规则；
- 不涉及真实计算的模板处理。

openpyxl：
- 可以读写公式文本；
- **不会真实计算公式**。

## 4. Office

### Windows
- Excel
- WPS

### UOS/Linux
- LibreOffice Calc

工作簿真实公式重算必须使用真实 Office。

如果功能需要：
- 真实计算；
- 原生渲染结果；
- 经验证不可替代的原生复制语义；

才进入 Office 层。

## 5. 工作簿计算硬边界

禁止：
- Python 自制不完整公式引擎冒充 Office；
- openpyxl 冒充公式计算；
- 仅依赖旧 cached value 当作本轮计算结果。

Office 计算流程必须：
1. 打开；
2. 重算；
3. 保存；
4. 可靠关闭；
5. 再由 Python 提取结果。

## 6. Fast OOXML

`systems/s3_central_statistics/fast_ooxml.py` 属于 Direct OOXML 高性能实现。

必须：
- workbook.xml / rels 用 XML parser；
- Relationship target 用 POSIX package path 解析；
- 输出前检查结构完整性；
- 用户错误说人话；
- 技术日志保留 rId、Target、part 等诊断信息。

## 7. Sheet 删除

Direct OOXML 物理删除 Sheet 时至少同步考虑：
- workbook.xml
- workbook.xml.rels
- `[Content_Types].xml`
- sheet rels
- definedNames
- bookViews
- 相关 drawing/table/comment 等依赖

若保留 Sheet 引用了待删除 Sheet：
- 不得静默制造 `#REF!`
- 自动清理空 Sheet 且存在未解除依赖时，保留该 Sheet（及仍被引用的名称）并隐藏；明确要求硬删除时阻止操作并报告依赖
- 不自动改业务公式

### 7.1 工作表身份与顺序字段（铁律）

- SheetName / rId / worksheet part 用于识别「删的是谁」（工作表身份）；
- old_index / new_index 用于维护依赖工作表顺序的字段
  （`localSheetId`、bookViews 的 `firstSheet`/`activeTab`）；
- **`localSheetId` 不是工作表身份**，它只是 workbook.xml 当前顺序中的位置，
  禁止把它当身份用于判定，也禁止用逐次 `-=1` 的写法回移（一次删除多张表
  会累计偏差）；必须先建立删除集合的 old→new 映射，再一次性重算。

### 7.2 Sheet-local definedNames（铁律）

Direct OOXML 删除 Sheet 时必须维护 Sheet-local definedNames
（`_xlnm._FilterDatabase` / `_xlnm.Print_Area` / `_xlnm.Print_Titles`）：

- 被物理删除 Sheet 的 local definedName（owner=被删表）应同步删除；如果保留公式仍引用该名称，不物理删除 Sheet，按自动清理规则隐藏并保留名称；
- 属于保留 Sheet 的 localSheetId 必须根据新的 Sheet 顺序重映射；
- 应结合 `localSheetId` 对应的 owner 与名称公式文本中的 SheetName 做
  交叉验证；两者矛盾视为作用域/引用异常，不得自动猜测改写；
- 无法证明安全的复杂 definedName（自定义名称、跨表复杂引用、外部引用、
  动态公式）必须保守处理：与被删表无关则保留；疑似引用待删空表时，
  自动删空表场景优先把该空表降级为隐藏，不得生成可能损坏的工作簿；
- 引用「明确业务删除」表（如内嵌配置页）且无法挽留的名称一并移除并记日志；
- 用户提示不得出现 localSheetId / definedName / rId / part 等底层术语；
  技术日志保留 defined_name / owner / old/new localSheetId / action / reason。

### 7.3 definedName 公式依赖保护（铁律）

Direct OOXML 删除工作表及其 definedName 前，必须确认最终保留工作表没有继续引用该名称；无法证明安全时不得物理删除，也不得静默制造 `#NAME?`。处置方式固定为：自动清理场景保留工作表和依赖名称并隐藏；用户明确要求硬删除时阻止并报告，由用户决定如何处理依赖。

- 名称引用检测按 Excel 名称语义做 token 边界 + 大小写不敏感匹配
  （``Rate`` 不得误判 ``InterestRate``；``myconfigvalue`` 命中
  ``MyConfigValue``），禁止裸 substring；
- 无人使用的名称可随表安全删除；自动清理时若有公式引用名称，保留工作表和名称并隐藏；明确业务硬删时阻止并报告；
- 绝不自动重写业务公式（改引用目标、内联名称都属于业务公式重写）；
- 保留表公式识别不到时宁可保守，不得产出 #NAME? 工作簿。

删除后的完整性验证不得降低标准：sheet↔relationship↔part 无悬空、
被删表 part 不再被引用、被删表的 local definedName 不存在、保留表
localSheetId 与新顺序一致且不越界、剩余名称不指向已删除的 Sheet、
输出可被 openpyxl 重开并可再次进入 FastOoxmlCompiler。

## 8. 条件格式

规则是否触发与显示颜色不是一回事。

纯 Python/OOXML 可做已验证规则的静态求值；
真实显示效果若必须验证，则进入对应平台 Office 层。

统一状态：
- TRUE
- FALSE
- UNSUPPORTED
- ERROR

不支持的规则不得伪装成 FALSE。
