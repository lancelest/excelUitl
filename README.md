# Excel 导出工具

企业级 Excel 导出库，基于 Apache POI，支持百万级数据导出、多级表头、条件样式、内置图表。

## 功能特性

- **数据导出** — 支持实体类和 Map 两种数据源，单 Sheet 最大 1,048,576 行
- **样式系统** — 样式模板 + 自定义样式 + 动态样式（StyleProvider），样式指纹缓存避免内存溢出
- **多级表头** — 支持任意层级表头和单元格合并
- **数据转换** — DataConverter 函数式接口，枚举转中文、状态码映射、日期格式化等
- **条件样式** — StyleProvider 根据值/行数据动态着色，支持斑马纹、阈值高亮等
- **内置图表** — 簇状/堆积/百分比堆积柱图，支持数据标签、网格线、系列颜色
- **Web 导出** — 内置 HttpServletResponse 导出，自动处理中文文件名编码

## 快速开始

### Maven 依赖

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.projectlombok</groupId>
    <artifactId>lombok</artifactId>
    <version>1.18.30</version>
    <scope>provided</scope>
</dependency>
<dependency>
    <groupId>org.apache.commons</groupId>
    <artifactId>commons-lang3</artifactId>
    <version>3.14.0</version>
</dependency>
<dependency>
    <groupId>org.apache.commons</groupId>
    <artifactId>commons-collections4</artifactId>
    <version>4.4</version>
</dependency>
```

### 三行代码导出

```java
SheetConfig<User> config = ExcelExportUtils.createSimpleTemplate(
    "用户信息", dataList,
    new String[]{"id", "name", "department"},
    new String[]{"ID", "姓名", "部门"}
);
ExcelExportUtils.exportSingleSheet(new FileOutputStream("用户信息.xlsx"), config);
```

---

## 核心概念

### 配置类层次

```
SheetConfig<T>              ← Sheet 配置（顶层入口）
├── List<HeaderConfig>      ← 表头（支持多级 + 合并）
│   └── CellStyleConfig     ← 表头样式
├── List<ColumnConfig>      ← 列配置
│   ├── CellStyleConfig     ← 静态列样式
│   ├── DataConverter       ← 数据转换器
│   └── StyleProvider       ← 动态样式提供器
├── CellStyleConfig         ← 默认数据行样式（兜底）
└── ChartConfig             ← 图表配置（可选）
```

### 样式优先级

```
StyleProvider（动态）> 列 styleConfig（静态）> Sheet defaultDataStyle（兜底）
```

---

## 使用指南

### 1. 基础导出

```java
// 实体类数据源
SheetConfig<User> config = new SheetConfig<User>()
    .setSheetName("员工表")
    .setHeaders(Collections.singletonList(
        new HeaderConfig()
            .addColumnNames("ID", "姓名", "部门")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
    ))
    .setColumnConfigs(Arrays.asList(
        new ColumnConfig().setFieldName("id").setWidth(8),
        new ColumnConfig().setFieldName("name").setWidth(15),
        new ColumnConfig().setFieldName("department").setWidth(15)
    ))
    .setDataList(userList);

ExcelExportUtils.export(new FileOutputStream("员工表.xlsx"), config);
```

### 2. 多级表头

```java
// 一级表头
HeaderConfig h1 = new HeaderConfig()
    .addColumnNames("基本信息", "基本信息", "绩效数据", "绩效数据")
    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
    .addMergeRegion(0, 0, 0, 1)   // "基本信息" 横向合并
    .addMergeRegion(0, 0, 2, 3);  // "绩效数据" 横向合并

// 二级表头
HeaderConfig h2 = new HeaderConfig()
    .addColumnNames("姓名", "部门", "评分", "项目数")
    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

SheetConfig<User> config = new SheetConfig<User>()
    .setHeaders(Arrays.asList(h1, h2))  // 按从上到下排列
    .setColumnConfigs(columns)
    .setDataList(dataList)
    .setFreezeRow(2);  // 冻结两行表头
```

效果：
```
┌──────────────┬──────────────┐
│   基本信息    │   绩效数据    │
├──────┬───────┼──────┬───────┤
│ 姓名 │ 部门  │ 评分 │ 项目数 │
├──────┼───────┼──────┼───────┤
│ 张三 │ 技术部 │  95  │   5   │
└──────┴───────┴──────┴───────┘
```

### 3. 数据转换（DataConverter）

```java
// 枚举值转中文
new ColumnConfig()
    .setFieldName("gender")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "未知";
        return ((Integer) value) == 1 ? "男" : "女";
    });

// 小数转百分比
new ColumnConfig()
    .setFieldName("rate")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "0%";
        return String.format("%.2f%%", ((Number) value).doubleValue() * 100);
    });
```

### 4. 条件样式（StyleProvider）

```java
// 薪资 > 20000 高亮
new ColumnConfig()
    .setFieldName("salary")
    .setStyleProvider((value, rowData, rowIndex) -> {
        if (value != null && ((Number) value).doubleValue() > 20000) {
            return StyleTemplate.DATA.toStyleConfig()
                .setRgbBackgroundColor(StyleTemplate.rgb(255, 255, 200))
                .setBold(true);
        }
        return null;  // null → 降级使用列 styleConfig
    });
```

### 5. 内置图表

```java
// 簇状柱图（默认）
ChartConfig chart = new ChartConfig()
    .setTitle("部门绩效对比")
    .setCategoryColumn("department")
    .addSeries("performanceScore", "绩效评分")
    .addSeries("salary", "薪资");

// 堆积柱图
ChartConfig stackedChart = new ChartConfig()
    .setTitle("季度业绩")
    .setCategoryColumn("quarter")
    .addSeries("productA", "产品A")
    .addSeries("productB", "产品B")
    .setBarGrouping(ChartConfig.BarGrouping.STACKED);

// 附加到 Sheet
sheetConfig.setChartConfig(chart);
```

**图表配置项一览：**

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `title` | 图表标题 | 空 |
| `categoryColumn` | X 轴类别列字段名 | 必填 |
| `seriesList` | Y 轴数据系列 | 必填 |
| `barChart` | 是否条形图（水平） | false |
| `barGrouping` | CLUSTERED / STACKED / PERCENT_STACKED | CLUSTERED |
| `showDataLabel` | 是否显示数据标签 | true |
| `fontSize` | 图表字体大小（磅） | 10 |
| `showMajorGridlines` | 是否显示主要网格线 | true |
| `majorGridlineColor` | 网格线颜色（十六进制） | "D9D9D9" |
| `seriesColors` | 系列颜色列表 | POI 默认配色 |
| `legendPosition` | 图例位置 | BOTTOM |
| `categoryAxisRotation` | X 轴标签旋转角度 | null（水平） |

### 6. Web 导出

```java
// 自定义文件名
ExcelExportUtils.exportToResponse(response, "月度报表.xlsx", config);

// 自动加时间戳 → 月度报表_20260412_193000.xlsx
ExcelExportUtils.exportToResponseWithTimestamp(response, "月度报表", config);
```

### 7. 多 Sheet 导出

```java
ExcelExportUtils.export(os, sheet1, sheet2, sheet3);
```

---

## 样式模板

### 表头样式

| 模板 | 说明 | 背景色 | 字体色 |
|------|------|--------|--------|
| `HEADER` | 表头样式 | 红色 | 白色、加粗 |

### 数据样式

| 模板 | 说明 | 背景色 | 字体色 |
|------|------|--------|--------|
| `DATA` | 数据样式 | 白色 | 黑色、边框 |

### 使用方式

```java
// 直接使用
.setStyleConfig(StyleTemplate.HEADER.toStyleConfig())

// 基于模板修改（深拷贝，不影响模板本身）
StyleTemplate.HEADER.toStyleConfig()
    .setRgbBackgroundColor(StyleTemplate.rgb(0, 112, 192));  // 改为蓝色

// 自定义 RGB 颜色
.setRgbBackgroundColor(StyleTemplate.rgb(255, 0, 0))   // 红色
.setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128)) // 灰色
```

---

## 数字 / 日期格式

```java
.setNumberFormat("#,##0")       // 千分位：12,345
.setNumberFormat("0.00")        // 两位小数：12.35
.setNumberFormat("0.00%")       // 百分比：85.00%
.setNumberFormat("0.00E+00")    // 科学计数法
.setDateFormat("yyyy-MM-dd")    // 日期格式
.setDateFormat("yyyy年MM月dd日") // 自定义日期
```

---

## 最佳实践

| 实践 | 说明 |
|------|------|
| 使用样式模板 | `StyleTemplate.HEADER.toStyleConfig()` 代替手动配置 10+ 属性 |
| 合理设置列宽 | 文本 12-15，金额 18-20，日期 14 |
| 设置 batchSize | 小数据 100-500，大数据 1000-5000 |
| 处理 null 值 | 在 DataConverter 中处理，避免空单元格 |
| try-with-resources | 确保流正确关闭 |
| 异常捕获 | 捕获 ExcelExportException 和 IOException |

---

## 性能

| 指标 | 数值 |
|------|------|
| 最大行数 | 1,048,576 |
| 最大列数 | 16,384 |
| 导出速度 | ~1000 行/秒 |
| 样式缓存 | 指纹去重，相同样式复用同一 POI CellStyle |
| 反射缓存 | Method / Field 只查找一次，60+ 字段场景从 1700ms 降至 50ms |

---

## 异常处理

```java
try {
    ExcelExportUtils.export(os, config);
} catch (ExcelExportException e) {
    System.err.println("错误码: " + e.getErrorCode());
    System.err.println("错误信息: " + e.getMessage());
} catch (IOException e) {
    System.err.println("文件操作失败: " + e.getMessage());
}
```

| 错误码 | 说明 |
|--------|------|
| `NULL_OUTPUT_STREAM` | 输出流为空 |
| `EMPTY_SHEET_CONFIG` | Sheet 配置为空 |
| `DATA_EXCEED_LIMIT` | 数据量超过最大行数 |
| `FIELD_NOT_FOUND` | 字段不存在 |
| `DATA_CONVERT_ERROR` | 数据转换失败 |
| `STYLE_PROVIDER_ERROR` | 动态样式计算失败 |
| `EXPORT_ERROR` | 导出失败 |

---

## 常见问题

**Q: 导出的 Excel 打开提示格式错误？**
确保流正确关闭，使用 try-with-resources。

**Q: 自定义 RGB 颜色不生效？**
使用 `setRgbBackgroundColor(StyleTemplate.rgb(r, g, b))`，不要用 `setBackgroundColor(short)`，short 会溢出。

**Q: StyleProvider 返回的样式互相影响？**
`StyleTemplate.toStyleConfig()` 返回深拷贝，不会互相影响。

**Q: Web 导出中文文件名乱码？**
使用内置的 `exportToResponse()` 方法，已自动处理 URLEncoder 编码。

**Q: 图表打开后需要修复？**
v4.0 已通过 ZIP 后处理架构解决 POI 图表 XML 序列化问题，如遇到请升级到 v4.0.4。

---

## 版本历史

### v4.0.4（2026-04-12）

- 新增堆积柱图（STACKED / PERCENT_STACKED），通过 `BarGrouping` 枚举配置
- 新增主要网格线配置（`showMajorGridlines` + `majorGridlineColor`）
- 废弃 `setStacked(boolean)`，保留向后兼容

### v4.0.3（2026-04-11）

- 图表字体大小可配置（`fontSize`，默认 10）

### v4.0.0（2026-03-25）

- 新增图表功能（簇状柱图/条形图）
- XSSFWorkbook 替代 SXSSFWorkbook（图表不支持流式写入）
- ZIP 后处理架构解决 POI 图表 XML 序列化问题

### v3.0.5（2026-03-25）

- 反射缓存优化（60+ 字段场景从 1700ms 降至 50ms）

### v3.0.4（2026-03-25）

- 合并区域文字自动归位

### v3.0.3（2026-03-25）

- 表头列宽自适应

### v3.0.2（2026-03-22）

- 新增 StyleProvider 动态样式接口
- 新增 Web 导出方法

### v3.0.1（2026-03-22）

- 修复 `toStyleConfig()` 共享引用 Bug
- 修复 `rgb()` 颜色溢出 Bug

### v3.0.0（2026-03-22）

- 新增 StyleTemplate 样式模板
- 优化样式缓存策略

---

## 许可证

MIT License
