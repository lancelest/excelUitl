# Excel 导出工具 API 完整文档

> 版本：v4.0.4
> 更新日期：2026-04-12

---

## 目录

- [1. 核心配置类](#1-核心配置类)
  - [1.1 SheetConfig](#11-sheetconfig)
  - [1.2 HeaderConfig](#12-headerconfig)
  - [1.3 ColumnConfig](#13-columnconfig)
  - [1.4 CellStyleConfig](#14-cellstyleconfig)
  - [1.5 StyleTemplate](#15-styletemplate)
  - [1.6 ChartConfig](#16-chartconfig)
- [2. 核心工具类](#2-核心工具类)
  - [2.1 ExcelExportUtils](#21-excelexportutils)
  - [2.2 DataConverter](#22-dataconverter)
  - [2.3 StyleProvider](#23-styleprovider)
- [3. 异常类](#3-异常类)
- [4. 完整使用示例](#4-完整使用示例)
- [5. 样式缓存机制](#5-样式缓存机制)
- [6. 常见问题](#6-常见问题)

---

## 1. 核心配置类

### 1.1 SheetConfig

Sheet 配置类，定义单个 Sheet 的所有属性。顶层入口。

#### 属性列表

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `sheetName` | String | "Sheet1" | Sheet 名称 |
| `headers` | List\<HeaderConfig\> | 空列表 | 表头配置列表，按从上到下排列 |
| `columnConfigs` | List\<ColumnConfig\> | 空列表 | 列配置列表 |
| `dataList` | List\<T\> | null | 数据列表（支持实体类和 Map） |
| `freezeRow` | int | 0 | 冻结前 N 行，0 = 不冻结 |
| `freezeCol` | int | 0 | 冻结前 N 列，0 = 不冻结 |
| `batchSize` | int | 100 | 批量刷新间隔（行数） |
| `defaultRowHeight` | short | 400 | 默认行高（twips，1/20 磅） |
| `displayGridlines` | boolean | true | 是否显示网格线 |
| `autoFilter` | boolean | false | 是否启用自动筛选 |
| `defaultDataStyle` | CellStyleConfig | null | 默认数据行样式（兜底） |
| `chartConfig` | ChartConfig | null | 图表配置，null = 不生成图表 |

#### 使用示例

```java
SheetConfig<User> config = new SheetConfig<User>()
    .setSheetName("员工报表")
    .setHeaders(Arrays.asList(header1, header2))
    .setColumnConfigs(columns)
    .setDataList(dataList)
    .setFreezeRow(2)
    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
    .setBatchSize(500)
    .setChartConfig(chartConfig);
```

---

### 1.2 HeaderConfig

表头配置类，支持多级表头和单元格合并。

#### 属性列表

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `columnNames` | List\<String\> | 空列表 | 列名列表，长度应与列数一致 |
| `styleConfig` | CellStyleConfig | null | 表头样式 |
| `height` | short | 400 | 表头行高（twips） |
| `mergeRegions` | List\<MergeRegion\> | 空列表 | 合并区域列表 |
| `level` | int | 0 | 表头层级索引（多级表头时自动设置） |

#### 方法

| 方法 | 说明 |
|------|------|
| `addColumnName(String)` | 添加一个列名 |
| `addColumnNames(String...)` | 批量添加列名 |
| `addMergeRegion(int startRow, int endRow, int startCol, int endCol)` | 添加合并区域 |
| `createHorizontalMerge(int startCol, int endCol)` | 当前行内横向合并 |

> 合并区域坐标相对于当前 HeaderConfig 所在行，0-based。

#### MergeRegion 内部类

| 字段 | 类型 | 说明 |
|------|------|------|
| `startRow` | int | 起始行（相对） |
| `endRow` | int | 结束行（相对） |
| `startCol` | int | 起始列（0-based） |
| `endCol` | int | 结束列（0-based） |

---

### 1.3 ColumnConfig

列配置类，定义每列的字段映射、宽度、样式、数据转换规则。

#### 属性列表

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `fieldName` | String | null | 字段名（实体类属性或 Map key） |
| `width` | int | 20 | 列宽（字符数 × 256） |
| `styleConfig` | CellStyleConfig | null | 静态列样式 |
| `styleProvider` | StyleProvider | null | 动态样式（优先级 > styleConfig） |
| `numberFormat` | String | null | 数字格式（如 `#,##0.00`、`0.00%`） |
| `dateFormat` | String | null | 日期格式（如 `yyyy-MM-dd`） |
| `hidden` | boolean | false | 是否隐藏 |
| `dataConverter` | DataConverter | null | 数据转换器 |

#### 使用示例

```java
new ColumnConfig()
    .setFieldName("salary")
    .setWidth(18)
    .setNumberFormat("#,##0")
    .setStyleConfig(StyleTemplate.DATA.toStyleConfig())
    .setStyleProvider((value, rowData, rowIndex) -> {
        if (((Number) value).doubleValue() > 20000) {
            return StyleTemplate.DATA.toStyleConfig()
                .setRgbBackgroundColor(StyleTemplate.rgb(255, 255, 200));
        }
        return null;
    })
    .setDataConverter((value, rowData, rowIndex) -> value == null ? 0 : value);
```

---

### 1.4 CellStyleConfig

单元格样式配置类，定义字体、颜色、边框、对齐等所有样式属性。

#### 属性列表

| 类别 | 属性 | 类型 | 默认值 | 说明 |
|------|------|------|--------|------|
| **颜色** | `backgroundColor` | short | -1 | 背景色（IndexedColors 索引） |
| | `rgbBackgroundColor` | byte[] | null | RGB 背景色（优先级更高） |
| | `fontColor` | short | -1 | 字体颜色 |
| **字体** | `bold` | boolean | false | 加粗 |
| | `italic` | boolean | false | 斜体 |
| | `strikeout` | boolean | false | 删除线 |
| | `underline` | byte | 0 | 下划线（0/1/2） |
| | `fontSize` | short | 10 | 字号（磅） |
| | `fontName` | String | "微软雅黑" | 字体名称 |
| **对齐** | `horizontalAlignment` | HorizontalAlignment | CENTER | 水平对齐 |
| | `verticalAlignment` | VerticalAlignment | CENTER | 垂直对齐 |
| **边框** | `borderTop` | BorderStyle | THIN | 上边框 |
| | `borderBottom` | BorderStyle | THIN | 下边框 |
| | `borderLeft` | BorderStyle | THIN | 左边框 |
| | `borderRight` | BorderStyle | THIN | 右边框 |
| **其他** | `wrapText` | boolean | false | 自动换行 |
| | `indentation` | short | 0 | 缩进 |
| | `rotation` | short | 0 | 旋转角度（-90~90） |
| | `locked` | boolean | true | 锁定单元格 |
| | `hidden` | boolean | false | 隐藏公式 |

#### 方法

| 方法 | 说明 |
|------|------|
| `setAllBorders(BorderStyle)` | 统一设置四边边框 |

---

### 1.5 StyleTemplate

样式模板枚举，提供预定义样式，避免重复配置。

#### 枚举列表

| 模板 | 说明 | 背景色 | 字体色 |
|------|------|--------|--------|
| `HEADER` | 表头样式 | 红色 | 白色、加粗 |
| `DATA` | 数据样式 | 白色 | 黑色、边框 |

#### 方法

| 方法 | 返回值 | 说明 |
|------|--------|------|
| `toStyleConfig()` | CellStyleConfig | 返回深拷贝（修改不影响模板） |
| `copyStyleConfig()` | CellStyleConfig | 同 toStyleConfig() |
| `rgb(int r, int g, int b)` | byte[] | 创建 RGB 颜色数组 |

#### 使用方式

```java
// 直接使用
StyleTemplate.HEADER.toStyleConfig()

// 基于模板修改（深拷贝）
StyleTemplate.HEADER.toStyleConfig()
    .setRgbBackgroundColor(StyleTemplate.rgb(0, 112, 192));

// 自定义 RGB
.setRgbBackgroundColor(StyleTemplate.rgb(255, 0, 0))   // 红
.setRgbBackgroundColor(StyleTemplate.rgb(0, 255, 0))   // 绿
.setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128)) // 灰
```

---

### 1.6 ChartConfig

图表配置类，在 Sheet 数据结束后自动生成柱状图/条形图。

#### 属性列表

| 类别 | 属性 | 类型 | 默认值 | 说明 |
|------|------|------|--------|------|
| **基础** | `title` | String | null | 图表标题，空则不显示 |
| | `gapRows` | int | 2 | 数据行与图表之间的间隔行数 |
| **字体** | `fontSize` | int | 10 | 图表字体大小（磅） |
| **数据** | `categoryColumn` | String | 必填 | X 轴类别列字段名 |
| | `seriesList` | List\<SeriesConfig\> | 空列表 | Y 轴数据系列 |
| **外观** | `barChart` | boolean | false | true = 条形图（水平） |
| | `barGrouping` | BarGrouping | CLUSTERED | 分组方式 |
| | `showDataLabel` | boolean | true | 是否显示数据标签 |
| | `legendPosition` | LegendPosition | BOTTOM | 图例位置 |
| | `seriesColors` | List\<String\> | 空列表 | 系列颜色（十六进制） |
| **网格线** | `showMajorGridlines` | boolean | true | 是否显示主要网格线 |
| | `majorGridlineColor` | String | "D9D9D9" | 网格线颜色 |
| **坐标轴** | `categoryAxisRotation` | Integer | null | X 轴标签旋转角度 |
| | `categoryAxisTitle` | String | null | X 轴标题 |
| | `valueAxisTitle` | String | null | Y 轴标题 |
| | `valueAxisMin` | Double | null | Y 轴最小值 |
| | `valueAxisMax` | Double | null | Y 轴最大值 |
| | `valueAxisUnit` | Double | null | Y 轴刻度间隔 |

#### 内部枚举

**BarGrouping（柱图分组方式）**

| 值 | 说明 |
|------|------|
| `CLUSTERED` | 簇状柱图（默认） |
| `STACKED` | 堆积柱图 |
| `PERCENT_STACKED` | 百分比堆积柱图 |

**LegendPosition（图例位置）**

| 值 | 说明 |
|------|------|
| `TOP` | 顶部 |
| `BOTTOM` | 底部（默认） |
| `LEFT` | 左侧 |
| `RIGHT` | 右侧 |
| `NONE` | 不显示 |

#### SeriesConfig（数据系列）

| 字段 | 类型 | 说明 |
|------|------|------|
| `fieldName` | String | 字段名（对应 ColumnConfig.fieldName） |
| `seriesName` | String | 系列名称（显示在图例中） |

#### 方法

| 方法 | 说明 |
|------|------|
| `addSeries(String fieldName, String seriesName)` | 添加数据系列 |

#### 废弃方法

| 方法 | 替代方案 |
|------|---------|
| `setStacked(boolean)` | `setBarGrouping(BarGrouping.STACKED)` |
| `isStacked()` | `getBarGrouping() != CLUSTERED` |

#### 使用示例

```java
// 簇状柱图
new ChartConfig()
    .setTitle("部门绩效对比")
    .setCategoryColumn("department")
    .addSeries("performanceScore", "绩效评分")
    .addSeries("salary", "薪资");

// 堆积柱图
new ChartConfig()
    .setTitle("季度业绩")
    .setCategoryColumn("quarter")
    .addSeries("productA", "产品A")
    .addSeries("productB", "产品B")
    .setBarGrouping(ChartConfig.BarGrouping.STACKED)
    .setShowDataLabel(false);  // 堆积时建议关闭

// 自定义颜色和网格线
new ChartConfig()
    .setTitle("销售趋势")
    .setCategoryColumn("month")
    .addSeries("sales", "销售额")
    .setSeriesColors(Arrays.asList("FF6B6B"))
    .setMajorGridlineColor("E0E0E0")
    .setFontSize(11);

// 附加到 Sheet
sheetConfig.setChartConfig(chartConfig);
```

---

## 2. 核心工具类

### 2.1 ExcelExportUtils

核心导出工具类。

#### 方法列表

| 方法 | 说明 | 使用场景 |
|------|------|--------|
| `export(OutputStream, SheetConfig<?>...)` | 导出到输出流 | 文件导出 |
| `exportSingleSheet(OutputStream, SheetConfig<?>)` | 单 Sheet 导出（语义别名） | 简化单 Sheet 场景 |
| `exportMultiSheets(OutputStream, List<SheetConfig<?>>)` | 多 Sheet 导出 | 多 Sheet 场景 |
| `exportToResponse(HttpServletResponse, String, SheetConfig<?>...)` | 导出到 HTTP 响应 | Web 导出（自定义文件名） |
| `exportToResponseWithTimestamp(HttpServletResponse, String, SheetConfig<?>...)` | 导出并自动加时间戳 | Web 导出（防冲突） |
| `createSimpleTemplate(...)` | 快速创建标准配置 | 快速生成简单报表 |

#### 方法签名

```java
// 通用导出
public static void export(OutputStream outputStream, SheetConfig<?>... sheetConfigs)
    throws ExcelExportException

// Web 导出（自定义文件名）
public static void exportToResponse(HttpServletResponse response, String fileName,
                                    SheetConfig<?>... sheetConfigs)
    throws ExcelExportException, IOException

// Web 导出（自动加时间戳）
public static void exportToResponseWithTimestamp(HttpServletResponse response, String baseName,
                                                   SheetConfig<?>... sheetConfigs)
    throws ExcelExportException, IOException

// 快速创建标准配置
public static <T> SheetConfig<T> createSimpleTemplate(
    String sheetName, List<T> dataList,
    String[] fieldNames, String[] columnNames)
```

#### 使用示例

```java
// 导出到文件
try (OutputStream os = new FileOutputStream("报表.xlsx")) {
    ExcelExportUtils.export(os, sheetConfig);
}

// 多 Sheet
ExcelExportUtils.export(os, sheet1, sheet2, sheet3);

// Web 导出
ExcelExportUtils.exportToResponse(response, "月度报表.xlsx", config);

// Web 导出 + 时间戳 → 月度报表_20260412_193000.xlsx
ExcelExportUtils.exportToResponseWithTimestamp(response, "月度报表", config);
```

---

### 2.2 DataConverter

数据转换器函数式接口，在写入单元格前对原始数据进行自定义转换。

```java
@FunctionalInterface
public interface DataConverter {
    Object convert(Object value, Object rowData, int rowIndex);
}
```

| 参数 | 说明 |
|------|------|
| `value` | 原始字段值（getter / 反射获取） |
| `rowData` | 整行数据对象（可访问其他字段） |
| `rowIndex` | 行索引（从 0 开始，表头不计入） |
| **返回值** | 转换后的值（写入单元格） |

#### 典型场景

| 场景 | 示例 |
|------|------|
| 枚举转中文 | `1 → "男"` |
| 状态码转描述 | `PENDING → "待处理"` |
| 小数转百分比 | `0.85 → "85.00%"` |
| 日期格式化 | `Date → "2026年04月12日"` |
| 多字段组合 | `firstName + lastName → "张三"` |
| 空值处理 | `null → "—"` |

#### 注意事项

- 转换器返回值类型不限（Number、Date、String 等）
- 抛出异常会包装为 `ExcelExportException`，中断导出
- 执行时机早于 `setCellValueSafe()`
- 建议处理 null 值

---

### 2.3 StyleProvider

动态样式提供器函数式接口，根据值/行数据/行索引动态决定样式。

```java
@FunctionalInterface
public interface StyleProvider {
    CellStyleConfig provide(Object value, Object rowData, int rowIndex);
}
```

| 参数 | 说明 |
|------|------|
| `value` | 原始字段值（DataConverter **执行前**） |
| `rowData` | 整行数据对象 |
| `rowIndex` | 行索引（从 0 开始） |
| **返回值** | 样式配置；返回 `null` 则降级使用列 `styleConfig` |

#### 样式优先级

```
StyleProvider（动态）
    ↓ 返回 null 时降级
列 styleConfig（静态）
    ↓ 没有时降级
Sheet defaultDataStyle（兜底）
```

#### 典型场景

| 场景 | 示例 |
|------|------|
| 阈值高亮 | 薪资 > 20000 → 黄色 + 灰底 |
| 状态着色 | "异常" → 红色，"正常" → 绿色 |
| 多字段判断 | 部门="技术" 且 绩效>90 → 绿色 |
| 斑马纹 | 偶数行浅蓝底色 |

#### 与 DataConverter 的区别

| | DataConverter | StyleProvider |
|--|-------------|---------------|
| 改变值 | ✅ | ❌ |
| 改变样式 | ❌ | ✅ |
| 接收的值 | 原始值 | 原始值（同） |
| 可同时配置 | ✅ | ✅ |

---

## 3. 异常类

### ExcelExportException

```java
public class ExcelExportException extends RuntimeException {
    private String errorCode;
    // 构造方法: (message), (message, errorCode), (message, cause), (message, errorCode, cause)
}
```

#### 错误代码

| 错误码 | 说明 |
|--------|------|
| `NULL_OUTPUT_STREAM` | 输出流为空 |
| `EMPTY_SHEET_CONFIG` | Sheet 配置为空 |
| `NULL_SHEET_CONFIG` | 第 N 个 Sheet 配置为空 |
| `EMPTY_COLUMN_CONFIG` | 列配置为空 |
| `DATA_EXCEED_LIMIT` | 数据量超过最大行数 |
| `COLUMN_EXCEED_LIMIT` | 列数超过最大列数 |
| `EMPTY_FIELD_NAME` | 字段名为空 |
| `FIELD_NOT_FOUND` | 字段不存在 |
| `FIELD_ACCESS_DENIED` | 字段访问权限不足 |
| `DATA_CONVERT_ERROR` | 数据转换失败 |
| `STYLE_PROVIDER_ERROR` | 动态样式计算失败 |
| `MERGE_ERROR` | 合并单元格失败 |
| `FLUSH_ERROR` | 刷新数据失败 |
| `EXPORT_ERROR` | 导出失败 |

---

## 4. 完整使用示例

### 4.1 基础单 Sheet 导出

```java
public class BasicExportExample {
    public static void main(String[] args) throws Exception {
        List<User> dataList = Arrays.asList(
            new User("张三", "技术部", 15000),
            new User("李四", "销售部", 12000),
            new User("王五", "人事部", 8000)
        );

        List<ColumnConfig> columns = Arrays.asList(
            new ColumnConfig().setFieldName("name").setWidth(12)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()),
            new ColumnConfig().setFieldName("dept").setWidth(12)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()),
            new ColumnConfig().setFieldName("salary").setWidth(15)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig())
        );

        HeaderConfig header = new HeaderConfig()
            .addColumnNames("姓名", "部门", "薪资")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        SheetConfig<User> sheet = new SheetConfig<User>()
            .setSheetName("员工报表")
            .setHeaders(Collections.singletonList(header))
            .setColumnConfigs(columns)
            .setDataList(dataList)
            .setFreezeRow(1)
            .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig());

        try (OutputStream os = new FileOutputStream("员工报表.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }
    }
}
```

### 4.2 多级表头 + 图表

```java
public class ChartExportExample {
    public static void main(String[] args) throws Exception {
        List<Employee> data = generateData(10);

        // 表头
        HeaderConfig h1 = new HeaderConfig()
            .addColumnNames("基本信息", "基本信息", "绩效数据")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
            .addMergeRegion(0, 0, 0, 1);
        HeaderConfig h2 = new HeaderConfig()
            .addColumnNames("姓名", "部门", "评分")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        // 列
        List<ColumnConfig> columns = Arrays.asList(
            new ColumnConfig().setFieldName("name").setWidth(12),
            new ColumnConfig().setFieldName("department").setWidth(12),
            new ColumnConfig().setFieldName("score").setWidth(10).setNumberFormat("0.00")
        );

        // 图表
        ChartConfig chart = new ChartConfig()
            .setTitle("部门绩效对比")
            .setCategoryColumn("department")
            .addSeries("score", "绩效评分")
            .setBarGrouping(ChartConfig.BarGrouping.CLUSTERED)
            .setSeriesColors(Arrays.asList("4472C4"));

        // Sheet
        SheetConfig<Employee> sheet = new SheetConfig<Employee>()
            .setSheetName("绩效报表")
            .setHeaders(Arrays.asList(h1, h2))
            .setColumnConfigs(columns)
            .setDataList(data)
            .setFreezeRow(2)
            .setChartConfig(chart);

        try (OutputStream os = new FileOutputStream("绩效图表.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }
    }
}
```

### 4.3 Web 场景导出

```java
@RestController
@RequestMapping("/api/export")
public class ExportController {

    @GetMapping("/report")
    public void exportReport(HttpServletResponse response) {
        try {
            ExcelExportUtils.exportToResponseWithTimestamp(
                response, "月度报表", buildSheetConfig()
            );
        } catch (ExcelExportException | IOException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
    }
}
```

---

## 5. 样式缓存机制

工具内部使用**样式指纹**（fingerprint）缓存 POI CellStyle 对象。相同样式自动复用，避免内存溢出。

### 缓存键格式

```
header_0|bg:10|rgb:null|fc:9|bold:true|fs:11|fn:微软雅黑|...
data_salary|rgb:128,128,128|fc:10|bold:false|...
```

### 关键点

| 要点 | 说明 |
|------|------|
| 深拷贝 | `StyleTemplate.toStyleConfig()` 每次返回新对象，不会互相影响 |
| 自动纳入 | StyleProvider 返回的样式也会自动缓存 |
| 安全 | 不同样式配置不会相互干扰 |
| 反射缓存 | Method / Field 只查找一次（ConcurrentHashMap） |

---

## 6. 常见问题

### Q1: 自定义 RGB 颜色不生效？

```java
// ❌ 错误：setBackgroundColor 接收 short，RGB 值会溢出
.setBackgroundColor(StyleTemplate.rgb(255, 0, 0))

// ✅ 正确：使用 setRgbBackgroundColor
.setRgbBackgroundColor(StyleTemplate.rgb(255, 0, 0))
```

### Q2: StyleProvider 返回的样式互相影响？

`StyleTemplate.toStyleConfig()` 已是深拷贝，不会。确保不要在回调外共享同一个 CellStyleConfig 实例并修改。

### Q3: 图表打开后需要修复？

v4.0.4 已彻底解决 POI 图表 XML 序列化问题（通过 ZIP 后处理架构）。如仍遇到，请确认版本 >= v4.0.4。

### Q4: 合并单元格后数据错位？

MergeRegion 的 startRow/endRow 是**相对于当前 HeaderConfig**，不是绝对行号。多级表头中每个 HeaderConfig 的坐标都从 0 开始。

### Q5: DataConverter 和 StyleProvider 冲突？

不冲突。DataConverter 改变值，StyleProvider 改变样式。可同时配置。注意 StyleProvider 拿到的是 DataConverter **执行前**的原始值。

### Q6: Web 导出中文文件名乱码？

使用内置 `exportToResponse()` 方法，已自动处理 URLEncoder。

### Q7: 大数据导出内存溢出？

设置合理 batchSize（大数据 1000-5000），确保使用 try-with-resources 关闭流。

---

## 版本历史

| 版本 | 日期 | 主要变更 |
|------|------|--------|
| v4.0.4 | 2026-04-12 | 堆积柱图（BarGrouping 枚举）；主要网格线配置 |
| v4.0.3 | 2026-04-11 | 图表字体大小可配置 |
| v4.0.0 | 2026-03-25 | 图表功能；XSSFWorkbook；ZIP 后处理架构 |
| v3.0.5 | 2026-03-25 | 反射缓存优化 |
| v3.0.2 | 2026-03-22 | StyleProvider；Web 导出方法 |
| v3.0.1 | 2026-03-22 | 修复 toStyleConfig 共享引用；修复 rgb 溢出 |
| v3.0.0 | 2026-03-22 | StyleTemplate；样式指纹缓存 |

---

**文档结束**

如有问题，请参考 `README.md` 或源码 Javadoc 注释。
