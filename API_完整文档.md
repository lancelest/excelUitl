# Excel导出工具 API 完整文档

> 版本：v3.0.2
> 更新日期：2026-03-22

---

## 目录

- [1. 核心配置类](#1-核心配置类)
  - [1.1 SheetConfig](#11-sheetconfig)
  - [1.2 HeaderConfig](#12-headerconfig)
  - [1.3 ColumnConfig](#13-columnconfig)
  - [1.4 CellStyleConfig](#14-cellstyleconfig)
  - [1.5 StyleTemplate](#15-styletemplate)
- [2. 核心工具类](#2-核心工具类)
  - [2.1 ExcelExportUtils](#21-excelexportutils)
  - [2.2 DataConverter](#22-dataconverter)
  - [2.3 StyleProvider](#23-styleprovider)
- [3. 异常类](#3-异常类)
- [4. 完整使用示例](#4-完整使用示例)
  - [4.1 基础单Sheet导出](#41-基础单sheet导出)
  - [4.2 多Sheet导出](#42-多sheet导出)
  - [4.3 多级表头](#43-多级表头)
  - [4.4 DataConverter 数据转换](#44-dataconverter-数据转换)
  - [4.5 StyleProvider 条件样式](#45-styleprovider-条件样式)
  - [4.6 Web 场景导出](#46-web-场景导出)
- [5. 样式缓存机制](#5-样式缓存机制)
- [6. 性能优化建议](#6-性能优化建议)
- [7. 常见问题](#7-常见问题)

---

## 1. 核心配置类

### 1.1 SheetConfig

Sheet配置类，用于定义单个Sheet的所有属性。

#### 属性列表

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `sheetName` | String | 必填 | Sheet名称，支持中文 |
| `headers` | List<HeaderConfig> | 空列表 | 表头配置列表，支持多级表头 |
| `columnConfigs` | List<ColumnConfig> | 必填 | 列配置列表 |
| `dataList` | List<T> | 空列表 | 数据列表（支持实体类和Map） |
| `freezeRow` | int | 0 | 冻结行数（表头行数） |
| `freezeCol` | int | 0 | 冻结列数 |
| `autoFilter` | boolean | true | 是否启用自动筛选 |
| `displayGridlines` | boolean | true | 是否显示网格线 |
| `defaultRowHeight` | short | 300 | 默认行高（单位：1/20点） |
| `defaultDataStyle` | CellStyleConfig | null | 默认数据单元格样式 |
| `batchSize` | int | 100 | 批量刷新间隔（行数） |

#### 使用示例

```java
SheetConfig<User> sheet = new SheetConfig<User>()
    .setSheetName("员工报表")
    .setHeaders(Arrays.asList(header1, header2))  // 多级表头
    .setColumnConfigs(columns)                    // 列配置
    .setDataList(dataList)                       // 数据列表
    .setFreezeRow(2)                           // 冻结前2行
    .setFreezeCol(1)                           // 冻结前1列
    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig()) // 默认样式
    .setBatchSize(500);                         // 每500行刷新一次
```

---

### 1.2 HeaderConfig

表头配置类，支持多级表头和单元格合并。

#### 属性列表

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `columnNames` | List<String> | 空列表 | 列名列表 |
| `styleConfig` | CellStyleConfig | null | 表头样式 |
| `height` | short | 400 | 表头行高 |
| `mergeRegions` | List<MergeRegion> | 空列表 | 合并区域配置 |

#### 合并区域配置

```java
public static class MergeRegion {
    private int startRow;  // 合并起始行（相对于当前表头行）
    private int endRow;    // 合并结束行
    private int startCol;  // 合并起始列
    private int endCol;    // 合并结束列
}
```

#### 使用示例

```java
// 两级表头示例
HeaderConfig header1 = new HeaderConfig()
    .addColumnNames("序号", "基本信息", "基本信息", "绩效数据", "备注")
    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
    .addMergeRegion(0, 0, 1, 1)  // 第2-3列合并（基本信息）
    .addMergeRegion(0, 0, 2, 2); // 第4列合并（绩效数据）

HeaderConfig header2 = new HeaderConfig()
    .addColumnNames("", "姓名", "部门", "绩效评分", "备注")
    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

// 配置到Sheet
sheet.setHeaders(Arrays.asList(header1, header2));
```

---

### 1.3 ColumnConfig

列配置类，定义每列的属性、样式、数据转换规则。

#### 属性列表

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `fieldName` | String | 必填 | 字段名（实体类字段或Map的key） |
| `width` | int | 20 | 列宽（字符数，实际宽度 = 字符数 × 256） |
| `styleConfig` | CellStyleConfig | null | 列样式（静态） |
| `styleProvider` | StyleProvider | null | 动态样式提供器（优先级高于styleConfig） |
| `numberFormat` | String | null | 数字格式（如：`#,##0.00`、`0.00%`） |
| `dateFormat` | String | null | 日期格式（如：`yyyy-MM-dd`） |
| `hidden` | boolean | false | 是否隐藏该列 |
| `dataConverter` | DataConverter | null | 数据转换器 |

#### 使用示例

```java
ColumnConfig salaryColumn = new ColumnConfig()
    .setFieldName("salary")           // 字段名
    .setWidth(15)                    // 列宽
    .setNumberFormat("#,##0")         // 千分位格式
    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()) // 静态样式
    .setStyleProvider((value, rowData, rowIndex) -> {  // 动态样式
        if (((Number) value).doubleValue() > 20000) {
            return StyleTemplate.DATA.toStyleConfig()
                .setFontColor(IndexedColors.YELLOW.getIndex())
                .setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128));
        }
        return StyleTemplate.DATA.toStyleConfig();
    })
    .setDataConverter((value, rowData, rowIndex) -> {  // 数据转换
        return value == null ? 0 : value;
    });
```

---

### 1.4 CellStyleConfig

单元格样式配置类，定义字体、边框、对齐、背景等所有样式属性。

#### 属性列表

| 类别 | 属性 | 类型 | 默认值 | 说明 |
|------|------|------|--------|------|
| **颜色** | `backgroundColor` | short | -1 | 背景色（IndexedColors索引） |
| | `rgbBackgroundColor` | byte[] | null | RGB背景色（优先级高于backgroundColor） |
| | `fontColor` | short | -1 | 字体颜色 |
| **字体** | `bold` | boolean | false | 是否加粗 |
| | `italic` | boolean | false | 是否倾斜 |
| | `strikeout` | boolean | false | 是否删除线 |
| | `underline` | byte | 0 | 下划线（0=无，1=单下划线，2=双下划线） |
| | `fontSize` | short | 11 | 字体大小（单位：磅） |
| | `fontName` | String | "微软雅黑" | 字体名称 |
| **对齐** | `horizontalAlignment` | HorizontalAlignment | CENTER | 水平对齐方式 |
| | `verticalAlignment` | VerticalAlignment | CENTER | 垂直对齐方式 |
| **边框** | `borderTop` | BorderStyle | THIN | 上边框 |
| | `borderBottom` | BorderStyle | THIN | 下边框 |
| | `borderLeft` | BorderStyle | THIN | 左边框 |
| | `borderRight` | BorderStyle | THIN | 右边框 |
| **其他** | `wrapText` | boolean | false | 是否自动换行 |
| | `indentation` | short | 0 | 缩进 |
| | `rotation` | short | 0 | 旋转角度（-90到90度） |
| | `locked` | boolean | true | 是否锁定单元格 |
| | `hidden` | boolean | false | 是否隐藏公式 |

#### 使用示例

```java
CellStyleConfig style = new CellStyleConfig()
    .setFontColor(IndexedColors.WHITE.getIndex())
    .setRgbBackgroundColor(StyleTemplate.rgb(255, 0, 0))  // 红色RGB背景
    .setBold(true)
    .setFontSize((short) 12)
    .setFontName("宋体")
    .setHorizontalAlignment(HorizontalAlignment.CENTER)
    .setVerticalAlignment(VerticalAlignment.CENTER)
    .setAllBorders(BorderStyle.THIN)  // 四边细边框
    .setWrapText(true);
```

---

### 1.5 StyleTemplate

样式模板枚举，提供预定义的常用样式，避免重复配置。

#### 枚举列表

| 模板 | 说明 | 背景色 | 字体色 |
|------|------|--------|--------|
| `HEADER` | 表头样式 | 红色 | 白色、加粗、居中 |
| `DATA` | 数据样式 | 白色 | 黑色、居中、边框 |

#### 使用方法

```java
// 1. 直接使用模板
CellStyleConfig headerStyle = StyleTemplate.HEADER.toStyleConfig();

// 2. 基于模板修改（深拷贝，不影响模板本身）
CellStyleConfig customHeader = StyleTemplate.HEADER.toStyleConfig()
    .setRgbBackgroundColor(StyleTemplate.rgb(0, 255, 0));  // 绿色背景

// 3. 使用RGB颜色
setRgbBackgroundColor(StyleTemplate.rgb(255, 0, 0))  // 红色
setRgbBackgroundColor(StyleTemplate.rgb(0, 255, 0))  // 绿色
setRgbBackgroundColor(StyleTemplate.rgb(0, 0, 255))  // 蓝色
setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128))  // 灰色
```

#### 重要说明

- `toStyleConfig()` 返回的是**深拷贝**，修改返回值不影响模板本身
- `rgb()` 返回 `byte[]`，需要配合 `setRgbBackgroundColor()` 使用
- 优先级：`rgbBackgroundColor` > `backgroundColor`

---

## 2. 核心工具类

### 2.1 ExcelExportUtils

核心导出工具类，提供多种导出方式。

#### 方法列表

| 方法 | 说明 | 使用场景 |
|------|------|--------|
| `export(OutputStream, SheetConfig...)` | 导出到输出流 | 通用导出、文件导出 |
| `exportToResponse(HttpServletResponse, String, SheetConfig...)` | 导出到HTTP响应（自定义文件名） | Web导出，文件名自己填 |
| `exportToResponseWithTimestamp(HttpServletResponse, String, SheetConfig...)` | 导出到HTTP响应（自动加时间戳） | Web导出，避免文件名冲突 |
| `createSimpleTemplate(...)` | 创建简单模板配置 | 快速生成标准报表 |

---

#### 方法1：export（通用导出）

```java
/**
 * 导出Excel到输出流
 * @param outputStream 输出流
 * @param sheetConfigs Sheet配置（支持单个或多个）
 */
public static void export(OutputStream outputStream, SheetConfig<?>... sheetConfigs)
    throws ExcelExportException
```

**使用示例：**

```java
// 导出到文件
try (OutputStream os = new FileOutputStream("员工报表.xlsx")) {
    ExcelExportUtils.export(os, sheetConfig);
}

// 导出到多个Sheet
try (OutputStream os = new FileOutputStream("综合报表.xlsx")) {
    ExcelExportUtils.export(os, sheet1, sheet2, sheet3);
}
```

---

#### 方法2：exportToResponse（Web导出，自定义文件名）

```java
/**
 * 导出到HTTP响应
 * @param response HTTP响应
 * @param fileName 文件名（完整，如"员工报表.xlsx"）
 * @param sheetConfigs Sheet配置
 */
public static void exportToResponse(HttpServletResponse response, String fileName,
                                    SheetConfig<?>... sheetConfigs)
    throws ExcelExportException, IOException
```

**使用示例：**

```java
public void exportReport(HttpServletResponse response) {
    try {
        ExcelExportUtils.exportToResponse(
            response,
            "月度报表_v2.xlsx",  // 完整文件名
            buildSheetConfig()
        );
    } catch (ExcelExportException | IOException e) {
        response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
    }
}
```

---

#### 方法3：exportToResponseWithTimestamp（Web导出，自动加时间戳）

```java
/**
 * 导出到HTTP响应（自动加时间戳）
 * @param response HTTP响应
 * @param baseName 基础文件名（不含扩展名，如"员工报表"）
 * @param sheetConfigs Sheet配置
 * 文件名格式：{baseName}_yyyyMMdd_HHmmss.xlsx
 */
public static void exportToResponseWithTimestamp(HttpServletResponse response, String baseName,
                                                   SheetConfig<?>... sheetConfigs)
    throws ExcelExportException, IOException
```

**使用示例：**

```java
public void exportReport(HttpServletResponse response) {
    try {
        // 文件名自动变成：员工报表_20260322_170845.xlsx
        ExcelExportUtils.exportToResponseWithTimestamp(
            response,
            "员工报表",
            buildSheetConfig()
        );
    } catch (ExcelExportException | IOException e) {
        response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
    }
}
```

---

#### 方法4：createSimpleTemplate（快速创建标准报表）

```java
/**
 * 创建简单模板配置
 * @param sheetName Sheet名称
 * @param dataList 数据列表
 * @param fieldNames 字段名数组
 * @param columnNames 列名数组
 */
public static <T> SheetConfig<T> createSimpleTemplate(
    String sheetName, List<T> dataList,
    String[] fieldNames, String[] columnNames)
```

**使用示例：**

```java
SheetConfig<User> sheet = ExcelExportUtils.createSimpleTemplate(
    "员工报表",
    userList,
    new String[]{"seq", "name", "department", "salary"},
    new String[]{"序号", "姓名", "部门", "薪资"}
);

try (OutputStream os = new FileOutputStream("快速报表.xlsx")) {
    ExcelExportUtils.export(os, sheet);
}
```

---

### 2.2 DataConverter

数据转换器接口，用于在写入Excel前对原始数据进行自定义转换。

#### 接口定义

```java
@FunctionalInterface
public interface DataConverter {
    /**
     * 转换数据
     * @param value 原始值（通过getter/反射获取的字段值）
     * @param rowData 整行数据对象（可用于访问其他字段）
     * @param rowIndex 行索引（从0开始）
     * @return 转换后的值（会被设置到单元格中）
     */
    Object convert(Object value, Object rowData, int rowIndex);
}
```

#### 使用场景

| 场景 | 示例 |
|------|------|
| **枚举转中文** | `1 → "男"`, `0 → "女"` |
| **状态码转描述** | `PENDING → "待处理"`, `COMPLETED → "已完成"` |
| **数字格式化** | `0.85 → "85%"`, `12345.6 → "12,345.60"` |
| **日期格式化** | `2024-03-22 → "2024年03月22日"` |
| **多字段组合** | `firstName + lastName → "张三"` |
| **空值处理** | `null → "—"`, `null → "0"` |

#### 完整示例

```java
// 场景1：枚举值转中文
new ColumnConfig()
    .setFieldName("gender")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "未知";
        return ((Integer) value) == 1 ? "男" : "女";
    });

// 场景2：状态码转描述
new ColumnConfig()
    .setFieldName("status")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "";
        switch (value.toString()) {
            case "PENDING": return "待处理";
            case "DEVELOPING": return "开发中";
            case "COMPLETED": return "已完成";
            default: return value.toString();
        }
    });

// 场景3：小数转百分比
new ColumnConfig()
    .setFieldName("rate")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "0%";
        double rate = ((Number) value).doubleValue();
        return String.format("%.2f%%", rate * 100);
    });

// 场景4：日期格式化
new ColumnConfig()
    .setFieldName("createTime")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "";
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
        return sdf.format(value);
    });

// 场景5：多字段组合
new ColumnConfig()
    .setFieldName("fullName")
    .setDataConverter((value, rowData, rowIndex) -> {
        User user = (User) rowData;
        return user.getLastName() + " " + user.getFirstName();
    });

// 场景6：空值处理
new ColumnConfig()
    .setFieldName("remark")
    .setDataConverter((value, rowData, rowIndex) -> {
        return value == null ? "—" : value.toString();
    });
```

#### 注意事项

- 转换器返回的值可以是任意类型（Number、Date、String等）
- 转换器抛出异常会导致整个导出失败，包装为 `ExcelExportException`
- 转换器执行早于 `setCellValueSafe()`，影响后续类型判断
- 建议在转换器中处理 `null` 值，避免单元格为空

---

### 2.3 StyleProvider

动态样式提供器接口，根据单元格值、整行数据或行索引，动态决定样式。

#### 接口定义

```java
@FunctionalInterface
public interface StyleProvider {
    /**
     * 根据单元格值动态返回样式配置
     * @param value 原始字段值（未经DataConverter转换）
     * @param rowData 整行数据对象
     * @param rowIndex 行索引（从0开始）
     * @return 样式配置，返回null则使用列默认styleConfig
     */
    CellStyleConfig provide(Object value, Object rowData, int rowIndex);
}
```

#### 样式优先级

```
StyleProvider（动态）
    ↓ 返回null时降级
列 styleConfig（静态）
    ↓ 没有时降级
Sheet defaultDataStyle（兜底）
```

#### 使用场景

| 场景 | 示例 |
|------|------|
| **数值超阈值** | 薪资>20000 → 黄色字体+灰色背景 |
| **状态着色** | 异常→红色，正常→绿色 |
| **多字段判断** | 技术部 且 绩效>90 → 绿色高亮 |
| **奇偶行区分** | 偶数行浅蓝色背景（斑马纹） |

#### 完整示例

```java
// 场景1：薪资 > 20000 → 黄色字体 + 灰色背景
new ColumnConfig()
    .setFieldName("salary")
    .setStyleProvider((value, rowData, rowIndex) -> {
        if (value != null && ((Number) value).doubleValue() > 20000) {
            return StyleTemplate.DATA.toStyleConfig()
                .setFontColor(IndexedColors.YELLOW.getIndex())
                .setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128))
                .setBold(true);
        }
        return StyleTemplate.DATA.toStyleConfig();
    });

// 场景2：状态列根据值着色
new ColumnConfig()
    .setFieldName("status")
    .setStyleProvider((value, rowData, rowIndex) -> {
        CellStyleConfig base = StyleTemplate.DATA.toStyleConfig();
        if ("异常".equals(value)) {
            return base.setFontColor(IndexedColors.RED.getIndex()).setBold(true);
        } else if ("正常".equals(value)) {
            return base.setFontColor(IndexedColors.GREEN.getIndex());
        }
        return base;
    });

// 场景3：访问整行数据做多字段联合判断
new ColumnConfig()
    .setFieldName("score")
    .setStyleProvider((value, rowData, rowIndex) -> {
        Employee emp = (Employee) rowData;
        if ("技术部".equals(emp.getDepartment()) && value != null
                && ((Number) value).doubleValue() > 90) {
            return StyleTemplate.DATA.toStyleConfig()
                .setRgbBackgroundColor(StyleTemplate.rgb(198, 239, 206)) // 浅绿
                .setFontColor(IndexedColors.DARK_GREEN.getIndex());
        }
        return StyleTemplate.DATA.toStyleConfig();
    });

// 场景4：奇偶行区分（斑马纹）
ColumnConfig.StyleProvider zebraStyle = (value, rowData, rowIndex) -> {
    if (rowIndex % 2 == 0) {
        return StyleTemplate.DATA.toStyleConfig()
            .setRgbBackgroundColor(StyleTemplate.rgb(235, 245, 255)); // 浅蓝
    }
    return StyleTemplate.DATA.toStyleConfig();
};

// 对所有列统一应用斑马纹
columns.forEach(col -> col.setStyleProvider(zebraStyle));
```

#### 性能注意事项

| 注意点 | 说明 |
|------|------|
| 避免频繁new | 相同样式会通过**指纹缓存**自动复用，放心使用`toStyleConfig()` |
| 返回null可降级 | 返回null时自动降级到列的`styleConfig`，不会出错 |
| 大数据量 | 回调逻辑尽量简单，避免复杂计算（如数据库查询） |

#### 与DataConverter的区别

| | DataConverter | StyleProvider |
|--|-------------|---------------|
| **改变内容** | ✅ 改变**值** | ❌ 不改变值 |
| **改变样式** | ❌ 不改变样式 | ✅ 改变**样式** |
| **值时机** | 转换后的值 | **转换前**的原始值 |
| **可同时配置** | ✅ 可以 | ✅ 可以 |

---

## 3. 异常类

### ExcelExportException

Excel导出的统一异常类，包装所有导出过程中的异常。

#### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `message` | String | 异常描述 |
| `errorCode` | String | 错误代码（用于国际化或日志分析） |
| `cause` | Throwable | 原始异常 |

#### 错误代码列表

| 错误代码 | 说明 |
|---------|------|
| `NULL_OUTPUT_STREAM` | 输出流为空 |
| `EMPTY_SHEET_CONFIG` | Sheet配置为空 |
| `NULL_SHEET_CONFIG` | 第N个Sheet配置为空 |
| `EMPTY_COLUMN_CONFIG` | 列配置为空 |
| `DATA_EXCEED_LIMIT` | 数据量超过最大行数限制 |
| `COLUMN_EXCEED_LIMIT` | 列数超过最大列数限制 |
| `EMPTY_FIELD_NAME` | 字段名称为空 |
| `FIELD_NOT_FOUND` | 字段不存在于类中 |
| `FIELD_ACCESS_DENIED` | 字段访问权限不足 |
| `DATA_CONVERT_ERROR` | 数据转换失败 |
| `STYLE_PROVIDER_ERROR` | 动态样式计算失败 |
| `MERGE_ERROR` | 合并单元格失败 |
| `FLUSH_ERROR` | 刷新数据失败 |
| `EXPORT_ERROR` | 导出失败 |

#### 使用示例

```java
try {
    ExcelExportUtils.exportToResponse(response, "报表.xlsx", configs);
} catch (ExcelExportException e) {
    response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
    System.err.println("错误代码: " + e.getErrorCode());
    System.err.println("错误信息: " + e.getMessage());
    e.printStackTrace();
}
```

---

## 4. 完整使用示例

### 4.1 基础单Sheet导出

```java
public class BasicExportExample {
    public static void main(String[] args) throws Exception {
        // 1. 准备数据
        List<User> dataList = Arrays.asList(
            new User("张三", "技术部", 15000),
            new User("李四", "销售部", 12000),
            new User("王五", "人事部", 8000)
        );

        // 2. 配置列
        List<ColumnConfig> columns = new ArrayList<>();
        columns.add(new ColumnConfig().setFieldName("name").setWidth(12)
            .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("dept").setWidth(12)
            .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("salary").setWidth(15)
            .setNumberFormat("#,##0")
            .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

        // 3. 配置表头
        HeaderConfig header = new HeaderConfig()
            .addColumnNames("姓名", "部门", "薪资")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        // 4. 配置Sheet
        SheetConfig<User> sheet = new SheetConfig<User>()
            .setSheetName("员工报表")
            .setHeaders(Collections.singletonList(header))
            .setColumnConfigs(columns)
            .setDataList(dataList)
            .setFreezeRow(1)
            .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
            .setBatchSize(100);

        // 5. 导出
        try (OutputStream os = new FileOutputStream("员工报表.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }

        System.out.println("导出成功！");
    }
}
```

---

### 4.2 多Sheet导出

```java
public class MultiSheetExample {
    public static void main(String[] args) throws Exception {
        // 1. 准备三个Sheet的数据
        SheetConfig<User> sheet1 = buildEmployeeSheet();
        SheetConfig<Map<String, Object>> sheet2 = buildTechSheet();
        SheetConfig<Map<String, Object>> sheet3 = buildManagerSheet();

        // 2. 导出到一个Excel文件
        try (OutputStream os = new FileOutputStream("综合报表.xlsx")) {
            ExcelExportUtils.export(os, sheet1, sheet2, sheet3);
        }

        System.out.println("多Sheet导出成功！");
    }
}
```

---

### 4.3 多级表头

```java
public class MultiLevelHeaderExample {
    public static void main(String[] args) throws Exception {
        List<User> dataList = generateData(50);

        // 一级表头（第1行）
        HeaderConfig header1 = new HeaderConfig()
            .addColumnNames("序号", "基本信息", "基本信息", "绩效数据", "备注")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
            .addMergeRegion(0, 0, 1, 2)  // 第2-3列合并
            .addMergeRegion(0, 0, 3, 3); // 第4列合并

        // 二级表头（第2行）
        HeaderConfig header2 = new HeaderConfig()
            .addColumnNames("", "姓名", "部门", "评分", "备注")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        // 配置Sheet
        SheetConfig<User> sheet = new SheetConfig<User>()
            .setSheetName("多级表头报表")
            .setHeaders(Arrays.asList(header1, header2))
            .setColumnConfigs(buildColumns())
            .setDataList(dataList)
            .setFreezeRow(2);  // 冻结前2行

        try (OutputStream os = new FileOutputStream("多级表头.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }
    }
}
```

---

### 4.4 DataConverter 数据转换

```java
public class DataConverterExample {
    public static void main(String[] args) throws Exception {
        List<Map<String, Object>> dataList = Arrays.asList(
            Map.of("name", "张三", "gender", 1, "status", "PENDING"),
            Map.of("name", "李四", "gender", 0, "status", "COMPLETED")
        );

        List<ColumnConfig> columns = new ArrayList<>();

        // 列1：性别枚举转中文
        columns.add(new ColumnConfig()
            .setFieldName("gender")
            .setDataConverter((value, rowData, rowIndex) -> {
                if (value == null) return "未知";
                return ((Integer) value) == 1 ? "男" : "女";
            }));

        // 列2：状态码转描述
        columns.add(new ColumnConfig()
            .setFieldName("status")
            .setDataConverter((value, rowData, rowIndex) -> {
                if (value == null) return "";
                switch (value.toString()) {
                    case "PENDING": return "待处理";
                    case "COMPLETED": return "已完成";
                    default: return value.toString();
                }
            }));

        // 列3：姓名
        columns.add(new ColumnConfig().setFieldName("name"));

        HeaderConfig header = new HeaderConfig()
            .addColumnNames("性别", "状态", "姓名")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        SheetConfig<Map<String, Object>> sheet = new SheetConfig<>()
            .setSheetName("数据转换演示")
            .setHeaders(Collections.singletonList(header))
            .setColumnConfigs(columns)
            .setDataList(dataList);

        try (OutputStream os = new FileOutputStream("数据转换.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }
    }
}
```

---

### 4.5 StyleProvider 条件样式

```java
public class StyleProviderExample {
    public static void main(String[] args) throws Exception {
        List<User> dataList = generateData(100);

        // 列1：薪资 > 20000 高亮
        ColumnConfig salaryColumn = new ColumnConfig()
            .setFieldName("salary")
            .setWidth(15)
            .setNumberFormat("#,##0")
            .setStyleProvider((value, rowData, rowIndex) -> {
                if (value != null && ((Number) value).doubleValue() > 20000) {
                    return StyleTemplate.DATA.toStyleConfig()
                        .setFontColor(IndexedColors.YELLOW.getIndex())
                        .setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128))
                        .setBold(true);
                }
                return StyleTemplate.DATA.toStyleConfig();
            });

        // 列2：绩效分数着色
        ColumnConfig scoreColumn = new ColumnConfig()
            .setFieldName("score")
            .setWidth(12)
            .setNumberFormat("0.00")
            .setStyleProvider((value, rowData, rowIndex) -> {
                if (value != null) {
                    double score = ((Number) value).doubleValue();
                    CellStyleConfig base = StyleTemplate.DATA.toStyleConfig();
                    if (score < 70) {
                        return base.setFontColor(IndexedColors.RED.getIndex()).setBold(true);
                    } else if (score >= 90) {
                        return base.setFontColor(IndexedColors.DARK_GREEN.getIndex())
                            .setRgbBackgroundColor(StyleTemplate.rgb(198, 239, 206));
                    }
                }
                return StyleTemplate.DATA.toStyleConfig();
            });

        // 列3：奇偶行斑马纹
        ColumnConfig.StyleProvider zebraStyle = (value, rowData, rowIndex) -> {
            if (rowIndex % 2 == 0) {
                return StyleTemplate.DATA.toStyleConfig()
                    .setRgbBackgroundColor(StyleTemplate.rgb(235, 245, 255));
            }
            return StyleTemplate.DATA.toStyleConfig();
        };

        List<ColumnConfig> columns = new ArrayList<>();
        columns.add(new ColumnConfig().setFieldName("name")
            .setStyleProvider(zebraStyle));
        columns.add(scoreColumn);
        columns.add(salaryColumn);

        HeaderConfig header = new HeaderConfig()
            .addColumnNames("姓名", "绩效", "薪资")
            .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        SheetConfig<User> sheet = new SheetConfig<User>()
            .setSheetName("条件样式演示")
            .setHeaders(Collections.singletonList(header))
            .setColumnConfigs(columns)
            .setDataList(dataList)
            .setFreezeRow(1);

        try (OutputStream os = new FileOutputStream("条件样式.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }
    }
}
```

---

### 4.6 Web 场景导出

```java
@RestController
@RequestMapping("/api/export")
public class ExportController {

    /**
     * 方式1：自定义文件名导出
     */
    @GetMapping("/fixed-name")
    public void exportFixedName(HttpServletResponse response) {
        try {
            ExcelExportUtils.exportToResponse(
                response,
                "月度报表_v2.xlsx",  // 固定文件名
                buildSheetConfig()
            );
        } catch (ExcelExportException | IOException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
    }

    /**
     * 方式2：自动加时间戳导出
     */
    @GetMapping("/timestamp")
    public void exportWithTimestamp(HttpServletResponse response) {
        try {
            // 文件名自动变成：员工报表_20260322_170845.xlsx
            ExcelExportUtils.exportToResponseWithTimestamp(
                response,
                "员工报表",
                buildSheetConfig()
            );
        } catch (ExcelExportException | IOException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
    }

    /**
     * 方式3：多Sheet导出
     */
    @GetMapping("/multi")
    public void exportMulti(HttpServletResponse response) {
        try {
            List<SheetConfig<?>> configs = Arrays.asList(
                buildEmployeeSheet(),
                buildTechSheet(),
                buildManagerSheet()
            );

            ExcelExportUtils.exportToResponseWithTimestamp(
                response,
                "综合报表",
                configs.toArray(new SheetConfig[0])
            );
        } catch (ExcelExportException | IOException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
    }

    // 构建Sheet配置的方法省略...
}
```

---

## 5. 样式缓存机制

### 缓存原理

工具类使用**样式指纹机制**，相同样式的配置会自动复用同一个 POI `CellStyle` 对象，避免内存溢出。

### 样式指纹生成

```java
private static String generateStyleFingerprint(String cacheKey, CellStyleConfig config) {
    StringBuilder sb = new StringBuilder(cacheKey);
    sb.append("|bg:").append(config.getBackgroundColor());
    sb.append("|rgb:").append(config.getRgbBackgroundColor()); // RGB颜色
    sb.append("|fc:").append(config.getFontColor());
    sb.append("|bold:").append(config.isBold());
    sb.append("|fs:").append(config.getFontSize());
    sb.append("|fn:").append(config.getFontName());
    sb.append("|ha:").append(config.getHorizontalAlignment());
    sb.append("|va:").append(config.getVerticalAlignment());
    sb.append("|bt:").append(config.getBorderTop());
    sb.append("|bb:").append(config.getBorderBottom());
    sb.append("|bl:").append(config.getBorderLeft());
    sb.append("|br:").append(config.getBorderRight());
    sb.append("|wrap:").append(config.isWrapText());
    // ... 其他属性
    return sb.toString();
}
```

### 缓存键格式

```
header_0|bg:10|fc:9|bold:true|fs:11|fn:微软雅黑|...
data_salary|rgb:128,128,128|fc:10|bold:false|...
```

### 优势

| 优势 | 说明 |
|------|------|
| **内存优化** | 相同样式只创建一次 POI Style 对象 |
| **性能提升** | 避免重复创建样式，减少 GC 压力 |
| **自动复用** | 开发者无需手动管理缓存，框架自动处理 |

### 注意事项

- `StyleProvider` 返回的样式也会自动纳入缓存
- 使用 `StyleTemplate.toStyleConfig()` 会自动生成指纹
- 不同样式的配置不会相互干扰

---

## 6. 性能优化建议

### 6.1 批量刷新

```java
// 小数据量（< 1000行）：100-500
.setBatchSize(100)

// 大数据量（> 10000行）：1000-5000
.setBatchSize(2000)
```

### 6.2 列宽设置

| 数据类型 | 推荐宽度 |
|---------|---------|
| 序号 | 8-10 |
| 文本 | 12-15 |
| 金额 | 18-20 |
| 日期 | 14-16 |
| 百分比 | 10-12 |
| 长文本 | 20-30 |

### 6.3 避免重复创建对象

```java
// ✅ 推荐：使用模板深拷贝
.setStyleProvider((value, rowData, rowIndex) -> {
    return StyleTemplate.DATA.toStyleConfig()  // 每次都会自动复用
        .setFontColor(IndexedColors.RED.getIndex());
})

// ❌ 不推荐：每次都 new CellStyleConfig
.setStyleProvider((value, rowData, rowIndex) -> {
    return new CellStyleConfig()
        .setFontColor(IndexedColors.RED.getIndex());
})
```

### 6.4 数据转换优化

```java
// ✅ 推荐：在 DataConverter 中处理格式
.setDataConverter((value, rowData, rowIndex) -> {
    return value == null ? "—" : value.toString();
})

// ❌ 不推荐：依赖 Excel 自动格式
// 可能导致显示不符合预期
```

---

## 7. 常见问题

### Q1：导出的Excel打开后样式丢失？

**原因**：未设置 `styleConfig` 或返回 `null`。

**解决**：

```java
// 列级别
.setColumnConfig(column
    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()))

// Sheet级别
.setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
```

---

### Q2：自定义RGB颜色不生效？

**原因**：使用了 `setBackgroundColor(short)` 而不是 `setRgbBackgroundColor(byte[])`。

**解决**：

```java
// ❌ 错误
.setBackgroundColor(StyleTemplate.rgb(255, 0, 0))  // short溢出

// ✅ 正确
.setRgbBackgroundColor(StyleTemplate.rgb(255, 0, 0))
```

---

### Q3：StyleProvider 修改了一个影响所有列？

**原因**：返回的是共享引用。

**解决**：

```java
// StyleTemplate.toStyleConfig() 已经是深拷贝，无需担心
.setStyleProvider((value, rowData, rowIndex) -> {
    return StyleTemplate.DATA.toStyleConfig()  // 每次都是新对象
        .setFontColor(IndexedColors.RED.getIndex());
})
```

---

### Q4：大数据导出内存溢出？

**原因**：一次性加载所有数据到内存。

**解决**：

```java
// 1. 使用 SXSSFWorkbook（已内置）
// 2. 设置合理的 batchSize
.setBatchSize(2000)

// 3. 分页查询数据（每次只查一批）
for (int page = 0; page < totalPages; page++) {
    List<Data> batch = queryByPage(page, pageSize);
    dataList.addAll(batch);
    // 批量刷新会自动释放内存
}
```

---

### Q5：Web导出文件名乱码？

**原因**：未正确编码文件名。

**解决**：

```java
// 使用 exportToResponse 自动处理（已内置 URLEncoder）
ExcelExportUtils.exportToResponse(response, "中文文件名.xlsx", configs);

// 手动处理时
response.setHeader("Content-Disposition",
    "attachment; filename=" + URLEncoder.encode("中文文件名.xlsx", "UTF-8"));
```

---

### Q6：合并单元格后数据错位？

**原因**：`MergeRegion` 的行列索引相对于当前表头行。

**解决**：

```java
// 正确：相对于当前表头行
.addMergeRegion(0, 0, 1, 2)  // 当前表头行第0行，合并第1-2列

// 错误：使用绝对行索引
.addMergeRegion(1, 1, 1, 2)  // 错误
```

---

### Q7：DataConverter 和 StyleProvider 冲突吗？

**答案**：不冲突，互不影响。

- `DataConverter` 改变的是**值**
- `StyleProvider` 改变的是**样式**
- 可同时配置，独立执行

---

### Q8：如何隐藏某些列？

**方法**：

```java
new ColumnConfig()
    .setFieldName("hiddenField")
    .setHidden(true)  // 隐藏该列
```

---

### Q9：冻结行列不生效？

**原因**：`freezeRow` 和 `freezeCol` 设置错误。

**解决**：

```java
// 冻结前2行（两个表头行）
.setFreezeRow(2)

// 冻结前1列（序号列）
.setFreezeCol(1)
```

---

### Q10：导出速度太慢？

**排查**：

1. **检查 `batchSize`**：是否设置太小（如 10）
2. **检查 DataConverter**：是否有复杂计算或数据库查询
3. **检查数据量**：是否超过单次导出合理范围
4. **检查网络**：Web 场景下网络带宽是否足够

---

## 附录

### A. 依赖配置

```xml
<dependencies>
    <!-- Apache POI -->
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

    <!-- Commons 工具 -->
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-lang3</artifactId>
        <version>3.12.0</version>
    </dependency>
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-collections4</artifactId>
        <version>4.4</version>
    </dependency>

    <!-- Lombok -->
    <dependency>
        <groupId>org.projectlombok</groupId>
        <artifactId>lombok</artifactId>
        <version>1.18.20</version>
        <scope>provided</scope>
    </dependency>
</dependencies>
```

### B. POI 版本建议

当前版本：POI 4.1.2

建议升级至：**POI 5.2.5**

升级优势：
- 性能提升 20%~30%
- 内存占用降低 15%
- 更多 API 支持
- Bug 修复

### C. 版本历史

| 版本 | 日期 | 主要变更 |
|------|------|--------|
| v3.0.2 | 2026-03-22 | 新增 `StyleProvider` 动态样式接口；新增 Web 导出方法 |
| v3.0.1 | 2026-03-22 | 修复 `toStyleConfig()` 共享引用 Bug；修复 `rgb()` 颜色溢出 Bug |
| v3.0.0 | 2026-03-22 | 移除 RowStyleProvider；移除斑马纹样式；优化样式缓存；新增样式模板 |

---

**文档结束**

如有问题，请参考 `README.md` 或源码注释。
