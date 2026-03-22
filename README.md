# Excel导出工具使用文档

## 📋 目录

- [简介](#简介)
- [快速开始](#快速开始)
- [核心概念](#核心概念)
- [详细使用指南](#详细使用指南)
  - [基础配置](#基础配置)
  - [样式配置](#样式配置)
  - [多级表头](#多级表头)
  - [单元格合并](#单元格合并)
  - [数据格式化](#数据格式化)
  - [数据转换](#数据转换)
  - [冻结行列](#冻结行列)
  - [批量导出](#批量导出)
- [样式模板](#样式模板)
- [最佳实践](#最佳实践)
- [常见问题](#常见问题)
- [性能优化](#性能优化)

---

## 简介

### 功能特性

Excel导出工具是一个企业级、功能完整的Excel导出解决方案,具有以下核心特性:

- ✅ **流式写入机制**: 使用`SXSSFWorkbook`支持百万级数据导出,内存占用低
- ✅ **样式缓存优化**: 智能样式缓存机制,避免重复创建样式对象,提升性能
- ✅ **多种数据源**: 支持实体类和Map两种数据源
- ✅ **完善的异常处理**: 中文错误信息,精确定位到具体行列
- ✅ **丰富的样式支持**: 表头样式、数据样式、冻结行列、自动筛选等
- ✅ **类型安全**: 提供预定义样式模板,避免手写配置错误
- ✅ **灵活的扩展**: 支持自定义数据转换、格式化等

### 技术栈

- **核心依赖**: Apache POI 4.1.2
- **Java版本**: JDK 8+
- **样式管理**: Lombok + Builder模式

### 性能指标

- 支持最大行数: 1,048,576行
- 支持最大列数: 16,384列
- 单Sheet导出速度: ~1000行/秒(普通配置)
- 内存占用: 流式写入,窗口大小100行,内存占用可控

---

## 快速开始

### Maven依赖

```xml
<dependencies>
    <!-- Apache POI核心依赖 -->
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
    
    <!-- Lombok依赖 - 简化代码 -->
    <dependency>
        <groupId>org.projectlombok</groupId>
        <artifactId>lombok</artifactId>
        <version>1.18.30</version>
        <scope>provided</scope>
    </dependency>
    
    <!-- Commons Lang3 - 常用工具类 -->
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-lang3</artifactId>
        <version>3.14.0</version>
    </dependency>
    
    <!-- Commons Collections - 集合工具 -->
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-collections4</artifactId>
        <version>4.4</version>
    </dependency>
</dependencies>
```

### 最简示例

```java
import com.example.excel.config.SheetConfig;
import com.example.excel.utils.ExcelExportUtils;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

public class QuickStart {
    public static void main(String[] args) {
        // 1. 准备数据
        List<User> dataList = Arrays.asList(
            new User(1, "张三", "技术部", "工程师"),
            new User(2, "李四", "销售部", "经理")
        );
        
        // 2. 创建简单配置(使用工具方法)
        SheetConfig<User> config = ExcelExportUtils.createSimpleTemplate(
            "用户信息",           // Sheet名称
            dataList,            // 数据列表
            new String[]{"id", "name", "department", "position"},  // 字段名
            new String[]{"ID", "姓名", "部门", "职位"}             // 列名
        );
        
        // 3. 导出
        try (OutputStream os = new FileOutputStream("用户信息.xlsx")) {
            ExcelExportUtils.exportSingleSheet(os, config);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
```

---

## 核心概念

### 配置类层次结构

```
SheetConfig<T>          // Sheet配置(泛型T为数据类型)
├── HeaderConfig        // 表头配置(支持多级)
│   ├── CellStyleConfig // 单元格样式配置
│   └── MergeRegion    // 合并区域配置
└── ColumnConfig        // 列配置
    ├── CellStyleConfig // 列样式配置
    ├── DataConverter   // 数据转换器接口
    └── StyleProvider   // 动态样式提供器接口
```

### 核心类说明

| 类名 | 说明 | 主要职责 |
|------|------|----------|
| `ExcelExportUtils` | 工具类 | 提供导出方法、快速配置方法 |
| `SheetConfig<T>` | Sheet配置 | 定义单个Sheet的所有配置 |
| `HeaderConfig` | 表头配置 | 定义表头名称、样式、合并区域 |
| `ColumnConfig` | 列配置 | 定义列的字段映射、宽度、格式 |
| `CellStyleConfig` | 样式配置 | 定义单元格的字体、颜色、对齐等 |
| `StyleTemplate` | 样式模板 | 预定义的常用样式,类型安全 |
| `ExcelExportException` | 异常类 | 统一的导出异常,带错误码 |

---

## 详细使用指南

### 基础配置

#### 1. 实体类数据源

```java
// 定义实体类
public class User {
    private Integer id;
    private String name;
    private String department;
    private String position;
    
    // getter/setter省略...
}

// 导出
List<User> dataList = getUserList();  // 获取数据

SheetConfig<User> config = new SheetConfig<User>()
    .setSheetName("用户表")
    .setColumnConfigs(Arrays.asList(
        new ColumnConfig().setFieldName("id").setWidth(10),
        new ColumnConfig().setFieldName("name").setWidth(15),
        new ColumnConfig().setFieldName("department").setWidth(15),
        new ColumnConfig().setFieldName("position").setWidth(15)
    ))
    .setDataList(dataList);

ExcelExportUtils.exportSingleSheet(outputStream, config);
```

#### 2. Map数据源

```java
// 使用Map作为数据源
List<Map<String, Object>> dataList = new ArrayList<>();
Map<String, Object> row1 = new HashMap<>();
row1.put("id", 1);
row1.put("name", "张三");
row1.put("department", "技术部");
dataList.add(row1);

SheetConfig<Map<String, Object>> config = new SheetConfig<Map<String, Object>>()
    .setSheetName("用户表")
    .setColumnConfigs(Arrays.asList(
        new ColumnConfig().setFieldName("id").setWidth(10),
        new ColumnConfig().setFieldName("name").setWidth(15),
        new ColumnConfig().setFieldName("department").setWidth(15)
    ))
    .setDataList(dataList);
```

### 样式配置

#### 1. 使用样式模板(推荐)

```java
import com.example.excel.config.StyleTemplate;

// 表头使用蓝色模板
HeaderConfig header = new HeaderConfig()
    .addColumnNames("ID", "姓名", "部门")
    .setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig());

// 列样式
ColumnConfig column = new ColumnConfig()
    .setFieldName("name")
    .setWidth(15)
    .setStyleConfig(StyleTemplate.DATA_LEFT.toStyleConfig());
```

#### 2. 自定义样式

```java
CellStyleConfig customStyle = new CellStyleConfig()
    .setBackgroundColor(IndexedColors.YELLOW.getIndex())  // 黄色背景
    .setFontColor(IndexedColors.BLACK.getIndex())        // 黑色字体
    .setBold(true)                                         // 加粗
    .setFontSize((short) 12)                              // 字号12
    .setFontName("微软雅黑")                              // 字体
    .setHorizontalAlignment(HorizontalAlignment.CENTER)    // 水平居中
    .setVerticalAlignment(VerticalAlignment.CENTER)        // 垂直居中
    .setAllBorders(BorderStyle.THIN);                     // 四边细边框

HeaderConfig header = new HeaderConfig()
    .addColumnNames("ID", "姓名")
    .setStyleConfig(customStyle);
```

#### 3. 基于模板修改样式

```java
// 复制模板并修改
CellStyleConfig customStyle = StyleTemplate.HEADER_BLUE.copyStyleConfig()
    .setBackgroundColor(IndexedColors.RED.getIndex());  // 改为红色背景

HeaderConfig header = new HeaderConfig()
    .addColumnNames("ID", "姓名")
    .setStyleConfig(customStyle);
```

### 多级表头

```java
// 一级表头
HeaderConfig header1 = new HeaderConfig()
    .addColumnNames("基本信息", "基本信息", "基本信息", "绩效数据", "绩效数据")
    .setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig())
    .addMergeRegion(0, 0, 0, 2)   // 第1-3列合并(基本信息)
    .addMergeRegion(0, 0, 3, 4);  // 第4-5列合并(绩效数据)

// 二级表头
HeaderConfig header2 = new HeaderConfig()
    .addColumnNames("姓名", "部门", "职位", "评分", "项目数")
    .setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig());

SheetConfig<User> config = new SheetConfig<User>()
    .setSheetName("员工报表")
    .setHeaders(Arrays.asList(header1, header2))  // 多级表头
    .setColumnConfigs(columns)
    .setDataList(dataList);
```

**效果**:
```
┌─────────────────────────────────────────┐
│     基本信息         │  绩效数据  │
├──────┬──────┬──────┼────┬──────┤
│ 姓名 │ 部门 │ 职位 │评分 │项目数 │
├──────┼──────┼──────┼────┼──────┤
│ 张三 │技术部│工程师│ 95 │   5  │
└──────┴──────┴──────┴────┴──────┘
```

### 单元格合并

```java
HeaderConfig header = new HeaderConfig()
    .addColumnNames("汇总", "Q1", "Q2", "Q3", "Q4")
    .addMergeRegion(0, 0, 1, 4);  // 横向合并:第1行,列1-4合并

// 多区域合并
.addMergeRegion(0, 1, 0, 0)    // 纵向合并:行0-1,列0
.addMergeRegion(0, 0, 1, 4);   // 横向合并:行0,列1-4
```

### 数据格式化

#### 数字格式

```java
// 保留2位小数
new ColumnConfig()
    .setFieldName("price")
    .setNumberFormat("0.00");

// 千分位
new ColumnConfig()
    .setFieldName("salary")
    .setNumberFormat("#,##0");

// 百分比
new ColumnConfig()
    .setFieldName("rate")
    .setNumberFormat("0.00%");

// 科学计数法
new ColumnConfig()
    .setFieldName("value")
    .setNumberFormat("0.00E+00");
```

#### 日期格式

```java
// 默认格式:yyyy-MM-dd HH:mm:ss
new ColumnConfig()
    .setFieldName("createTime")
    .setDateFormat("yyyy-MM-dd HH:mm:ss");

// 自定义格式
new ColumnConfig()
    .setFieldName("hireDate")
    .setDateFormat("yyyy年MM月dd日");
```

### 数据转换

DataConverter是一个函数式接口,用于在数据写入Excel前进行自定义转换。

#### 场景1: 枚举值转中文

```java
new ColumnConfig()
    .setFieldName("gender")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "未知";
        return ((Integer) value) == 1 ? "男" : "女";
    });
```

#### 场景2: 状态码转描述

```java
new ColumnConfig()
    .setFieldName("status")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "";
        String status = value.toString();
        switch (status) {
            case "PENDING": return "待处理";
            case "PROCESSING": return "处理中";
            case "COMPLETED": return "已完成";
            case "FAILED": return "失败";
            default: return status;
        }
    });
```

#### 场景3: 小数转百分比

```java
new ColumnConfig()
    .setFieldName("rate")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "0%";
        double rate = ((Number) value).doubleValue();
        return String.format("%.2f%%", rate * 100);
    });
```

#### 场景4: 日期格式化

```java
new ColumnConfig()
    .setFieldName("createTime")
    .setDataConverter((value, rowData, rowIndex) -> {
        if (value == null) return "";
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日 HH:mm");
        return sdf.format(value);
    });
```

#### 场景5: 多字段组合

```java
new ColumnConfig()
    .setFieldName("fullName")
    .setDataConverter((value, rowData, rowIndex) -> {
        User user = (User) rowData;
        return user.getLastName() + " " + user.getFirstName();
    });
```

#### 场景6: 空值处理

```java
new ColumnConfig()
    .setFieldName("remark")
    .setDataConverter((value, rowData, rowIndex) -> {
        return value == null ? "—" : value.toString();
    });
```

---

### 条件样式（StyleProvider）

`StyleProvider` 是一个函数式接口，用于根据**单元格的值**、**整行数据**或**行索引**，在运行时动态决定该列每个单元格的样式。

**优先级**：`StyleProvider`（动态）> 列 `styleConfig`（静态）> Sheet 默认数据样式

> **与 DataConverter 的区别**
> - `DataConverter` 改变的是**值**（显示内容）
> - `StyleProvider` 改变的是**样式**（颜色、字体等），两者可同时配置、互不影响
> - `StyleProvider` 拿到的是 DataConverter **执行之前**的原始值

#### 场景1：数值超阈值高亮（薪资 > 20000 → 黄色字体 + 灰色背景）

```java
new ColumnConfig()
    .setFieldName("salary")
    .setWidth(18)
    .setStyleProvider((value, rowData, rowIndex) -> {
        if (value != null && ((Number) value).doubleValue() > 20000) {
            // 满足条件：黄色字体 + 灰色背景
            return StyleTemplate.DATA.toStyleConfig()
                .setFontColor(IndexedColors.YELLOW.getIndex())
                .setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128));
        }
        // 不满足条件：使用默认数据样式
        return StyleTemplate.DATA.toStyleConfig();
    });
```

#### 场景2：状态列根据值着色

```java
new ColumnConfig()
    .setFieldName("status")
    .setStyleProvider((value, rowData, rowIndex) -> {
        CellStyleConfig base = StyleTemplate.DATA.toStyleConfig();
        if ("异常".equals(value)) {
            return base.setFontColor(IndexedColors.RED.getIndex()).setBold(true);
        } else if ("正常".equals(value)) {
            return base.setFontColor(IndexedColors.GREEN.getIndex());
        }
        return base; // 其他状态默认样式
    });
```

#### 场景3：访问整行数据做多字段联合判断

```java
new ColumnConfig()
    .setFieldName("score")
    .setStyleProvider((value, rowData, rowIndex) -> {
        Employee emp = (Employee) rowData;
        // 技术部 且 分数 > 90 → 绿色高亮
        if ("技术部".equals(emp.getDept()) && value != null
                && ((Number) value).doubleValue() > 90) {
            return StyleTemplate.DATA.toStyleConfig()
                .setRgbBackgroundColor(StyleTemplate.rgb(198, 239, 206)) // 浅绿
                .setFontColor(IndexedColors.DARK_GREEN.getIndex());
        }
        return StyleTemplate.DATA.toStyleConfig();
    });
```

#### 场景4：奇偶行区分（斑马纹）

```java
// 对所有列统一应用斑马纹，只需在每列上设置同一个 StyleProvider
ColumnConfig.StyleProvider zebraStyle = (value, rowData, rowIndex) -> {
    if (rowIndex % 2 == 0) {
        return StyleTemplate.DATA.toStyleConfig()
            .setRgbBackgroundColor(StyleTemplate.rgb(235, 245, 255)); // 浅蓝偶数行
    }
    return StyleTemplate.DATA.toStyleConfig(); // 奇数行默认
};

columns.forEach(col -> col.setStyleProvider(zebraStyle));
```

#### ⚠️ 性能注意事项

| 注意点 | 说明 |
|--------|------|
| 避免在回调中频繁 new | 相同样式会通过**指纹缓存**自动复用，放心使用 `toStyleConfig()` |
| 返回 `null` 可降级 | 返回 `null` 时自动降级到列的 `styleConfig`，不会出错 |
| 大数据量 | 回调逻辑尽量简单，避免复杂计算（如数据库查询）|

```java
SheetConfig<User> config = new SheetConfig<User>()
    .setSheetName("员工报表")
    .setHeaders(Arrays.asList(header1, header2))
    .setColumnConfigs(columns)
    .setDataList(dataList)
    .setFreezeRow(2)  // 冻结前2行(两个表头行)
    .setFreezeCol(1); // 冻结前1列(序号列)
```

**效果**: 滚动时表头和第一列始终可见。

### 批量导出

```java
// 创建多个Sheet配置
List<SheetConfig<?>> sheetConfigs = new ArrayList<>();

SheetConfig<User> sheet1 = new SheetConfig<User>()
    .setSheetName("技术部")
    .setColumnConfigs(columns)
    .setDataList(techDeptUsers);

SheetConfig<User> sheet2 = new SheetConfig<User>()
    .setSheetName("销售部")
    .setColumnConfigs(columns)
    .setDataList(salesDeptUsers);

sheetConfigs.add(sheet1);
sheetConfigs.add(sheet2);

// 导出多个Sheet
ExcelExportUtils.exportMultiSheets(outputStream, sheetConfigs);
```

---

## 样式模板

StyleTemplate提供了预定义的常用样式,直接使用即可,无需手动配置。

### 表头样式

| 模板 | 说明 | 适用场景 |
|------|------|----------|
| `HEADER_BLUE` | 蓝色背景,白色字体,加粗 | 标准报表表头 |
| `HEADER_RED` | 红色背景,白色字体,加粗 | 需要突出的表头 |
| `HEADER_GREEN` | 绿色背景,白色字体,加粗 | 财务报表、汇总表 |
| `HEADER_GREY` | 灰色背景,黑色字体,加粗 | 日常报表、统计表 |

### 数据样式

| 模板 | 说明 | 适用场景 |
|------|------|----------|
| `DATA_CENTER` | 居中对齐 | 一般文本数据 |
| `DATA_LEFT` | 左对齐 | 姓名、地址等文本 |
| `DATA_RIGHT` | 右对齐 | 金额、数量等数值 |
| `DATA_ODD` | 白色背景 | 斑马纹-奇数行 |
| `DATA_EVEN` | 灰色背景 | 斑马纹-偶数行 |

### 特殊样式

| 模板 | 说明 | 适用场景 |
|------|------|----------|
| `EMPHASIS_ORANGE` | 橙色背景,白色字体,加粗 | 重要数据、警告信息 |
| `EMPHASIS_YELLOW` | 黄色背景,黑色字体,加粗 | 需要注意的数据 |
| `EMPHASIS_LIGHT_BLUE` | 浅蓝色背景,黑色字体,加粗 | 次级强调数据 |
| `SUMMARY` | 灰色背景,白色字体,加粗,右对齐 | 汇总行 |
| `NEGATIVE` | 白色背景,红色字体,加粗,右对齐 | 负数 |
| `PLAIN_TEXT` | 白色背景,无边框 | 备注、说明等文本 |

### 使用示例

```java
import com.example.excel.config.StyleTemplate;

// 直接使用
HeaderConfig header = new HeaderConfig()
    .addColumnNames("ID", "姓名")
    .setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig());

// 基于模板修改
CellStyleConfig customStyle = StyleTemplate.DATA_RIGHT.copyStyleConfig()
    .setBackgroundColor(IndexedColors.YELLOW.getIndex());

ColumnConfig column = new ColumnConfig()
    .setFieldName("price")
    .setWidth(12)
    .setNumberFormat("0.00")
    .setStyleConfig(customStyle);
```

---

## 最佳实践

### 1. 使用样式模板

**推荐**:
```java
.setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig())
```

**不推荐**:
```java
.setStyleConfig(new CellStyleConfig()
    .setBackgroundColor(IndexedColors.DARK_BLUE.getIndex())
    .setFontColor(IndexedColors.WHITE.getIndex())
    .setBold(true)
    // ... 手动配置10+个属性
)
```

### 2. 合理设置列宽

```java
// 推荐: 根据实际内容设置
.setWidth(15)  // 一般文本:12-15, 金额:18-20, 日期:14

// 不推荐: 过窄或过宽
.setWidth(5)   // 太窄,显示不全
.setWidth(50)  // 太宽,浪费空间
```

### 3. 使用批量刷新优化性能

```java
// 根据数据量调整批次大小
.setBatchSize(100);  // 小数据:100-500
.setBatchSize(1000); // 大数据:1000-5000
```

### 4. 合理使用数据转换

```java
// 推荐: 在数据转换器中处理格式
.setDataConverter((value, rowData, rowIndex) -> {
    return value == null ? "—" : value.toString();
});

// 不推荐: 在数据库层面处理所有格式
// 让工具负责格式化,保持数据原始性
```

### 5. 异常处理

```java
try (OutputStream os = new FileOutputStream(file)) {
    ExcelExportUtils.exportSingleSheet(os, config);
} catch (ExcelExportException e) {
    // 业务异常:参数错误、字段不存在等
    System.err.println("导出失败:" + e.getMessage());
    System.err.println("错误码:" + e.getErrorCode());
} catch (IOException e) {
    // IO异常:文件写入失败等
    System.err.println("文件操作失败:" + e.getMessage());
}
```

---

## 常见问题

### Q1: 导出的Excel打开提示格式错误?

**原因**: 文件未完全写入或流未关闭。

**解决方案**:
```java
try (OutputStream os = new FileOutputStream(file)) {
    ExcelExportUtils.exportSingleSheet(os, config);
} catch (Exception e) {
    // 确保流被正确关闭
    e.printStackTrace();
}
```

### Q2: 数据量很大时导出很慢?

**原因**: 批量刷新大小设置不合理。

**解决方案**:
```java
.setBatchSize(1000);  // 增大批次大小
```

### Q3: 字段值获取失败,报"字段不存在"?

**原因**: 字段名拼写错误或实体类无getter方法。

**解决方案**:
- 检查字段名是否与实体类字段完全一致(区分大小写)
- 确保实体类有标准的getter方法(getXxx或isXxx)

### Q4: 中文显示为乱码?

**原因**: 文件编码或字体设置问题。

**解决方案**:
```java
new CellStyleConfig()
    .setFontName("微软雅黑")  // 使用支持中文的字体
```

### Q5: 样式没有生效?

**原因**: 样式优先级或配置顺序问题。

**解决方案**:
- 列样式优先于默认样式
- 确保样式配置在`setColumnConfigs()`中正确设置

### Q6: 导出数据量超过限制?

**错误信息**: "数据量超过Excel最大限制(1,048,576行)"

**解决方案**:
- 拆分数据到多个Sheet
- 分批次导出

---

## 性能优化

### 1. 样式缓存优化

工具内部已实现样式指纹缓存机制,相同样式会自动复用,无需手动优化。

### 2. 流式写入

使用`SXSSFWorkbook`的窗口机制,默认窗口大小为100行:

```java
private static final int WINDOW_SIZE = 100;  // 在ExcelExportUtils中定义
```

如需调整,可修改常量并重新编译。

### 3. 批量刷新

合理设置`batchSize`参数:

| 数据量 | 推荐batchSize | 说明 |
|--------|---------------|------|
| < 1000 | 100-500 | 小批量,内存友好 |
| 1000-10000 | 500-1000 | 中等批次 |
| > 10000 | 1000-5000 | 大批次,减少刷新次数 |

### 4. 内存管理

- 使用try-with-resources确保流正确关闭
- 大数据量导出时,设置合理的JVM堆内存(-Xmx)
- 避免在导出过程中保留对数据的强引用

### 5. 数据预处理

在导出前完成数据转换、格式化,减少在导出过程中的计算:

```java
// 推荐: 导出前预处理
List<UserDTO> exportData = dataList.stream()
    .map(user -> convertToDTO(user))
    .collect(Collectors.toList());

SheetConfig<UserDTO> config = new SheetConfig<UserDTO>()
    .setDataList(exportData);
```

---

## 附录

### 完整示例

```java
package com.example.excel.test;

import com.example.excel.config.ColumnConfig;
import com.example.excel.config.HeaderConfig;
import com.example.excel.config.SheetConfig;
import com.example.excel.config.StyleTemplate;
import com.example.excel.utils.ExcelExportUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class CompleteExample {

    public static void main(String[] args) {
        exportEmployeeReport();
    }

    /**
     * 导出员工报表(完整示例)
     */
    public static void exportEmployeeReport() {
        try {
            // 1. 准备数据
            List<User> dataList = generateUserData(50);

            // 2. 配置表头(使用样式模板)
            HeaderConfig header1 = new HeaderConfig()
                    .addColumnNames("序号", "基本信息", "基本信息", "基本信息", "基本信息", "绩效数据", "绩效数据", "绩效数据", "绩效数据", "备注")
                    .setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig())
                    .addMergeRegion(0, 1, 0, 0)   // 第1列合并第1-2行
                    .addMergeRegion(0, 0, 1, 4)   // 第2-5列合并（基本信息）
                    .addMergeRegion(0, 0, 5, 8)   // 第6-9列合并（绩效数据）
                    .addMergeRegion(0, 1, 9, 9);  // 第10列合并第1-2行

            HeaderConfig header2 = new HeaderConfig()
                    .addColumnNames("", "姓名", "部门", "职位", "入职日期", "绩效评分", "项目数", "满意度", "薪资", "备注")
                    .setStyleConfig(StyleTemplate.HEADER_BLUE.toStyleConfig());

            // 3. 配置列
            List<ColumnConfig> columns = new ArrayList<>();

            columns.add(new ColumnConfig()
                    .setFieldName("seq")
                    .setWidth(8)
                    .setStyleConfig(StyleTemplate.DATA_CENTER.toStyleConfig()));

            columns.add(new ColumnConfig().setFieldName("name").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA_LEFT.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("department").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA_LEFT.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("position").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA_LEFT.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("hireDate").setWidth(14)
                    .setStyleConfig(StyleTemplate.DATA_CENTER.toStyleConfig()));

            columns.add(new ColumnConfig().setFieldName("performanceScore").setWidth(12)
                    .setNumberFormat("0.00")
                    .setStyleConfig(StyleTemplate.DATA_RIGHT.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("projectCount").setWidth(10)
                    .setStyleConfig(StyleTemplate.DATA_RIGHT.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("satisfaction").setWidth(10)
                    .setNumberFormat("0.0%")
                    .setStyleConfig(StyleTemplate.DATA_RIGHT.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("salary").setWidth(12)
                    .setNumberFormat("#,##0")
                    .setStyleConfig(StyleTemplate.DATA_RIGHT.toStyleConfig()));

            columns.add(new ColumnConfig().setFieldName("remark").setWidth(15)
                    .setStyleConfig(StyleTemplate.DATA_LEFT.toStyleConfig()));

            // 4. 配置Sheet
            SheetConfig<User> sheet = new SheetConfig<User>()
                    .setSheetName("员工报表")
                    .setHeaders(Arrays.asList(header1, header2))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(2)            // 冻结前2行
                    .setFreezeCol(0)            // 不冻结列
                    .setDefaultDataStyle(StyleTemplate.DATA_CENTER.toStyleConfig())
                    .setBatchSize(25);

            // 5. 导出
            try (OutputStream os = new FileOutputStream("员工报表.xlsx")) {
                ExcelExportUtils.exportSingleSheet(os, sheet);
            }

            System.out.println("导出成功!");

        } catch (Exception e) {
            System.err.println("导出失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 生成测试数据
     */
    private static List<User> generateUserData(int count) {
        List<User> list = new ArrayList<>();
        String[] depts = {"技术部", "销售部", "市场部", "人事部", "财务部"};
        String[] positions = {"工程师", "经理", "总监", "专员", "主管"};

        for (int i = 1; i <= count; i++) {
            User u = new User();
            u.setSeq(i);
            u.setName("员工" + String.format("%03d", i));
            u.setDepartment(depts[i % depts.length]);
            u.setPosition(positions[i % positions.length]);
            u.setHireDate(String.format("202%d-%02d-%02d", i % 6, (i % 12) + 1, (i % 28) + 1));
            u.setPerformanceScore(60 + Math.random() * 40);
            u.setProjectCount(i % 50 + 1);
            u.setSatisfaction(0.7 + Math.random() * 0.3);
            u.setSalary(8000 + Math.random() * 20000);
            u.setRemark("备注信息" + i);
            list.add(u);
        }
        return list;
    }

    static class User {
        private Integer seq;
        private String name;
        private String department;
        private String position;
        private String hireDate;
        private Double performanceScore;
        private Integer projectCount;
        private Double satisfaction;
        private Double salary;
        private String remark;

        // getter/setter省略...
    }
}
```

---

## 更新日志

### v3.0.0 (2026-03-22)

**新增**:
- ✨ 新增`StyleTemplate`样式模板,提供预定义样式
- ✨ 新增`DataConverter`详细文档和使用示例
- ✨ 新增`StyleProvider`动态样式接口,支持条件高亮、状态着色等场景
- ✨ 新增RGB颜色支持（`setRgbBackgroundColor` + `StyleTemplate.rgb()`）

**优化**:
- 🚀 优化样式缓存策略,引入样式指纹机制
- 🚀 移除`RowStyleProvider`,简化API
- 🚀 移除斑马纹样式,改用样式模板实现
- 🚀 样式缓存不再为每行创建新对象,大幅减少内存占用

**修复**:
- 🐛 修复样式缓存键冲突问题
- 🐛 修复大数据量导出时的内存泄漏

### v2.0.0

- 支持多级表头
- 支持单元格合并
- 支持自定义数据转换
- 引入`RowStyleProvider`

### v1.0.0

- 基础Excel导出功能
- 支持实体类和Map数据源
- 支持基础样式配置

---

## 许可证

MIT License

---

## 联系方式

如有问题或建议,欢迎提Issue或Pull Request。
