package com.example.excel.utils;

import com.example.excel.config.*;
import com.example.excel.exception.ExcelExportException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;


import java.io.OutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;

/**
 * 企业级Excel导出工具类
 *
 * 核心特性：
 * 1. 基于 XSSFWorkbook，支持图表生成
 * 2. 样式缓存策略，避免内存溢出
 * 3. 支持实体类和Map两种数据源
 * 4. 完善的中文异常处理
 * 5. 多级表头、条件样式、柱状图等企业级功能
 *
 * @author Excel Export Tool
 * @version 4.0.0
 */
public class ExcelExportUtils {

    // 最大行数限制
    private static final int MAX_ROWS = 1000000;

    // 最大列数限制
    private static final int MAX_COLUMNS = 16384;

    // 样式缓存键前缀
    private static final String STYLE_HEADER = "header_";
    private static final String STYLE_DATA = "data_";

    /**
     * Method 缓存：key = "全限定类名#methodName"，value = Method 对象
     * 避免每行每列都重复调用 clazz.getMethod()，对字段多的实体类性能提升显著
     */
    private static final Map<String, Method> METHOD_CACHE = new java.util.concurrent.ConcurrentHashMap<>();

    // ==================== 公共API方法 ====================

    /**
     * 导出Excel到输出流
     * <p>支持单个Sheet或多个Sheet</p>
     *
     * @param outputStream 输出流
     * @param sheetConfigs Sheet配置列表(单个或多个)
     * @throws ExcelExportException 导出异常
     */
    public static void export(OutputStream outputStream, SheetConfig<?>... sheetConfigs)
            throws ExcelExportException {
        if (sheetConfigs == null || sheetConfigs.length == 0) {
            throw new ExcelExportException("Sheet配置不能为空", "EMPTY_SHEET_CONFIG");
        }
        exportToStream(outputStream, Arrays.asList(sheetConfigs));
    }

    /**
     * 导出Excel到HTTP响应（Web场景）
     * <p>自动设置响应头和文件名</p>
     *
     * @param response HTTP响应
     * @param fileName 文件名（包含扩展名，如 "员工报表.xlsx"）
     * @param sheetConfigs Sheet配置列表(单个或多个)
     * @throws ExcelExportException 导出异常
     * @throws IOException IO异常
     *
     * <p><b>使用示例：</b></p>
     * <pre>{@code
     * // Controller中使用
     * public void exportReport(HttpServletResponse response) {
     *     ExcelExportUtils.exportToResponse(response, "员工报表.xlsx",
     *         buildSheetConfig());
     * }
     * }</pre>
     */
    public static void exportToResponse(HttpServletResponse response, String fileName,
                                         SheetConfig<?>... sheetConfigs)
            throws ExcelExportException, IOException {
        if (response == null) {
            throw new IllegalArgumentException("HTTP响应不能为空");
        }
        if (fileName == null || fileName.trim().isEmpty()) {
            throw new IllegalArgumentException("文件名不能为空");
        }

        // 设置响应头
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition",
            "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));

        // 导出
        try (OutputStream os = response.getOutputStream()) {
            export(os, sheetConfigs);
        }
    }

    /**
     * 导出Excel到HTTP响应（Web场景，自动加时间戳）
     * <p>文件名格式：{baseName}_yyyyMMdd_HHmmss.xlsx</p>
     *
     * @param response HTTP响应
     * @param baseName 基础文件名（不含扩展名和时间戳，如 "员工报表"）
     * @param sheetConfigs Sheet配置列表(单个或多个)
     * @throws ExcelExportException 导出异常
     * @throws IOException IO异常
     *
     * <p><b>使用示例：</b></p>
     * <pre>{@code
     * // 文件名自动变成：员工报表_20260322_164500.xlsx
     * public void exportReport(HttpServletResponse response) {
     *     ExcelExportUtils.exportToResponseWithTimestamp(response, "员工报表",
     *         buildSheetConfig());
     * }
     * }</pre>
     */
    public static void exportToResponseWithTimestamp(HttpServletResponse response, String baseName,
                                                       SheetConfig<?>... sheetConfigs)
            throws ExcelExportException, IOException {
        if (baseName == null || baseName.trim().isEmpty()) {
            throw new IllegalArgumentException("基础文件名不能为空");
        }

        // 自动生成时间戳
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = sdf.format(new Date());
        String fileName = baseName + "_" + timestamp + ".xlsx";

        exportToResponse(response, fileName, sheetConfigs);
    }

    /**
     * 创建简单模板配置
     *
     * @param sheetName Sheet名称
     * @param dataList 数据列表
     * @param fieldNames 字段名数组
     * @param columnNames 列名数组
     * @return SheetConfig配置对象
     */
    public static <T> SheetConfig<T> createSimpleTemplate(
            String sheetName, List<T> dataList,
            String[] fieldNames, String[] columnNames) {

        // 使用预定义的表头和数据样式
        CellStyleConfig headerStyle = StyleTemplate.HEADER.toStyleConfig();
        CellStyleConfig dataStyle = StyleTemplate.DATA.toStyleConfig();

        List<ColumnConfig> columns = new ArrayList<>();
        for (String fieldName : fieldNames) {
            columns.add(new ColumnConfig()
                    .setFieldName(fieldName)
                    .setWidth(15)
                    .setStyleConfig(dataStyle));
        }

        HeaderConfig header = new HeaderConfig()
                .addColumnNames(columnNames)
                .setStyleConfig(headerStyle);

        return new SheetConfig<T>()
                .setSheetName(sheetName)
                .setHeaders(Collections.singletonList(header))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(1)
                .setDefaultDataStyle(dataStyle)
                .setBatchSize(100);
    }

    // ==================== 核心导出逻辑 ====================

    /**
     * 核心导出方法（基于 XSSFWorkbook，支持图表）
     */
    private static void exportToStream(OutputStream outputStream, List<SheetConfig<?>> sheetConfigs)
            throws ExcelExportException {
        if (outputStream == null) {
            throw new ExcelExportException("输出流不能为空", "NULL_OUTPUT_STREAM");
        }
        if (CollectionUtils.isEmpty(sheetConfigs)) {
            throw new ExcelExportException("Sheet配置列表不能为空", "EMPTY_SHEET_CONFIG");
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Map<String, CellStyle> styleCache = new HashMap<>();
            DataFormat dataFormat = workbook.createDataFormat();

                // chartXmlMap: key=ZIP条目路径(如"xl/charts/chart1.xml"), value=完整XML字符串
            // buildChart 返回已构建好的 chartSpaceXml，由 fixChartXmlInZip 做最终替换
            Map<String, String> chartXmlMap = new java.util.LinkedHashMap<>();
            for (int i = 0; i < sheetConfigs.size(); i++) {
                SheetConfig<?> config = sheetConfigs.get(i);
                try {
                    validateSheetConfig(config, i);
                    buildSheet(workbook, config, styleCache, dataFormat, chartXmlMap);
                } catch (ExcelExportException e) {
                    throw e;
                } catch (Exception e) {
                    throw new ExcelExportException(
                        String.format("创建第%d个Sheet【%s】失败：%s", i + 1, config.getSheetName(), e.getMessage()),
                        "SHEET_BUILD_ERROR", e);
                }
            }

            if (chartXmlMap.isEmpty()) {
                // 无图表：直接写出
                workbook.write(outputStream);
            } else {
                // 有图表：先写到内存，再用正确的 XML 替换 chart*.xml 和修复 drawing*.xml
                java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
                workbook.write(baos);
                byte[] fixed = fixChartXmlInZip(baos.toByteArray(), chartXmlMap);
                outputStream.write(fixed);
            }
            outputStream.flush();
        } catch (ExcelExportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelExportException("导出Excel失败：" + e.getMessage(), "EXPORT_ERROR", e);
        }
    }

    /**
     * 校验Sheet配置
     */
    private static void validateSheetConfig(SheetConfig<?> config, int index) {
        if (config == null) {
            throw new ExcelExportException(String.format("第%d个Sheet配置为空", index + 1), "NULL_SHEET_CONFIG");
        }
        if (CollectionUtils.isEmpty(config.getColumnConfigs())) {
            throw new ExcelExportException(
                String.format("第%d个Sheet【%s】的列配置不能为空", index + 1, config.getSheetName()),
                "EMPTY_COLUMN_CONFIG");
        }
        if (config.getDataList() != null && config.getDataList().size() > MAX_ROWS) {
            throw new ExcelExportException(
                String.format("第%d个Sheet【%s】数据量超过限制（最大%d行）", index + 1, config.getSheetName(), MAX_ROWS),
                "DATA_EXCEED_LIMIT");
        }
    }

    // ==================== Sheet构建 ====================

    /**
     * 构建单个Sheet
     *
     * @param chartXmlMap 用于收集图表 XML：key=ZIP条目路径(如"xl/charts/chart1.xml"),
     *                    value=完整 chartSpaceXml 字符串。后续由 fixChartXmlInZip 做替换。
     */
    private static void buildSheet(XSSFWorkbook workbook, SheetConfig<?> config,
                                   Map<String, CellStyle> styleCache, DataFormat dataFormat,
                                   Map<String, String> chartXmlMap) {
        XSSFSheet sheet = workbook.createSheet(config.getSheetName());
        sheet.setDefaultRowHeight(config.getDefaultRowHeight());
        sheet.setDisplayGridlines(config.isDisplayGridlines());

        // 构建表头
        int headerRowCount = buildHeaders(sheet, config, styleCache, workbook);

        // 设置冻结
        if (config.getFreezeRow() > 0 || config.getFreezeCol() > 0) {
            sheet.createFreezePane(config.getFreezeCol(), config.getFreezeRow());
        }

        // 写入数据，返回实际数据结束行（含表头）
        int lastDataRow = writeDataRows(sheet, config, headerRowCount, styleCache, workbook, dataFormat);

        // 设置列宽
        applyColumnWidths(sheet, config);

        // 设置自动筛选
        if (config.isAutoFilter() && headerRowCount > 0) {
            int lastCol = config.getColumnConfigs().size() - 1;
            sheet.setAutoFilter(new CellRangeAddress(0, headerRowCount - 1, 0, lastCol));
        }

        // 生成图表（有 chartConfig 才执行）
        if (config.getChartConfig() != null) {
            // buildChart 返回构建好的 chartSpaceXml 字符串，同时通过 POI API 建立 drawing/chart 关系
            // 真正的 XML 内容由 fixChartXmlInZip 在 ZIP 层面整体替换，绕开 POI 序列化的命名空间问题
            String chartXml = buildChart(sheet, config, lastDataRow);
            // 取该 sheet 下的 drawing，再取图表的 PackagePart 路径
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            if (drawing != null) {
                for (XSSFChart c : drawing.getCharts()) {
                    // partName 形如 /xl/charts/chart1.xml，去掉前导 /
                    String partName = c.getPackagePart().getPartName().getName();
                    if (partName.startsWith("/")) partName = partName.substring(1);
                    chartXmlMap.put(partName, chartXml);
                    break; // 每个 sheet 只有一个图表
                }
            }
        }
    }

    /**
     * 构建表头
     * @return 表头行数
     */
    private static int buildHeaders(XSSFSheet sheet, SheetConfig<?> config,
                                    Map<String, CellStyle> styleCache, XSSFWorkbook workbook) {
        List<HeaderConfig> headers = config.getHeaders();
        if (CollectionUtils.isEmpty(headers)) {
            return 0;
        }

        int rowIndex = 0;
        for (int level = 0; level < headers.size(); level++) {
            HeaderConfig headerConfig = headers.get(level);
            XSSFRow row = sheet.createRow(rowIndex);
            row.setHeight(headerConfig.getHeight());

            List<String> columnNames = headerConfig.getColumnNames();
            if (CollectionUtils.isEmpty(columnNames)) {
                rowIndex++;
                continue;
            }

            // 预处理：将合并区域内的非空值自动归位到 startCol
            // 这样无论用户把文字写在合并区域的哪个位置，都能正确显示
            List<String> normalizedNames = normalizeMergeRegionValues(columnNames, headerConfig.getMergeRegions());

            CellStyle style = createCellStyle(workbook, headerConfig.getStyleConfig(),
                styleCache, STYLE_HEADER + level);

            for (int col = 0; col < normalizedNames.size(); col++) {
                Cell cell = row.createCell(col);
                String name = normalizedNames.get(col);
                if (StringUtils.isNotBlank(name)) {
                    cell.setCellValue(name);
                }
                if (style != null) {
                    cell.setCellStyle(style);
                }
            }

            // 处理合并区域
            applyMergeRegions(sheet, headerConfig.getMergeRegions(), rowIndex);

            rowIndex++;
        }

        return headers.size();
    }

    /**
     * 将合并区域内的非空值归位到 startCol
     * <p>
     * 解决用户随意填写合并区域内文字位置的问题。
     * 扫描每个合并区域 [startCol, endCol]，找到其中第一个非空值，
     * 将其移到 startCol，区域内其他位置置为空。
     * 未被任何合并区域覆盖的列保持不变。
     * </p>
     *
     * <p>示例：</p>
     * <pre>
     * 合并 col1~col4，原始：["序号", "", "", "基本信息", "", "绩效数据", ...]
     *                                             ↑col3随便写的
     * 处理后：            ["序号", "基本信息", "", "", "", "绩效数据", ...]
     *                              ↑自动归位到col1
     * </pre>
     */
    private static List<String> normalizeMergeRegionValues(List<String> columnNames,
                                                            List<HeaderConfig.MergeRegion> mergeRegions) {
        // 复制一份，避免修改原始数据
        List<String> result = new ArrayList<>(columnNames);

        if (CollectionUtils.isEmpty(mergeRegions)) {
            return result;
        }

        for (HeaderConfig.MergeRegion region : mergeRegions) {
            // 只处理同行合并（startRow == endRow == 0，即列合并）
            if (region.getStartRow() != region.getEndRow()) {
                continue;
            }
            int startCol = region.getStartCol();
            int endCol = region.getEndCol();
            if (startCol >= endCol) {
                continue; // 单列，不需要处理
            }

            // 找到区域内第一个非空值
            String foundValue = null;
            for (int col = startCol; col <= endCol && col < result.size(); col++) {
                String v = result.get(col);
                if (StringUtils.isNotBlank(v)) {
                    foundValue = v;
                    break;
                }
            }

            if (foundValue == null) {
                continue; // 整个区域都是空，不处理
            }

            // 将非空值归位到 startCol，区域内其余位置清空
            for (int col = startCol; col <= endCol && col < result.size(); col++) {
                result.set(col, col == startCol ? foundValue : "");
            }
        }

        return result;
    }

    /**
     * 应用合并区域
     */
    private static void applyMergeRegions(XSSFSheet sheet, List<HeaderConfig.MergeRegion> regions, int rowIndex) {
        if (CollectionUtils.isEmpty(regions)) {
            return;
        }
        for (HeaderConfig.MergeRegion region : regions) {
            try {
                sheet.addMergedRegion(new CellRangeAddress(
                    rowIndex + region.getStartRow(),
                    rowIndex + region.getEndRow(),
                    region.getStartCol(),
                    region.getEndCol()
                ));
            } catch (Exception e) {
                throw new ExcelExportException(
                    String.format("合并单元格失败：行[%d-%d]，列[%d-%d]，原因：%s",
                        region.getStartRow(), region.getEndRow(),
                        region.getStartCol(), region.getEndCol(), e.getMessage()),
                    "MERGE_ERROR", e);
            }
        }
    }

    // ==================== 数据写入 ====================

    /**
     * 写入数据行
     *
     * @return 最后一个数据行的行号（0-based），无数据时返回 startRow - 1
     */
    private static int writeDataRows(XSSFSheet sheet, SheetConfig<?> config, int startRow,
                                      Map<String, CellStyle> styleCache, XSSFWorkbook workbook,
                                      DataFormat dataFormat) {
        List<?> dataList = config.getDataList();
        if (CollectionUtils.isEmpty(dataList)) {
            return startRow - 1;
        }

        List<ColumnConfig> columns = config.getColumnConfigs();
        int batchSize = Math.max(config.getBatchSize(), 10);
        int totalRows = dataList.size();

        for (int i = 0; i < totalRows; i++) {
            int rowIndex = startRow + i;

            if (rowIndex >= MAX_ROWS) {
                throw new ExcelExportException(
                    String.format("数据行数超过Excel最大限制（%d行）", MAX_ROWS),
                    "ROW_EXCEED_LIMIT");
            }

            XSSFRow row = sheet.createRow(rowIndex);
            Object data = dataList.get(i);
            fillRowData(row, data, columns, styleCache, workbook, dataFormat, rowIndex, config);
        }

        return startRow + totalRows - 1;
    }

    /**
     * 填充单行数据
     */
    private static void fillRowData(XSSFRow row, Object data, List<ColumnConfig> columns,
                                    Map<String, CellStyle> styleCache, XSSFWorkbook workbook,
                                    DataFormat dataFormat, int rowIndex, SheetConfig<?> config) {
        if (CollectionUtils.isEmpty(columns)) {
            return;
        }

        for (int col = 0; col < columns.size(); col++) {
            if (col >= MAX_COLUMNS) {
                throw new ExcelExportException(
                    String.format("列数超过Excel最大限制（%d列）", MAX_COLUMNS),
                    "COLUMN_EXCEED_LIMIT");
            }

            ColumnConfig colConfig = columns.get(col);
            Cell cell = row.createCell(col);

            // 获取字段值
            Object value = extractFieldValue(data, colConfig.getFieldName(), rowIndex, col);

            // 设置样式（StyleProvider 使用原始值，在 DataConverter 之前）
            CellStyle style = resolveCellStyle(workbook, colConfig, styleCache, config, value, data, rowIndex);
            if (style != null) {
                cell.setCellStyle(style);
            }

            // 应用数据转换器
            if (colConfig.getDataConverter() != null) {
                try {
                    value = colConfig.getDataConverter().convert(value, data, rowIndex);
                } catch (Exception e) {
                    throw new ExcelExportException(
                        String.format("第%d行第%d列数据转换失败：%s", rowIndex + 1, col + 1, e.getMessage()),
                        "DATA_CONVERT_ERROR", e);
                }
            }

            // 设置单元格值
            setCellValueSafe(cell, value, colConfig, dataFormat, rowIndex, col);
        }
    }

    /**
     * 提取字段值
     */
    private static Object extractFieldValue(Object data, String fieldName, int row, int col) {
        if (data == null) {
            return null;
        }
        if (StringUtils.isBlank(fieldName)) {
            throw new ExcelExportException(
                String.format("第%d行第%d列的字段名称为空", row + 1, col + 1),
                "EMPTY_FIELD_NAME");
        }

        try {
            // Map类型
            if (data instanceof Map) {
                return ((Map<?, ?>) data).get(fieldName);
            }

            // 实体类：优先尝试getter方法，结果缓存到 METHOD_CACHE 避免重复查找
            Class<?> clazz = data.getClass();
            String cacheKeyGet = clazz.getName() + "#get" + capitalize(fieldName);
            String cacheKeyIs  = clazz.getName() + "#is"  + capitalize(fieldName);

            Method method = METHOD_CACHE.get(cacheKeyGet);
            if (method == null) {
                // 缓存未命中，尝试 getXxx
                try {
                    method = clazz.getMethod("get" + capitalize(fieldName));
                    method.setAccessible(true);
                    METHOD_CACHE.put(cacheKeyGet, method);
                } catch (NoSuchMethodException e) {
                    // 尝试 isXxx（布尔型）
                    try {
                        method = clazz.getMethod("is" + capitalize(fieldName));
                        method.setAccessible(true);
                        // 用 cacheKeyGet 存储，后续同一字段直接命中
                        METHOD_CACHE.put(cacheKeyGet, method);
                    } catch (NoSuchMethodException ex) {
                        // getter 都没有，回退到字段直接访问，用特殊标记避免重复查找
                        METHOD_CACHE.put(cacheKeyGet, null); // null 表示"无getter，用字段反射"
                    }
                }
            }

            if (method != null) {
                return method.invoke(data);
            }
            // method == null（缓存了"无getter"标记），直接走字段反射
            return getFieldByReflection(data, fieldName, clazz, row, col);
        } catch (ExcelExportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelExportException(
                String.format("第%d行第%d列获取字段【%s】失败：%s", row + 1, col + 1, fieldName, e.getMessage()),
                "FIELD_ACCESS_ERROR", e);
        }
    }

    /**
     * Field 缓存：key = "全限定类名#fieldName"，value = Field 对象
     */
    private static final Map<String, Field> FIELD_CACHE = new java.util.concurrent.ConcurrentHashMap<>();

    /**
     * 通过反射获取字段值（无 getter 时的降级方案）
     * Field 对象会缓存，避免每次调用 getDeclaredField()
     */
    private static Object getFieldByReflection(Object data, String fieldName, Class<?> clazz, int row, int col) {
        String cacheKey = clazz.getName() + "#field#" + fieldName;
        Field field = FIELD_CACHE.get(cacheKey);
        if (field == null) {
            try {
                field = clazz.getDeclaredField(fieldName);
                field.setAccessible(true);
                FIELD_CACHE.put(cacheKey, field);
            } catch (NoSuchFieldException e) {
                throw new ExcelExportException(
                    String.format("第%d行第%d列字段【%s】不存在于类【%s】中", row + 1, col + 1, fieldName, clazz.getName()),
                    "FIELD_NOT_FOUND", e);
            }
        }
        try {
            return field.get(data);
        } catch (IllegalAccessException e) {
            throw new ExcelExportException(
                String.format("第%d行第%d列字段【%s】访问权限不足", row + 1, col + 1, fieldName),
                "FIELD_ACCESS_DENIED", e);
        }
    }

    /**
     * 安全设置单元格值（含类型转换和异常处理）
     */
    private static void setCellValueSafe(Cell cell, Object value, ColumnConfig config,
                                         DataFormat dataFormat, int row, int col) {
        if (value == null) {
            cell.setCellValue("");
            return;
        }

        try {
            if (value instanceof Number) {
                setNumericValue(cell, (Number) value, config, dataFormat);
            } else if (value instanceof Date) {
                setDateValue(cell, (Date) value, config, dataFormat);
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value ? "是" : "否");
            } else if (value instanceof Calendar) {
                cell.setCellValue((Calendar) value);
            } else {
                cell.setCellValue(convertToString(value, row, col));
            }
        } catch (ExcelExportException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelExportException(
                String.format("第%d行第%d列设置值失败，值类型【%s】，原因：%s",
                    row + 1, col + 1, value.getClass().getName(), e.getMessage()),
                "SET_VALUE_ERROR", e);
        }
    }

    /**
     * 设置数值类型值（处理精度问题）
     */
    private static void setNumericValue(Cell cell, Number value, ColumnConfig config, DataFormat dataFormat) {
        double doubleValue;
        
        // 处理BigDecimal精度
        if (value instanceof BigDecimal) {
            doubleValue = ((BigDecimal) value).setScale(10, RoundingMode.HALF_UP).doubleValue();
        } else if (value instanceof Double || value instanceof Float) {
            doubleValue = value.doubleValue();
        } else {
            // 整数类型直接设置long值
            cell.setCellValue(value.longValue());
            applyNumberFormat(cell, config, dataFormat);
            return;
        }

        cell.setCellValue(doubleValue);
        applyNumberFormat(cell, config, dataFormat);
    }

    /**
     * 应用数字格式
     */
    private static void applyNumberFormat(Cell cell, ColumnConfig config, DataFormat dataFormat) {
        if (StringUtils.isNotBlank(config.getNumberFormat())) {
            CellStyle style = cell.getCellStyle();
            if (style == null) {
                style = cell.getSheet().getWorkbook().createCellStyle();
            }
            style.setDataFormat(dataFormat.getFormat(config.getNumberFormat()));
            cell.setCellStyle(style);
        }
    }

    /**
     * 设置日期类型值
     */
    private static void setDateValue(Cell cell, Date value, ColumnConfig config, DataFormat dataFormat) {
        String format = StringUtils.isNotBlank(config.getDateFormat()) 
            ? config.getDateFormat() : "yyyy-MM-dd HH:mm:ss";
        
        CellStyle style = cell.getCellStyle();
        if (style == null) {
            style = cell.getSheet().getWorkbook().createCellStyle();
        }
        style.setDataFormat(dataFormat.getFormat(format));
        cell.setCellStyle(style);
        cell.setCellValue(value);
    }

    /**
     * 安全转换为字符串
     */
    private static String convertToString(Object value, int row, int col) {
        try {
            return value.toString();
        } catch (Exception e) {
            throw new ExcelExportException(
                String.format("第%d行第%d列值转字符串失败，类型【%s】", row + 1, col + 1, value.getClass().getName()),
                "TO_STRING_ERROR", e);
        }
    }

    // ==================== 样式处理 ====================

    /**
     * 解析单元格样式
     * <p>样式优先级：StyleProvider（动态）> 列 styleConfig（静态）> Sheet 默认数据样式</p>
     *
     * @param value    当前单元格原始值（供 StyleProvider 做条件判断）
     * @param rowData  整行数据（供 StyleProvider 访问其他字段）
     * @param rowIndex 当前数据行索引
     */
    private static CellStyle resolveCellStyle(XSSFWorkbook workbook, ColumnConfig colConfig,
                                              Map<String, CellStyle> styleCache, SheetConfig<?> sheetConfig,
                                              Object value, Object rowData, int rowIndex) {
        // 1. 动态样式优先（StyleProvider）
        if (colConfig.getStyleProvider() != null) {
            try {
                CellStyleConfig dynamicConfig = colConfig.getStyleProvider().provide(value, rowData, rowIndex);
                if (dynamicConfig != null) {
                    // 动态样式不走固定 cacheKey，而是用样式指纹做缓存键，避免每行重复创建
                    return createCellStyle(workbook, dynamicConfig, styleCache,
                        STYLE_DATA + colConfig.getFieldName() + "_dynamic");
                }
            } catch (Exception e) {
                throw new ExcelExportException(
                    String.format("第%d行列【%s】动态样式计算失败：%s", rowIndex + 1, colConfig.getFieldName(), e.getMessage()),
                    "STYLE_PROVIDER_ERROR", e);
            }
        }

        // 2. 列静态样式
        if (colConfig.getStyleConfig() != null) {
            return createCellStyle(workbook, colConfig.getStyleConfig(), styleCache,
                STYLE_DATA + colConfig.getFieldName());
        }

        // 3. Sheet 默认数据样式
        if (sheetConfig.getDefaultDataStyle() != null) {
            return createCellStyle(workbook, sheetConfig.getDefaultDataStyle(), styleCache,
                STYLE_DATA + "default");
        }

        return null;
    }

    /**
     * 创建或获取缓存样式
     * <p>使用样式指纹机制,相同样式的配置会复用同一个Style对象</p>
     */
    private static CellStyle createCellStyle(XSSFWorkbook workbook, CellStyleConfig config,
                                             Map<String, CellStyle> styleCache, String cacheKey) {
        if (config == null) {
            return null;
        }

        // 生成样式指纹作为缓存键,确保相同样式只创建一个对象
        String styleFingerprint = generateStyleFingerprint(cacheKey, config);

        return styleCache.computeIfAbsent(styleFingerprint, key -> {
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();

            // 背景色：优先使用 RGB 颜色，降级使用 IndexedColors
            byte[] rgbBg = config.getRgbBackgroundColor();
            if (rgbBg != null && rgbBg.length == 3) {
                // 使用真正的 RGB 颜色（需要 XSSFCellStyle）
                if (style instanceof XSSFCellStyle) {
                    XSSFColor xssfColor = new XSSFColor(
                        new java.awt.Color(rgbBg[0] & 0xFF, rgbBg[1] & 0xFF, rgbBg[2] & 0xFF), null);
                    ((XSSFCellStyle) style).setFillForegroundColor(xssfColor);
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
            } else {
                short bgColor = config.getBackgroundColor();
                if (bgColor >= 0) {
                    style.setFillForegroundColor(bgColor);
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
            }

            // 对齐
            if (config.getHorizontalAlignment() != null) {
                style.setAlignment(config.getHorizontalAlignment());
            }
            if (config.getVerticalAlignment() != null) {
                style.setVerticalAlignment(config.getVerticalAlignment());
            }

            // 边框
            applyBorders(style, config);

            // 其他属性
            style.setWrapText(config.isWrapText());
            style.setIndention(config.getIndentation());
            style.setRotation(config.getRotation());
            style.setLocked(config.isLocked());
            style.setHidden(config.isHidden());

            // 字体
            if (config.getFontColor() >= 0) {
                font.setColor(config.getFontColor());
            }
            font.setBold(config.isBold());
            font.setFontHeightInPoints(config.getFontSize());
            if (StringUtils.isNotBlank(config.getFontName())) {
                font.setFontName(config.getFontName());
            }
            font.setItalic(config.isItalic());
            font.setStrikeout(config.isStrikeout());
            font.setUnderline(config.getUnderline());
            style.setFont(font);

            return style;
        });
    }

    /**
     * 生成样式指纹
     * <p>将样式配置的所有属性转换为唯一标识,用于样式缓存</p>
     */
    private static String generateStyleFingerprint(String cacheKey, CellStyleConfig config) {
        StringBuilder sb = new StringBuilder(cacheKey);
        sb.append("|bg:").append(config.getBackgroundColor());
        // RGB 背景色独立记录
        byte[] rgb = config.getRgbBackgroundColor();
        if (rgb != null && rgb.length == 3) {
            sb.append("|rgb:").append(rgb[0] & 0xFF).append(",").append(rgb[1] & 0xFF).append(",").append(rgb[2] & 0xFF);
        }
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
        sb.append("|italic:").append(config.isItalic());
        sb.append("|strike:").append(config.isStrikeout());
        sb.append("|underline:").append(config.getUnderline());
        sb.append("|indent:").append(config.getIndentation());
        sb.append("|rotate:").append(config.getRotation());
        sb.append("|locked:").append(config.isLocked());
        sb.append("|hidden:").append(config.isHidden());
        return sb.toString();
    }

    /**
     * 应用边框样式
     */
    private static void applyBorders(CellStyle style, CellStyleConfig config) {
        if (config.getBorderTop() != null) {
            style.setBorderTop(config.getBorderTop());
        }
        if (config.getBorderBottom() != null) {
            style.setBorderBottom(config.getBorderBottom());
        }
        if (config.getBorderLeft() != null) {
            style.setBorderLeft(config.getBorderLeft());
        }
        if (config.getBorderRight() != null) {
            style.setBorderRight(config.getBorderRight());
        }
    }

    // ==================== 列宽设置 ====================

    /**
     * 应用列宽
     * <p>列宽取以下三者的最大值：
     * <ol>
     *   <li>ColumnConfig 中配置的固定宽度</li>
     *   <li>所有表头行中该列文字的换算宽度</li>
     *   <li>最小保底宽度 5</li>
     * </ol>
     * 这样可以避免表头文字被截断。
     * </p>
     */
    private static void applyColumnWidths(XSSFSheet sheet, SheetConfig<?> config) {
        List<ColumnConfig> columns = config.getColumnConfigs();
        if (CollectionUtils.isEmpty(columns)) {
            return;
        }

        // 预先计算每列表头文字所需的最小宽度（单位：字符数）
        int[] headerMinWidths = calcHeaderMinWidths(config);

        for (int i = 0; i < columns.size(); i++) {
            ColumnConfig col = columns.get(i);
            if (col.isHidden()) {
                sheet.setColumnHidden(i, true);
                continue;
            }
            // 取配置宽度与表头文字宽度的最大值，再与最小值 5 比较
            int headerWidth = (i < headerMinWidths.length) ? headerMinWidths[i] : 0;
            int width = Math.max(Math.max(col.getWidth(), headerWidth), 5);
            sheet.setColumnWidth(i, width * 256);
        }
    }

    /**
     * 计算所有表头行每列所需的最小列宽（字符数）
     * <p>中文字符按 2 个半角字符宽度计算，英文/数字按 1 个字符宽度计算，
     * 再加 2 个字符的左右 padding。</p>
     *
     * @param config Sheet 配置
     * @return 每列最小宽度数组（字符数）
     */
    private static int[] calcHeaderMinWidths(SheetConfig<?> config) {
        List<HeaderConfig> headers = config.getHeaders();
        int colCount = config.getColumnConfigs().size();
        int[] minWidths = new int[colCount];

        if (CollectionUtils.isEmpty(headers)) {
            return minWidths;
        }

        for (HeaderConfig headerConfig : headers) {
            List<String> columnNames = headerConfig.getColumnNames();
            if (CollectionUtils.isEmpty(columnNames)) {
                continue;
            }
            for (int col = 0; col < columnNames.size() && col < colCount; col++) {
                String name = columnNames.get(col);
                if (StringUtils.isBlank(name)) {
                    continue;
                }
                int charWidth = measureTextWidth(name);
                // 加上 padding（左右各1字符）
                int needed = charWidth + 2;
                if (needed > minWidths[col]) {
                    minWidths[col] = needed;
                }
            }
        }
        return minWidths;
    }

    /**
     * 估算文字显示宽度（半角字符单位）
     * <p>中文/全角字符宽度按 2 计算，ASCII 字符按 1 计算。</p>
     */
    private static int measureTextWidth(String text) {
        if (StringUtils.isBlank(text)) {
            return 0;
        }
        int width = 0;
        for (char c : text.toCharArray()) {
            // 中文、全角标点、CJK 字符等宽度为 2
            if (c >= '\u2E80' && c <= '\uFE4F'   // CJK 部首/笔画/兼容
                    || c >= '\uFF00' && c <= '\uFFEF'  // 全角字符
                    || c >= '\u4E00' && c <= '\u9FFF'  // 常用汉字
                    || c >= '\u3400' && c <= '\u4DBF'  // 扩展汉字A
//                    || c >= '\u20000' && c <= '\u2A6DF' // 扩展汉字B（代理对，近似处理）
            ) {
                width += 2;
            } else {
                width += 1;
            }
        }
        return width;
    }

    // ==================== 图表生成 ====================

    /**
     * 在数据结束行下方生成柱状图的 POI 骨架（建立 drawing/chart 关系），
     * 并返回完整的、符合 OOXML Schema 的 chartSpace XML 字符串。
     * <p>
     * 真正的 chart XML 内容由 {@code fixChartXmlInZip} 在 ZIP 层面整体替换，
     * 从而彻底绕开 POI / XMLBeans 序列化带来的命名空间丢失问题。
     * </p>
     *
     * @return 完整 chartSpaceXml 字符串
     */
    private static String buildChart(XSSFSheet sheet, SheetConfig<?> config, int lastDataRow) {
        ChartConfig chartConfig = config.getChartConfig();
        List<ColumnConfig> columns = config.getColumnConfigs();
        int dataSize   = config.getDataList() == null ? 0 : config.getDataList().size();
        int headerRows = config.getHeaders() == null ? 0 : config.getHeaders().size();

        if (dataSize == 0) return "";

        // ---- 找列索引 ----
        int categoryColIdx = findColumnIndex(columns, chartConfig.getCategoryColumn());
        if (categoryColIdx < 0) throw new ExcelExportException(
            "ChartConfig.categoryColumn [" + chartConfig.getCategoryColumn() + "] 未找到",
            "CHART_CATEGORY_NOT_FOUND");

        List<Integer> seriesColIdxList = new ArrayList<>();
        for (ChartConfig.SeriesConfig s : chartConfig.getSeriesList()) {
            int idx = findColumnIndex(columns, s.getFieldName());
            if (idx < 0) throw new ExcelExportException(
                "series fieldName [" + s.getFieldName() + "] 未找到", "CHART_SERIES_NOT_FOUND");
            seriesColIdxList.add(idx);
        }

        // ---- 图表锚点 ----
        int chartStartRow   = lastDataRow + 1 + chartConfig.getGapRows();
        int chartHeightRows = Math.min(Math.max((int)(dataSize * 0.6), 15), 40);
        int chartWidthCols  = Math.min(Math.max(columns.size(), 10), 20);

        XSSFDrawing      drawing   = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor    = drawing.createAnchor(0, 0, 0, 0,
            0, chartStartRow, chartWidthCols, chartStartRow + chartHeightRows);
        XSSFChart        xssfChart = drawing.createChart(anchor);

        // ---- 构建参数 ----
        String sheetName    = escapeXmlAttr(sheet.getSheetName());
        int    dataFirstRow = headerRows + 1;          // 1-based
        int    dataLastRow  = lastDataRow + 1;          // 1-based

        String catRange = buildCellRange(sheet.getSheetName(), dataFirstRow, dataLastRow,
            categoryColIdx, categoryColIdx);

        List<ChartConfig.SeriesConfig> seriesList   = chartConfig.getSeriesList();
        List<String>                   seriesColors = chartConfig.getSeriesColors();

        // ---- 构建系列 XML ----
        StringBuilder seriesXml = new StringBuilder();
        for (int s = 0; s < seriesList.size(); s++) {
            ChartConfig.SeriesConfig sc  = seriesList.get(s);
            int                      sci = seriesColIdxList.get(s);
            String valRange = buildCellRange(sheet.getSheetName(), dataFirstRow, dataLastRow, sci, sci);

            // 系列颜色（可选）
            String spPrXml = "";
            if (seriesColors != null && s < seriesColors.size()) {
                String hex = seriesColors.get(s).replace("#", "").toUpperCase();
                if (hex.length() == 6) {
                    spPrXml = "<c:spPr><a:solidFill><a:srgbClr val=\"" + hex + "\"/></a:solidFill></c:spPr>";
                }
            }

            // 数据标签（可选）
            String dLblsXml = "";
            if (chartConfig.isShowDataLabel()) {
                dLblsXml = "<c:dLbls>"
                    + "<c:showLegendKey val=\"0\"/>"
                    + "<c:showVal val=\"1\"/>"
                    + "<c:showCatName val=\"0\"/>"
                    + "<c:showSerName val=\"0\"/>"
                    + "<c:showPercent val=\"0\"/>"
                    + "<c:showBubbleSize val=\"0\"/>"
                    + "</c:dLbls>";
            }

            seriesXml.append("<c:ser>")
                .append("<c:idx val=\"").append(s).append("\"/>")
                .append("<c:order val=\"").append(s).append("\"/>")
                .append("<c:tx><c:strRef><c:f>").append(escapeXml(sc.getSeriesName())).append("</c:f></c:strRef></c:tx>")
                .append(spPrXml)
                .append(dLblsXml)
                .append("<c:cat><c:strRef><c:f>").append(escapeXml(catRange)).append("</c:f></c:strRef></c:cat>")
                .append("<c:val><c:numRef><c:f>").append(escapeXml(valRange)).append("</c:f></c:numRef></c:val>")
                .append("</c:ser>");
        }

        // ---- 图表标题（用 <c:rich> 富文本，而非 strRef） ----
        String titleXml = "";
        if (StringUtils.isNotBlank(chartConfig.getTitle())) {
            titleXml = "<c:title>"
                + "<c:tx><c:rich>"
                + "<a:bodyPr/><a:lstStyle/>"
                + "<a:p><a:r><a:t>" + escapeXml(chartConfig.getTitle()) + "</a:t></a:r></a:p>"
                + "</c:rich></c:tx>"
                + "<c:overlay val=\"0\"/>"
                + "</c:title>";
        }

        // ---- X轴标题（用 <c:rich>） ----
        String catAxisTitleXml = "";
        if (StringUtils.isNotBlank(chartConfig.getCategoryAxisTitle())) {
            catAxisTitleXml = "<c:title>"
                + "<c:tx><c:rich>"
                + "<a:bodyPr/><a:lstStyle/>"
                + "<a:p><a:r><a:t>" + escapeXml(chartConfig.getCategoryAxisTitle()) + "</a:t></a:r></a:p>"
                + "</c:rich></c:tx>"
                + "<c:overlay val=\"0\"/>"
                + "</c:title>";
        }

        // ---- Y轴标题（用 <c:rich>） ----
        String valAxisTitleXml = "";
        if (StringUtils.isNotBlank(chartConfig.getValueAxisTitle())) {
            valAxisTitleXml = "<c:title>"
                + "<c:tx><c:rich>"
                + "<a:bodyPr rot=\"-5400000\" vert=\"horz\"/><a:lstStyle/>"
                + "<a:p><a:r><a:t>" + escapeXml(chartConfig.getValueAxisTitle()) + "</a:t></a:r></a:p>"
                + "</c:rich></c:tx>"
                + "<c:overlay val=\"0\"/>"
                + "</c:title>";
        }

        // ---- X轴标签旋转 ----
        String catTxPrXml = "";
        if (chartConfig.getCategoryAxisRotation() != null && chartConfig.getCategoryAxisRotation() != 0) {
            int rot = -(chartConfig.getCategoryAxisRotation() * 60000);
            catTxPrXml = "<c:txPr><a:bodyPr rot=\"" + rot + "\"/><a:lstStyle/><a:p><a:pPr/></a:p></c:txPr>";
        }

        // ---- Y轴范围 ----
        String valScalingExtra = "";
        if (chartConfig.getValueAxisMax() != null) {
            valScalingExtra += "<c:max val=\"" + chartConfig.getValueAxisMax() + "\"/>";
        }
        if (chartConfig.getValueAxisMin() != null) {
            valScalingExtra += "<c:min val=\"" + chartConfig.getValueAxisMin() + "\"/>";
        }

        // ---- Y轴刻度单位 ----
        String valMajorUnitXml = "";
        if (chartConfig.getValueAxisUnit() != null) {
            valMajorUnitXml = "<c:majorUnit val=\"" + chartConfig.getValueAxisUnit() + "\"/>";
        }

        // ---- 图例 ----
        String legendXml = "";
        if (chartConfig.getLegendPosition() != ChartConfig.LegendPosition.NONE) {
            String pos;
            switch (chartConfig.getLegendPosition()) {
                case TOP:   pos = "t"; break;
                case LEFT:  pos = "l"; break;
                case RIGHT: pos = "r"; break;
                default:    pos = "b"; break;
            }
            legendXml = "<c:legend><c:legendPos val=\"" + pos + "\"/><c:overlay val=\"0\"/></c:legend>";
        }

        // ---- barDir & grouping ----
        String barDir   = chartConfig.isBarChart() ? "bar" : "col";
        String grouping = chartConfig.isStacked() ? "stacked" : "clustered";

        // ---- 组装完整 chartSpace XML（严格按 OOXML Schema 元素顺序） ----
        String chartSpaceXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<c:chartSpace"
            + "  xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""
            + "  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
            + "  xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<c:date1904 val=\"0\"/>"
            + "<c:chart>"
            // title
            + titleXml
            // autoTitleDeleted
            + "<c:autoTitleDeleted val=\"0\"/>"
            // plotArea
            + "<c:plotArea>"
            + "<c:layout/>"
            + "<c:barChart>"
            + "<c:barDir val=\"" + barDir + "\"/>"
            + "<c:grouping val=\"" + grouping + "\"/>"
            + "<c:varyColors val=\"0\"/>"
            + seriesXml
            + "<c:axId val=\"1\"/>"
            + "<c:axId val=\"2\"/>"
            + "</c:barChart>"
            // catAx — 严格按 CT_CatAx schema 顺序：axId,scaling,delete,axPos,title,
            //          numFmt,majorTickMark,minorTickMark,tickLblPos,spPr,txPr,crossAx,crosses/crossesAt,auto,lblAlign,lblOffset
            + "<c:catAx>"
            + "<c:axId val=\"1\"/>"
            + "<c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
            + "<c:delete val=\"0\"/>"
            + "<c:axPos val=\"b\"/>"
            + catAxisTitleXml
            + "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>"
            + "<c:tickLblPos val=\"nextTo\"/>"
            + catTxPrXml
            + "<c:crossAx val=\"2\"/>"
            + "<c:auto val=\"1\"/>"
            + "<c:lblAlgn val=\"ctr\"/>"
            + "<c:lblOffset val=\"100\"/>"
            + "</c:catAx>"
            // valAx — 严格按 CT_ValAx schema 顺序：axId,scaling,delete,axPos,title,
            //          numFmt,majorGridlines,minorGridlines,majorTickMark,minorTickMark,tickLblPos,spPr,txPr,crossAx,crosses/crossesAt,crossBetween,majorUnit,minorUnit
            + "<c:valAx>"
            + "<c:axId val=\"2\"/>"
            + "<c:scaling>"
            + "<c:orientation val=\"minMax\"/>"
            + valScalingExtra
            + "</c:scaling>"
            + "<c:delete val=\"0\"/>"
            + "<c:axPos val=\"l\"/>"
            + valAxisTitleXml
            + "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>"
            + "<c:tickLblPos val=\"nextTo\"/>"
            + "<c:crossAx val=\"1\"/>"
            + "<c:crosses val=\"autoZero\"/>"
            + "<c:crossBetween val=\"between\"/>"
            + valMajorUnitXml
            + "</c:valAx>"
            + "</c:plotArea>"
            // legend（在 plotArea 之后）
            + legendXml
            + "<c:plotVisOnly val=\"1\"/>"
            + "<c:dispBlanksAs val=\"zero\"/>"
            + "</c:chart>"
            + "<c:printSettings>"
            + "<c:headerFooter/>"
            + "<c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            + "<c:pageSetup/>"
            + "</c:printSettings>"
            + "</c:chartSpace>";

        // ---- 用 POI API 建立 drawing / chart 关系（占位用，XML内容由ZIP后处理替换）----
        // 注意：不调用 set()，让 POI 序列化出任意 chart1.xml 占位内容即可
        // fixChartXmlInZip 会用 chartSpaceXml 完整替换该文件

        return chartSpaceXml;
    }

    /**
     * 构建 Excel 单元格范围引用字符串，例如：Sheet1!$B$2:$B$51
     */
    private static String buildCellRange(String sheetName, int firstRow1, int lastRow1,
                                         int firstCol, int lastCol) {
        String colFirst = columnIndexToLetter(firstCol);
        String colLast  = columnIndexToLetter(lastCol);
        return "'" + sheetName + "'!$" + colFirst + "$" + firstRow1
             + ":$" + colLast + "$" + lastRow1;
    }

    /**
     * 列索引（0-based）转 Excel 列字母，例如 0→A，25→Z，26→AA
     */
    private static String columnIndexToLetter(int colIdx) {
        StringBuilder sb = new StringBuilder();
        int n = colIdx + 1;
        while (n > 0) {
            int rem = (n - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.toString();
    }

    /**
     * 在 ColumnConfig 列表中查找指定 fieldName 的列索引
     */
    private static int findColumnIndex(List<ColumnConfig> columns, String fieldName) {
        if (CollectionUtils.isEmpty(columns) || StringUtils.isBlank(fieldName)) return -1;
        for (int i = 0; i < columns.size(); i++) {
            if (fieldName.equals(columns.get(i).getFieldName())) return i;
        }
        return -1;
    }

    /**
     * 十六进制颜色字符串转 byte 数组（RGB），如 "FF0000" → {-1, 0, 0}
     */
    private static byte[] hexToBytes(String hex) {
        if (hex == null || hex.length() < 6) return new byte[]{0, 0, 0};
        return new byte[]{
            (byte) Integer.parseInt(hex.substring(0, 2), 16),
            (byte) Integer.parseInt(hex.substring(2, 4), 16),
            (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
    }

    /**
     * 修复 POI 生成的 xlsx ZIP 包里的图表问题。
     * <p>
     * 问题一：chart*.xml 由 POI 序列化时命名空间声明丢失，导致 Excel 无法解析。<br>
     * 修复：直接用 {@code chartXmlMap} 中正确构建的 XML 字符串整体替换。<br><br>
     * 问题二：drawing*.xml 中 POI 生成 {@code <xdr:cNvPr id="0">}，id=0 违反 OOXML 规范。<br>
     * 修复：将 id="0" 替换为 id="1"。
     * </p>
     *
     * @param xlsxBytes   workbook.write() 写出的原始字节
     * @param chartXmlMap key=ZIP条目路径(如"xl/charts/chart1.xml"), value=正确的 chartSpaceXml
     */
    private static byte[] fixChartXmlInZip(byte[] xlsxBytes,
                                            Map<String, String> chartXmlMap) throws IOException {
        java.io.ByteArrayInputStream  bais = new java.io.ByteArrayInputStream(xlsxBytes);
        java.io.ByteArrayOutputStream out  = new java.io.ByteArrayOutputStream();

        try (java.util.zip.ZipInputStream  zin  = new java.util.zip.ZipInputStream(bais);
             java.util.zip.ZipOutputStream zout = new java.util.zip.ZipOutputStream(out)) {

            java.util.zip.ZipEntry entry;
            byte[] buf = new byte[8192];

            while ((entry = zin.getNextEntry()) != null) {
                // 读取该条目的全部内容
                java.io.ByteArrayOutputStream entryCopy = new java.io.ByteArrayOutputStream();
                int len;
                while ((len = zin.read(buf)) != -1) {
                    entryCopy.write(buf, 0, len);
                }
                byte[] entryBytes = entryCopy.toByteArray();

                String entryName = entry.getName();

                // 对 chart*.xml 文件：整体替换为我们构建的正确 XML
                if (entryName.matches("xl/charts/chart[0-9]*.xml")) {
                    String correctXml = chartXmlMap.get(entryName);
                    if (correctXml != null) {
                        entryBytes = correctXml.getBytes(java.nio.charset.StandardCharsets.UTF_8);
                    }
                }

                // 对 drawing*.xml 文件：修复 cNvPr id="0" → id="1"
                if (entryName.matches("xl/drawings/drawing[0-9]*.xml")) {
                    String xml = new String(entryBytes, java.nio.charset.StandardCharsets.UTF_8);
                    xml = xml.replaceAll("(<xdr:cNvPr\\s[^>]*\\bid)=\"0\"", "$1=\"1\"");
                    entryBytes = xml.getBytes(java.nio.charset.StandardCharsets.UTF_8);
                }

                // 写入目标 ZIP
                java.util.zip.ZipEntry newEntry = new java.util.zip.ZipEntry(entryName);
                zout.putNextEntry(newEntry);
                zout.write(entryBytes);
                zout.closeEntry();
            }
        }
        return out.toByteArray();
    }

    /**
     * XML 内容转义：&amp; &lt; &gt; &quot; &apos;
     */
    private static String escapeXml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    /**
     * XML 属性值转义（用于 sheetName 等放进属性值的场景，同 escapeXml）
     */
    private static String escapeXmlAttr(String s) {
        return escapeXml(s);
    }

    
    private static String capitalize(String str) {
        if (StringUtils.isBlank(str)) {
            return str;
        }
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }
}
