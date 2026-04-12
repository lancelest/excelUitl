package com.example.excel.test;

import com.example.excel.config.*;
import com.example.excel.model.User;
import com.example.excel.utils.ExcelExportUtils;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

/**
 * Excel导出测试类
 * 
 * @author Excel Export Tool
 * @version 3.0.0
 */
public class ExcelExportTest {

    public static void main(String[] args) {
        System.out.println("========== Excel导出工具测试 ==========\n");

        // 测试1：基础多级表头导出
//        exportExcel();

        // 测试2：StyleProvider 条件样式示例
//        testStyleProvider();

        // 测试3：图表导出（柱状图）
//        testChartExport();

        // 测试4：堆积柱图 & 百分比堆积柱图
        testStackedChartExport();

        System.out.println("========================================");
    }

    /**
     * 导出Excel
     */
    private static void exportExcel() {
        try {
            // 1. 准备数据（50条）
            List<User> dataList = generateUserData(50);

            // 2. 使用样式模板定义表头样式：红色背景、白色字体、加粗
            HeaderConfig header1 = new HeaderConfig()
                    .addColumnNames("序号", "基本信息", "", "", "", "绩效数据", "", "", "", "备注")
                    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
                    .addMergeRegion(0, 1, 0, 0)   // 第1列合并第1-2行
                    .addMergeRegion(0, 0, 1, 4)   // 第2-5列合并（基本信息）
                    .addMergeRegion(0, 0, 5, 8) ;  // 第6-9列合并（绩效数据）
                    //.addMergeRegion(0, 1, 9, 9);  // 第10列合并第1-2行

            // 3. 配置二级表头（第2行）
            HeaderConfig header2 = new HeaderConfig()
                    .addColumnNames("", "姓名", "部门", "职位", "入职日期", "绩效评分", "项目数", "满意度", "薪资", "备注")
                    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig()
                     .setRgbBackgroundColor(StyleTemplate.rgb(0, 255, 0)));

            // 4. 配置列（共10列）
            List<ColumnConfig> columns = new ArrayList<>();

            // 第1列：序号（居中）
            columns.add(new ColumnConfig()
                    .setFieldName("seq").setWidth(8).setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            // 第2-5列：基本信息（左对齐）
            columns.add(new ColumnConfig().setFieldName("name").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("department").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("position").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("hireDate").setWidth(14)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

            // 第6-9列：绩效数据（右对齐，带格式）
            columns.add(new ColumnConfig().setFieldName("performanceScore").setWidth(12)
                    .setNumberFormat("0.00")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("projectCount").setWidth(10)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("satisfaction").setWidth(10)
                    .setNumberFormat("0.0%")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("salary").setWidth(12)
                    .setNumberFormat("#,##0")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

            // 第10列：备注（左对齐）
            columns.add(new ColumnConfig().setFieldName("remark").setWidth(15)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

            // 5. 配置Sheet
            SheetConfig<User> sheet = new SheetConfig<User>()
                    .setSheetName("员工报表")
                    .setHeaders(Arrays.asList(header1, header2))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(3)            // 冻结前2行（两个表头行）
                    .setFreezeCol(0)            // 不冻结列
                    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())  // 默认数据样式
                    .setBatchSize(25);

SheetConfig<User> sheet1 = new SheetConfig<User>()
                    .setSheetName("员工报表1")
                    .setHeaders(Arrays.asList(header1, header2))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(3)            // 冻结前2行（两个表头行）
                    .setFreezeCol(0)            // 不冻结列
                    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())  // 默认数据样式
                    .setBatchSize(25);

            // 6. 导出
            File file = new File("D:\\excelData\\员工报表v1.xlsx");
            File parent = file.getParentFile();
            if (parent != null && !parent.exists()) {
                parent.mkdirs();
            }

            try (OutputStream os = new FileOutputStream(file)) {
                ExcelExportUtils.export(os, sheet,sheet1);
            }

        } catch (Exception e) {
            System.err.println("  ✗ 测试失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 测试StyleProvider动态样式
     * 演示4个典型场景：数值阈值、状态着色、多字段判断、奇偶行
     */
    private static void testStyleProvider() {
        System.out.println("  开始测试：StyleProvider 条件样式");
        try {
            // 生成包含各种情况的测试数据
            List<User> dataList = generateUserData(30);

            // ========== 场景1：薪资 > 20000 → 黄色字体 + 灰色背景 ==========
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

            // ========== 场景2：绩效分数 < 70 → 红色加粗，>= 90 → 绿色 ==========
            ColumnConfig scoreColumn = new ColumnConfig()
                .setFieldName("performanceScore")
                .setWidth(12)
                .setNumberFormat("0.00")
                .setStyleProvider((value, rowData, rowIndex) -> {
                    if (value != null) {
                        double score = ((Number) value).doubleValue();
                        CellStyleConfig base = StyleTemplate.DATA.toStyleConfig();
                        if (score < 70) {
                            // 不达标：红色加粗
                            return base.setFontColor(IndexedColors.RED.getIndex()).setBold(true);
                        } else if (score >= 90) {
                            // 优秀：绿色
                            return base.setFontColor(IndexedColors.DARK_GREEN.getIndex())
                                .setRgbBackgroundColor(StyleTemplate.rgb(198, 239, 206)); // 浅绿背景
                        }
                    }
                    return StyleTemplate.DATA.toStyleConfig();
                });

            // ========== 场景3：满意度 > 0.9 → 紫色高亮（多字段判断）==========
            ColumnConfig satisfactionColumn = new ColumnConfig()
                .setFieldName("satisfaction")
                .setWidth(12)
                .setNumberFormat("0.0%")
                .setStyleProvider((value, rowData, rowIndex) -> {
                    User emp = (User) rowData;
                    if ("技术部".equals(emp.getDepartment())
                            && value != null
                            && ((Number) value).doubleValue() > 0.9) {
                        return StyleTemplate.DATA.toStyleConfig()
                            .setFontColor(IndexedColors.GREEN.getIndex())
                            .setRgbBackgroundColor(StyleTemplate.rgb(230, 230, 250))
                            .setBold(true);
                    }
                    return StyleTemplate.DATA.toStyleConfig();
                });

            // ========== 场景4：奇偶行区分（斑马纹）==========
            ColumnConfig.StyleProvider zebraStyle = (value, rowData, rowIndex) -> {
                if (rowIndex % 2 == 0) {
                    return StyleTemplate.DATA.toStyleConfig()
                        .setRgbBackgroundColor(StyleTemplate.rgb(235, 245, 255)); // 浅蓝
                }
                return StyleTemplate.DATA.toStyleConfig(); // 奇数行默认
            };

            // 构建列配置（前3列应用斑马纹，后3列分别应用条件样式）
            List<ColumnConfig> columns = new ArrayList<>();
            columns.add(new ColumnConfig()
                .setFieldName("seq").setWidth(8).setStyleProvider(zebraStyle));
            columns.add(new ColumnConfig()
                .setFieldName("name").setWidth(12).setStyleProvider(zebraStyle));
            columns.add(new ColumnConfig()
                .setFieldName("department").setWidth(12).setStyleProvider(zebraStyle));

            columns.add(scoreColumn);    // 绩效：按分数着色
            columns.add(satisfactionColumn); // 满意度：多字段判断
            columns.add(salaryColumn);   // 薪资：阈值高亮

            // 表头
            HeaderConfig header = new HeaderConfig()
                .addColumnNames("", "姓名", "部门", "绩效评分", "满意度", "薪资")
                .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

            // 2. 使用样式模板定义表头样式：红色背景、白色字体、加粗
            HeaderConfig header1 = new HeaderConfig()
                    .addColumnNames("序号", "基本信息", "基本信息",  "绩效数据", "绩效数据", "绩效数据")
                    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig())
                    .addMergeRegion(0, 1, 0, 0)   // 第1列合并第1-2行
                    .addMergeRegion(0, 0, 1, 3)   // 第2-5列合并（基本信息）
                    .addMergeRegion(0, 0, 4, 5)   // 第6-9列合并（绩效数据）
                    ;  // 第10列合并第1-2行

            // 3. 配置二级表头（第2行）
            HeaderConfig header2 = new HeaderConfig()
                    .addColumnNames("序号", "姓名", "部门", "绩效评分", "满意度", "薪资")
                    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig()
                            .setRgbBackgroundColor(StyleTemplate.rgb(0, 255, 0)));

            // 配置Sheet
            SheetConfig<User> sheet = new SheetConfig<User>()
                .setSheetName("条件样式演示")
                .setHeaders(Arrays.asList(header1, header2))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(3)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(50);

            // 导出
            File file = new File("D:\\excelData\\条件样式演示.xlsx");
            File parent = file.getParentFile();
            if (parent != null && !parent.exists()) {
                parent.mkdirs();
            }

            try (OutputStream os = new FileOutputStream(file)) {
                ExcelExportUtils.export(os, sheet);
            }

            System.out.println("  ✓ 条件样式导出成功，文件：" + file.getAbsolutePath());
            System.out.println("    - 薪资 > 20000 → 黄色字体+灰色背景+加粗");
            System.out.println("    - 绩效 < 70 → 红色加粗，>= 90 → 绿色");
            System.out.println("    - 技术部且满意度 > 0.9 → 紫色高亮");
            System.out.println("    - 偶数行 → 浅蓝色背景（斑马纹）");

        } catch (Exception e) {
            System.err.println("  ✗ StyleProvider 测试失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 测试图表导出
     * 演示：在员工数据下方自动生成绩效评分 + 薪资柱状图
     * - X轴：部门（department 字段）
     * - Y轴系列1：绩效评分（performanceScore）
     * - Y轴系列2：项目数（projectCount）
     * - 标签旋转 45°，显示数据标签，自动位置
     */
    private static void testChartExport() {
        System.out.println("  开始测试：图表导出");
        try {
            // 1. 准备数据（20条，覆盖多个部门）
            List<User> dataList = generateUserData(20);

            // 2. 图表配置
            ChartConfig chartConfig = new ChartConfig()
                    .setTitle("员工绩效概览")
                    .setCategoryColumn("name")           // X轴用姓名
                    .addSeries("performanceScore", "绩效评分")
                    .addSeries("projectCount", "项目数")
                    .addSeries("satisfaction", "满意度")
                    .setShowDataLabel(true)
                    .setCategoryAxisRotation(45)          // X轴标签旋转45°
                    .setCategoryAxisTitle("员工姓名")
                    .setValueAxisTitle("数值")
                    .setLegendPosition(ChartConfig.LegendPosition.BOTTOM)
                    // 系列颜色：蓝色 + 橙色
                    .setSeriesColors(Arrays.asList("4472C4", "ED7D31"))
                    .setFontSize(10)                     // 设置字体大小为12号
                    .setGapRows(2);

            // 3. 列配置（与 exportExcel 保持一致的字段）
            List<ColumnConfig> columns = new ArrayList<>();
            columns.add(new ColumnConfig().setFieldName("seq").setWidth(8)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("name").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("department").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("position").setWidth(12)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("hireDate").setWidth(14)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("performanceScore").setWidth(12)
                    .setNumberFormat("0.00")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("projectCount").setWidth(10)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("satisfaction").setWidth(10)
                    .setNumberFormat("0.0%")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("salary").setWidth(12)
                    .setNumberFormat("#,##0")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("remark").setWidth(15)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

            // 4. 表头（单级，简洁）
            HeaderConfig header = new HeaderConfig()
                    .addColumnNames("序号", "姓名", "部门", "职位", "入职日期",
                            "绩效评分", "项目数", "满意度", "薪资", "备注")
                    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

            // 5. Sheet配置，挂上 chartConfig
            SheetConfig<User> sheet = new SheetConfig<User>()
                    .setSheetName("员工绩效图表")
                    .setHeaders(Arrays.asList(header))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(2)
                    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                    .setBatchSize(50)
                    .setChartConfig(chartConfig);   // ← 关键：设置图表配置

            // 6. 导出
            File file = new File("D:\\excelData\\员工绩效-图表.xlsx");
            File parent = file.getParentFile();
            if (parent != null && !parent.exists()) {
                parent.mkdirs();
            }
            try (OutputStream os = new FileOutputStream(file)) {
                ExcelExportUtils.export(os, sheet);
            }

            System.out.println("  ✓ 图表导出成功，文件：" + file.getAbsolutePath());
            System.out.println("    - X轴：员工姓名（旋转45°）");
            System.out.println("    - 系列1：绩效评分（蓝色）");
            System.out.println("    - 系列2：项目数（橙色）");
            System.out.println("    - 显示数据标签，图例在底部");

        } catch (Exception e) {
            System.err.println("  ✗ 图表测试失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 测试堆积柱图 & 百分比堆积柱图
     * 演示：同一份数据用三种分组方式导出，方便对比效果
     *  - Sheet1：簇状柱图（CLUSTERED，原有默认）
     *  - Sheet2：堆积柱图（STACKED）
     *  - Sheet3：百分比堆积柱图（PERCENT_STACKED）
     */
    private static void testStackedChartExport() {
        System.out.println("  开始测试：堆积柱图 & 百分比堆积柱图");
        try {
            List<User> dataList = generateUserData(15);

            // ---- 公用列配置 ----
            List<ColumnConfig> columns = new ArrayList<>();
            columns.add(new ColumnConfig().setFieldName("seq").setWidth(8)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("name").setWidth(14)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("performanceScore").setWidth(12)
                    .setNumberFormat("0.00")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("projectCount").setWidth(10)
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
            columns.add(new ColumnConfig().setFieldName("satisfaction").setWidth(10)
                    .setNumberFormat("0.0%")
                    .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

            HeaderConfig header = new HeaderConfig()
                    .addColumnNames("序号", "姓名", "绩效评分", "项目数", "满意度")
                    .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

            // ---- 工厂：根据分组方式生成 ChartConfig ----
            // Sheet1：簇状（默认，CLUSTERED）
            ChartConfig clusteredChart = new ChartConfig()
                    .setTitle("绩效概览 - 簇状柱图")
                    .setCategoryColumn("name")
                    .addSeries("performanceScore", "绩效评分")
                    .addSeries("projectCount", "项目数")
                    .setBarGrouping(ChartConfig.BarGrouping.CLUSTERED)
                    .setCategoryAxisRotation(45)
                    .setShowDataLabel(true)
                    .setLegendPosition(ChartConfig.LegendPosition.BOTTOM)
                    .setFontSize(10);

            // Sheet2：堆积（STACKED）
            ChartConfig stackedChart = new ChartConfig()
                    .setTitle("绩效概览 - 堆积柱图")
                    .setCategoryColumn("name")
                    .addSeries("performanceScore", "绩效评分")
                    .addSeries("projectCount", "项目数")
                    .setBarGrouping(ChartConfig.BarGrouping.STACKED)
                    .setCategoryAxisRotation(45)
                    .setShowDataLabel(false)      // 堆积柱图数据标签容易重叠，关掉
                    .setLegendPosition(ChartConfig.LegendPosition.BOTTOM)
                    .setSeriesColors(Arrays.asList("4472C4", "ED7D31"))
                    .setFontSize(10);

            // Sheet3：百分比堆积（PERCENT_STACKED）
            ChartConfig pctStackedChart = new ChartConfig()
                    .setTitle("绩效概览 - 百分比堆积柱图")
                    .setCategoryColumn("name")
                    .addSeries("performanceScore", "绩效评分")
                    .addSeries("projectCount", "项目数")
                    .setBarGrouping(ChartConfig.BarGrouping.PERCENT_STACKED)
                    .setCategoryAxisRotation(45)
                    .setShowDataLabel(false)
                    .setLegendPosition(ChartConfig.LegendPosition.RIGHT)
                    .setSeriesColors(Arrays.asList("70AD47", "FFC000"))
                    .setFontSize(10);

            // ---- 三个 Sheet ----
            SheetConfig<User> sheet1 = new SheetConfig<User>()
                    .setSheetName("簇状柱图")
                    .setHeaders(Arrays.asList(header))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(2)
                    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                    .setBatchSize(50)
                    .setChartConfig(clusteredChart);

            SheetConfig<User> sheet2 = new SheetConfig<User>()
                    .setSheetName("堆积柱图")
                    .setHeaders(Arrays.asList(header))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(2)
                    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                    .setBatchSize(50)
                    .setChartConfig(stackedChart);

            SheetConfig<User> sheet3 = new SheetConfig<User>()
                    .setSheetName("百分比堆积")
                    .setHeaders(Arrays.asList(header))
                    .setColumnConfigs(columns)
                    .setDataList(dataList)
                    .setFreezeRow(2)
                    .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                    .setBatchSize(50)
                    .setChartConfig(pctStackedChart);

            File file = new File("D:\\excelData\\堆积柱图对比.xlsx");
            File parent = file.getParentFile();
            if (parent != null && !parent.exists()) {
                parent.mkdirs();
            }
            try (OutputStream os = new FileOutputStream(file)) {
                ExcelExportUtils.export(os, sheet1, sheet2, sheet3);
            }

            System.out.println("  ✓ 堆积柱图导出成功，文件：" + file.getAbsolutePath());
            System.out.println("    - Sheet1：簇状柱图（CLUSTERED）");
            System.out.println("    - Sheet2：堆积柱图（STACKED）");
            System.out.println("    - Sheet3：百分比堆积柱图（PERCENT_STACKED）");

        } catch (Exception e) {
            System.err.println("  ✗ 堆积柱图测试失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 生成用户数据
     */
    private static List<User> generateUserData(int count) {
        List<User> list = new ArrayList<>();
        String[] depts = {"技术部", "销售部", "市场部", "人事部", "财务部"};
        String[] positions = {"工程师", "经理", "总监", "专员", "主管"};
        Random r = new Random();

        for (int i = 1; i <= count; i++) {
            User u = new User();
            u.setSeq(i);
            u.setName("北京天坛医院员工" + String.format("%03d", i));
            u.setDepartment(depts[r.nextInt(depts.length)]);
            u.setPosition(positions[r.nextInt(positions.length)]);
            u.setHireDate(String.format("202%d-%02d-%02d", r.nextInt(6), r.nextInt(12) + 1, r.nextInt(28) + 1));
            u.setPerformanceScore(60 + r.nextDouble() * 40);
            u.setProjectCount(r.nextInt(50) + 1);
            u.setSatisfaction(0.7 + r.nextDouble() * 0.3);
            u.setSalary(8000 + r.nextDouble() * 20000);
            u.setRemark("备注信息" + i);
            list.add(u);
        }
        return list;
    }
}
