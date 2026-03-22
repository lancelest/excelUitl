package com.example.excel.demo;

import com.example.excel.config.ColumnConfig;
import com.example.excel.config.HeaderConfig;
import com.example.excel.config.SheetConfig;
import com.example.excel.config.StyleTemplate;
import com.example.excel.model.User;
import com.example.excel.utils.ExcelExportUtils;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Random;

/**
 * 实际使用场景演示
 * 演示type=1/2/3/ALL的不同报表生成
 */
public class ReportExportDemo {

    public static void main(String[] args) {
        System.out.println("========================================");
        System.out.println("Excel导出工具 - 多报表场景演示");
        System.out.println("========================================");

        // 模拟参数
        String type = args.length > 0 ? args[0] : "ALL";
        System.out.println("导出类型: " + type);

        try {
            // 根据type导出不同报表
            switch (type) {
                case "1":
                    exportEmployeeReport();
                    break;
                case "2":
                    exportTechReport();
                    break;
                case "3":
                    exportManagerReport();
                    break;
                case "ALL":
                default:
                    exportAllReports();
                    break;
            }

            System.out.println("========================================");
            System.out.println("导出完成!");
            System.out.println("========================================");

        } catch (Exception e) {
            System.err.println("导出失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 导出员工报表(type=1)
     */
    private static void exportEmployeeReport() throws Exception {
        System.out.println("\n[场景1] 导出员工报表...");

        // 准备数据
        List<User> dataList = generateEmployeeData(50);

        // 使用createSimpleTemplate快速创建
        SheetConfig<User> sheet = ExcelExportUtils.createSimpleTemplate(
                "员工报表",
                dataList,
                new String[]{"seq", "name", "department", "position", "hireDate"},
                new String[]{"序号", "姓名", "部门", "职位", "入职日期"}
        );

        // 导出
        try (OutputStream os = new FileOutputStream("员工报表.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }

        System.out.println("  ✓ 员工报表导出成功");
    }

    /**
     * 导出技术报表(type=2)
     */
    private static void exportTechReport() throws Exception {
        System.out.println("\n[场景2] 导出技术报表...");

        // 准备数据
        List<DataRow> dataList = generateTechData(30);

        // 配置列
        List<ColumnConfig> columns = new ArrayList<>();
        columns.add(new ColumnConfig().setFieldName("projectName").setWidth(20)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("developer").setWidth(12)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("status").setWidth(10)
                .setDataConverter((value, rowData, rowIndex) -> {
                    if (value == null) return "—";
                    String status = value.toString();
                    switch (status) {
                        case "PENDING": return "待开发";
                        case "DEVELOPING": return "开发中";
                        case "TESTING": return "测试中";
                        case "COMPLETED": return "已完成";
                        case "DEPLOYED": return "已上线";
                        default: return status;
                    }
                })
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("completionRate").setWidth(12)
                .setNumberFormat("0.00%")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("linesOfCode").setWidth(12)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

        // 配置表头
        HeaderConfig header = new HeaderConfig()
                .addColumnNames("项目名称", "开发人员", "状态", "完成率", "代码行数")
                .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        // 配置Sheet
        SheetConfig<DataRow> sheet = new SheetConfig<DataRow>()
                .setSheetName("技术报表")
                .setHeaders(Collections.singletonList(header))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);

        // 导出
        try (OutputStream os = new FileOutputStream("技术报表.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }

        System.out.println("  ✓ 技术报表导出成功");
    }

    /**
     * 导出管理层报表(type=3)
     */
    private static void exportManagerReport() throws Exception {
        System.out.println("\n[场景3] 导出管理层报表...");

        // 准备数据
        List<DataRow> dataList = generateManagerData(10);

        // 配置列
        List<ColumnConfig> columns = new ArrayList<>();
        columns.add(new ColumnConfig().setFieldName("department").setWidth(15)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("headCount").setWidth(10)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("totalSalary").setWidth(15)
                .setNumberFormat("#,##0.00")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("avgPerformance").setWidth(12)
                .setNumberFormat("0.00")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("projectCount").setWidth(10)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

        // 配置表头
        HeaderConfig header = new HeaderConfig()
                .addColumnNames("部门", "人数", "总薪资", "平均绩效", "项目数")
                .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        // 配置Sheet
        SheetConfig<DataRow> sheet = new SheetConfig<DataRow>()
                .setSheetName("管理层报表")
                .setHeaders(Collections.singletonList(header))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);

        // 导出
        try (OutputStream os = new FileOutputStream("管理层报表.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }

        System.out.println("  ✓ 管理层报表导出成功");
    }

    /**
     * 导出所有报表(type=ALL)
     * 生成一个Excel文件,包含3个Sheet
     */
    private static void exportAllReports() throws Exception {
        System.out.println("\n[场景4] 导出所有报表(包含3个Sheet)...");

        // 准备数据
        List<User> employeeData = generateEmployeeData(50);
        List<DataRow> techData = generateTechData(30);
        List<DataRow> managerData = generateManagerData(10);

        // 员工报表
        SheetConfig<User> sheet1 = ExcelExportUtils.createSimpleTemplate(
                "员工报表",
                employeeData,
                new String[]{"seq", "name", "department", "position", "hireDate"},
                new String[]{"序号", "姓名", "部门", "职位", "入职日期"}
        );

        // 技术报表
        List<ColumnConfig> techColumns = new ArrayList<>();
        techColumns.add(new ColumnConfig().setFieldName("projectName").setWidth(20)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        techColumns.add(new ColumnConfig().setFieldName("developer").setWidth(12)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        techColumns.add(new ColumnConfig().setFieldName("status").setWidth(10)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        techColumns.add(new ColumnConfig().setFieldName("completionRate").setWidth(12)
                .setNumberFormat("0.00%")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        techColumns.add(new ColumnConfig().setFieldName("linesOfCode").setWidth(12)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        HeaderConfig techHeader = new HeaderConfig()
                .addColumnNames("项目名称", "开发人员", "状态", "完成率", "代码行数")
                .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());
        SheetConfig<DataRow> sheet2 = new SheetConfig<DataRow>()
                .setSheetName("技术报表")
                .setHeaders(Collections.singletonList(techHeader))
                .setColumnConfigs(techColumns)
                .setDataList(techData)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);

        // 管理层报表
        List<ColumnConfig> managerColumns = new ArrayList<>();
        managerColumns.add(new ColumnConfig().setFieldName("department").setWidth(15)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        managerColumns.add(new ColumnConfig().setFieldName("headCount").setWidth(10)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        managerColumns.add(new ColumnConfig().setFieldName("totalSalary").setWidth(15)
                .setNumberFormat("#,##0.00")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        managerColumns.add(new ColumnConfig().setFieldName("avgPerformance").setWidth(12)
                .setNumberFormat("0.00")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        managerColumns.add(new ColumnConfig().setFieldName("projectCount").setWidth(10)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        HeaderConfig managerHeader = new HeaderConfig()
                .addColumnNames("部门", "人数", "总薪资", "平均绩效", "项目数")
                .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());
        SheetConfig<DataRow> sheet3 = new SheetConfig<DataRow>()
                .setSheetName("管理层报表")
                .setHeaders(Collections.singletonList(managerHeader))
                .setColumnConfigs(managerColumns)
                .setDataList(managerData)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);

        // 导出所有Sheet到一个Excel文件
        try (OutputStream os = new FileOutputStream("综合报表.xlsx")) {
            ExcelExportUtils.export(os, sheet1, sheet2, sheet3);
        }

        System.out.println("  ✓ 综合报表导出成功(包含3个Sheet)");
    }

    /**
     * 演示自定义RGB颜色
     */
    private static void exportWithCustomColor() throws Exception {
        System.out.println("\n[演示5] 自定义RGB颜色...");

        List<User> dataList = generateEmployeeData(10);

        SheetConfig<User> sheet = ExcelExportUtils.createSimpleTemplate(
                "自定义颜色",
                dataList,
                new String[]{"name"},
                new String[]{"姓名"}
        );

        // 修改表头颜色为rgb(255, 100, 100) 浅红色
        sheet.getHeaders().get(0).setStyleConfig(
                StyleTemplate.HEADER.toStyleConfig()
                        .setRgbBackgroundColor(StyleTemplate.rgb(255, 100, 100))
        );

        // 修改数据颜色为rgb(240, 248, 255) 淡蓝色
        sheet.setDefaultDataStyle(
                StyleTemplate.DATA.toStyleConfig()
                        .setRgbBackgroundColor(StyleTemplate.rgb(240, 248, 255))
        );

        try (OutputStream os = new FileOutputStream("自定义颜色.xlsx")) {
            ExcelExportUtils.export(os, sheet);
        }

        System.out.println("  ✓ 自定义颜色导出成功");
        System.out.println("    - 表头颜色: rgb(255, 100, 100)");
        System.out.println("    - 数据颜色: rgb(240, 248, 255)");
    }

    // ==================== 数据生成方法 ====================

    private static List<User> generateEmployeeData(int count) {
        List<User> list = new ArrayList<>();
        String[] depts = {"技术部", "销售部", "市场部", "人事部", "财务部"};
        String[] positions = {"工程师", "经理", "总监", "专员", "主管"};
        Random r = new Random();

        for (int i = 1; i <= count; i++) {
            User u = new User();
            u.setSeq(i);
            u.setName("员工" + String.format("%03d", i));
            u.setDepartment(depts[r.nextInt(depts.length)]);
            u.setPosition(positions[r.nextInt(positions.length)]);
            u.setHireDate(String.format("202%d-%02d-%02d", r.nextInt(6), r.nextInt(12) + 1, r.nextInt(28) + 1));
            list.add(u);
        }
        return list;
    }

    private static List<DataRow> generateTechData(int count) {
        List<DataRow> list = new ArrayList<>();
        String[] projects = {"电商系统", "移动App", "数据分析", "CRM系统", "OA系统"};
        String[] developers = {"张三", "李四", "王五", "赵六", "孙七"};
        String[] statuses = {"PENDING", "DEVELOPING", "TESTING", "COMPLETED", "DEPLOYED"};
        Random r = new Random();

        for (int i = 1; i <= count; i++) {
            DataRow row = new DataRow();
            row.put("projectName", projects[i % projects.length] + "_" + i);
            row.put("developer", developers[r.nextInt(developers.length)]);
            row.put("status", statuses[r.nextInt(statuses.length)]);
            row.put("completionRate", 0.5 + r.nextDouble() * 0.5);
            row.put("linesOfCode", 1000 + r.nextInt(50000));
            list.add(row);
        }
        return list;
    }

    private static List<DataRow> generateManagerData(int count) {
        List<DataRow> list = new ArrayList<>();
        String[] depts = {"技术部", "销售部", "市场部", "人事部", "财务部", "客服部"};
        Random r = new Random();

        for (int i = 0; i < count; i++) {
            DataRow row = new DataRow();
            int headCount = 5 + r.nextInt(50);
            row.put("department", depts[i % depts.length]);
            row.put("headCount", headCount);
            row.put("totalSalary", headCount * (10000 + r.nextDouble() * 15000));
            row.put("avgPerformance", 70 + r.nextDouble() * 30);
            row.put("projectCount", 3 + r.nextInt(20));
            list.add(row);
        }
        return list;
    }

    /**
     * 通用数据行(用于Map数据源)
     */
    static class DataRow {
        private java.util.Map<String, Object> data = new java.util.HashMap<>();

        public void put(String key, Object value) {
            data.put(key, value);
        }

        public Object get(String key) {
            return data.get(key);
        }
    }
}
