package com.example.excel.controller;

import com.example.excel.config.*;
import com.example.excel.exception.ExcelExportException;
import com.example.excel.model.User;
import com.example.excel.utils.ExcelExportUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.*;

/**
 * Excel导出Controller
 * 模拟实际使用场景,提供RESTful API接口给前端调用
 *
 * @author Excel Export Tool
 * @version 3.0.0
 */
public class ExcelExportController {

    /**
     * 导出报表接口
     *
     * @param type 报表类型: 1-员工报表, 2-技术报表, 3-管理层报表, ALL-所有报表
     * @param response HTTP响应
     */
    public void exportReport(String type, HttpServletResponse response) {
        try {
            // 使用新的Web导出方法，自动处理响应头和文件名
            ExcelExportUtils.exportToResponseWithTimestamp(
                response,
                getReportName(type),  // 基础文件名，自动加时间戳
                buildSheetConfigs(type).toArray(new SheetConfig[0])
            );
        } catch (ExcelExportException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            System.err.println("导出失败: " + e.getMessage());
        } catch (IOException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            System.err.println("IO异常: " + e.getMessage());
        }
    }

    /**
     * 根据报表类型构建Sheet配置
     *
     * @param type 报表类型
     * @return Sheet配置列表
     */
    private List<SheetConfig<?>> buildSheetConfigs(String type) {
        List<SheetConfig<?>> configs = new ArrayList<>();

        // type=ALL,生成所有报表
        if ("ALL".equalsIgnoreCase(type)) {
            configs.add(buildEmployeeSheet());
            configs.add(buildTechSheet());
            configs.add(buildManagerSheet());
        }
        // type=1,员工报表
        else if ("1".equals(type)) {
            configs.add(buildEmployeeSheet());
        }
        // type=2,技术报表
        else if ("2".equals(type)) {
            configs.add(buildTechSheet());
        }
        // type=3,管理层报表
        else if ("3".equals(type)) {
            configs.add(buildManagerSheet());
        }

        return configs;
    }

    /**
     * 获取报表基础名称（不含扩展名和时间戳）
     *
     * @param type 报表类型
     * @return 基础文件名
     */
    private String getReportName(String type) {
        switch (type) {
            case "1":
                return "员工报表";
            case "2":
                return "技术报表";
            case "3":
                return "管理层报表";
            case "ALL":
                return "综合报表";
            default:
                return "报表";
        }
    }

    // ==================== Sheet构建方法 ====================

    /**
     * 构建员工报表
     */
    private SheetConfig<User> buildEmployeeSheet() {
        // 准备数据
        List<User> dataList = generateEmployeeData(50);

        // 配置列
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
        columns.add(new ColumnConfig().setFieldName("salary").setWidth(12)
                .setNumberFormat("#,##0")
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));

        // 配置表头
        HeaderConfig header = new HeaderConfig()
                .addColumnNames("序号", "姓名", "部门", "职位", "入职日期", "薪资")
                .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());

        return new SheetConfig<User>()
                .setSheetName("员工报表")
                .setHeaders(Collections.singletonList(header))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);
    }

    /**
     * 构建技术报表
     */
    private SheetConfig<Map<String, Object>> buildTechSheet() {
        // 准备数据
        List<Map<String, Object>> dataList = generateTechData(30);

        // 配置列
        List<ColumnConfig> columns = new ArrayList<>();
        columns.add(new ColumnConfig().setFieldName("projectName").setWidth(20)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("developer").setWidth(12)
                .setStyleConfig(StyleTemplate.DATA.toStyleConfig()));
        columns.add(new ColumnConfig().setFieldName("status").setWidth(10)
                .setDataConverter(this::convertStatus)
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

        return new SheetConfig<Map<String, Object>>()
                .setSheetName("技术报表")
                .setHeaders(Collections.singletonList(header))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);
    }

    /**
     * 构建管理层报表
     */
    private SheetConfig<Map<String, Object>> buildManagerSheet() {
        // 准备数据
        List<Map<String, Object>> dataList = generateManagerData(10);

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

        return new SheetConfig<Map<String, Object>>()
                .setSheetName("管理层报表")
                .setHeaders(Collections.singletonList(header))
                .setColumnConfigs(columns)
                .setDataList(dataList)
                .setFreezeRow(1)
                .setDefaultDataStyle(StyleTemplate.DATA.toStyleConfig())
                .setBatchSize(100);
    }

    // ==================== 数据转换器 ====================

    /**
     * 状态转换
     */
    private Object convertStatus(Object value, Object rowData, int rowIndex) {
        if (value == null) {
            return "—";
        }
        String status = value.toString();
        switch (status) {
            case "PENDING":
                return "待开发";
            case "DEVELOPING":
                return "开发中";
            case "TESTING":
                return "测试中";
            case "COMPLETED":
                return "已完成";
            case "DEPLOYED":
                return "已上线";
            default:
                return status;
        }
    }

    // ==================== 数据生成方法 ====================

    /**
     * 生成员工数据
     */
    private List<User> generateEmployeeData(int count) {
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
            u.setSalary(8000 + r.nextDouble() * 20000);
            list.add(u);
        }
        return list;
    }

    /**
     * 生成技术数据
     */
    private List<Map<String, Object>> generateTechData(int count) {
        List<Map<String, Object>> list = new ArrayList<>();
        String[] projects = {"电商系统", "移动App", "数据分析", "CRM系统", "OA系统"};
        String[] developers = {"张三", "李四", "王五", "赵六", "孙七"};
        String[] statuses = {"PENDING", "DEVELOPING", "TESTING", "COMPLETED", "DEPLOYED"};
        Random r = new Random();

        for (int i = 1; i <= count; i++) {
            Map<String, Object> row = new HashMap<>();
            row.put("projectName", projects[i % projects.length] + "_" + i);
            row.put("developer", developers[r.nextInt(developers.length)]);
            row.put("status", statuses[r.nextInt(statuses.length)]);
            row.put("completionRate", 0.5 + r.nextDouble() * 0.5);
            row.put("linesOfCode", 1000 + r.nextInt(50000));
            list.add(row);
        }
        return list;
    }

    /**
     * 生成管理层数据
     */
    private List<Map<String, Object>> generateManagerData(int count) {
        List<Map<String, Object>> list = new ArrayList<>();
        String[] depts = {"技术部", "销售部", "市场部", "人事部", "财务部", "客服部"};
        Random r = new Random();

        for (int i = 0; i < count; i++) {
            Map<String, Object> row = new HashMap<>();
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
}
