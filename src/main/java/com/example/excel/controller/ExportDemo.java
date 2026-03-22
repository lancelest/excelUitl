package com.example.excel.controller;

import com.example.excel.utils.ExcelExportUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * 实际使用场景演示
 * 模拟前端调用导出接口
 *
 * @author Excel Export Tool
 * @version 3.0.0
 */
public class ExportDemo {

    public static void main(String[] args) {
        System.out.println("========================================");
        System.out.println("Excel导出工具 - 实际使用场景演示");
        System.out.println("========================================");

        // 场景1: 导出单个报表(员工报表)
        exportSingleReport();

        // 场景2: 导出多个报表(type=ALL)
        exportAllReports();

        // 场景3: 演示RGB颜色自定义
        exportWithCustomColor();

        System.out.println("========================================");
        System.out.println("所有演示完成!");
        System.out.println("========================================");
    }

    /**
     * 场景1: 导出单个报表
     * <p>type=1,只生成员工报表</p>
     */
    private static void exportSingleReport() {
        System.out.println("\n【场景1】导出单个报表(type=1,员工报表)");

        // 创建Controller实例
        ExcelExportController controller = new ExcelExportController();

        // 模拟HTTP响应(实际场景中由Spring MVC提供)
        MockHttpServletResponse response = new MockHttpServletResponse();

        // 调用导出接口
        controller.exportReport("1", (HttpServletResponse) response);

        // 保存文件
        saveToFile(response.getData(), "员工报表.xlsx");

        System.out.println("  ✓ 员工报表导出成功");
    }

    /**
     * 场景2: 导出多个报表
     * <p>type=ALL,生成所有报表(员工报表、技术报表、管理层报表)</p>
     */
    private static void exportAllReports() {
        System.out.println("\n【场景2】导出多个报表(type=ALL)");

        // 创建Controller实例
        ExcelExportController controller = new ExcelExportController();

        // 模拟HTTP响应
        MockHttpServletResponse response = new MockHttpServletResponse();

        // 调用导出接口,生成所有报表
        controller.exportReport("ALL", (HttpServletResponse) response);

        // 保存文件
        saveToFile(response.getData(), "综合报表.xlsx");

        System.out.println("  ✓ 综合报表导出成功(包含3个Sheet)");
    }

    /**
     * 场景3: 使用自定义RGB颜色
     * <p>演示如何使用StyleTemplate.rgb()方法自定义颜色</p>
     */
    private static void exportWithCustomColor() {
        System.out.println("\n【场景3】自定义RGB颜色");

        try {
            // 使用ExcelExportUtils直接导出
            // 这里演示如何使用自定义颜色

            // 创建简单模板
            com.example.excel.config.SheetConfig<User> sheet =
                    ExcelExportUtils.createSimpleTemplate(
                            "自定义颜色演示",
                            generateTestData(10),
                            new String[]{"name", "age", "department"},
                            new String[]{"姓名", "年龄", "部门"}
                    );

            // 修改表头颜色为自定义RGB(255, 100, 100) - 浅红色
            sheet.getHeaders().get(0).setStyleConfig(
                    com.example.excel.config.StyleTemplate.HEADER.toStyleConfig()
                            .setRgbBackgroundColor(com.example.excel.config.StyleTemplate.rgb(255, 100, 100))
            );

            // 修改数据颜色为自定义RGB(240, 248, 255) - 淡蓝色
            sheet.setDefaultDataStyle(
                    com.example.excel.config.StyleTemplate.DATA.toStyleConfig()
                            .setRgbBackgroundColor(com.example.excel.config.StyleTemplate.rgb(240, 248, 255))
            );

            // 导出
            OutputStream os = new FileOutputStream("自定义颜色演示.xlsx");
            ExcelExportUtils.export(os, sheet);
            os.close();

            System.out.println("  ✓ 自定义颜色演示导出成功");
            System.out.println("  - 表头颜色: rgb(255, 100, 100) 浅红色");
            System.out.println("  - 数据颜色: rgb(240, 248, 255) 淡蓝色");

        } catch (Exception e) {
            System.out.println("  ✗ 导出失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 保存数据到文件
     */
    private static void saveToFile(byte[] data, String fileName) {
        try (OutputStream os = new FileOutputStream(fileName)) {
            os.write(data);
            System.out.println("  文件已保存: " + fileName);
        } catch (Exception e) {
            System.out.println("  ✗ 保存失败: " + e.getMessage());
        }
    }

    /**
     * 生成测试数据
     */
    private static java.util.List<User> generateTestData(int count) {
        java.util.List<User> list = new java.util.ArrayList<>();
        for (int i = 1; i <= count; i++) {
            User u = new User();
            u.setSeq(i);
            u.setName("测试用户" + i);
            list.add(u);
        }
        return list;
    }

    /**
     * 模拟HTTP响应
     */
    static class MockHttpServletResponse implements HttpServletResponse {
        private byte[] data;
        private int status = 200;
        private String contentType;
        private String characterEncoding;
        private java.util.Map<String, java.util.List<String>> headers = new java.util.HashMap<>();

        public byte[] getData() {
            return data;
        }

        @Override
        public void setStatus(int status) {
            this.status = status;
        }

        @Override
        public int getStatus() {
            return status;
        }

        @Override
        public void setContentType(String contentType) {
            this.contentType = contentType;
        }

        @Override
        public String getContentType() {
            return contentType;
        }

        @Override
        public void setCharacterEncoding(String encoding) {
            this.characterEncoding = encoding;
        }

        @Override
        public String getCharacterEncoding() {
            return characterEncoding;
        }

        @Override
        public void setHeader(String name, String value) {
            java.util.List<String> values = new java.util.ArrayList<>();
            values.add(value);
            headers.put(name, values);
        }

        @Override
        public void addHeader(String name, String value) {
            headers.computeIfAbsent(name, k -> new java.util.ArrayList<>()).add(value);
        }

        @Override
        public javax.servlet.ServletOutputStream getOutputStream() throws java.io.IOException {
            return new javax.servlet.ServletOutputStream() {
                private java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
                
                @Override
                public boolean isReady() {
                    return true;
                }
                
                @Override
                public void setWriteListener(javax.servlet.WriteListener writeListener) {
                    // 不需要实现
                }
                
                @Override
                public void write(int b) throws java.io.IOException {
                    baos.write(b);
                }
                
                @Override
                public void flush() throws java.io.IOException {
                    data = baos.toByteArray();
                }
            };
        }

        @Override
        public java.io.PrintWriter getWriter() throws java.io.IOException {
            return new java.io.PrintWriter(new java.io.ByteArrayOutputStream());
        }

        @Override
        public void setContentLength(int len) {}

        @Override
        public void setContentLengthLong(long len) {}

        @Override
        public void addCookie(javax.servlet.http.Cookie cookie) {}

        @Override
        public boolean containsHeader(String name) {
            return headers.containsKey(name);
        }

        @Override
        public String encodeURL(String url) {
            return url;
        }

        @Override
        public String encodeRedirectURL(String url) {
            return url;
        }

        @Override
        public void sendError(int sc, String msg) throws java.io.IOException {
            this.status = sc;
        }

        @Override
        public void sendError(int sc) throws java.io.IOException {
            this.status = sc;
        }

        @Override
        public void sendRedirect(String location) throws java.io.IOException {
            // 模拟重定向
        }

        @Override
        public void setDateHeader(String name, long date) {
            setHeader(name, String.valueOf(date));
        }

        @Override
        public void addDateHeader(String name, long date) {
            addHeader(name, String.valueOf(date));
        }

        @Override
        public void setIntHeader(String name, int value) {
            setHeader(name, String.valueOf(value));
        }

        @Override
        public void addIntHeader(String name, int value) {
            addHeader(name, String.valueOf(value));
        }

        @Override
        public String getHeader(String name) {
            java.util.List<String> values = headers.get(name);
            return (values != null && !values.isEmpty()) ? values.get(0) : null;
        }

        @Override
        public java.util.Collection<String> getHeaders(String name) {
            return headers.getOrDefault(name, java.util.Collections.emptyList());
        }

        @Override
        public java.util.Collection<String> getHeaderNames() {
            return headers.keySet();
        }

        @Override
        public String encodeUrl(String url) {
            return url;
        }

        @Override
        public String encodeRedirectUrl(String url) {
            return url;
        }

        @Override
        public void setStatus(int sc, String sm) {
            this.status = sc;
        }

        @Override
        public int getBufferSize() {
            return 8192;
        }

        @Override
        public void setBufferSize(int size) {}

        @Override
        public void flushBuffer() throws java.io.IOException {}

        @Override
        public void resetBuffer() {}

        @Override
        public boolean isCommitted() {
            return false;
        }

        @Override
        public void reset() {
            data = null;
            status = 200;
            headers.clear();
        }

        @Override
        public void setLocale(java.util.Locale loc) {}

        @Override
        public java.util.Locale getLocale() {
            return java.util.Locale.getDefault();
        }
    }

    /**
     * 用户实体(复用)
     */
    public static class User {
        private Integer seq;
        private String name;
        private String age;
        private String department;

        public Integer getSeq() {
            return seq;
        }

        public void setSeq(Integer seq) {
            this.seq = seq;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getAge() {
            return age;
        }

        public void setAge(String age) {
            this.age = age;
        }

        public String getDepartment() {
            return department;
        }

        public void setDepartment(String department) {
            this.department = department;
        }
    }
}
