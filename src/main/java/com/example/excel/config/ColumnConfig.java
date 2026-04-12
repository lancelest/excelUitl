package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 列配置类，定义 Excel 中每一列的字段映射、宽度、样式、数据转换规则。
 *
 * <p><b>核心字段关系：</b></p>
 * <ul>
 *   <li>{@code fieldName} — 绑定数据源字段（实体类属性名或 Map 的 key）</li>
 *   <li>{@code styleConfig} — 静态样式；{@code styleProvider} — 动态样式（优先级更高）</li>
 *   <li>{@code dataConverter} — 数据转换，在写入单元格前执行</li>
 * </ul>
 *
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@Accessors(chain = true)
public class ColumnConfig {

    /** 字段名，对应实体类属性名或 Map 的 key */
    private String fieldName;

    /** 列宽（字符数），实际像素宽度 = 字符数 × 256，默认 20 */
    private int width = 20;

    /** 列静态样式配置，优先级低于 StyleProvider */
    private CellStyleConfig styleConfig;

    /** 数字格式，如 {@code #,##0.00}、{@code 0.00%}、{@code 0.00E+00} */
    private String numberFormat;

    /** 日期格式，如 {@code yyyy-MM-dd HH:mm:ss} */
    private String dateFormat;

    /** 是否隐藏该列 */
    private boolean hidden = false;

    /** 数据转换器，在值写入单元格前执行自定义转换 */
    private DataConverter dataConverter;

    /** 动态样式提供器，优先级高于 styleConfig，为 null 时退回使用 styleConfig */
    private StyleProvider styleProvider;

    /**
     * 数据转换器接口
     * 
     * <p><b>功能说明：</b></p>
     * 用于在将数据写入Excel单元格之前,对原始数据进行自定义转换。
     * 支持类型转换、格式化、数据映射、业务逻辑处理等场景。
     * 
     * <p><b>适用场景：</b></p>
     * <ul>
     *   <li>枚举值转中文（如：1 → "男", 0 → "女"）</li>
     *   <li>状态码转描述（如：PENDING → "待处理"）</li>
     *   <li>数据格式化（如：小数转百分比，日期格式转换）</li>
     *   <li>字段映射（如：多个字段组合显示）</li>
     *   <li>空值处理（如：null → "—"）</li>
     *   <li>复杂业务逻辑（如：根据其他字段计算衍生值）</li>
     * </ul>
     * 
     * <p><b>使用示例：</b></p>
     * <pre>{@code
     * // 示例1：枚举值转中文
     * new ColumnConfig()
     *     .setFieldName("gender")
     *     .setDataConverter((value, rowData, rowIndex) -> {
     *         if (value == null) return "未知";
     *         return ((Integer) value) == 1 ? "男" : "女";
     *     });
     * 
     * // 示例2：小数转百分比
     * new ColumnConfig()
     *     .setFieldName("rate")
     *     .setDataConverter((value, rowData, rowIndex) -> {
     *         if (value == null) return "0%";
     *         double rate = ((Number) value).doubleValue();
     *         return String.format("%.2f%%", rate * 100);
     *     });
     * 
     * // 示例3：日期格式化
     * new ColumnConfig()
     *     .setFieldName("createTime")
     *     .setDataConverter((value, rowData, rowIndex) -> {
     *         if (value == null) return "";
     *         SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
     *         return sdf.format(value);
     *     });
     * 
     * // 示例4：多个字段组合显示
     * new ColumnConfig()
     *     .setFieldName("fullName")
     *     .setDataConverter((value, rowData, rowIndex) -> {
     *         User user = (User) rowData;
     *         return user.getLastName() + " " + user.getFirstName();
     *     });
     * }</pre>
     * 
     * <p><b>注意事项：</b></p>
     * <ul>
     *   <li>转换器返回的值会被设置到单元格中,可以是任意类型</li>
     *   <li>如果转换器抛出异常,会包装为ExcelExportException并中断导出</li>
     *   <li>转换器的执行顺序早于setCellValueSafe(),所以会影响后续的类型判断</li>
     *   <li>对于null值,建议在转换器中处理返回默认值,避免单元格为空</li>
     * </ul>
     */
    @FunctionalInterface
    public interface DataConverter {
        /**
         * 转换数据
         * 
         * @param value 原始值（通过getter/反射获取的字段值）
         * @param rowData 整行数据对象（可用于访问其他字段）
         * @param rowIndex 行索引（从0开始,可用于实现基于行号的逻辑）
         * @return 转换后的值（会被设置到单元格中）
         */
        Object convert(Object value, Object rowData, int rowIndex);
    }

    /**
     * 动态样式提供器接口
     *
     * <p><b>功能说明：</b></p>
     * 用于根据单元格的值、所在行数据或行索引，动态决定该单元格的样式。
     * 优先级高于列的静态 {@code styleConfig}，适用于条件高亮、状态着色等场景。
     *
     * <p><b>适用场景：</b></p>
     * <ul>
     *   <li>数值超阈值高亮（如：薪资 &gt; 20000 → 黄色字体+灰色背景）</li>
     *   <li>状态着色（如：状态为"异常" → 红色字体）</li>
     *   <li>奇偶行区分（如：偶数行浅蓝色背景）</li>
     *   <li>多条件复合判断（如：部门为"技术部" 且 绩效 &gt; 90 → 绿色）</li>
     * </ul>
     *
     * <p><b>使用示例：</b></p>
     * <pre>{@code
     * // 示例1：薪资超过20000时黄色字体+灰色背景
     * new ColumnConfig()
     *     .setFieldName("salary")
     *     .setStyleProvider((value, rowData, rowIndex) -> {
     *         if (value != null && ((Number) value).doubleValue() > 20000) {
     *             return StyleTemplate.DATA.toStyleConfig()
     *                 .setFontColor(IndexedColors.YELLOW.getIndex())
     *                 .setRgbBackgroundColor(StyleTemplate.rgb(128, 128, 128));
     *         }
     *         return StyleTemplate.DATA.toStyleConfig(); // 普通样式
     *     });
     *
     * // 示例2：状态列根据值着色
     * new ColumnConfig()
     *     .setFieldName("status")
     *     .setStyleProvider((value, rowData, rowIndex) -> {
     *         CellStyleConfig base = StyleTemplate.DATA.toStyleConfig();
     *         if ("异常".equals(value)) {
     *             return base.setFontColor(IndexedColors.RED.getIndex()).setBold(true);
     *         } else if ("正常".equals(value)) {
     *             return base.setFontColor(IndexedColors.GREEN.getIndex());
     *         }
     *         return base;
     *     });
     *
     * // 示例3：奇偶行区分
     * new ColumnConfig()
     *     .setFieldName("name")
     *     .setStyleProvider((value, rowData, rowIndex) -> {
     *         if (rowIndex % 2 == 0) {
     *             return StyleTemplate.DATA.toStyleConfig()
     *                 .setRgbBackgroundColor(StyleTemplate.rgb(235, 245, 255)); // 浅蓝
     *         }
     *         return StyleTemplate.DATA.toStyleConfig();
     *     });
     * }</pre>
     *
     * <p><b>注意事项：</b></p>
     * <ul>
     *   <li>返回 {@code null} 时自动降级使用列的 {@code styleConfig}</li>
     *   <li>每个单元格都会调用一次，对大数据量要注意性能（避免在内部 new 过多对象）</li>
     *   <li>与 {@code DataConverter} 独立，互不影响，可同时配置</li>
     *   <li>{@code value} 是经过 DataConverter 转换之前的原始值</li>
     * </ul>
     */
    @FunctionalInterface
    public interface StyleProvider {
        /**
         * 根据单元格值动态返回样式配置
         *
         * @param value    原始字段值（未经 DataConverter 转换）
         * @param rowData  整行数据对象（可访问其他字段）
         * @param rowIndex 行索引（从0开始，表头不计入）
         * @return 样式配置，返回 null 则使用列默认 styleConfig
         */
        CellStyleConfig provide(Object value, Object rowData, int rowIndex);
    }
}
