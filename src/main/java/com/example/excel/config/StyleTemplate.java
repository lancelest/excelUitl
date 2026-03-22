package com.example.excel.config;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 样式模板类
 * <p>
 * 提供预定义的常用样式模板,避免重复配置,增强类型安全。
 * 所有的样式模板都可以直接使用,也可以作为基础样式进行修改。
 * </p>
 *
 * <p><b>使用示例：</b></p>
 * <pre>{@code
 * // 直接使用预定义模板
 * ColumnConfig column = new ColumnConfig()
 *     .setFieldName("name")
 *     .setStyleConfig(StyleTemplate.HEADER.toStyleConfig());
 *
 * // 使用自定义RGB颜色
 * ColumnConfig column = new ColumnConfig()
 *     .setFieldName("name")
 *     .setStyleConfig(StyleTemplate.HEADER.toStyleConfig()
 *         .setBackgroundColor(StyleTemplate.rgb(255, 0, 0)));
 * }</pre>
 *
 * @author Excel Export Tool
 * @version 3.0.0
 */
public enum StyleTemplate {

    /**
     * 表头样式 - 红色背景,白色字体,加粗
     * <p>适用场景：所有报表的表头</p>
     */
    HEADER(new CellStyleConfig()
            .setBackgroundColor(IndexedColors.RED.getIndex())
            .setFontColor(IndexedColors.WHITE.getIndex())
            .setBold(true)
            .setFontSize((short) 11)
            .setFontName("微软雅黑")
            .setHorizontalAlignment(HorizontalAlignment.CENTER)
            .setVerticalAlignment(VerticalAlignment.CENTER)
            .setAllBorders(BorderStyle.THIN)),

    /**
     * 数据样式 - 白色背景,黑色字体,居中对齐,带边框
     * <p>适用场景：所有数据单元格</p>
     */
    DATA(new CellStyleConfig()
            .setBackgroundColor(IndexedColors.WHITE.getIndex())
            .setFontColor(IndexedColors.BLACK.getIndex())
            .setBold(false)
            .setFontSize((short) 11)
            .setFontName("微软雅黑")
            .setHorizontalAlignment(HorizontalAlignment.CENTER)
            .setVerticalAlignment(VerticalAlignment.CENTER)
            .setAllBorders(BorderStyle.THIN));

    /**
     * 样式配置对象
     */
    private final CellStyleConfig styleConfig;

    /**
     * 私有构造函数
     */
    StyleTemplate(CellStyleConfig styleConfig) {
        this.styleConfig = styleConfig;
    }

    /**
     * 获取样式配置对象的深拷贝
     * <p>
     * 每次调用都会返回一个全新的拷贝，修改返回值不会影响模板本身，
     * 也不会影响其他通过同一模板创建的配置对象。
     * </p>
     *
     * @return 样式配置对象的深拷贝
     */
    public CellStyleConfig toStyleConfig() {
        return copyStyleConfig();
    }

    /**
     * 获取样式配置对象的深拷贝
     * <p>修改拷贝不会影响模板本身</p>
     *
     * @return 样式配置对象的拷贝
     */
    public CellStyleConfig copyStyleConfig() {
        CellStyleConfig copy = new CellStyleConfig();
        copy.setBackgroundColor(styleConfig.getBackgroundColor());
        // 深拷贝 RGB 颜色数组，避免共享引用
        if (styleConfig.getRgbBackgroundColor() != null) {
            byte[] src = styleConfig.getRgbBackgroundColor();
            copy.setRgbBackgroundColor(new byte[]{src[0], src[1], src[2]});
        }
        copy.setFontColor(styleConfig.getFontColor());
        copy.setBold(styleConfig.isBold());
        copy.setFontSize(styleConfig.getFontSize());
        copy.setFontName(styleConfig.getFontName());
        copy.setHorizontalAlignment(styleConfig.getHorizontalAlignment());
        copy.setVerticalAlignment(styleConfig.getVerticalAlignment());
        copy.setBorderTop(styleConfig.getBorderTop());
        copy.setBorderBottom(styleConfig.getBorderBottom());
        copy.setBorderLeft(styleConfig.getBorderLeft());
        copy.setBorderRight(styleConfig.getBorderRight());
        copy.setWrapText(styleConfig.isWrapText());
        copy.setItalic(styleConfig.isItalic());
        copy.setStrikeout(styleConfig.isStrikeout());
        copy.setUnderline(styleConfig.getUnderline());
        copy.setIndentation(styleConfig.getIndentation());
        copy.setRotation(styleConfig.getRotation());
        copy.setLocked(styleConfig.isLocked());
        copy.setHidden(styleConfig.isHidden());
        return copy;
    }

    /**
     * 创建RGB颜色字节数组，用于 {@link CellStyleConfig#setRgbBackgroundColor(byte[])}
     * <p>
     * POI 的 IndexedColors 颜色有限，使用此方法可以设置任意 RGB 颜色。
     * </p>
     *
     * <p><b>使用示例：</b></p>
     * <pre>{@code
     * // 绿色背景
     * .setRgbBackgroundColor(StyleTemplate.rgb(0, 255, 0))
     *
     * // 自定义蓝色
     * .setRgbBackgroundColor(StyleTemplate.rgb(30, 144, 255))
     * }</pre>
     *
     * @param red   红色分量 (0-255)
     * @param green 绿色分量 (0-255)
     * @param blue  蓝色分量 (0-255)
     * @return byte数组 {r, g, b}
     * @throws IllegalArgumentException 如果颜色值超出范围
     */
    public static byte[] rgb(int red, int green, int blue) {
        if (red < 0 || red > 255) {
            throw new IllegalArgumentException("红色分量必须在0-255之间: " + red);
        }
        if (green < 0 || green > 255) {
            throw new IllegalArgumentException("绿色分量必须在0-255之间: " + green);
        }
        if (blue < 0 || blue > 255) {
            throw new IllegalArgumentException("蓝色分量必须在0-255之间: " + blue);
        }
        return new byte[]{(byte) red, (byte) green, (byte) blue};
    }
}
