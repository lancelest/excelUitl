package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 单元格样式配置类，定义字体、颜色、边框、对齐、背景等所有样式属性。
 *
 * <p>所有字段均支持链式调用。优先级：{@code rgbBackgroundColor} > {@code backgroundColor}。</p>
 *
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@Accessors(chain = true)
public class CellStyleConfig {

    /** 背景颜色索引（IndexedColors），-1 表示不设置；rgbBackgroundColor 优先级更高 */
    private short backgroundColor = -1;

    /** RGB 背景颜色，优先级高于 backgroundColor，格式 {@code byte[]{r, g, b}} */
    private byte[] rgbBackgroundColor = null;

    /** 字体颜色索引（IndexedColors），-1 表示不设置 */
    private short fontColor = -1;

    /** 是否加粗 */
    private boolean bold = false;

    /** 字体大小（单位：磅），默认 10 */
    private short fontSize = 10;

    /** 字体名称，默认 "微软雅黑" */
    private String fontName = "微软雅黑";

    /** 水平对齐方式，默认居中 */
    private HorizontalAlignment horizontalAlignment = HorizontalAlignment.CENTER;

    /** 垂直对齐方式，默认居中 */
    private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;

    /** 上边框样式，默认 THIN */
    private BorderStyle borderTop = BorderStyle.THIN;

    /** 下边框样式，默认 THIN */
    private BorderStyle borderBottom = BorderStyle.THIN;

    /** 左边框样式，默认 THIN */
    private BorderStyle borderLeft = BorderStyle.THIN;

    /** 右边框样式，默认 THIN */
    private BorderStyle borderRight = BorderStyle.THIN;

    /** 是否自动换行 */
    private boolean wrapText = false;

    /** 是否斜体 */
    private boolean italic = false;

    /** 是否删除线 */
    private boolean strikeout = false;

    /** 下划线类型：0=无，1=单下划线，2=双下划线 */
    private byte underline = 0;

    /** 缩进级别 */
    private short indentation = 0;

    /** 文本旋转角度（-90 到 90 度） */
    private short rotation = 0;

    /** 是否锁定单元格（保护工作表时生效） */
    private boolean locked = true;

    /** 是否隐藏公式 */
    private boolean hidden = false;

    /**
     * 统一设置四边边框样式
     *
     * @param borderStyle 边框样式
     * @return this
     */
    public CellStyleConfig setAllBorders(BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
        return this;
    }
}
