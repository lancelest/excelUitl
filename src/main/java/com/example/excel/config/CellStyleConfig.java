package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 单元格样式配置类
 * 用于定义Excel单元格的各种样式属性
 * 
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@Accessors(chain = true)
public class CellStyleConfig {

    // 背景颜色索引（-1表示不设置；如果 rgbBackgroundColor 不为null，则优先使用RGB）
    private short backgroundColor = -1;

    // RGB背景颜色（优先级高于 backgroundColor，格式为 byte[]{r, g, b}）
    private byte[] rgbBackgroundColor = null;

    // 字体颜色索引（-1表示不设置）
    private short fontColor = -1;

    // 是否加粗
    private boolean bold = false;

    // 字体大小（单位：磅）
    private short fontSize = 11;

    // 字体名称
    private String fontName = "微软雅黑";

    // 水平对齐方式
    private HorizontalAlignment horizontalAlignment = HorizontalAlignment.CENTER;

    // 垂直对齐方式
    private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;

    // 上边框样式
    private BorderStyle borderTop = BorderStyle.THIN;

    // 下边框样式
    private BorderStyle borderBottom = BorderStyle.THIN;

    // 左边框样式
    private BorderStyle borderLeft = BorderStyle.THIN;

    // 右边框样式
    private BorderStyle borderRight = BorderStyle.THIN;

    // 是否自动换行
    private boolean wrapText = false;

    // 是否倾斜
    private boolean italic = false;

    // 是否删除线
    private boolean strikeout = false;

    // 下划线类型（0：无，1：单下划线，2：双下划线）
    private byte underline = 0;

    // 缩进
    private short indentation = 0;

    // 旋转角度（-90到90度）
    private short rotation = 0;

    // 是否锁定单元格
    private boolean locked = true;

    // 是否隐藏公式
    private boolean hidden = false;

    /**
     * 设置四边边框样式
     */
    public CellStyleConfig setAllBorders(BorderStyle borderStyle) {
        this.borderTop = borderStyle;
        this.borderBottom = borderStyle;
        this.borderLeft = borderStyle;
        this.borderRight = borderStyle;
        return this;
    }
}
