package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.List;

/**
 * 图表配置类
 * <p>
 * 用于在 Sheet 数据结束后自动生成柱状图（或条形图）。
 * 图表位置和尺寸根据数据行数动态计算，调用者无需关心坐标。
 * </p>
 *
 * <p><b>基本用法：</b></p>
 * <pre>{@code
 * ChartConfig chart = new ChartConfig()
 *     .setTitle("部门绩效对比")
 *     .setCategoryColumn("department")
 *     .addSeries("performanceScore", "绩效评分")
 *     .addSeries("salary", "薪资");
 *
 * sheetConfig.setChartConfig(chart);
 * }</pre>
 *
 * @author Excel Export Tool
 * @version 4.0.0
 */
@Data
@Accessors(chain = true)
public class ChartConfig {

    // ==================== 基础信息 ====================

    /**
     * 图表标题，为空则不显示标题
     */
    private String title;

    /**
     * 数据结束行与图表顶部之间的间隔行数，默认 2 行
     */
    private int gapRows = 2;

    // ==================== 字体设置 ====================

    /**
     * 图表字体大小（磅数），默认 10
     * <p>应用于标题、坐标轴标题、坐标轴标签、图例等</p>
     */
    private int fontSize = 10;

    // ==================== 数据绑定 ====================

    /**
     * X 轴类别列的字段名（对应 ColumnConfig.fieldName）
     * <p>该列的值将作为柱图 X 轴的标签</p>
     */
    private String categoryColumn;

    /**
     * Y 轴数据系列列表（可多个系列，每个系列对应一组柱子）
     */
    private List<SeriesConfig> seriesList = new ArrayList<>();

    /**
     * 添加一个数据系列
     *
     * @param fieldName  字段名（对应 ColumnConfig.fieldName）
     * @param seriesName 系列名称（图例显示的文字）
     * @return this（链式调用）
     */
    public ChartConfig addSeries(String fieldName, String seriesName) {
        this.seriesList.add(new SeriesConfig(fieldName, seriesName));
        return this;
    }

    // ==================== 外观设置 ====================

    /**
     * 是否为条形图（水平方向），默认 false（垂直柱图）
     */
    private boolean barChart = false;

    /**
     * 柱图分组方式，默认簇状（CLUSTERED）
     * <ul>
     *   <li>{@link BarGrouping#CLUSTERED}       — 簇状柱图（默认）</li>
     *   <li>{@link BarGrouping#STACKED}         — 堆积柱图</li>
     *   <li>{@link BarGrouping#PERCENT_STACKED} — 百分比堆积柱图</li>
     * </ul>
     */
    private BarGrouping barGrouping = BarGrouping.CLUSTERED;

    /**
     * @deprecated 请改用 {@link #setBarGrouping(BarGrouping)}，此字段仅保留向后兼容
     */
    @Deprecated
    public ChartConfig setStacked(boolean stacked) {
        this.barGrouping = stacked ? BarGrouping.STACKED : BarGrouping.CLUSTERED;
        return this;
    }

    /** @deprecated 请使用 {@link #getBarGrouping()} 判断 */
    @Deprecated
    public boolean isStacked() {
        return barGrouping == BarGrouping.STACKED || barGrouping == BarGrouping.PERCENT_STACKED;
    }

    /**
     * 是否显示数据标签（柱子上方显示数值），默认 true
     */
    private boolean showDataLabel = true;

    /**
     * 图例位置，默认底部
     */
    private LegendPosition legendPosition = LegendPosition.BOTTOM;

    /**
     * 各系列颜色（十六进制，如 "FF0000"）
     * <p>按 seriesList 顺序对应，不设置则使用 POI 默认配色</p>
     */
    private List<String> seriesColors = new ArrayList<>();

    // ==================== 网格线设置 ====================

    /**
     * 是否显示主要网格线（Y 轴水平网格线），默认 true
     */
    private boolean showMajorGridlines = true;

    /**
     * 主要网格线颜色（十六进制，如 "D9D9D9"），默认浅灰
     * <p>设为 null 则使用 Excel 默认颜色</p>
     */
    private String majorGridlineColor = "D9D9D9";

    // ==================== 坐标轴设置 ====================

    /**
     * X 轴标签旋转角度（度数，顺时针为正，如 45 表示顺时针 45°）
     * <p>为 null 则使用 Excel 默认（水平显示）</p>
     */
    private Integer categoryAxisRotation = null;

    /**
     * X 轴标题文字，为空则不显示
     */
    private String categoryAxisTitle;

    /**
     * Y 轴标题文字，为空则不显示
     */
    private String valueAxisTitle;

    /**
     * Y 轴最小值，为 null 则由 Excel 自动计算
     */
    private Double valueAxisMin = null;

    /**
     * Y 轴最大值，为 null 则由 Excel 自动计算
     */
    private Double valueAxisMax = null;

    /**
     * Y 轴刻度间隔，为 null 则由 Excel 自动计算
     */
    private Double valueAxisUnit = null;

    // ==================== 内部枚举 ====================

    /**
     * 柱图分组方式枚举
     */
    public enum BarGrouping {
        /** 簇状柱图（默认） */
        CLUSTERED,
        /** 堆积柱图 */
        STACKED,
        /** 百分比堆积柱图 */
        PERCENT_STACKED
    }

    /**
     * 图例位置枚举
     */
    public enum LegendPosition {
        TOP, BOTTOM, LEFT, RIGHT, NONE
    }

    /**
     * 数据系列配置
     */
    @Data
    public static class SeriesConfig {
        /** 字段名，对应 ColumnConfig.fieldName */
        private final String fieldName;
        /** 系列名称，显示在图例中 */
        private final String seriesName;

        public SeriesConfig(String fieldName, String seriesName) {
            this.fieldName = fieldName;
            this.seriesName = seriesName;
        }
    }
}
