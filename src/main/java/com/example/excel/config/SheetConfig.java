package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.List;

/**
 * Sheet 配置类，定义单个 Sheet 页的所有属性。
 *
 * <p><b>典型用法：</b></p>
 * <pre>{@code
 * SheetConfig<User> config = new SheetConfig<User>()
 *     .setSheetName("员工报表")
 *     .setHeaders(Arrays.asList(header1, header2))
 *     .setColumnConfigs(columns)
 *     .setDataList(dataList)
 *     .setFreezeRow(2)
 *     .setChartConfig(chartConfig);  // 可选：附加图表
 * }</pre>
 *
 * @param <T> 数据类型（支持实体类或 {@code Map<String, Object>}）
 * @author Excel Export Tool
 * @version 4.0.0
 */
@Data
@Accessors(chain = true)
public class SheetConfig<T> {

    /** Sheet 名称，默认 "Sheet1" */
    private String sheetName = "Sheet1";

    /** 表头配置列表，按从上到下顺序排列，支持多级表头 */
    private List<HeaderConfig> headers = new ArrayList<>();

    /** 列配置列表，定义每列的字段映射、宽度、样式等 */
    private List<ColumnConfig> columnConfigs = new ArrayList<>();

    /** 数据列表，支持实体类和 Map 两种类型 */
    private List<T> dataList;

    /** 冻结行数（冻结前 N 行），0 表示不冻结 */
    private int freezeRow = 0;

    /** 冻结列数（冻结前 N 列），0 表示不冻结 */
    private int freezeCol = 0;

    /** 批量刷新间隔（行数），影响内存占用和写入性能 */
    private int batchSize = 100;

    /** 默认行高（单位：twips，1/20 磅），默认 400 */
    private short defaultRowHeight = 400;

    /** 是否显示网格线，默认 true */
    private boolean displayGridlines = true;

    /** 是否启用自动筛选，默认 false */
    private boolean autoFilter = false;

    /** 默认数据行样式，无列级样式时的兜底样式 */
    private CellStyleConfig defaultDataStyle;

    /** 图表配置，为 null 则不生成图表 */
    private ChartConfig chartConfig;

    /**
     * 添加列配置
     *
     * @param columnConfig 列配置
     * @return this
     */
    public SheetConfig<T> addColumnConfig(ColumnConfig columnConfig) {
        if (this.columnConfigs == null) {
            this.columnConfigs = new ArrayList<>();
        }
        this.columnConfigs.add(columnConfig);
        return this;
    }
}
