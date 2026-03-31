package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.List;

/**
 * Sheet配置类
 * 用于定义Excel中单个Sheet页的所有配置信息
 * 
 * @param <T> 数据类型（支持实体类或Map）
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@Accessors(chain = true)
public class SheetConfig<T> {

    // Sheet名称
    private String sheetName = "Sheet1";

    // 表头配置列表（支持多级表头）
    private List<HeaderConfig> headers = new ArrayList<>();

    // 列配置列表
    private List<ColumnConfig> columnConfigs = new ArrayList<>();

    // 数据列表（支持泛型）
    private List<T> dataList;

    // 冻结行数（前N行冻结）
    private int freezeRow = 0;

    // 冻结列数（前N列冻结）
    private int freezeCol = 0;

    // 每批处理的数据量（用于流式写入优化内存）
    private int batchSize = 100;

    // 默认行高（单位：twips）
    private short defaultRowHeight = 400;

    // 是否显示网格线
    private boolean displayGridlines = true;

    // 是否自动筛选
    private boolean autoFilter = false;

    // 默认数据行样式
    private CellStyleConfig defaultDataStyle;

    // 图表配置（为 null 则不生成图表）
    private ChartConfig chartConfig;

    /**
     * 添加列配置
     */
    public SheetConfig<T> addColumnConfig(ColumnConfig columnConfig) {
        if (this.columnConfigs == null) {
            this.columnConfigs = new ArrayList<>();
        }
        this.columnConfigs.add(columnConfig);
        return this;
    }
}
