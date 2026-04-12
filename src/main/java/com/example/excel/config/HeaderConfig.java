package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 表头配置类，支持多级表头和单元格合并。
 *
 * <p>多级表头用法：创建多个 HeaderConfig 对象，按从上到下的顺序加入 SheetConfig.headers 列表。
 * 合并区域坐标相对于当前 HeaderConfig 所在行，0-based。</p>
 *
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@Accessors(chain = true)
public class HeaderConfig {

    /** 表头名称列表，长度应与列数一致 */
    private List<String> columnNames = new ArrayList<>();

    /** 表头样式配置 */
    private CellStyleConfig styleConfig;

    /** 表头行高（单位：twips，1/20 磅），默认 400（约 20 磅） */
    private short height = 400;

    /** 合并区域列表 */
    private List<MergeRegion> mergeRegions = new ArrayList<>();

    /** 表头层级索引（多级表头时自动设置，0 = 第一级） */
    private int level = 0;

    /**
     * 合并区域配置，坐标相对于当前 HeaderConfig 所在行
     */
    @Data
    public static class MergeRegion {
        /** 合并起始行（相对于当前表头行，0-based） */
        private int startRow;
        /** 合并结束行（相对于当前表头行，0-based） */
        private int endRow;
        /** 合并起始列（0-based） */
        private int startCol;
        /** 合并结束列（0-based） */
        private int endCol;

        public MergeRegion(int startRow, int endRow, int startCol, int endCol) {
            this.startRow = startRow;
            this.endRow = endRow;
            this.startCol = startCol;
            this.endCol = endCol;
        }
    }

    /**
     * 添加一个表头名称
     *
     * @param columnName 列名
     * @return this
     */
    public HeaderConfig addColumnName(String columnName) {
        if (this.columnNames == null) {
            this.columnNames = new ArrayList<>();
        }
        this.columnNames.add(columnName);
        return this;
    }

    /**
     * 批量添加表头名称
     *
     * @param columnNames 列名数组
     * @return this
     */
    public HeaderConfig addColumnNames(String... columnNames) {
        if (this.columnNames == null) {
            this.columnNames = new ArrayList<>();
        }
        this.columnNames.addAll(Arrays.asList(columnNames));
        return this;
    }

    /**
     * 添加合并区域
     *
     * @param startRow 起始行（相对于当前表头行，0-based）
     * @param endRow   结束行（相对于当前表头行，0-based）
     * @param startCol 起始列（0-based）
     * @param endCol   结束列（0-based）
     * @return this
     */
    public HeaderConfig addMergeRegion(int startRow, int endRow, int startCol, int endCol) {
        if (this.mergeRegions == null) {
            this.mergeRegions = new ArrayList<>();
        }
        this.mergeRegions.add(new MergeRegion(startRow, endRow, startCol, endCol));
        return this;
    }

    /**
     * 创建当前行内的横向合并（等价于 addMergeRegion(0, 0, startCol, endCol)）
     *
     * @param startCol 起始列（0-based）
     * @param endCol   结束列（0-based）
     * @return this
     */
    public HeaderConfig createHorizontalMerge(int startCol, int endCol) {
        return addMergeRegion(0, 0, startCol, endCol);
    }
}
