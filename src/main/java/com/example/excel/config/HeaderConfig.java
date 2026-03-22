package com.example.excel.config;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 表头配置类
 * 用于定义Excel表头的属性和样式
 * 支持多级表头和复杂合并场景
 * 
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@Accessors(chain = true)
public class HeaderConfig {

    // 表头名称列表（支持多级表头）
    private List<String> columnNames = new ArrayList<>();

    // 表头样式配置
    private CellStyleConfig styleConfig;

    // 表头高度（单位：twips，1/20磅）
    private short height = 400;

    // 合并区域列表
    private List<MergeRegion> mergeRegions = new ArrayList<>();

    // 表头层级（用于多级表头）
    private int level = 0;

    /**
     * 合并区域配置
     */
    @Data
    public static class MergeRegion {
        // 起始行（相对于表头起始行）
        private int startRow;
        // 结束行（相对于表头起始行）
        private int endRow;
        // 起始列
        private int startCol;
        // 结束列
        private int endCol;

        public MergeRegion(int startRow, int endRow, int startCol, int endCol) {
            this.startRow = startRow;
            this.endRow = endRow;
            this.startCol = startCol;
            this.endCol = endCol;
        }
    }

    /**
     * 添加表头名称
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
     */
    public HeaderConfig addMergeRegion(int startRow, int endRow, int startCol, int endCol) {
        if (this.mergeRegions == null) {
            this.mergeRegions = new ArrayList<>();
        }
        this.mergeRegions.add(new MergeRegion(startRow, endRow, startCol, endCol));
        return this;
    }

    /**
     * 创建简单的横向合并（合并当前行的指定列范围）
     */
    public HeaderConfig createHorizontalMerge(int startCol, int endCol) {
        return addMergeRegion(0, 0, startCol, endCol);
    }
}
