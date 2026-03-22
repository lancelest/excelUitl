package com.example.excel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * 用户实体类
 * 用于演示Excel导出功能
 * 
 * @author Excel Export Tool
 * @version 2.0.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Accessors(chain = true)
public class User {
    // 序号（第一列，合并两行）
    private Integer seq;
    
    // === 基本信息组（第2-5列，一级表头合并） ===
    private String name;        // 姓名
    private String department;  // 部门
    private String position;    // 职位
    private String hireDate;    // 入职日期
    
    // === 绩效数据组（第6-9列，一级表头合并） ===
    private Double performanceScore;  // 绩效评分
    private Integer projectCount;     // 项目数
    private Double satisfaction;      // 满意度
    private Double salary;            // 薪资
    
    // 备注（第10列，合并两行）
    private String remark;
}
