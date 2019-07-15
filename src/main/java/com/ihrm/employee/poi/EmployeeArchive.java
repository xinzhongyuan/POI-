package com.ihrm.employee.poi;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


import java.io.Serializable;
import java.util.Date;



@Data
@AllArgsConstructor
@NoArgsConstructor
public class EmployeeArchive implements Serializable {
    private static final long serialVersionUID = 5768915936056289957L;
    /**
     * ID
     */

    @ExcelAttribute(sort = 1)
    private String id;
    /**
     * 操作人
     */
    @ExcelAttribute(sort = 2)
    private String opUser;
    /**
     * 月份
     */
    @ExcelAttribute(sort = 3)
    private String month;
    /**
     * 企业ID
     */
    @ExcelAttribute(sort = 4)
    private String companyId;
    /**
     * 总人数
     */
    @ExcelAttribute(sort = 5)
    private Integer totals;
    /**
     * 在职人数
     */
    @ExcelAttribute(sort = 6)
    private Integer payrolls;
    /**
     * 离职人数
     */
    @ExcelAttribute(sort = 7)
    private Integer departures;
    /**
     * 数据
     */
    @ExcelAttribute(sort = 8)
    private String data;
    /**
     * 创建时间
     */
    @ExcelAttribute(sort = 9)
    private Date createTime;
}
