package com.ihrm.employee.controller;

import com.ihrm.employee.employee.EmployeeReportResult;


import com.ihrm.employee.poi.DownloadUtils;
import com.ihrm.employee.poi.ExcelExportUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;


@RestController
@CrossOrigin
@RequestMapping("/employees")
public class EmployeeController extends BaseController {

    /**
     * 当月人事报表导出
     *  参数：
     *      年月-月（2018-02%）
     */
    @RequestMapping(value = "/export/{month}", method = RequestMethod.GET)
    public void export(@PathVariable String month,HttpServletRequest  request,HttpServletResponse resonse) throws Exception {
        this.response=resonse;

        //1.获取报表数据
        List<EmployeeReportResult> list = new ArrayList<>();
        EmployeeReportResult result=new EmployeeReportResult();
        result.setTitle("宅急送");
        for (int i = 0; i <100000; i++) {
            list.add(result);
        }
        //2.构造Excel
        //创建工作簿
        //SXSSFWorkbook : 百万数据报表
      //  Workbook wb = new XSSFWorkbook();
       SXSSFWorkbook wb = new SXSSFWorkbook(100); //阈值，内存中的对象数量最大数量
        //构造sheet
        Sheet sheet = wb.createSheet();
        //创建行
        //标题
        String [] titles = "编号,姓名,手机,最高学历,国家地区,护照号,籍贯,生日,属相,入职时间,离职类型,离职原因,离职时间".split(",");
        //处理标题
        Row row = sheet.createRow(0);
        int titleIndex=0;
        for (String title : titles) {
            Cell cell = row.createCell(titleIndex++);
            cell.setCellValue(title);
        }
        int rowIndex = 1;
        Cell cell=null;
        for(int i=0;i<10;i++){
        for (EmployeeReportResult employeeReportResult : list) {
            row = sheet.createRow(rowIndex++);
            // 编号,
            cell = row.createCell(0);
            cell.setCellValue(employeeReportResult.getTitle());
        }
        }
        //3.完成下载
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        wb.write(os);
        new DownloadUtils().download(os,response,month+"人事报表.xlsx");
    }

    /**
     * 采用模板打印的形式完成报表生成
     *      模板
     *  参数：
     *      年月-月（2018-02%）
     *
     *      sxssf对象不支持模板打印
     */
    @RequestMapping(value = "/export2/{month}", method = RequestMethod.GET)
    public void export(@PathVariable String month,HttpServletResponse response) throws Exception {
        //1.获取报表数据
        List<EmployeeReportResult> list = new ArrayList<>();
        EmployeeReportResult result=new EmployeeReportResult();
        result.setTitle("宅急送");
        for (int i = 0; i <10; i++) {
            list.add(result);
        }
        //2.加载模板
        Resource resource = new ClassPathResource("excel-template/hr-demo.xlsx");
        FileInputStream fis = new FileInputStream(resource.getFile());
        ExcelExportUtil<EmployeeReportResult> exports = new ExcelExportUtil<>(EmployeeReportResult.class, 2, 2);
        exports.export(response, fis, list, "人事报表.xlsx");
    }


    /**
     * 采用模板打印的形式完成报表生成
     *      模板
     *  参数：
     *      年月-月（2018-02%）
     *      sxssf对象不支持模板打印
     */
    @RequestMapping(value = "/export3/{month}", method = RequestMethod.GET)
    public void export2(@PathVariable String month,HttpServletResponse response) throws Exception {
        //1.获取报表数据
        List<EmployeeReportResult> list = new ArrayList<>();
        EmployeeReportResult result=new EmployeeReportResult();
        result.setTitle("宅急送");
        for (int i = 0; i <10; i++) {
            list.add(result);
        }
        //2.加载模板
        Resource resource = new ClassPathResource("excel-template/hr-demo.xlsx");
        FileInputStream fis = new FileInputStream(resource.getFile());

        //3.根据模板创建工作簿
        Workbook wb = new XSSFWorkbook(fis);
        //4.读取工作表
        Sheet sheet = wb.getSheetAt(0);
        //5.抽取公共样式
        Row row = sheet.getRow(2);
        CellStyle styles [] = new CellStyle[row.getLastCellNum()];
        for(int i=0;i<row.getLastCellNum();i++) {
            Cell cell = row.getCell(i);
            styles[i] = cell.getCellStyle();
        }
        //6.构造单元格
        int rowIndex = 2;
        Cell cell=null;
        for(int i=0;i<1000;i++) {
            for (EmployeeReportResult employeeReportResult : list) {
                row = sheet.createRow(rowIndex++);
                // 编号,
                cell = row.createCell(0);
                cell.setCellValue(employeeReportResult.getTitle());
                cell.setCellStyle(styles[0]);
               /* // 姓名,
                cell = row.createCell(1);
                cell.setCellValue(employeeReportResult.getUsername());
                cell.setCellStyle(styles[1]);
                // 手机,
                cell = row.createCell(2);
                cell.setCellValue(employeeReportResult.getMobile());
                cell.setCellStyle(styles[2]);
                // 最高学历,
                cell = row.createCell(3);
                cell.setCellValue(employeeReportResult.getTheHighestDegreeOfEducation());
                cell.setCellStyle(styles[3]);
                // 国家地区,
                cell = row.createCell(4);
                cell.setCellValue(employeeReportResult.getNationalArea());
                cell.setCellStyle(styles[4]);
                // 护照号,
                cell = row.createCell(5);
                cell.setCellValue(employeeReportResult.getPassportNo());
                cell.setCellStyle(styles[5]);
                // 籍贯,
                cell = row.createCell(6);
                cell.setCellValue(employeeReportResult.getNativePlace());
                cell.setCellStyle(styles[6]);
                // 生日,
                cell = row.createCell(7);
                cell.setCellValue(employeeReportResult.getBirthday());
                cell.setCellStyle(styles[7]);
                // 属相,
                cell = row.createCell(8);
                cell.setCellValue(employeeReportResult.getZodiac());
                cell.setCellStyle(styles[8]);
                // 入职时间,
                cell = row.createCell(9);
                cell.setCellValue(employeeReportResult.getTimeOfEntry());
                cell.setCellStyle(styles[9]);
                // 离职类型,
                cell = row.createCell(10);
                cell.setCellValue(employeeReportResult.getTypeOfTurnover());
                cell.setCellStyle(styles[10]);
                // 离职原因,
                cell = row.createCell(11);
                cell.setCellValue(employeeReportResult.getReasonsForLeaving());
                cell.setCellStyle(styles[11]);
                // 离职时间
                cell = row.createCell(12);
                cell.setCellValue(employeeReportResult.getResignationTime());
                cell.setCellStyle(styles[12]);*/
            }
        }
        //7.下载
        //3.完成下载
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        wb.write(os);
        new DownloadUtils().download(os,response,month+"人事报表--非工具类版本.xlsx");
    }
}



