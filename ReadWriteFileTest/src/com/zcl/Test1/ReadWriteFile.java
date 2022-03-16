package com.zcl.Test1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class ReadWriteFile {
    /**
     * 封装快速合并行和列的方法
     *
     * @param sheet    工作表
     * @param firstRow 开始行
     * @param lastRow  结束行
     * @param firstCol 开始列
     * @param lastCol  结束列
     */
    public static void cellAddress(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        /**
         * 列数和行数都是由索引组成的
         * firstRow 区域中第一个单元格的行号
         * lastRow 区域中最后一个单元格的行号
         * firstCol 区域中第一个单元格的列号
         * lastCol 区域中最后一个单元格的列号
         */
        CellRangeAddress cellAddresses = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(cellAddresses);
        ;
    }

    /**
     * 创建一个单元格并以某种方式对齐它
     *
     * @param wb     工作簿
     * @param row    在其中创建单元格的行
     * @param value  填充单元格的内容
     * @param column 创建单元格的列号
     */
    private static void createCell(Workbook wb, Row row, int column, String value) {
        Cell cell = row.createCell(column); // 根据行创建的列
        cell.setCellValue(value); // 对单元格赋值
        CellStyle cellStyle = wb.createCellStyle(); // 创建单元格样式表
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 创建标题头部
     *
     * @param workbook 接收的工作簿
     * @param sheet    接收的工作表
     */
    public static void headerTitle(XSSFWorkbook workbook, XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(0); // 第一行数据
        XSSFRow row1 = sheet.createRow(1); // 第二行数据

        // ----------------- 合并行和列【创建方法】 开始 ---------------------
        cellAddress(sheet, 0, 1, 0, 0); // 工号行
        cellAddress(sheet, 0, 1, 1, 1); // 姓名行
        cellAddress(sheet, 0, 1, 2, 2); // 部门
        cellAddress(sheet, 0, 0, 3, 8); // 合并工资列
        cellAddress(sheet, 0, 0, 9, 10); // 扣款列
        cellAddress(sheet, 0, 0, 11, 17); // 个人缴纳部分
        cellAddress(sheet, 0, 0, 18, 24); // 企业缴纳部分
        cellAddress(sheet, 0, 1, 25, 25); // 个税金额
        cellAddress(sheet, 0, 1, 26, 26); // 应发工资
        cellAddress(sheet, 0, 1, 27, 27); // 实发工资
        cellAddress(sheet, 0, 1, 28, 28); // 企业支出成本
        // ----------------- 创建表头一二行的信息 开始 ---------------------
        XSSFCell cell = row.createCell(0); // 创建第一列
        XSSFCell cell2 = row.createCell(1); // 创建第二列
        // 设置单元格靠底部居中
        // row.createCell(0).setCellValue("工号"); // 根据行创建的列

        createCell(workbook, row, 0, "工号");
        createCell(workbook, row, 1, "姓名");
        createCell(workbook, row, 2, "部门");
        createCell(workbook, row, 3, "工资");
        createCell(workbook, row, 9, "扣款");
        createCell(workbook, row, 11, "个人缴纳部分");
        createCell(workbook, row, 18, "企业缴纳部分");
        createCell(workbook, row, 25, "个税金额");
        createCell(workbook, row, 26, "应发工资");
        createCell(workbook, row, 27, "实发工资");
        createCell(workbook, row, 28, "企业支持成本");

        // ---------- 第二行数据填写 工资部分 开始
        createCell(workbook, row1, 3, "底薪");
        createCell(workbook, row1, 4, "岗位工资");
        createCell(workbook, row1, 5, "绩效奖金");
        createCell(workbook, row1, 6, "全勤奖金");
        createCell(workbook, row1, 7, "交通补助");
        createCell(workbook, row1, 8, "通信补助");
        // ---------- 第二行数据填写 扣款部分 开始
        createCell(workbook, row1, 9, "考勤扣除");
        createCell(workbook, row1, 10, "违规处罚");
        // ---------- 第二行数据填写 个人缴纳部分 开始
        createCell(workbook, row1, 11, "养老");
        createCell(workbook, row1, 12, "医疗");
        createCell(workbook, row1, 13, "失业");
        createCell(workbook, row1, 14, "工伤");
        createCell(workbook, row1, 15, "生育");
        createCell(workbook, row1, 16, "公积金");
        createCell(workbook, row1, 17, "合计");
        // ---------- 第二行数据填写 企业缴纳部分 开始
        createCell(workbook, row1, 18, "养老");
        createCell(workbook, row1, 19, "医疗");
        createCell(workbook, row1, 20, "失业");
        createCell(workbook, row1, 21, "工伤");
        createCell(workbook, row1, 22, "生育");
        createCell(workbook, row1, 23, "公积金");
        createCell(workbook, row1, 24, "合计");
    }

    /**
     * 创建一个类可以共同访问的集合储存去读取表的新数据
     */
    public static ArrayList FileSum = new ArrayList<>();

    /**
     * 创建一个公共遍历读取集合和添加部门信息的方法
     *
     * @param data    读取数据集合
     * @param index   插入集合的位置
     * @param Section 插入部门的信息
     */
    public static void forList(List<List<Object>> data, int index, String Section) {
        for (List<Object> objects : data.subList(1, data.size())) {
            objects.add(index, Section); // 手动指定索引插入部门信息
            // 声明变量来交换【补助和扣款】的位置
            Object o7 = objects.get(7);
            objects.set(7, objects.get(9));
            objects.set(9, o7);

            Object o8 = objects.get(8);
            objects.set(8, objects.get(10));
            objects.set(10, o8);

            // objects：[1001, 沈逸春, 研发部, 4000, 2000, 500, 150, 200, 150, 0, 0, null, null, ]
            FileSum.add(objects);
        }
    }

    /**
     * 程序入口主方法
     *
     * @param args
     * @throws Exception 异常捕获
     */
    public static void main(String[] args) throws Exception {
        // 创建读取文件对象
        ExcelUtils excelUtils = new ExcelUtils();
        // 读取指定文件[一次要读取6张表的数据合并]
        List<List<Object>> research = excelUtils.readExcelFirstSheet(new File("ReadWriteFileTest\\resource\\研发部-薪酬表.xlsx")); // 1
        List<List<Object>> lists = excelUtils.readExcelFirstSheet(new File("ReadWriteFileTest\\resource\\大客户部-薪酬表.xlsx")); // 2
        List<List<Object>> bazaar = excelUtils.readExcelFirstSheet(new File("ReadWriteFileTest\\resource\\市场部-薪酬表.xlsx")); // 3
        List<List<Object>> market = excelUtils.readExcelFirstSheet(new File("ReadWriteFileTest\\resource\\销售部-薪酬表.xlsx")); // 4
        List<List<Object>> calculate = excelUtils.readExcelFirstSheet(new File("ReadWriteFileTest\\resource\\员工五险一金申报表.xlsx")); // 工资基数表

        // FIXME 将上面的集合全部去掉第一行读取到一个新的集合中储存，并将部门的信息添加进集合
        forList(research, 2, "研发部");
        forList(lists, 2, "大客户部");
        forList(bazaar, 2, "市场部");
        forList(market, 2, "销售部");

        // FIXME 遍历五险一金的表去除表头行进行传递数据


        readWriteFileData(FileSum, calculate);
        System.out.println("文件读取成功，请查看本工程项目下的.xlsx文件");
    }

    /**
     * 读取指定的文件并写出新的文件
     *
     * @param FileSum 读取所有文件的对象
     */
    public static void readWriteFileData(List<List<Object>> FileSum, List<List<Object>> calculate) {
        // 1、创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 2、创建工作表
        XSSFSheet sheet = workbook.createSheet("支付表");
        // 3、调用表头创建方法
        headerTitle(workbook, sheet); // 已经创建0和1行，赋值是数据需要在第三行开始

        // 声明一个记录员工资和扣款数据的数组
        double[] collData = new double[8];
        // 设置自动创建行的行号
        int rowIndex = 2;
        // 企业支持成本
        double CS = 0.0;
        // 应发工资 - 五险一金 = 实发工资
        double salarys = 0.0; // 应发工资
        double personages = 0.0; // 个人五险一金

        // 4、遍历集合
        for (List<Object> objects : FileSum) {
            // 自动创建行数据 rowIndex为创建行的索引
            XSSFRow row = sheet.createRow(rowIndex); // 创建行号，跳过表头2两行
            // [1001, 沈逸春, 研发部, 4000, 2000, 500, 150, 200, 150, 0, 0, null, null, ] 目前是可以遍历之后通过get(index)缩影的方法获取到每一集合中的具体数值
            // 使用FileSum.size()会出现缩影越界
            // System.out.println("集合的元素："+objects.size()); 14个元素
            for (int i = 0; i < objects.size(); i++) {
                // FIXME 遍历写入员工的基本信息
                {
                    // 非空判断
                    if (objects.get(i) == null || objects.get(i) == "") {
                        // 4024 田浩思 研发部 3200 1200 700 150 0 50 200 150 null null
                        continue;
                    }
                    // 调用列的方法写入员工的基本数据信息
                    createCell(workbook, row, i, objects.get(i).toString());
                }

                // FIXME 声明一个集合分别储存所有工资和扣款，根据判断条件获取到值储存进26列的应发工资
                {
                    // 循环获取到值储存进去
                    if (i >= 3 && i <= 10) {
                        // 跳过前面的三列数据
                        collData[i - 3] = Double.parseDouble(objects.get(i).toString());
                    }
                    // 遍历计算集合中应发工资的值
                    double salary = 0.0;
                    for (int j = 0; j < collData.length; j++) {
                        if (j >= 6 && j <= 7) {
                            salary -= collData[j];
                        } else {
                            salary += collData[j];
                        }
                    }
                    // 判断获取到应发的员工工资
                    if (i == 10) {
                        CS = salary;
                        // 实发工资的计算：应发-五险一金
                        salarys = salary;
                    }
                    createCell(workbook, row, 26, String.valueOf(salary));
                }

                // FIXME 计算个人和企业所交的五险一金

                /**
                 * 思路：根据读取的五险一金的表获得工资基数按照给定的公式进行计算
                 * -	养老保险：单位，20%，个人，8%
                 * -	医疗保险：单位，8%，个人，2%；
                 * -	失业保险：单位，2%，个人，1%；
                 * -	工伤保险：单位，0.5%，个人不用缴费；
                 * -	生育保险：单位，0.7%，个人不用缴费；
                 * -    住房公积金缴纳比例有 8%、10%、12%三档每人具体公积金缴纳比例数据也记录在“员工五险一金申报表.xlsx”中。
                 */

                // 获取到遍历的工资基数数据
                List<Object> objects1 = calculate.get(rowIndex - 1); // 一条数据遍历11次，已去掉表头
                int startIndex = 11;
                // 创建一个基本工资的变量
                double num = Double.parseDouble(objects1.get(2).toString());
                // 创建一个计算个人缴纳部分的金额
                double personage = 0.0;

                // TODO 声明一个个人五险一金承担比例,-1代表的是公积金
                {
                    double[] nums = {0.08, 0.02, 0.01, 0, 0, -1};
                    for (int z = 0; z < nums.length; z++) {
                        BigDecimal a1 = BigDecimal.valueOf(num);
                        BigDecimal b1 = BigDecimal.valueOf(nums[z]);
                        BigDecimal c1 = a1.multiply(b1);
                        double sum = c1.doubleValue();
                        // 公积金
                        if (nums[z] == -1) {
                            sum = num * Double.parseDouble(objects1.get(3).toString());
                        }
                        if (nums[z] == -2) {
                            for (int zz = 0; zz < z; zz++) {
                                // 个人缴纳
                                sum += num * nums[zz];
                            }
                        }
                        if (z <= 5) {
                            personage = sum + personage;

                            personages = personage;
                        }
                        System.out.println("五险一金个人缴纳合计："+personages);
                        createCell(workbook, row, startIndex + z, String.valueOf(sum));
                        createCell(workbook, row, 17, String.valueOf(personages)); // 五险一金个人缴纳合计

                    }
                    startIndex += nums.length;
                }

                // TODO 企业支付公积金未完成计算
                {
                    double[] nums = {0.2, 0.08, 0.02, 0.005, 0.007, 0, -1};
                    for (int j = 0; j < nums.length; j++) {
                        BigDecimal a1 = BigDecimal.valueOf(num);
                        BigDecimal b1 = BigDecimal.valueOf(nums[j]);
                        BigDecimal c1 = a1.multiply(b1); // divide() 除运算 0.5
                        double sum = c1.doubleValue();
                        if (nums[j] == -1) {
                            // 公积金企业不用出
                            sum = 0.0;
                            // 计算合计
                            for (int z = 0; z < j; z++) {
                                sum += num * nums[z];
                            }
                        }
                        createCell(workbook, row, startIndex + j + 1, String.valueOf(sum));
                        // 根据企业缴纳的五险一金计算出总的企业支出成本
                        if (j == 6) {
                            CS += sum;
                        }
                    }
                }
            }
            // TODO 计算个税金额
            // 个税金额
            double PIT = 0.0;

            if(salarys <= 3000){
                PIT = salarys * 0.3;
            }else if(salarys > 3000 && salarys <= 12000){
                PIT = salarys * 0.1;
            }else if(salarys > 1200 && salarys <= 25000){
                PIT = salarys * 0.2;
            }else if(salarys > 25000 && salarys <= 35000){
                PIT = salarys * 0.25;
            }else if(salarys > 35000 && salarys <= 55000){
                PIT = salarys * 0.3;
            }else if(salarys > 55000 && salarys <= 80000){
                PIT = salarys * 0.35;
            }else if(salarys > 80000){
                PIT = salarys * 0.45;
            }

            // 个税金额
            createCell(workbook, row, 25, String.valueOf(PIT));
            // 实发工资
            createCell(workbook, row, 27, String.valueOf(salarys-personages));
            // Excel企业支出成本列
            createCell(workbook, row, 28, String.valueOf(CS));
            // 控制行数的自增
            rowIndex++;
        }


        // 6、写出文件位置
        try (FileOutputStream fos = new FileOutputStream(new File("企业员工月度工资成本支付表.xlsx"))) {
            workbook.write(fos); // 输出文件
            // 关闭输出流
            fos.close();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
