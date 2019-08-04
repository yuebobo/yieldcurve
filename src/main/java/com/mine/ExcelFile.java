package com.mine;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static java.math.BigDecimal.ROUND_HALF_DOWN;


/**
 *
 * 时间 : 2018/8/19.
 */
public class ExcelFile {

    private static final BigDecimal RATIO_VALUE    = new BigDecimal(0.5);

    public static void dealFile(String path) throws FileNotFoundException {
        System.out.println("给定文件路径");
        System.out.println(path);
        String basePath;
        if (path.contains("\\")) {
            basePath = path.substring(0, path.lastIndexOf("\\"));
        } else if (path.contains("/")) {
            basePath = path.substring(0, path.lastIndexOf("/"));
        } else {
            throw new FileNotFoundException();
        }
        System.out.println("基本路径：" + basePath);
        String outFilePath = basePath + "\\out_"+System.currentTimeMillis()+"..xlsx";
        dealFile(path,outFilePath);
    }

    public static void dealFile(String modelPath ,String outPath){
        FileInputStream in = null;
        FileOutputStream out = null;
        try {
            in = new FileInputStream(modelPath);
            out = new FileOutputStream(outPath);
            XSSFWorkbook excel = new XSSFWorkbook(in);
            dealExcel(excel);
            excel.write(out);
            out.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$  文件未找到 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
        } catch (IOException e1) {
            e1.printStackTrace();
            System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$  IO异常 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    /**
     * 处理excel
     * @param excel
     */
    private static void dealExcel(XSSFWorkbook excel){
        XSSFSheet sheet = excel.getSheetAt(0);
        Iterator it = sheet.rowIterator();
        List<XSSFRow> rowList = new ArrayList<>(100);

        //iterator 的数据装换到list里 进行处理
        iteratorToList(it,rowList);

        //A 点力
        BigDecimal pointAforce = new BigDecimal(Util.getValueFromXssfcell(rowList.get(0).getCell(2)));
        // A点位移
        BigDecimal pointAdisplacement = new BigDecimal(Util.getValueFromXssfcell(rowList.get(0).getCell(1)));
        System.out.println("A点  位移；"+pointAdisplacement + " 力："+pointAforce);

        //计算斜率，平均斜率，寻找C点
        int pointCPosition = calculateSlope(rowList, pointAforce, pointAdisplacement);

        //C点力
        BigDecimal ponitCforce = new BigDecimal(Util.getValueFromXssfcell(rowList.get(pointCPosition).getCell(2)));

        //B点位置寻找，力值为C点力值的一半 的最近的点
        BigDecimal halfCforce = ponitCforce.divide(new BigDecimal(2),3,ROUND_HALF_DOWN);
        int pointBPosition = calculatePointB(rowList, halfCforce);

        //计算A-B拟合直线
        String equation = insertLineAToB(pointAforce, pointAdisplacement, pointBPosition, rowList);

        //C点位置
        XSSFCell cell = rowList.get(0).createCell(6);
        cell.setCellValue("C: "+(pointCPosition+2));
        System.out.println("C点 位置："+(pointCPosition+2));

        //B点位置
        cell = rowList.get(0).createCell(7);
        cell.setCellValue("B: "+(pointBPosition+2));
        System.out.println("C点 位置："+(pointBPosition+2));

        //方程式
        cell = rowList.get(0).createCell(8);
        cell.setCellValue(equation);
        System.out.println("直线方程为：" + equation);

    }

    private static String insertLineAToB(BigDecimal forceA,BigDecimal displacementA,int positionB,List<XSSFRow> rowList) {
        if (positionB < 0) {
            return "";
        }
        //B 点力
        BigDecimal forceB = new BigDecimal(Util.getValueFromXssfcell(rowList.get(positionB).getCell(2)));
        // B点位移
        BigDecimal displacementB = new BigDecimal(Util.getValueFromXssfcell(rowList.get(positionB).getCell(1)));

        // y = kx + b
        // k = (fb -fa)/(db - da)
        BigDecimal k = forceB.subtract(forceA).divide(displacementB.subtract(displacementA), 3, ROUND_HALF_DOWN);
        // b = y - kx
        BigDecimal b = forceA.subtract(k.multiply(displacementA));

        String equation = "y = " + k + "x + " + b;
        return equation;
        //录入到excel表里

    }


    /**
     * 寻找B点
     * @param rowList
     * @param force
     * @return
     */
    private static int calculatePointB(List<XSSFRow> rowList,BigDecimal force){
        int size = rowList.size();
        BigDecimal forceBefore = new BigDecimal(0);
        BigDecimal forceAfter;
        for (int i = 0;i < size;i++){
            forceAfter =  new BigDecimal(Util.getValueFromXssfcell(rowList.get(i).getCell(2)));
            if (force.compareTo(forceBefore) >= 0 && force.compareTo(forceAfter) <= 0){
                BigDecimal subBefore = force.subtract(forceBefore);
                BigDecimal subAfter = forceAfter.subtract(force);
                if (subBefore.compareTo(subAfter) <= 0){
                    return i-1;
                }else {
                    return i;
                }
            }
            forceBefore =  forceAfter;
        }
        return -1;
    }


    /**
     * 依次计算各点的斜率，斜率的平均值（已给定点为起点）
     * @param rowList
     * @param pointAforce
     * @param pointAdisplacement
     */
    private static int calculateSlope(List<XSSFRow> rowList,BigDecimal pointAforce,BigDecimal pointAdisplacement){
        BigDecimal force;
        BigDecimal splacement;

        BigDecimal preForce = pointAforce;
        BigDecimal preSplacement = pointAdisplacement;

        BigDecimal slope;
        BigDecimal aveSlope;
        BigDecimal ratio;
        BigDecimal sumSlope = new BigDecimal(0);

        int count = 0;
        int pointCPosition = 0;

        int size = rowList.size();
        for (int i = 1; i < size; i++ ) {
            splacement = new BigDecimal(Util.getValueFromXssfcell(rowList.get(i).getCell(1)));
            force = new BigDecimal(Util.getValueFromXssfcell(rowList.get(i).getCell(2)));
            //过滤掉异常的数据
            if (splacement.compareTo(preSplacement) < 0) continue;
            if (force.compareTo(preForce) < 0) continue;

            //斜率计算（基于A点）
            slope = force.subtract(pointAforce).divide(splacement.subtract(pointAdisplacement),3,ROUND_HALF_DOWN);
            //累计
            sumSlope = sumSlope.add(slope);
            //计算斜率平均值
            aveSlope = sumSlope.divide(new BigDecimal(++count),3,ROUND_HALF_DOWN);
            //斜率除以平均斜率
            System.out.println("slope:"+slope);
            System.out.println("aveSlope:"+aveSlope);

            ratio = slope.divide(aveSlope,3,ROUND_HALF_DOWN);

            rowList.get(i).createCell(3);
            rowList.get(i).createCell(4);
            rowList.get(i).createCell(5);

            //值填入表格中
            rowList.get(i).getCell(3).setCellValue(Util.getPrecisionString(slope.toString(),3));
            rowList.get(i).getCell(4).setCellValue(Util.getPrecisionString(aveSlope.toString(),3));
            rowList.get(i).getCell(5).setCellValue(Util.getPrecisionString(ratio.toString(),3));

            //寻找C点
            if (ratio.compareTo(RATIO_VALUE) <= 0 && pointCPosition == 0){
                XSSFCellStyle cellStyle = rowList.get(i).getCell(5).getCellStyle();
                cellStyle.setTopBorderColor(new XSSFColor(java.awt.Color.RED));
                pointCPosition = i;
                System.out.println("C点 ："+i);
                System.out.println("force :"+force + " ,splacement :"+splacement);
            }

            preForce  = force;
            preSplacement = splacement;
        }
        //位置加上表头 和A点  多的两行
        return pointCPosition;
    }







    /**
     * iterator 的数据装换到list里
     * @param it
     * @param rowList
     */
    private static void iteratorToList(Iterator<Row> it,List<XSSFRow> rowList){
        XSSFRow row = (XSSFRow) it.next();
        XSSFCell cell = row.getCell(1);
        String value = Util.getValueFromXssfcell(cell);
        //判断首行是否是数据
        try {
            Double d = Double.valueOf(value);
            //没有发生异常，则加入首行数据
            rowList.add(row);
            System.out.println("首行就为数据");
        }catch (NumberFormatException e){
            System.out.println("首行不为数据");
        }

        while (it.hasNext()){
           row = (XSSFRow) it.next();
           if (row == null) return;
           cell = row.getCell(1);
           if (cell == null) return;
           value = Util.getValueFromXssfcell(cell);
           if (value == null || value == "") return;
           rowList.add(row);
        }
    }
}