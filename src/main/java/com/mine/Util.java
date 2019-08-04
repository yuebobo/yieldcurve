package com.mine;

import org.apache.poi.xssf.usermodel.XSSFCell;

import java.math.BigDecimal;


public class Util {


	public static void printArray(Object[] objects){
		int i = 0;
		for (Object object : objects){
			System.out.println( "   "+(++i)+"      "+object.toString());
		}
	}

	/**
	 * 四舍五入 返回String类型
	 * @param value
	 * @param precision 保留几位小数
	 * @return
	 */
	public static String getPrecisionString(String value,int precision){
		if(value == null || value == "") return "0";
		double v =Math.abs(Double.valueOf(value));
		if(precision < 0) return String.format("%.0f",v); 
		return String.format("%." + precision + "f",v); 
	}

	/**
	 * 四舍五入
	 * @param value
	 * @param precision
	 * @return
	 */
	public static String getPrecisionString(Double value,int precision){
		if(value == null ) return "0";
		if(precision < 0) return String.format("%.0f",value); 
		return String.format("%." + precision + "f",value); 
	}

	/**
	 * 四舍五入 返回Double类型
	 * @param value
	 * @param precision 保留几位小数
	 * @return
	 */
	public static Double getPrecisionDouble(String value,int precision){
		if(value == null || value == "") return 0D;
		double v =Math.abs(Double.valueOf(value));
		if(precision < 1)	return Double.valueOf(String.format("%.0f",v)); 
		return Double.valueOf((String.format("%." + precision + "f",v))); 
	}

	/**
	 * 对两个值进行相乘，在除以2
	 * @param value1
	 * @param value2
	 * @return
	 */
	public static String multiplyAndHalf(String value1,String value2){
		BigDecimal v1 = new BigDecimal(value1);
		BigDecimal v2 = new BigDecimal(value2);
		BigDecimal v = v1.multiply(v2); 
		return getPrecisionString(v.divide(new BigDecimal(2)).toString(),0);
	}

	/**
	 * 获取excel单元格里的值
	 * @param cell
	 * @return
	 */
	public static String getValueFromXssfcell(XSSFCell cell){
		if (cell == null) return null;
		try {
			return cell.getNumericCellValue()+"";
		}catch (Exception e){
			try {
				return cell.getStringCellValue();
			}catch (Exception e1){
				return cell.getRawValue();
			}
		}
	}

}
