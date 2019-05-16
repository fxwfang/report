package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import report.ExcelExportHelper;

public class TestExcelExportHelper {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		String modelPath = System.getProperty("user.dir")+"\\data\\业务系统数据model.xlsx";
		String writePath = System.getProperty("user.dir")+"\\data\\业务系统数据data.xlsx";
		System.out.println(modelPath);
		System.out.println(writePath);
		
		
		XSSFWorkbook workBook = new XSSFWorkbook(new FileInputStream(new File(modelPath)));
		
		// -------- 标题样式设置 -------
		XSSFCellStyle titleStyle = ExcelExportHelper.getDefaultTitleStyle(workBook);
		// -------- 内容样式设置 -------
		XSSFCellStyle bodyStyle = ExcelExportHelper.getDefaultBodyStyle(workBook);
		
		XSSFCellStyle[][] s = new XSSFCellStyle[6][2];
		for (int i = 0; i < 6; i++) {
			for (int j = 0; j < 2; j++) {
				if (i == 0) {
					s[i][j] = titleStyle;
				} else {
					s[i][j] = bodyStyle;
				}
			}
		}

		// 普通单元格颜色
		String[][] dataValue = new String[3][2];
		dataValue[0][0] = "交易量（笔）";
		dataValue[0][1] = "支行数量";
		dataValue[1][0] = "（1，8]";
		dataValue[1][1] = "80000";
		dataValue[2][0] = "（8，16]";
		dataValue[2][1] = "99000";

		String[][] dataValue2 = new String[5][2];
		dataValue2[0][0] = "交易量（笔）";
		dataValue2[0][1] = "支行数量（个）";
		dataValue2[1][0] = "（1，8]";
		dataValue2[1][1] = "80";
		dataValue2[2][0] = "（8，16]";
		dataValue2[2][1] = "99";
		dataValue2[3][0] = "（16，30]";
		dataValue2[3][1] = "99";
		dataValue2[4][0] = "（30，100]";
		dataValue2[4][1] = "99";
		
		//table 1
		String[][] d3 = new String[1][1];
		d3[0][0] = "4月24日共787实体网点做了34561笔交易，平均每个网点44笔交易。2130个柜员号做了对客交易，平均每个柜员号做16交易，其中813的交易量大于16笔，1317个柜员号的交易量小于等于16笔。";
		int[][] cellLenArr =new int[1][1];
		cellLenArr[0][0] = 15;
		int[][] cellHeightArr =new int[1][1];
		cellHeightArr[0][0] = 3;
		XSSFCellStyle[][] styleArr = new XSSFCellStyle[1][1];
		
		//重新设置样式
		XSSFCellStyle titleStyle1 = ExcelExportHelper.getDefaultTitleStyle(workBook);
		titleStyle1.setAlignment(HorizontalAlignment.LEFT);
		titleStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
		titleStyle1.setFillPattern(FillPatternType.NO_FILL); // ALT_BARS
		XSSFFont font = workBook.createFont();
		font.setColor(HSSFColor.BLACK.index);
		font.setFontName("宋体");
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		titleStyle1.setFont(font);
		styleArr[0][0] = titleStyle1;
		workBook = ExcelExportHelper.createTable(workBook, 1, 1, 1, d3,cellLenArr, cellHeightArr, styleArr);
		
		//table2
		int[][] defaultCellLenArr = ExcelExportHelper.getDefaultCellLenArr(dataValue, 2);
		int[][] defaultCellHeightArr = ExcelExportHelper.getDefaultCellHeightArr(dataValue, 2);
		ExcelExportHelper.createTable(workBook, 1, 5, 1, dataValue, defaultCellLenArr, defaultCellHeightArr, s);
		
		//table3
		defaultCellLenArr = ExcelExportHelper.getDefaultCellLenArr(dataValue2, 2);
		defaultCellHeightArr = ExcelExportHelper.getDefaultCellHeightArr(dataValue2, 2);
		ExcelExportHelper.createTable(workBook, 1, 5, 7, dataValue2, defaultCellLenArr, defaultCellHeightArr, s);

		
		ExcelExportHelper.writeExcel(new File(writePath), workBook);
		
	}
}
