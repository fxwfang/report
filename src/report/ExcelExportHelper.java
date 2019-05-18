package report;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Pattern;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelExportHelper {

	// 获取开始行行标
	public static int getStartRow(int startRow, int rowIndex, int colIndex, int[][] cellHeightArr) {
		if (rowIndex - 1 < 0) { // 第一行
			return startRow;
		}
		for (int i = 0; i < rowIndex; i++) {
			startRow += cellHeightArr[i][colIndex];
		}
		return startRow;
	}

	// 获取结束行的行标
	public static int getEndRow(int startRow, int rowIndex, int colIndex, int[][] cellHeightArr) {
		int endRow = startRow;
		for (int i = 0; i <= rowIndex; i++) {
			endRow += cellHeightArr[i][colIndex];
		}
		return endRow - 1;
	}

	// 获取开始列的列标
	public static int getStartCol(int startColumn, int rowIndex, int colIndex, int[][] cellLenArr) {
		if (colIndex - 1 < 0) {
			return startColumn;
		}
		for (int i = 0; i < colIndex; i++) {
			startColumn += cellLenArr[rowIndex][i];
		}
		return startColumn;
	}

	// 获取结束列的列标
	public static int getEndCol(int startColumn, int rowIndex, int colIndex, int[][] cellLenArr) {
		int endCol = startColumn;
		for (int i = 0; i <= colIndex; i++) {
			endCol += cellLenArr[rowIndex][i];
		}
		return endCol - 1;
	}
	
	/**
	 * 创建表格
	 * @param workBook
	 * @param sheetIndex
	 * @param startRow
	 * @param startColumn
	 * @param dataValue
	 * @param cellLenArr
	 * @param cellHeightArr
	 * @param styleArr
	 * @return
	 */
	@SuppressWarnings("deprecation")
	public static XSSFWorkbook createTable(XSSFWorkbook workBook, int sheetIndex, int startRow, int startColumn,
			String[][] dataValue, int[][] cellLenArr, int[][] cellHeightArr, XSSFCellStyle[][] styleArr) {
		
		int sheetNum = workBook.getNumberOfSheets();
		XSSFSheet sheet = null;
		if( sheetIndex > sheetNum -1){
			sheet = workBook.createSheet("报表数据");
		}else{
			sheet = workBook.getSheetAt(sheetIndex);
		}
		
		// 行数 和 列数
		int rowNum = dataValue.length;
		int colNum = dataValue[0].length;

		// 先合并单元格
		for (int i = 0; i < rowNum; i++) {
			int startColPoint = startColumn; // 列指针
			for (int j = 0; j < colNum; j++) {
				int endColPoint = startColPoint +  cellLenArr[i][j] - 1;
				int startRowPoint = getStartRow(startRow, i, j, cellHeightArr);
				int endRowPoint = getEndRow(startRow, i, j, cellHeightArr);
				CellRangeAddress region = new CellRangeAddress(startRowPoint, endRowPoint, startColPoint, endColPoint);
				sheet.addMergedRegion(region);
				startColPoint = endColPoint + 1;
			}
		}
		// 赋值
		for (int i = 0; i < rowNum; i++) {
			for (int j = 0; j < colNum; j++) { // 一共创建的单元格数量
				// 每个合并单元格的 开始行 结束行 开始列 结束列
				int startRowPoint = getStartRow(startRow, i, j, cellHeightArr);
				int endRowPoint = getEndRow(startRow, i, j, cellHeightArr);
				int newCellFlag = 0; // 标识每个合并单元格的第一个位置

				for (int v = startRowPoint; v <= endRowPoint; v++) { // 遍历行
					XSSFRow row = sheet.getRow(v);
					if (sheet.getRow(v) == null) {
						row = sheet.createRow(v);
					}
					int startColPoint = getStartCol(startColumn, i, j, cellLenArr); // 列控制
					int endColPoint = getEndCol(startColumn, i, j, cellLenArr);
					for (int z = startColPoint; z <= endColPoint; z++) { // 遍历列
						XSSFCell cell = row.createCell(z);
						if(hasFindStyle(styleArr, i, j))
							cell.setCellStyle(styleArr[i][j]);

						String value = dataValue[i][j];
						if (isNumber(value)) {
							cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
						}
						if (newCellFlag == 0) { //存在合并单元格的情况时,只赋值给合并单元格的第一个小单元格
							if (isNumber(value)) {
								cell.setCellValue(Double.valueOf(value));
							}else{
								cell.setCellValue(value);
							}
						}
						newCellFlag++;
					}
				}
			}
		}
		return workBook;
	}
	
	
	public static boolean hasFindStyle( XSSFCellStyle[][] styleArr,int rowIndex,int colIndex){
		try {
			XSSFCellStyle style = styleArr[rowIndex][colIndex];
			if(style == null){
				return false;
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	/**
	 * 写入
	 * @param file 写入哪个Excel文件
	 * @param workBook 
	 * @throws IOException
	 */
	public static void writeExcel(File file,XSSFWorkbook workBook) throws IOException{
		FileOutputStream fileOut = new FileOutputStream(file);
		workBook.write(fileOut);
		fileOut.close();
	}
	
	/**
	 * 整数或小数匹配
	 * 
	 * @param value
	 * @return
	 */
	public static boolean isNumber(String value) {
		if(value == null){
			return false;
		}
		Pattern pattern = Pattern.compile("^(\\-|\\+)?\\d+(\\.\\d+)?$");
		return pattern.matcher(value).matches();
	}
	
	/**
	 * 获取默认的标题头样式
	 * @param workBook
	 * @return
	 */
	public static XSSFCellStyle getDefaultTitleStyle(XSSFWorkbook workBook){
		// -------- 标题样式设置 -------
		XSSFCellStyle titleStyle = workBook.createCellStyle();

		//字体
		XSSFFont font = workBook.createFont();
		font.setColor(HSSFColor.WHITE.index);
		font.setFontName("宋体");
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		titleStyle.setFont(font);
		//自动换行
		titleStyle.setWrapText(true); 
		
		// 边框
		titleStyle.setBorderBottom(CellStyle.BORDER_THIN);
		titleStyle.setBorderTop(CellStyle.BORDER_THIN);
		titleStyle.setBorderLeft(CellStyle.BORDER_THIN);
		titleStyle.setBorderRight(CellStyle.BORDER_THIN);
		//背景色
		titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // ALT_BARS
		titleStyle.setFillForegroundColor(new XSSFColor(new Color(95, 73, 122)));
		
	//		titleStyle.setAlignment(HSSFCellStyle.ALIGN_GENERAL); //没生效,用下面两行替换
			titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			titleStyle.setAlignment(HorizontalAlignment.CENTER);
		
		return titleStyle;
	}
	
	/**
	 * 获取默认的Body样式
	 * @param workBook
	 * @return
	 */
	public static XSSFCellStyle getDefaultBodyStyle(XSSFWorkbook workBook){
		// -------- 内容样式设置 -------
		XSSFCellStyle bodyStyle = workBook.createCellStyle();
		bodyStyle.setBorderBottom(CellStyle.BORDER_THIN);
		bodyStyle.setBorderTop(CellStyle.BORDER_THIN);
		bodyStyle.setBorderLeft(CellStyle.BORDER_THIN);
		bodyStyle.setBorderRight(CellStyle.BORDER_THIN);
		
		//自动换行
		bodyStyle.setWrapText(true); 

		// 两个一起设置才能成功设置颜色
//		bodyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // ALT_BARS
//		bodyStyle.setFillForegroundColor(new XSSFColor(new Color(95, 73, 122)));
		
		//居中
		bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		
		//千分位设置
//		short format = workBook.createDataFormat().getFormat("#,##0");
//		bodyStyle.setDataFormat(format);
		
		return bodyStyle;
	}
	

	/**
	 * 获取默认的行宽数组
	 * @param dataValue
	 * @return
	 */
	public static int[][] getDefaultCellHeightArr(String[][] dataValue,int defaultLen){
		int dRLen = dataValue.length;
		int dCLent = dataValue[0].length;
		int[][] cellHeightArr = new int[dRLen][dCLent];
		
		if(defaultLen <= 0){
			defaultLen =1;
		}
		for(int i = 0; i< dRLen ;i++){
			for(int j=0; j< dCLent ;j++){
			    
				cellHeightArr[i][j] = defaultLen;
			}
		}
		return cellHeightArr;
	}
	
	/**
	 * 获取默认的列宽数组
	 * @param dataValue
	 * @return
	 */
	public static int[][] getDefaultCellLenArr(String[][] dataValue,int defaultLen){
		int dRLen = dataValue.length;
		int dCLent = dataValue[0].length;
		int[][] cellLenArr = new int[dRLen][dCLent];
	
		if(defaultLen <= 0){
			defaultLen =1;
		}
		for(int i = 0; i< dRLen ;i++){
			for(int j=0; j< dCLent ;j++){
				cellLenArr[i][j] = defaultLen;
			}
		}
		return cellLenArr;
	}

}
