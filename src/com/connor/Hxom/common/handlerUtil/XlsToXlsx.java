package com.connor.Hxom.common.handlerUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/***
 * 将.xls文件装换成.xlsx文件，但不能保证样式完全一样
 * 
 * @author Administrator
 *
 */
public class XlsToXlsx {
	/**
	 * 创建新的Xlsx文件
	 * 
	 * @param fileName
	 *            需要创建的文件名
	 * @param source
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	private File transform(String fileName, File source) throws IOException, InvalidFormatException {
		// 这个对象可以解析xls和xlsx文件，管他xls还是xlsx文件，统一转换成xlsx文件
		Workbook src_wb = WorkbookFactory.create(new FileInputStream(source));
		// 需要新建一个能解析xlsx文件的对象，否则会出现格式错误甚至文件损坏
		Workbook target_wb = new XSSFWorkbook();
		// 输出文件
		File out = new File(fileName);
		// 输出文件流
		FileOutputStream outputStream = new FileOutputStream(out);
		// 目标excel文件的
		int sheetNum = src_wb.getNumberOfSheets();
		XSSFSheet[] target = new XSSFSheet[sheetNum];
		Sheet[] src = new Sheet[sheetNum];
		for (int i = 0; i < sheetNum; i++) {
			// 获取一个表
			src[i] = src_wb.getSheetAt(i);
			String src_SheetName = src[i].getSheetName();
			// 需要新建的sheet页
			target[i] = ((XSSFWorkbook) target_wb).createSheet(src_SheetName);
			int rowNum = src[i].getLastRowNum();
			for (int j = 0; j <= rowNum; j++) {
				Row src_Row = src[i].getRow(j);
				Row target_Row = target[i].createRow(j);
				target_Row.setHeight(src_Row.getHeight());
				// 一页行数
				int cellNum = src_Row.getLastCellNum();
				// 创建单元格
				for (int k = 0; k <= cellNum; k++) {
					target_Row.createCell(k);
				}
				// 数据填充
				for (int k = 0; k < cellNum; k++) {
					Cell src_cell = src_Row.getCell(k);
					src_cell.getCellStyle();
					XSSFCell target_cell = (XSSFCell) target_Row.getCell(k);
					target_cell.getCellStyle();
					switch (src_cell.getCellType()) {
					case HSSFCell.CELL_TYPE_STRING:
						target_cell.setCellValue(src_cell.getStringCellValue());
						break;
					case HSSFCell.CELL_TYPE_BOOLEAN:
						target_cell.setCellValue(src_cell.getBooleanCellValue());
						break;
					case HSSFCell.CELL_TYPE_NUMERIC:
						target_cell.setCellValue(src_cell.getNumericCellValue());
						break;
					case HSSFCell.CELL_TYPE_FORMULA:
						target_cell.setCellValue(src_cell.getCellFormula());
						break;
					case HSSFCell.CELL_TYPE_BLANK:
						target_cell.setCellValue(src_cell.getStringCellValue());
						break;
					case HSSFCell.CELL_TYPE_ERROR:
						target_cell.setCellValue(src_cell.getErrorCellValue());
						break;
					}
				}
			}
		}
		for (int i = 0; i < target.length; i++) {
			int Regions = src[i].getNumMergedRegions();
			for (int k = 0; k < Regions; k++) {
				CellRangeAddress cra = src[i].getMergedRegion(k);
				target[i].addMergedRegion(cra);
			}
		}
		target_wb.write(outputStream);
		return out;
	}

	public static void main(String[] args) throws InvalidFormatException, IOException {
		String path = System.getProperty("user.home");
		path += "\\Teamcenter\\temp\\";// 构成临时文件
		String fileName = path + "专用工艺装备组合夹具明细表AA.xlsx";
		String sourcePath = path + "x\\专用工艺装备组合夹具明细表.xls";
		// 源文件
		File source = new File(sourcePath);
		// 需要转换到的目标文件
		File target = new XlsToXlsx().transform(fileName, source);
		System.out.println("结束");
	}
}
