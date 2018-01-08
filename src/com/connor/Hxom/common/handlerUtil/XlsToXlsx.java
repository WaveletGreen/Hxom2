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
 * ��.xls�ļ�װ����.xlsx�ļ��������ܱ�֤��ʽ��ȫһ��
 * 
 * @author Administrator
 *
 */
public class XlsToXlsx {
	/**
	 * �����µ�Xlsx�ļ�
	 * 
	 * @param fileName
	 *            ��Ҫ�������ļ���
	 * @param source
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	private File transform(String fileName, File source) throws IOException, InvalidFormatException {
		// ���������Խ���xls��xlsx�ļ�������xls����xlsx�ļ���ͳһת����xlsx�ļ�
		Workbook src_wb = WorkbookFactory.create(new FileInputStream(source));
		// ��Ҫ�½�һ���ܽ���xlsx�ļ��Ķ��󣬷������ָ�ʽ���������ļ���
		Workbook target_wb = new XSSFWorkbook();
		// ����ļ�
		File out = new File(fileName);
		// ����ļ���
		FileOutputStream outputStream = new FileOutputStream(out);
		// Ŀ��excel�ļ���
		int sheetNum = src_wb.getNumberOfSheets();
		XSSFSheet[] target = new XSSFSheet[sheetNum];
		Sheet[] src = new Sheet[sheetNum];
		for (int i = 0; i < sheetNum; i++) {
			// ��ȡһ����
			src[i] = src_wb.getSheetAt(i);
			String src_SheetName = src[i].getSheetName();
			// ��Ҫ�½���sheetҳ
			target[i] = ((XSSFWorkbook) target_wb).createSheet(src_SheetName);
			int rowNum = src[i].getLastRowNum();
			for (int j = 0; j <= rowNum; j++) {
				Row src_Row = src[i].getRow(j);
				Row target_Row = target[i].createRow(j);
				target_Row.setHeight(src_Row.getHeight());
				// һҳ����
				int cellNum = src_Row.getLastCellNum();
				// ������Ԫ��
				for (int k = 0; k <= cellNum; k++) {
					target_Row.createCell(k);
				}
				// �������
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
		path += "\\Teamcenter\\temp\\";// ������ʱ�ļ�
		String fileName = path + "ר�ù���װ����ϼо���ϸ��AA.xlsx";
		String sourcePath = path + "x\\ר�ù���װ����ϼо���ϸ��.xls";
		// Դ�ļ�
		File source = new File(sourcePath);
		// ��Ҫת������Ŀ���ļ�
		File target = new XlsToXlsx().transform(fileName, source);
		System.out.println("����");
	}
}
