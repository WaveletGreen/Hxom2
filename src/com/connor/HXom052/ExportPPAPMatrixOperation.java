package com.connor.HXom052;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.teamcenter.rac.aif.AbstractAIFApplication;
import com.teamcenter.rac.aif.AbstractAIFOperation;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class ExportPPAPMatrixOperation extends AbstractAIFOperation {

	private TCSession session;
	private AbstractAIFApplication app;
	/** 供方提供 */
	private ArrayList<PPAPBean> providerBean = null;
	/** 顾客提供 */
	private ArrayList<PPAPBean> customerBean = null;
	/** 模板最大行数 */
	private int maxRwo = 0;
	// 测试用
	private final boolean debug = false;
	/** 导出模板 */
	private File exportFile = null;
	/** 尾部行数 */
	private int endCount = 0;
	/** 头部行数 */
	private int headCount = 6;
	private File path = null;
	private JFileChooser jFileChooser;

	public ExportPPAPMatrixOperation() {
		super();
		this.app = AIFUtility.getCurrentApplication();
		this.session = (TCSession) app.getSession();
	}

	@Override
	public void executeOperation() throws Exception {
		openSelector();
		if (null == path) {
			return;
		}
		providerBean = new ArrayList<>();
		getSearchPPAP("SearchGFPPAP", providerBean);
		if (debug) {
			printfBean(providerBean);
		}
		customerBean = new ArrayList<>();
		getSearchPPAP("SearchPPAPTJQD", customerBean);
		if (debug) {
			printfBean(customerBean);
		}
		exportBean();
	}

	/**
	 * 测试用
	 * 
	 * @param providerBean
	 * @throws InvocationTargetException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 * @throws SecurityException
	 * @throws NoSuchMethodException
	 */
	private void printfBean(ArrayList<PPAPBean> providerBean) throws NoSuchMethodException, SecurityException,
			IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		for (PPAPBean bean : providerBean) {
			System.out.print(bean.getIndex() + "\t");
			bean._printBean();
		}

	}

	/**
	 * 打开选择文件夹进行导出
	 */
	private void openSelector() {
		jFileChooser = new JFileChooser();
		jFileChooser.setDialogTitle("请选择存放的文件夹");
		FileSystemView fsv = FileSystemView.getFileSystemView();
		// 当前用户的桌面路径
		String deskPath = fsv.getHomeDirectory().getPath();
		this.jFileChooser.setCurrentDirectory(new File(deskPath));// 文件选择器的初始目录定为当前用户桌面
		this.jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		int state = jFileChooser.showOpenDialog(null);
		jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);// 只能选择目录
		if (state == JFileChooser.APPROVE_OPTION) {
			path = jFileChooser.getSelectedFile();
		}
	}

	/***
	 * 从查询构建器中查询到的ItemRevision，存放到bean中,这个bean是通用的233
	 * 
	 * @param searcher
	 * @param beanList
	 * @throws Exception
	 */
	private void getSearchPPAP(String searcher, ArrayList<PPAPBean> beanList) throws Exception {
		TCComponent[] component = session.search(searcher, new String[] { "名称" }, new String[] { "*" });
		int index = 1;
		for (TCComponent child : component) {
			PPAPBean bean = new PPAPBean();
			bean.setIndex(index);
			TCComponentItemRevision revision = (TCComponentItemRevision) child;
			String date_released = revision.getProperty("date_released");
			bean.setRatifyDate(date_released);// 获取批准日期
			TCComponent form = revision.getRelatedComponent("IMAN_master_form_rev");
			String[] prop = form.getProperties(PPAPBean.Attr);
			// 按顺序存放,采用反射，因此需要按顺序排放而且这个顺序是bean的setter
			for (int i = 0; i < PPAPBean.Attr.length; i++) {
				bean._setAttr(PPAPBean.Attr[i], prop[i]);
			}
			beanList.add(bean);
			index++;
		}
	}

	/**
	 * 原来有的xls需要删掉吗？
	 * 
	 * @throws IOException
	 * @throws TCException
	 * @throws InvocationTargetException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 * @throws SecurityException
	 * @throws NoSuchMethodException
	 */
	private void exportBean() throws IOException, TCException, NoSuchMethodException, SecurityException,
			IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		exportFile = new File(path + "\\PPAD矩阵表.xlsx");
		ArrayList<ArrayList<PPAPBean>> toSheet = new ArrayList<>();
		toSheet.add(providerBean);
		toSheet.add(customerBean);
		exportExcel(exportFile, toSheet);
	}

	/**
	 * 导出bean到excel中,其实我不知道我在写什么
	 *
	 * @param exportFile
	 *            模板
	 * @param toSheet
	 *            存放bean的集合
	 * @throws IOException
	 * @throws TCException
	 * @throws InvocationTargetException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 * @throws SecurityException
	 * @throws NoSuchMethodException
	 */
	private <T> void exportExcel(File exportFile, ArrayList<ArrayList<PPAPBean>> toSheet)
			throws IOException, TCException, NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException {
		InputStream input = this.getClass().getResourceAsStream("PPAP矩阵表.xlsx");
		FileOutputStream output = new FileOutputStream(exportFile); // 读取的文件路径
		XSSFWorkbook wb = new XSSFWorkbook(input);
		for (int in = 0; in < wb.getNumberOfSheets(); in++) {
			ArrayList<PPAPBean> beans = toSheet.get(in);
			XSSFSheet sheet = wb.getSheetAt(in);
			maxRwo = sheet.getLastRowNum() + 1 - headCount - endCount;
			int startRow = headCount;
			short cols;
			XSSFRow sourceRow = sheet.getRow(headCount);// 第6行开始作为模板
			boolean insert = false;
			// 导出项目过多，需要动态插入
			if (beans.size() > maxRwo) {
				insert = true;
			}
			// 动态插入
			// 负责动态添加单元格并合并，样式到下面再添加
			if (insert) {
				/*
				 * 比如有11个插入项，只有3行，则需要多添加（11-3*2）/2=2.5->升到3行，再多插入3行，
				 * 这三行在倒数第8行的上一行前插入 ，保证下面能获取到样式
				 */
				int inserter = (int) Math.ceil((double) (beans.size() - maxRwo));
				sheet.shiftRows(sheet.getLastRowNum() - endCount, sheet.getLastRowNum(), inserter);
				// 插入还不行，还要createRow才会有一个新的row，否则是null
				for (int j = 0; j < inserter; j++) {
					int index = sheet.getLastRowNum() - endCount + j - inserter;
					XSSFRow row = sheet.createRow(index);
					System.out.println(index);
					XSSFCell sourceCell = null;
					for (cols = sourceRow.getFirstCellNum(); cols < sourceRow.getLastCellNum(); cols++) {
						XSSFCell cell = null;
						cell = row.createCell(cols);
						sourceCell = sourceRow.getCell(cols);
						cell.setCellStyle(sourceCell.getCellStyle());
						cell.setCellType(XSSFCell.CELL_TYPE_STRING);
					}
				}
			}
			maxRwo = sheet.getLastRowNum() + 1 - headCount - endCount;
			for (int i = 0; i < beans.size(); i++) {
				PPAPBean bean = beans.get(i);
				XSSFRow row = sheet.getRow(startRow);
				XSSFCell sourceCell = null;
				for (cols = sourceRow.getFirstCellNum(); cols < sourceRow.getLastCellNum(); cols++) {
					XSSFCell cell = null;
					sourceCell = sourceRow.getCell(cols);
					cell = row.getCell(cols);
					T value = bean._getBeanAttr(PPAPBean.publicAttr[cols]);
					if (null == value) {
						cell.setCellValue("");
					} else {
						cell.setCellValue(String.valueOf(value));
					}
					cell.setCellStyle(sourceCell.getCellStyle());
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);// 文本格式
				}
				++startRow;

			}
		}
		wb.write(output);
		output.close();
		input.close();
		MessageBox.post("PPAP矩阵表导出成功", "成功", MessageBox.INFORMATION);
	}

}
