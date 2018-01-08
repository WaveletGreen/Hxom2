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
	/** �����ṩ */
	private ArrayList<PPAPBean> providerBean = null;
	/** �˿��ṩ */
	private ArrayList<PPAPBean> customerBean = null;
	/** ģ��������� */
	private int maxRwo = 0;
	// ������
	private final boolean debug = false;
	/** ����ģ�� */
	private File exportFile = null;
	/** β������ */
	private int endCount = 0;
	/** ͷ������ */
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
	 * ������
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
	 * ��ѡ���ļ��н��е���
	 */
	private void openSelector() {
		jFileChooser = new JFileChooser();
		jFileChooser.setDialogTitle("��ѡ���ŵ��ļ���");
		FileSystemView fsv = FileSystemView.getFileSystemView();
		// ��ǰ�û�������·��
		String deskPath = fsv.getHomeDirectory().getPath();
		this.jFileChooser.setCurrentDirectory(new File(deskPath));// �ļ�ѡ�����ĳ�ʼĿ¼��Ϊ��ǰ�û�����
		this.jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		int state = jFileChooser.showOpenDialog(null);
		jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);// ֻ��ѡ��Ŀ¼
		if (state == JFileChooser.APPROVE_OPTION) {
			path = jFileChooser.getSelectedFile();
		}
	}

	/***
	 * �Ӳ�ѯ�������в�ѯ����ItemRevision����ŵ�bean��,���bean��ͨ�õ�233
	 * 
	 * @param searcher
	 * @param beanList
	 * @throws Exception
	 */
	private void getSearchPPAP(String searcher, ArrayList<PPAPBean> beanList) throws Exception {
		TCComponent[] component = session.search(searcher, new String[] { "����" }, new String[] { "*" });
		int index = 1;
		for (TCComponent child : component) {
			PPAPBean bean = new PPAPBean();
			bean.setIndex(index);
			TCComponentItemRevision revision = (TCComponentItemRevision) child;
			String date_released = revision.getProperty("date_released");
			bean.setRatifyDate(date_released);// ��ȡ��׼����
			TCComponent form = revision.getRelatedComponent("IMAN_master_form_rev");
			String[] prop = form.getProperties(PPAPBean.Attr);
			// ��˳����,���÷��䣬�����Ҫ��˳���ŷŶ������˳����bean��setter
			for (int i = 0; i < PPAPBean.Attr.length; i++) {
				bean._setAttr(PPAPBean.Attr[i], prop[i]);
			}
			beanList.add(bean);
			index++;
		}
	}

	/**
	 * ԭ���е�xls��Ҫɾ����
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
		exportFile = new File(path + "\\PPAD�����.xlsx");
		ArrayList<ArrayList<PPAPBean>> toSheet = new ArrayList<>();
		toSheet.add(providerBean);
		toSheet.add(customerBean);
		exportExcel(exportFile, toSheet);
	}

	/**
	 * ����bean��excel��,��ʵ�Ҳ�֪������дʲô
	 *
	 * @param exportFile
	 *            ģ��
	 * @param toSheet
	 *            ���bean�ļ���
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
		InputStream input = this.getClass().getResourceAsStream("PPAP�����.xlsx");
		FileOutputStream output = new FileOutputStream(exportFile); // ��ȡ���ļ�·��
		XSSFWorkbook wb = new XSSFWorkbook(input);
		for (int in = 0; in < wb.getNumberOfSheets(); in++) {
			ArrayList<PPAPBean> beans = toSheet.get(in);
			XSSFSheet sheet = wb.getSheetAt(in);
			maxRwo = sheet.getLastRowNum() + 1 - headCount - endCount;
			int startRow = headCount;
			short cols;
			XSSFRow sourceRow = sheet.getRow(headCount);// ��6�п�ʼ��Ϊģ��
			boolean insert = false;
			// ������Ŀ���࣬��Ҫ��̬����
			if (beans.size() > maxRwo) {
				insert = true;
			}
			// ��̬����
			// ����̬��ӵ�Ԫ�񲢺ϲ�����ʽ�����������
			if (insert) {
				/*
				 * ������11�������ֻ��3�У�����Ҫ����ӣ�11-3*2��/2=2.5->����3�У��ٶ����3�У�
				 * �������ڵ�����8�е���һ��ǰ���� ����֤�����ܻ�ȡ����ʽ
				 */
				int inserter = (int) Math.ceil((double) (beans.size() - maxRwo));
				sheet.shiftRows(sheet.getLastRowNum() - endCount, sheet.getLastRowNum(), inserter);
				// ���뻹���У���ҪcreateRow�Ż���һ���µ�row��������null
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
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);// �ı���ʽ
				}
				++startRow;

			}
		}
		wb.write(output);
		output.close();
		input.close();
		MessageBox.post("PPAP��������ɹ�", "�ɹ�", MessageBox.INFORMATION);
	}

}
