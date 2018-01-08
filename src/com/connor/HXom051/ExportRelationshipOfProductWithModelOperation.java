package com.connor.HXom051;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.channels.FileChannel;
import java.util.ArrayList;

import javax.swing.JOptionPane;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.teamcenter.rac.aif.AbstractAIFApplication;
import com.teamcenter.rac.aif.AbstractAIFOperation;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentBOMWindowType;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class ExportRelationshipOfProductWithModelOperation extends AbstractAIFOperation {

	private TCSession session;
	private AbstractAIFApplication app;
	private InterfaceAIFComponent[] target;
	/** ֻ��ѡ�������汾���ܽ�����һ�� */
	private boolean pass = false;
	/** װ��ͼ�汾 */
	private TCComponentItemRevision HX3_ZPTRevision = null;
	/** ר�ù���װ����ϼҾ���ϸ��汾 */
	private TCComponentItemRevision HX3_GYZBJJMXBRevision = null;
	/** ��Ҫ������һЩ���� */
	private ArrayList<ExportBean> exportBean = null;
	/** ģ�����ڵ�Item��ѡ�� */
	private String preference = "TC_custom_exportRelationship";
	/** ��¼���е���BoMLine */
	private ArrayList<TCComponentBOMLine> targetBomLine = null;
	/** ģ��������� */
	private int maxRwo = 0;
	// �����ã�������Դ�
	private final boolean debug = false;
	/** ����ģ�� */
	private File exportFile = null;
	/** β������ */
	private int endCount = 8;
	/** ͷ������ */
	private int headCount = 2;
	/** ���Ի��ϸ�ʵ�����ݲ�һ����������Ժͷ��� */
	/** ������Ӣ�Ķ�Ӧ�����Ե�ʱ����Ӣ�Ŀ��ܳ���ZPT��LJT�������ͻ����������ģ�������� */
	private int modelselector = 0;
	private final String[] TEST_MODEL = { "ZPT", "LJT" };
	private final String[] PUBLISH_MODEL = { "װ��ͼ", "���ͼ" };

	public ExportRelationshipOfProductWithModelOperation() {
		super();
		this.app = AIFUtility.getCurrentApplication();
		this.session = (TCSession) app.getSession();
	}

	@Override
	public void executeOperation() throws Exception {
		if (checkout()) {
			getBoMLine();
			getProperties();
			exportBean();
		} else {
			return;
		}

	}

	/**
	 * ԭ���е�xls��Ҫɾ����
	 * 
	 * @throws IOException
	 * @throws TCException
	 */
	private void exportBean() throws IOException, TCException {
		// ����ģ��
		String path = System.getProperty("user.home");
		path += "\\Teamcenter\\temp";// ������ʱ�ļ�
		File file = new File(path);

		// û��ģ�壬��������һ������
		if (!checkModel(path, file)) {
			return;
		}
		// ��ģ��
		File tempFile = new File(path + "\\" + HX3_GYZBJJMXBRevision.getProperty("item_id") + ".xlsx");
		if (!tempFile.exists()) {
			copyFile(exportFile, tempFile);
		}
		exportExcel(exportFile, tempFile);

	}

	/**
	 * ����bean��excel��
	 * 
	 * @param exportFile
	 *            ģ��
	 * @param tempFile
	 *            Ŀ��excel��һ���ϴ���HX3_GYZBJJMXBRevision��
	 * @throws IOException
	 * @throws TCException
	 */
	private void exportExcel(File exportFile, File tempFile) throws IOException, TCException {
		InputStream input = new FileInputStream(exportFile);
		FileOutputStream output = new FileOutputStream(tempFile); // ��ȡ���ļ�·��
		XSSFWorkbook wb = new XSSFWorkbook(input);
		XSSFSheet sheet = wb.getSheetAt(0);
		maxRwo = sheet.getLastRowNum() - endCount - headCount;
		int startRow = headCount + 1;
		short cols;
		int nowRow = 0;
		XSSFRow sourceRow = sheet.getRow(2);
		ArrayList<XSSFRow> rows = new ArrayList<>();
		int insertStatue = 0;
		boolean insert = false;
		// �������������������Ҳ���
		// int capacity = (int) Math.ceil((double) exportBean.size() / 2);
		// ������Ŀ���࣬��Ҫ��̬����
		if (exportBean.size() > maxRwo * 2) {
			insertStatue = JOptionPane.showConfirmDialog(null, "������Ŀ����" + maxRwo * 2 + "��Ƿ�̬����?", "��������",
					JOptionPane.YES_NO_CANCEL_OPTION);
			if (insertStatue == JOptionPane.YES_OPTION) {
				insert = true;
			} else if (insertStatue == JOptionPane.NO_OPTION) {
				insert = false;
			} else if (insertStatue == JOptionPane.CANCEL_OPTION) {
				output.close();
				input.close();
				return;
			}
		}
		// ��̬����
		// ����̬��ӵ�Ԫ�񲢺ϲ�����ʽ�����������
		if (insert) {
			// ��̬����
			/*
			 * ������11�������ֻ��3�У�����Ҫ����ӣ�11-3*2��/2=2.5->����3�У��ٶ����3�У��������ڵ�����8�е���һ��ǰ����
			 * ����֤�����ܻ�ȡ����ʽ
			 */
			int inserter = (int) Math.ceil(((double) (exportBean.size() - maxRwo * 2)) / 2);
			sheet.shiftRows(sheet.getLastRowNum() - endCount + 1, sheet.getLastRowNum(), inserter);
			// ���뻹���У���ҪcreateRow�Ż���һ���µ�row��������null
			for (int i = 0; i < inserter; i++) {
				int index = sheet.getLastRowNum() - endCount + 1 + i - inserter;
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
			// ������
			// �ϲ���Ԫ��
			for (int i = sheet.getLastRowNum() - endCount - inserter + 1; i < sheet.getLastRowNum() - endCount
					+ 1; i++) {
				sheet.addMergedRegion(new CellRangeAddress(i, i, 2, 3));
				sheet.addMergedRegion(new CellRangeAddress(i, i, 4, 5));
				sheet.addMergedRegion(new CellRangeAddress(i, i, 6, 8));
				sheet.addMergedRegion(new CellRangeAddress(i, i, 9, 11));
				sheet.addMergedRegion(new CellRangeAddress(i, i, 13, 14));
				sheet.addMergedRegion(new CellRangeAddress(i, i, 15, 16));
				sheet.addMergedRegion(new CellRangeAddress(i, i, 17, 18));
			}
		}
		boolean toLeft = false;
		boolean toRight = false;
		Boolean isCreat = false;
		maxRwo = sheet.getLastRowNum() - endCount - headCount;

		if (exportBean.size() > 0) {
			XSSFCell src = sheet.getRow(0).getCell(14);
			XSSFCell productType = sheet.getRow(0).getCell(15);
			XSSFCell productName = sheet.getRow(1).getCell(15);
			productType.setCellValue(exportBean.get(0).getOfUsingNumber());
			productType.setCellStyle(src.getCellStyle());
			//String obName=HX3_ZPTRevision.getProperty("object_name");
			productName.setCellValue(HX3_ZPTRevision.getProperty("object_name"));
			productName.setCellStyle(src.getCellStyle());
			
		}

		for (int i = 0; i < exportBean.size(); i++) {
			ExportBean bean = exportBean.get(i);
			XSSFRow row = null;
			XSSFCell sourceCell = null;
			nowRow = i + startRow;
			toLeft = false;
			toRight = false;
			// ��ȡ��Math.ceil((double) exportBean.size() / 2 - 1
			if (i < maxRwo) {
				row = sheet.getRow(i + startRow);
				// ģ���ǹ̶���ʽ������Ҫ���
				toLeft = true;
				rows.add(row);
			} else if (i >= maxRwo && i < maxRwo * 2) {
				toRight = true;
				nowRow = nowRow - maxRwo - headCount - 1;
			}
			// ����ģ�����
			if (toLeft) {
				for (cols = sourceRow.getFirstCellNum(); cols < sourceRow.getLastCellNum(); cols++) {
					XSSFCell cell = null;
					if (isCreat) {
						cell = row.createCell(cols);
					} else {
						cell = row.getCell(cols);
					}
					sourceCell = sourceRow.getCell(cols);
					cell.setCellStyle(sourceCell.getCellStyle());
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);// �ı���ʽ
					switch (cols) {
					case 1:// ��1�У���0��ʼ���ϲ�֮��Ӻϲ���ĵ�һ�п�ʼ����
						cell.setCellValue(bean.getIndex());// д������
						break;
					case 2:// ��2��
						cell.setCellValue(bean.getSerialNumber());// д������
						break;
					case 4:// ��4��
						cell.setCellValue(bean.getName());// д������
						break;
					case 6:// ��6��
						cell.setCellValue(bean.getOfUsingNumber());// д������
						break;
					case 9:// ��9��
						cell.setCellValue(bean.getComment());// д������
						break;
					default:
						break;
					}
				}
			} else if (toRight) {
				for (cols = 12; cols < sourceRow.getLastCellNum(); cols++) {
					XSSFCell cell = null;
					XSSFRow xssfRow = rows.get(nowRow);
					cell = xssfRow.getCell(cols);
					sourceCell = sourceRow.getCell(cols);
					cell.setCellStyle(sourceCell.getCellStyle());
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);// �ı���ʽ
					switch (cols) {
					case 12:// ��1�У���0��ʼ���ϲ�֮��Ӻϲ���ĵ�һ�п�ʼ����
						cell.setCellValue(bean.getIndex());// д������
						break;
					case 13:// ��2��
						cell.setCellValue(bean.getSerialNumber());// д������
						break;
					case 15:// ��4��
						cell.setCellValue(bean.getName());// д������
						break;
					case 17:// ��6��
						cell.setCellValue(bean.getOfUsingNumber());// д������
						break;
					case 19:// ��9��
						cell.setCellValue(bean.getComment());// д������
						break;
					default:
						break;
					}
				}
			}
		}
		wb.write(output);
		output.close();
		input.close();
		System.out.println("-------WRITE EXCEL OVER-------");
		// �ϴ����ݼ�
		uploadAndDeleteHistoryDataset(tempFile);

	}

	/***
	 * �ϴ��ļ���ͬʱ��ɾ����ʷͬ���ļ�����ʱ�ļ�
	 * 
	 * @param tempFile
	 * @throws TCException
	 */
	private void uploadAndDeleteHistoryDataset(File tempFile) throws TCException {
		TCComponentDatasetType t = (TCComponentDatasetType) session.getTypeComponent("Dataset");
		TCComponentDataset dataset = t.create(tempFile.getName(), "", "MSExcelX");
		String pathx[] = { tempFile.getAbsolutePath() };
		String type[] = { "excel" };
		dataset.setFiles(pathx, type);
		TCComponent[] FromRevDataset = HX3_GYZBJJMXBRevision.getRelatedComponents("TC_Attaches");
		ArrayList<TCComponent> oldDataset = new ArrayList<>();
		for (int i = 0; i < FromRevDataset.length; i++) {
			if (FromRevDataset[i] instanceof TCComponentDataset) {
				if (FromRevDataset[i].toString().equals(tempFile.getName())) {
					oldDataset.add(FromRevDataset[i]);
				}
			}
		}
		// ɾ����ʷ�ļ�
		HX3_GYZBJJMXBRevision.remove("TC_Attaches", oldDataset.toArray(new TCComponent[oldDataset.size()]));
		HX3_GYZBJJMXBRevision.add("TC_Attaches", dataset);
		// ɾ����ʱ�ļ���ֻ����ģ��
		if (!debug) {
			tempFile.delete();
		}
		MessageBox.post("����ר�ù���װ����ϼо���ϸ��ɹ�������" + HX3_GYZBJJMXBRevision.getProperty("object_string") + "��", "�ɹ�",
				MessageBox.INFORMATION);

	}

	/**
	 * ����Ƿ���ģ��
	 * 
	 * @param path
	 *            ��·��
	 * @param file
	 *            ģ���ļ���
	 * @param exportFile
	 *            ������ģ��
	 * @return
	 * @throws TCException
	 * @throws IOException
	 */

	private boolean checkModel(String path, File file) throws TCException, IOException {
		// ������ѡ���ȡ����ģ���Item����ȷ��ģ��İ汾
		TCPreferenceService service = session.getPreferenceService();
		String targetID = service.getStringValue(preference);
		TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
		TCComponentItem targetItem = null;
		try {
			targetItem = itemType.findItems(targetID)[0];
		} catch (Exception e) {
			MessageBox.post("�������õ���ѡ��" + preference + "�Ҳ���ָ����Item:" + targetID + "������ϵϵͳ����Ա", "����", MessageBox.ERROR);
		}
		TCComponentItemRevision targetItemRevision = targetItem.getLatestItemRevision();
		TCComponentDataset xls = (TCComponentDataset) targetItemRevision.getRelatedComponent("TC_Attaches");
		if (null == xls) {
			MessageBox.post(
					targetItemRevision.getProperty("object_name") + "û����Ӧ��ģ�壬��ȷ��ģ���Ƿ���ɾ�������ع�ϵ��TC_Attaches������ϵϵͳ����Ա!", "",
					MessageBox.ERROR);
			return false;
		}
		File[] f = xls.getFiles("excel");// ģ�����ػر������ػ���
		// ��ȷ����û��ģ��
		if (f.length == 0) {
			MessageBox.post(targetItemRevision.getProperty("object_name") + "û����Ӧ��ģ�壬��ȷ��ģ���Ƿ���ɾ��������ϵϵͳ����Ա!", "����û��ģ��",
					MessageBox.ERROR);
			return false;
		}
		String fileName = f[0].getName();
		if (!fileName.endsWith(".xlsx")) {
			System.out.println("ģ�岻��xlsx��ʽ�������ģ���ʽ");
			MessageBox.post("����ʧ�ܣ�ģ���ʽ����", "����", MessageBox.ERROR);
			return false;
		}
		exportFile = new File(path + "\\" + fileName.substring(0, fileName.indexOf("."))
				+ targetItemRevision.getProperty("item_revision_id") + ".xlsx");
		// �ļ�����������Ҫ����
		if (!file.exists()) {
			file.mkdirs();
		}
		if (!exportFile.exists()) {
			copyFile(f[0], exportFile);
		}
		return true;
	}

	/***
	 * ͨ��nio�ļ��ܵ��������ļ�
	 * 
	 * @param inputFile
	 * @param outputFile
	 * @throws IOException
	 */
	private void copyFile(File inputFile, File outputFile) throws IOException {
		int length = 2097152;// 2M��С
		FileInputStream in = new FileInputStream(inputFile);
		FileOutputStream out = new FileOutputStream(outputFile);
		FileChannel inC = in.getChannel();
		FileChannel outC = out.getChannel();
		while (true) {
			if (inC.position() == inC.size()) {
				inC.close();
				outC.close();
				in.close();
				out.close();
				return;
			}
			if ((inC.size() - inC.position()) < 20971520)
				length = (int) (inC.size() - inC.position());
			else
				length = 20971520;
			inC.transferTo(inC.position(), length, outC);
			inC.position(inC.position() + length);
		}
	}

	/**
	 * ��ȡ���汾��item_id��α�ļ����°汾��item_id��object_name��item_comment��
	 * �����bean�����ڵ�����excel��
	 * 
	 * @throws Exception
	 */
	private void getProperties() throws Exception {
		int index = 1;
		for (TCComponentBOMLine BoMLine : targetBomLine) {
			TCComponentItemRevision revision = BoMLine.getItemRevision();
			// ʹ�����ͼ��
			String item_id = revision.getProperty("item_id");
			// ģ��ͼα�ļ��С�HX3_MUJT��(���ݹ�ϵ���ļ���)
			TCComponent MUJT = revision.getRelatedComponent("HX3_MUJT");
			if (MUJT != null) {
				TCComponent[] HX3_MUJT = revision.getRelatedComponents("HX3_MUJT");
				if (HX3_MUJT.length <= 0) {
					System.out.println("��ϵ�²������κ�GZT");
					continue;
				}
				for (TCComponent tcComponent : HX3_MUJT) {
					TCComponentItem item = (TCComponentItem) tcComponent;
					TCComponent target = tcComponent;
					ExportBean bean = new ExportBean();
					bean.setIndex(index);
					bean.setOfUsingNumber(item_id);
					bean.setSerialNumber(target.getProperty("item_id"));
					bean.setName(target.getProperty("object_name"));
					bean.setComment(item.getLatestItemRevision().getRelatedComponent("IMAN_master_form_rev")
							.getProperty("item_comment"));
					exportBean.add(bean);
					++index;
				}
			}
		}
		if (debug)
			for (ExportBean bean : exportBean) {
				System.out.println("��ţ�" + bean.getIndex() + "���:" + bean.getSerialNumber() + "\t���ƣ�" + bean.getName()
						+ "\tʹ�����ͼ�ţ�" + bean.getOfUsingNumber() + "\t��ע��" + bean.getComment());
			}

	}

	/**
	 * ����ѡ��װ��ͼ�汾�µ�BOM��ͼ���ҵ��ò�Ʒ����ص����ͼ�����ͼ(��BoMLine)����¼����Щͼֽ�ı�ţ����뵽�����еġ�ʹ�����ͼ�š�(
	 * ���ݹ�ϵ��Item_id)�У����ҵ���Щ���ͼ�����ͼ�µ�ģ��ͼα�ļ��С�HX3_MUJT��(���ݹ�ϵ���ļ���)����ȡ��Щģ��ͼHX3_GZT(
	 * �����ж��)��item_id��object_name��ģ��ͼ�汾��HX3_GZTRevisionMaster��item_comment��
	 * ���뵽�����еġ���š��������ơ�������ע����
	 */
	private void getBoMLine() {
		// ��ʼ������
		exportBean = new ArrayList<ExportBean>();
		targetBomLine = new ArrayList<TCComponentBOMLine>();
		String HX3_ZPTRevision_Item_id = null;
		if (pass) {
			try {
				HX3_ZPTRevision_Item_id = HX3_ZPTRevision.getProperty("item_id");
				// ��session�л�ȡtype
				TCComponentBOMWindowType type = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");// TCComponentBOMWindowType();
				// �bom window
				TCComponentBOMWindow window = type.create(null);
				// ���ö���bom��
				window.setWindowTopLine(HX3_ZPTRevision.getItem(), HX3_ZPTRevision, null, null);
				targetBomLine.add(window.getTopBOMLine());
				// ���㿴��û�н��ݹ�
				if (1 == modelselector) {
					System.out.println("��ϵ�������Խ׶�");
				} else if (0 == modelselector) {
					System.out.println("��ϵ���������׶�");
				}
				for (AIFComponentContext child : window.getTopBOMLine().getChildren()) {
					TCComponentBOMLine tcComponent = (TCComponentBOMLine) child.getComponent();
					targetBomLine.add(tcComponent);
					loopBom(tcComponent);
				}
			} catch (Exception e) {
				e.printStackTrace();
				MessageBox.post(HX3_ZPTRevision_Item_id + "��û����ͼ", "����", MessageBox.ERROR);
			}
		}
	}

	/**
	 * ����BoMLineѰ��Ŀ����BoMLine�����õݹ�
	 * 
	 * @param tcComponent
	 * @throws TCException
	 */
	private void loopBom(TCComponentBOMLine tcComponent) throws TCException {
		String type = "";
		for (AIFComponentContext subChild : tcComponent.getChildren()) {
			TCComponentBOMLine subBoMLine = (TCComponentBOMLine) subChild.getComponent();
			// �����������ģ�����Ӧ�Ļ�ZPT��LJT
			type = subBoMLine.getProperty("bl_item_object_type");
			if (1 == modelselector) {
				if (TEST_MODEL[0].equals(type)) {
					loopBom(subBoMLine);
				} else if (TEST_MODEL[1].equals(type)) {
					targetBomLine.add(subBoMLine);
				}
			} else if (0 == modelselector) {
				if (PUBLISH_MODEL[0].equals(type)) {
					loopBom(subBoMLine);
				} else if (PUBLISH_MODEL[1].equals(type)) {
					targetBomLine.add(subBoMLine);
				}
			}
		}
	}

	/**
	 * ��鵱ǰѡ�еĶ����Ƿ�ѡ����װ��ͼ�汾(HX3_ZPTRevision)��ר�ù���װ����ϼҾ���ϸ��汾(
	 * HX3_GYZBJJMXBRevision)
	 * 
	 * @throws Exception
	 */
	private boolean checkout() throws Exception {
		target = app.getTargetComponents();
		boolean got_HX3_ZPTRevision = false;
		boolean got_HX3_GYZBJJMXBRevision = false;
		/*
		 * ֻ��ͬʱѡ��װ��ͼ�汾(HX3_ZPTRevision)��ר�ù���װ����ϼҾ���ϸ��汾(HX3_GYZBJJMXBRevision)��
		 * �ſ��Ե��ò˵�������-������ר�ù���װ����ϼҾ���ϸ��
		 */
		if (2 == target.length) {
			for (int i = 0; i < target.length; i++) {
				if (debug)
					System.out.println(target[i].getProperty("object_type"));
				if ("GYZBJJMXB Revision".equals(target[i].getProperty("object_type"))) {
					got_HX3_ZPTRevision = true;
					HX3_GYZBJJMXBRevision = (TCComponentItemRevision) target[i];
				}
				if ("ZPT Revision".equals(target[i].getProperty("object_type"))) {
					got_HX3_GYZBJJMXBRevision = true;
					HX3_ZPTRevision = (TCComponentItemRevision) target[i];
				}
			}
		} else {
			MessageBox.post("��ѡ��װ��ͼ�汾��ר�ù���װ����ϼҾ���ϸ��汾���ٵ�������", "����", MessageBox.ERROR);
			return false;
		}
		// ѡ��������Ŀ��
		if (got_HX3_ZPTRevision && got_HX3_GYZBJJMXBRevision) {
			this.pass = true;
			return true;
		} else {
			MessageBox.post("��ѡ��װ��ͼ�汾��ר�ù���װ����ϼҾ���ϸ��汾���ٵ�������", "����", MessageBox.ERROR);
			return false;
		}
	}

}
