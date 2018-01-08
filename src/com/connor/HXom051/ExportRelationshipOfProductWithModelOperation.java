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
	/** 只有选中两个版本才能进行下一步 */
	private boolean pass = false;
	/** 装配图版本 */
	private TCComponentItemRevision HX3_ZPTRevision = null;
	/** 专用工艺装备组合家具明细表版本 */
	private TCComponentItemRevision HX3_GYZBJJMXBRevision = null;
	/** 需要到处的一些属性 */
	private ArrayList<ExportBean> exportBean = null;
	/** 模板所在的Item首选项 */
	private String preference = "TC_custom_exportRelationship";
	/** 记录所有的子BoMLine */
	private ArrayList<TCComponentBOMLine> targetBomLine = null;
	/** 模板最大行数 */
	private int maxRwo = 0;
	// 测试用，建议测试打开
	private final boolean debug = false;
	/** 导出模板 */
	private File exportFile = null;
	/** 尾部行数 */
	private int endCount = 8;
	/** 头部行数 */
	private int headCount = 2;
	/** 测试机上跟实机数据不一样，方便测试和发布 */
	/** 测试中英文对应，测试的时候用英文可能出现ZPT和LJT，交付客户方是用中文，方便测试 */
	private int modelselector = 0;
	private final String[] TEST_MODEL = { "ZPT", "LJT" };
	private final String[] PUBLISH_MODEL = { "装配图", "零件图" };

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
	 * 原来有的xls需要删掉吗？
	 * 
	 * @throws IOException
	 * @throws TCException
	 */
	private void exportBean() throws IOException, TCException {
		// 载入模板
		String path = System.getProperty("user.home");
		path += "\\Teamcenter\\temp";// 构成临时文件
		File file = new File(path);

		// 没有模板，不进行下一步操作
		if (!checkModel(path, file)) {
			return;
		}
		// 有模板
		File tempFile = new File(path + "\\" + HX3_GYZBJJMXBRevision.getProperty("item_id") + ".xlsx");
		if (!tempFile.exists()) {
			copyFile(exportFile, tempFile);
		}
		exportExcel(exportFile, tempFile);

	}

	/**
	 * 导出bean到excel中
	 * 
	 * @param exportFile
	 *            模板
	 * @param tempFile
	 *            目标excel，一会上传到HX3_GYZBJJMXBRevision上
	 * @throws IOException
	 * @throws TCException
	 */
	private void exportExcel(File exportFile, File tempFile) throws IOException, TCException {
		InputStream input = new FileInputStream(exportFile);
		FileOutputStream output = new FileOutputStream(tempFile); // 读取的文件路径
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
		// 最多的行数，会限制左右插入
		// int capacity = (int) Math.ceil((double) exportBean.size() / 2);
		// 导出项目过多，需要动态插入
		if (exportBean.size() > maxRwo * 2) {
			insertStatue = JOptionPane.showConfirmDialog(null, "导出项目超出" + maxRwo * 2 + "项，是否动态插入?", "超出限制",
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
		// 动态插入
		// 负责动态添加单元格并合并，样式到下面再添加
		if (insert) {
			// 动态生成
			/*
			 * 比如有11个插入项，只有3行，则需要多添加（11-3*2）/2=2.5->升到3行，再多插入3行，这三行在倒数第8行的上一行前插入
			 * ，保证下面能获取到样式
			 */
			int inserter = (int) Math.ceil(((double) (exportBean.size() - maxRwo * 2)) / 2);
			sheet.shiftRows(sheet.getLastRowNum() - endCount + 1, sheet.getLastRowNum(), inserter);
			// 插入还不行，还要createRow才会有一个新的row，否则是null
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
			// 插入行
			// 合并单元格
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
			// 获取行Math.ceil((double) exportBean.size() / 2 - 1
			if (i < maxRwo) {
				row = sheet.getRow(i + startRow);
				// 模板是固定格式，不需要添加
				toLeft = true;
				rows.add(row);
			} else if (i >= maxRwo && i < maxRwo * 2) {
				toRight = true;
				nowRow = nowRow - maxRwo - headCount - 1;
			}
			// 放在模板左边
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
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);// 文本格式
					switch (cols) {
					case 1:// 第1列，从0开始，合并之后从合并后的第一列开始插入
						cell.setCellValue(bean.getIndex());// 写入内容
						break;
					case 2:// 第2列
						cell.setCellValue(bean.getSerialNumber());// 写入内容
						break;
					case 4:// 第4列
						cell.setCellValue(bean.getName());// 写入内容
						break;
					case 6:// 第6列
						cell.setCellValue(bean.getOfUsingNumber());// 写入内容
						break;
					case 9:// 第9列
						cell.setCellValue(bean.getComment());// 写入内容
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
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);// 文本格式
					switch (cols) {
					case 12:// 第1列，从0开始，合并之后从合并后的第一列开始插入
						cell.setCellValue(bean.getIndex());// 写入内容
						break;
					case 13:// 第2列
						cell.setCellValue(bean.getSerialNumber());// 写入内容
						break;
					case 15:// 第4列
						cell.setCellValue(bean.getName());// 写入内容
						break;
					case 17:// 第6列
						cell.setCellValue(bean.getOfUsingNumber());// 写入内容
						break;
					case 19:// 第9列
						cell.setCellValue(bean.getComment());// 写入内容
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
		// 上传数据集
		uploadAndDeleteHistoryDataset(tempFile);

	}

	/***
	 * 上传文件，同时会删除历史同名文件和临时文件
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
		// 删除历史文件
		HX3_GYZBJJMXBRevision.remove("TC_Attaches", oldDataset.toArray(new TCComponent[oldDataset.size()]));
		HX3_GYZBJJMXBRevision.add("TC_Attaches", dataset);
		// 删除临时文件，只留下模板
		if (!debug) {
			tempFile.delete();
		}
		MessageBox.post("导出专用工艺装备组合夹具明细表成功，挂在" + HX3_GYZBJJMXBRevision.getProperty("object_string") + "下", "成功",
				MessageBox.INFORMATION);

	}

	/**
	 * 检查是否有模板
	 * 
	 * @param path
	 *            根路径
	 * @param file
	 *            模板文件夹
	 * @param exportFile
	 *            导出的模板
	 * @return
	 * @throws TCException
	 * @throws IOException
	 */

	private boolean checkModel(String path, File file) throws TCException, IOException {
		// 根据首选项获取到有模板的Item，并确定模板的版本
		TCPreferenceService service = session.getPreferenceService();
		String targetID = service.getStringValue(preference);
		TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
		TCComponentItem targetItem = null;
		try {
			targetItem = itemType.findItems(targetID)[0];
		} catch (Exception e) {
			MessageBox.post("根据配置的首选项" + preference + "找不到指定的Item:" + targetID + "，请联系系统管理员", "错误", MessageBox.ERROR);
		}
		TCComponentItemRevision targetItemRevision = targetItem.getLatestItemRevision();
		TCComponentDataset xls = (TCComponentDataset) targetItemRevision.getRelatedComponent("TC_Attaches");
		if (null == xls) {
			MessageBox.post(
					targetItemRevision.getProperty("object_name") + "没有相应的模板，请确认模板是否已删除。挂载关系是TC_Attaches，请联系系统管理员!", "",
					MessageBox.ERROR);
			return false;
		}
		File[] f = xls.getFiles("excel");// 模板下载回本地下载回来
		// 先确认有没有模板
		if (f.length == 0) {
			MessageBox.post(targetItemRevision.getProperty("object_name") + "没有相应的模板，请确认模板是否已删除。请联系系统管理员!", "错误，没有模板",
					MessageBox.ERROR);
			return false;
		}
		String fileName = f[0].getName();
		if (!fileName.endsWith(".xlsx")) {
			System.out.println("模板不是xlsx格式，请更新模板格式");
			MessageBox.post("导出失败，模板格式错误！", "错误", MessageBox.ERROR);
			return false;
		}
		exportFile = new File(path + "\\" + fileName.substring(0, fileName.indexOf("."))
				+ targetItemRevision.getProperty("item_revision_id") + ".xlsx");
		// 文件不存在则需要下载
		if (!file.exists()) {
			file.mkdirs();
		}
		if (!exportFile.exists()) {
			copyFile(f[0], exportFile);
		}
		return true;
	}

	/***
	 * 通过nio文件管道，复制文件
	 * 
	 * @param inputFile
	 * @param outputFile
	 * @throws IOException
	 */
	private void copyFile(File inputFile, File outputFile) throws IOException {
		int length = 2097152;// 2M大小
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
	 * 获取到版本的item_id，伪文件夹下版本的item_id，object_name，item_comment，
	 * 存放在bean中用于导出到excel表
	 * 
	 * @throws Exception
	 */
	private void getProperties() throws Exception {
		int index = 1;
		for (TCComponentBOMLine BoMLine : targetBomLine) {
			TCComponentItemRevision revision = BoMLine.getItemRevision();
			// 使用零件图号
			String item_id = revision.getProperty("item_id");
			// 模具图伪文件夹【HX3_MUJT】(根据关系找文件夹)
			TCComponent MUJT = revision.getRelatedComponent("HX3_MUJT");
			if (MUJT != null) {
				TCComponent[] HX3_MUJT = revision.getRelatedComponents("HX3_MUJT");
				if (HX3_MUJT.length <= 0) {
					System.out.println("关系下不存在任何GZT");
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
				System.out.println("序号：" + bean.getIndex() + "编号:" + bean.getSerialNumber() + "\t名称：" + bean.getName()
						+ "\t使用零件图号：" + bean.getOfUsingNumber() + "\t备注：" + bean.getComment());
			}

	}

	/**
	 * 遍历选中装配图版本下的BOM视图，找到该产品下相关的零件图和组件图(找BoMLine)，记录下这些图纸的编号，插入到报表中的”使用零件图号”(
	 * 根据关系找Item_id)中；再找到这些零件图和组件图下的模具图伪文件夹【HX3_MUJT】(根据关系找文件夹)，提取这些模具图HX3_GZT(
	 * 可能有多个)的item_id、object_name和模具图版本表单HX3_GZTRevisionMaster的item_comment，
	 * 插入到报表中的“编号”、“名称”、“备注”中
	 */
	private void getBoMLine() {
		// 初始化集合
		exportBean = new ArrayList<ExportBean>();
		targetBomLine = new ArrayList<TCComponentBOMLine>();
		String HX3_ZPTRevision_Item_id = null;
		if (pass) {
			try {
				HX3_ZPTRevision_Item_id = HX3_ZPTRevision.getProperty("item_id");
				// 从session中获取type
				TCComponentBOMWindowType type = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");// TCComponentBOMWindowType();
				// 搭建bom window
				TCComponentBOMWindow window = type.create(null);
				// 设置顶层bom行
				window.setWindowTopLine(HX3_ZPTRevision.getItem(), HX3_ZPTRevision, null, null);
				targetBomLine.add(window.getTopBOMLine());
				// 方便看有没有进递归
				if (1 == modelselector) {
					System.out.println("关系导出测试阶段");
				} else if (0 == modelselector) {
					System.out.println("关系导出发布阶段");
				}
				for (AIFComponentContext child : window.getTopBOMLine().getChildren()) {
					TCComponentBOMLine tcComponent = (TCComponentBOMLine) child.getComponent();
					targetBomLine.add(tcComponent);
					loopBom(tcComponent);
				}
			} catch (Exception e) {
				e.printStackTrace();
				MessageBox.post(HX3_ZPTRevision_Item_id + "下没有视图", "错误", MessageBox.ERROR);
			}
		}
	}

	/**
	 * 遍历BoMLine寻找目标子BoMLine，采用递归
	 * 
	 * @param tcComponent
	 * @throws TCException
	 */
	private void loopBom(TCComponentBOMLine tcComponent) throws TCException {
		String type = "";
		for (AIFComponentContext subChild : tcComponent.getChildren()) {
			TCComponentBOMLine subBoMLine = (TCComponentBOMLine) subChild.getComponent();
			// 交付的是中文，测试应改回ZPT和LJT
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
	 * 检查当前选中的对象是否都选中了装配图版本(HX3_ZPTRevision)和专用工艺装备组合家具明细表版本(
	 * HX3_GYZBJJMXBRevision)
	 * 
	 * @throws Exception
	 */
	private boolean checkout() throws Exception {
		target = app.getTargetComponents();
		boolean got_HX3_ZPTRevision = false;
		boolean got_HX3_GYZBJJMXBRevision = false;
		/*
		 * 只有同时选中装配图版本(HX3_ZPTRevision)和专用工艺装备组合家具明细表版本(HX3_GYZBJJMXBRevision)，
		 * 才可以调用菜单【报表】-【导出专用工艺装备组合家具明细表】
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
			MessageBox.post("请选择装配图版本和专用工艺装备组合家具明细表版本后，再导出报表", "错误", MessageBox.ERROR);
			return false;
		}
		// 选中了两个目标
		if (got_HX3_ZPTRevision && got_HX3_GYZBJJMXBRevision) {
			this.pass = true;
			return true;
		} else {
			MessageBox.post("请选择装配图版本和专用工艺装备组合家具明细表版本后，再导出报表", "错误", MessageBox.ERROR);
			return false;
		}
	}

}
