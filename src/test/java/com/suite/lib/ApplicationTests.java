package com.suite.lib;

import com.suite.lib.util.ExcelUtil;
import com.suite.lib.util.WordUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.util.*;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ApplicationTests {

	@Test
	public void contextLoads() {
	}

	/**
	 * 测试用文件的统一存放目录
	 */
	private String testDirectoryPath = "E:\\Computer Programming\\测试用文件夹\\FunctionUnitIntegrationProject";

	/**
	 * 测试ExcelUtil功能单件
	 * @throws Exception
	 */
	@Test
	public void useExcelUtil() throws Exception {
		String createExcelPath = testDirectoryPath + "\\ExcelUtil\\CreateExcel.xls";
		String createExcelPath1 = testDirectoryPath + "\\ExcelUtil\\CreateExcel.xlsx";
		String readExcelPath = createExcelPath1;
		//----- 生成Excel
		try {
			Date currentTime = new Date();
			//=== 生成数据 ===//
			Map<String, List<Map<String, Object>>> sheetMap = new LinkedHashMap<>();
			List<Map<String, Object>> sheet1 = new ArrayList<>();
			List<Map<String, Object>> sheet2 = new ArrayList<>();
			List<Map<String, Object>> sheet3 = new ArrayList<>();
			Map<String, Object> sheet1_row = new LinkedHashMap<>();
			Map<String, Object> sheet1_row1 = new LinkedHashMap<>();
			Map<String, Object> sheet1_row2 = new LinkedHashMap<>();
			Map<String, Object> sheet1_row3 = new LinkedHashMap<>();
			sheetMap.put("Sheet - 工作表1", sheet1);
			sheetMap.put("Sheet - 工作表2", sheet2);
			sheetMap.put("Sheet - 工作表3", sheet3);
			sheet1.add(sheet1_row);
			sheet1.add(sheet1_row1);
			sheet1.add(sheet1_row2);
			sheet1.add(sheet1_row3);
			sheet2.addAll(sheet1);
			sheet3.addAll(sheet1);

			sheet1_row.put("序号", "序号");
			sheet1_row.put("姓名", "姓名");
			sheet1_row.put("性别", "性别");
			sheet1_row.put("身份", "身份");
			sheet1_row.put("事件", "事件");
			sheet1_row.put("更新时间", "更新时间");

			sheet1_row1.put("序号", "1");
			sheet1_row1.put("姓名", "张三");
			sheet1_row1.put("性别", "男");
			sheet1_row1.put("身份", "身份1");
			sheet1_row1.put("事件", "事件1");
			sheet1_row1.put("更新时间", currentTime);

			sheet1_row2.put("序号", "2");
			sheet1_row2.put("姓名", "李四");
			sheet1_row2.put("性别", "男");
			sheet1_row2.put("身份", "身份2");
			sheet1_row2.put("事件", "事件1");
			sheet1_row2.put("更新时间", currentTime);

			sheet1_row3.put("序号", "3");
			sheet1_row3.put("姓名", "王五");
			sheet1_row3.put("性别", "女");
			sheet1_row3.put("身份", "身份1");
			sheet1_row3.put("事件", "事件4");
			sheet1_row3.put("更新时间", currentTime);

			//=== 生成输出流 ===//
			File file = new File(createExcelPath);
			if (!file.getParentFile().exists()) {
				file.getParentFile().mkdirs();
			}
			OutputStream os = new FileOutputStream(new File(createExcelPath));
			OutputStream os1 = new FileOutputStream(new File(createExcelPath1));
			//======//
			ExcelUtil.createExcelXls(sheetMap, os);
			ExcelUtil.createExcelXlsx(sheetMap, os1);
			System.out.println("useExcelUtil done!");
		} catch (Exception e) {
			e.printStackTrace();
		}
		//----- 读取Excel
		if (null != readExcelPath && new File(readExcelPath).exists()) {
			Map<String, List<Map<String, Object>>> sheetMap = ExcelUtil.readExcel(readExcelPath);
			System.out.println("ExcelUtil#readExcel " +
					((null != sheetMap && !sheetMap.isEmpty()) ? "complete!" : "undone!")
			);
		}
	}

	/**
	 * 测试WordUtil功能单件
	 * @throws Exception
	 */
	@Test
	public void useWordUtil() throws Exception {
		String createWordPath = testDirectoryPath + "\\WordUtil\\CreateWord.doc";
		String createWordPath1 = testDirectoryPath + "\\WordUtil\\CreateWord.docx";
		String readWordPath = createWordPath1;
		//----- 生成Word
		try {
			//=== 生成数据 ===//
			/* 数据装配这一块其实可以用多维数组来传递消息，以后优化代码可以考虑换用数组 */
			Map<String, List<Map<String, Object>>> wordData = new LinkedHashMap<>();
			List<Map<String, Object>> row = new ArrayList<>();
			List<Map<String, Object>> row1 = new ArrayList<>();
			List<Map<String, Object>> row2 = new ArrayList<>();
			List<Map<String, Object>> row3 = new ArrayList<>();

			Map<String, Object> row_cell1 = new LinkedHashMap<>();
			Map<String, Object> row_cell2 = new LinkedHashMap<>();
			Map<String, Object> row_cell3 = new LinkedHashMap<>();
			Map<String, Object> row_cell4 = new LinkedHashMap<>();

			Map<String, Object> row1_cell1 = new LinkedHashMap<>();
			Map<String, Object> row1_cell2 = new LinkedHashMap<>();
			Map<String, Object> row1_cell3 = new LinkedHashMap<>();
			Map<String, Object> row1_cell4 = new LinkedHashMap<>();

			Map<String, Object> row2_cell1 = new LinkedHashMap<>();
			Map<String, Object> row2_cell2 = new LinkedHashMap<>();
			Map<String, Object> row2_cell3 = new LinkedHashMap<>();
			Map<String, Object> row2_cell4 = new LinkedHashMap<>();

			Map<String, Object> row3_cell1 = new LinkedHashMap<>();
			Map<String, Object> row3_cell2 = new LinkedHashMap<>();
			Map<String, Object> row3_cell3 = new LinkedHashMap<>();
			Map<String, Object> row3_cell4 = new LinkedHashMap<>();

			wordData.put("header", row);
			wordData.put("Row1", row1);
			wordData.put("Row2", row2);
			wordData.put("Row3", row3);

			row.add(row_cell1);
			row.add(row_cell2);
			row.add(row_cell3);
			row.add(row_cell4);

			row1.add(row1_cell1);
			row1.add(row1_cell2);
			row1.add(row1_cell3);
			row1.add(row1_cell4);

			row2.add(row2_cell1);
			row2.add(row2_cell2);
			row2.add(row2_cell3);
			row2.add(row2_cell4);

			row3.add(row3_cell1);
			row3.add(row3_cell2);
			row3.add(row3_cell3);
			row3.add(row3_cell4);

			row_cell1.put("value", "序号");
			row_cell2.put("value", "姓名");
			row_cell3.put("value", "性别");
			row_cell4.put("value", "身份");

			row1_cell1.put("value", "1");
			row1_cell2.put("value", "张三");
			row1_cell3.put("value", "男");
			row1_cell4.put("value", "身份1");

			row2_cell1.put("value", "2");
			row2_cell2.put("value", "李四");
			row2_cell3.put("value", "男");
			row2_cell4.put("value", "身份2");

			row3_cell1.put("value", "3");
			row3_cell2.put("value", "王五");
			row3_cell3.put("value", "女");
			row3_cell4.put("value", "身份1");

			//=== 生成输出流 ===//
			File file = new File(createWordPath);
			if (!file.getParentFile().exists()) {
				file.getParentFile().mkdirs();
			}
			/*OutputStream os = new FileOutputStream(new File(createWordPath));*/
			OutputStream os1 = new FileOutputStream(new File(createWordPath1));
			//======//
			/*WordUtil.createWordDoc(sheetMap, os);*/
			WordUtil.createWordDocx(wordData, os1);
			System.out.println("useExcelUtil done!");
		} catch (Exception e) {
			e.printStackTrace();
		}
		//----- 读取Word
		/*if (null != readWordPath && new File(readWordPath).exists()) {
			Map<String, List<Map<String, Object>>> wordData = WordUtil.readWordDoc(readWordPath);
			System.out.println("WordUtil# " + ((null != wordData && !wordData.isEmpty()) ? "complete!" : "undone!"));
		}*/
	}

	@Test
	public void showURL() throws IOException {
		// 第一种：获取类加载的根路径   D:\git\daotie\daotie\target\classes
		File f = new File(this.getClass().getResource("/").getPath());
		System.out.println(f);

		// 获取当前类的所在工程路径; 如果不加“/”  获取当前类的加载目录  D:\git\daotie\daotie\target\classes\my
		File f2 = new File(this.getClass().getResource("").getPath());
		System.out.println(f2);

		// 第二种：获取项目路径    D:\git\daotie\daotie
		File directory = new File("");// 参数为空
		String courseFile = directory.getCanonicalPath();
		System.out.println(courseFile);


		// 第三种：  file:/D:/git/daotie/daotie/target/classes/
		URL xmlpath = this.getClass().getClassLoader().getResource("");
		System.out.println(xmlpath);


		// 第四种： D:\git\daotie\daotie
		System.out.println(System.getProperty("user.dir"));
		/*
		 * 结果： C:\Documents and Settings\Administrator\workspace\projectName
		 * 获取当前工程路径
		 */

		// 第五种：  获取所有的类路径 包括jar包的路径
		System.out.println(System.getProperty("java.class.path"));
	}

}
