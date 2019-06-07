package com.suite.lib.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

class ExcelCellStyle implements CellStyle {
	
	CellStyle cellStyle;
	
	ExcelCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
	
	/**
	 * 操作单元格样式 - 字体
	 * @param wb
	 * @param font
	 * @return
	 */
	protected boolean setFont(Workbook wb, Font font) {
		return setFont(wb, cellStyle, font);
	}
	
	/**
	 * 操作单元格样式 - 字体
	 * @param wb
	 * @param cellStyle
	 * @param font
	 * @return
	 */
	protected static boolean setFont(Workbook wb, CellStyle cellStyle, Font font) {
		try {
			if (wb instanceof HSSFWorkbook) {
				((HSSFCellStyle) cellStyle).setFont(font);
				return true;
			} else if (wb instanceof XSSFWorkbook) {
				((XSSFCellStyle) cellStyle).setFont(font);
				return true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return false;
	}

	@Override
	public short getIndex() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setDataFormat(short fmt) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getDataFormat() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public String getDataFormatString() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setFont(Font font) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFontIndex() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public int getFontIndexAsInt() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setHidden(boolean hidden) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getHidden() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void setLocked(boolean locked) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getLocked() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void setQuotePrefixed(boolean quotePrefix) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getQuotePrefixed() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void setAlignment(HorizontalAlignment align) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public HorizontalAlignment getAlignment() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public HorizontalAlignment getAlignmentEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setWrapText(boolean wrapped) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getWrapText() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment align) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public VerticalAlignment getVerticalAlignmentEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setRotation(short rotation) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getRotation() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setIndention(short indent) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getIndention() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setBorderLeft(BorderStyle border) {
		cellStyle.setBorderLeft(border);
	}

	@Override
	public BorderStyle getBorderLeft() {
		return cellStyle.getBorderLeft();
	}

	@Override
	public BorderStyle getBorderLeftEnum() {
		return cellStyle.getBorderLeftEnum();
	}

	@Override
	public void setBorderRight(BorderStyle border) {
		cellStyle.setBorderRight(border);
	}

	@Override
	public BorderStyle getBorderRight() {
		return cellStyle.getBorderRight();
	}

	@Override
	public BorderStyle getBorderRightEnum() {
		return cellStyle.getBorderRightEnum();
	}

	@Override
	public void setBorderTop(BorderStyle border) {
		cellStyle.setBorderTop(border);
	}

	@Override
	public BorderStyle getBorderTop() {
		return cellStyle.getBorderTop();
	}

	@Override
	public BorderStyle getBorderTopEnum() {
		return cellStyle.getBorderTopEnum();
	}

	@Override
	public void setBorderBottom(BorderStyle border) {
		cellStyle.setBorderBottom(border);
	}

	@Override
	public BorderStyle getBorderBottom() {
		return cellStyle.getBorderBottom();
	}

	@Override
	public BorderStyle getBorderBottomEnum() {
		return cellStyle.getBorderBottomEnum();
	}

	@Override
	public void setLeftBorderColor(short color) {
		cellStyle.setLeftBorderColor(color);
	}

	@Override
	public short getLeftBorderColor() {
		return cellStyle.getLeftBorderColor();
	}

	@Override
	public void setRightBorderColor(short color) {
		cellStyle.setRightBorderColor(color);
	}

	@Override
	public short getRightBorderColor() {
		return cellStyle.getRightBorderColor();
	}

	@Override
	public void setTopBorderColor(short color) {
		cellStyle.setTopBorderColor(color);
	}

	@Override
	public short getTopBorderColor() {
		return cellStyle.getTopBorderColor();
	}

	@Override
	public void setBottomBorderColor(short color) {
		cellStyle.setBottomBorderColor(color);
	}

	@Override
	public short getBottomBorderColor() {
		return cellStyle.getBottomBorderColor();
	}

	@Override
	public void setFillPattern(FillPatternType fp) {
		cellStyle.setFillPattern(fp);
	}

	@Override
	public FillPatternType getFillPattern() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public FillPatternType getFillPatternEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setFillBackgroundColor(short bg) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFillBackgroundColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public Color getFillBackgroundColorColor() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setFillForegroundColor(short bg) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFillForegroundColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public Color getFillForegroundColorColor() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void cloneStyleFrom(CellStyle source) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setShrinkToFit(boolean shrinkToFit) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getShrinkToFit() {
		// TODO Auto-generated method stub
		return false;
	}
	
}

class ExcelFont implements Font {
	Font font;
	
	ExcelFont(Font font) {
		this.font = font;
	}
	
	/**
	 * 操作Excel字体样式 - 是否粗体
	 * @param wb
	 * @param bold
	 * @return
	 */
	protected boolean setBold(Workbook wb, boolean bold) {
		return setBold(wb, font, bold);
	}
	
	/**
	 * 操作Excel字体样式 - 是否粗体
	 * @param wb
	 * @param font
	 * @param bold
	 * @return
	 */
	protected static boolean setBold(Workbook wb, Font font, boolean bold) {
		try {
			if (wb instanceof HSSFWorkbook) {
				((HSSFFont) font).setBold(bold);
				return true;
			} else if (wb instanceof XSSFWorkbook) {
				((XSSFFont) font).setBold(bold);
				return true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return false;
	}

	@Override
	public void setFontName(String name) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public String getFontName() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setFontHeight(short height) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setFontHeightInPoints(short height) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFontHeight() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public short getFontHeightInPoints() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setItalic(boolean italic) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getItalic() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void setStrikeout(boolean strikeout) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getStrikeout() {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void setColor(short color) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setTypeOffset(short offset) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getTypeOffset() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setUnderline(byte underline) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public byte getUnderline() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public int getCharSet() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setCharSet(byte charset) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setCharSet(int charset) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getIndex() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public int getIndexAsInt() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setBold(boolean bold) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public boolean getBold() {
		// TODO Auto-generated method stub
		return false;
	}
	
}

/**
 * 使用POI处理Excel文件
 * @description 
 * 	主要有四个属性:Workbook(工作表),Sheet(表单),Row(行), Cell(单元格).
 * @author Suite
 *
 */
public class ExcelUtil {
	//===== 读取Excel基础操作.思路是按照Workbook,Sheet,Row,Cell一层一层往下读取. =====//
	/**
	 * 初始化Workbook
	 * @param filePath
	 * @return
	 */
	protected static Workbook getReadWorkBookType(String filePath) {
		return getReadWorkBookType(filePath, new File(filePath));
	}
	
	/**
	 * 初始化Workbook
	 * @param filePath
	 * @return
	 */
	protected static Workbook getReadWorkBookType(String filePath, File file) {
        // xls-2003, xlsx-2007
        FileInputStream fis = null;
        try {
        	fis = new FileInputStream(file);
            if (filePath.toLowerCase().endsWith("xlsx")) {
                return new XSSFWorkbook(fis);
            } else if (filePath.toLowerCase().endsWith("xls")) {
                return new HSSFWorkbook(fis);
            } else {
              	// 抛出自定义的业务异常
              	/*throw OnlinePayErrorCode.EXCEL_ANALYZE_ERROR.convertToException("excel格式文件错误");*/
            	return null;
            }
        } catch (IOException e) {
        	//  抛出自定义的业务异常
            /*throw OnlinePayErrorCode.EXCEL_ANALYZE_ERROR.convertToException(e.getMessage());*/
        	e.printStackTrace();
        } finally {
            if (null != fis) {
            	IOUtils.closeQuietly(fis);
            }
        }
        return null;
	}
	
	/**
	 * 获取单元格的内容
	 * @description 
	 * 	· 如果excel中的数据是数字,会发现java中对应的变成了科学计数法的数据;
	 * 	· 所以在获取值的时候就要做一些特殊处理,这样就能保证获取的值是我想要的值.
	 * @param cell
	 * @return
	 */
	protected static String getCellStringVal(Cell cell) {
		if (null != cell) {
			cell.setCellType(CellType.STRING);
	        CellType cellType = cell.getCellType();
	        switch (cellType) {
	            case NUMERIC:
	                return cell.getStringCellValue();
	            case STRING:
	                return cell.getStringCellValue();
	            case BOOLEAN:
	                return String.valueOf(cell.getBooleanCellValue());
	            case FORMULA:
	                return cell.getCellFormula().trim();
	            case BLANK:
	                return "";
	            case ERROR:
	                return String.valueOf(cell.getErrorCellValue());
	            default:
	                return "";
	        }
		}
		return null;
    }

	/**
	 * 转换为单元格的内容对应的格式
	 * @param data
	 * @return
	 */
	protected static String toCellStringVal(Object data) {
		if (null != data) {
			// 基本类型处理
			if (data instanceof Boolean) {
				return String.valueOf(data);
			} else if (data instanceof Character) {
				return String.valueOf(data);
			} else if (data instanceof Double) {
				return String.valueOf(data);
			} else if (data instanceof Float) {
				return String.valueOf(data);
			} else if (data instanceof Integer) {
				return String.valueOf(data);
			} else if (data instanceof Long) {
				return String.valueOf(data);
			} else if (data instanceof Short) {
				return String.valueOf(data);
			} else if (data instanceof String) {
				return String.valueOf(data);
			}
			// 非基本类型处理
			if (data instanceof java.math.BigDecimal) {
				return ((java.math.BigDecimal) data).toPlainString();
			}
			if (data instanceof java.util.Date) {
				return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(data);
			}
		}
		return null;
	}

	/**
	 *	生成Excel基础操作
	 *	@description 主要包括:(1)创建工作簿Workbook,(2)创建工作表Sheet,(3)创建行对象Row,(4)创建单元格对象Cell,</br>
	 *	(5)设置单元格的内容,(6)设置单元格的样式,(7)输出,关闭流对象.
	 */
	//===== 生成Excel基础操作 =====//
	/**
	 * 创建工作簿Workbook
	 * @description xls - Version 2003, xlsx - Version 2007 or heigher
	 * @param version = 指定版本 {"2003" : 2003版本, other param : 2007及以上版本}
	 * @return
	 */
	protected static Workbook createWorkbook(String version) {
		Workbook workbook;
		if ("2003".equals(version)) {
			// xls - 2003
			workbook = new HSSFWorkbook();
		} else {
			// xlsx - 2007及以上版本
			workbook = new XSSFWorkbook();
		}
		return workbook;
	}

	/**
	 * 创建工作簿Workbook
	 * @description xls - Version 2003
	 * @return
	 */
	protected static Workbook createWorkbook() {
		return createWorkbook("2003");
	}

	/**
	 * 创建工作表Sheet
	 * @param wb - 工作簿对象
	 * @return
	 */
	protected static Sheet createSheet(Workbook wb) {
		return wb.createSheet();
	}

	/**
	 * 创建工作表Sheet
	 * @param wb - 工作簿对象
	 * @param sheetName - 工作表名称
	 * @return
	 */
	protected static Sheet createSheet(Workbook wb, String sheetName) {
		return wb.createSheet(sheetName);
	}

	/**
	 * 创建行对象Row
	 * @param sheet - 工作表对象
	 * @param rowNo - 行下标，从0开始
	 * @return
	 */
	protected static Row createRow(Sheet sheet, int rowNo) {
		return sheet.createRow(rowNo);
	}

	/**
	 * 创建单元格对象Cell
	 * @param row - 行对象
	 * @param cellNo - 单元格下标，从0开始
	 * @return
	 */
	protected static Cell createCell(Row row, int cellNo) {
		return row.createCell(cellNo);
	}
	
	/**
	 * 将工作簿输出到流对象，最后关闭流对象
	 * @param workbook - 工作簿对象
	 * @param os - 输出流对象
	 * @return
	 */
	protected static OutputStream outputWorkbook(Workbook workbook, OutputStream os) {
		try {
			workbook.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e1) {
			e1.printStackTrace();
		} finally {
			if (null != os) {
				IOUtils.closeQuietly(os);
			}
		}
		return os;
	}
	
	//===== 工具类的基本实现 =====//
	/**
	 * 解析Excel文件
	 * @param sourceFilePath
	 * @return
	 */
	public static Map<String, List<Map<String, Object>>> readExcel(String sourceFilePath) {
		Workbook workbook = getReadWorkBookType(sourceFilePath);
        return readExcel(workbook);
	}
	
	/**
	 * 解析Excel文件
	 * @param fileName
	 * @param file
	 * @return
	 */
	public static Map<String, List<Map<String, Object>>> readExcel(String fileName, File file) {
        Workbook workbook = getReadWorkBookType(fileName, file);
        return readExcel(workbook);
	}
	
	/**
	 * 解析Excel文件
	 * @description 
	 * 	· 把excel文件里每个Sheet页的数据解析后放入一个List<Map<String, Object>>>中.
	 * @param workbook
	 * @return - <b>sheetMap</b>
	 */
	protected static Map<String, List<Map<String, Object>>> readExcel(Workbook workbook) {
		Map<String, List<Map<String, Object>>> result = new HashMap<String, List<Map<String, Object>>>();
        try {
            for (Iterator<Sheet> iter = workbook.sheetIterator(); iter.hasNext();) {
            	//--- 遍历Sheet
            	Sheet sheet = iter.next();
            	// 装配当前Sheet的数据
            	String sheetName = sheet.getSheetName();
            	List<Map<String, Object>> sheetContent = readExcel_Sheet(sheet);
            	result.put(sheetName, sheetContent);
            }
        } catch (Exception e) {
        	e.printStackTrace();
        } finally {
            if (null != workbook) {
            	IOUtils.closeQuietly(workbook);
            }
        }
        return result;
    }
	
	/**
	 * 解析Excel文件的Sheet页内容
	 * @param sheet - 工作表对象
	 * @return 
	 */
	protected static List<Map<String, Object>> readExcel_Sheet(Sheet sheet) {
		List<Map<String, Object>> result = new ArrayList<>();
		try {// 读取当前Sheet的数据
			sheet.forEach((row) -> {
				result.add(readExcel_Row(row));
			});
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	/**
	 * 解析Excel文件的Row对象
	 * @param row - 工作表对象
	 * @return 
	 */
	protected static Map<String, Object> readExcel_Row(Row row) {
		Map<String, Object> result = new LinkedHashMap<>();
		try {// 读取当前Row的数据
			row.forEach((cell) -> {
				Map<String, Object> cellInfo = readExcel_Cell(cell);
				result.put((String) cellInfo.get("name"), cellInfo.get("value"));
			});
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	/**
	 * 解析Excel文件的Row对象
	 * @param row - 工作表对象
	 * @return 
	 */
	protected static Map<String, Object> readExcel_Cell(Cell cell) {
		Map<String, Object> result = new LinkedHashMap<>();
		try {//----- 读取当前Row的数据
			// 装配单元格对应的表头数据
			String cellName;
			Cell firstRowCell = cell.getRow().getSheet().getRow(0).getCell(cell.getColumnIndex());
			if (null == firstRowCell 
					|| null == (cellName = getCellStringVal(firstRowCell)) || "".equals(cellName.trim())) {
				cellName = "column" + (cell.getColumnIndex() + 1);
			}
			// 装配单元格数据
			result.put("name", cellName);
			result.put("value", getCellStringVal(cell));
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	/**
	 * 创建Excel文件 - xls格式
	 * @param sheetMap - 业务数据
	 * @param os - 输出流
	 * @return
	 */
	public static boolean createExcelXls(Map<String, List<Map<String, Object>>> sheetMap, OutputStream os) {
		return createExcel(sheetMap, os, null);
	}
	
	/**
	 * 创建Excel文件 - xlsx格式
	 * @param sheetMap - 业务数据
	 * @param os - 输出流
	 * @return
	 */
	public static boolean createExcelXlsx(Map<String, List<Map<String, Object>>> sheetMap, OutputStream os) {
		Map<String, String> params = new HashMap<>();
		params.put("version", "2007");
		return createExcel(sheetMap, os, params);
	}
	
	/**
	 * 创建Excel文件
	 * @param sheetMap - 业务数据
	 * @param os - 输出流
	 * @param params - 拓展参数集合
	 * @return
	 */
	public static boolean createExcel(Map<String, List<Map<String, Object>>> sheetMap, OutputStream os, Map<String, String> params) {
		String version = null;
		if (null != params) {
			version = params.get("version");
		}
		version = (null == version) ? "2003" : version;
		if (null != sheetMap && null != os) {
			try {
				Workbook workbook = createWorkbook(version);
				int[] sheetNo = new int[]{1};
				sheetMap.forEach((sheetName, sheetContent) -> {
					//--- 遍历创建Sheet
					Sheet sheet = createSheet(workbook, (!"".equals(sheetName.trim()) ? sheetName : "Sheet" + sheetNo[0]));
					final int[] rowNo = new int[]{0};
					sheetContent.forEach((rowMap) -> {
						//--- 遍历创建Row
						Row row = createRow(sheet, rowNo[0]);
						final int[] cellNo = new int[]{0};
						rowMap.forEach((columnName, cellValue) -> {
							//--- 遍历创建Cell
							Cell cell = createCell(row, cellNo[0]);
							//=== 设置单元格样式 ===//
							ExcelCellStyle cs = new ExcelCellStyle(cell.getCellStyle());
							if (0 == rowNo[0]) {
								// 表头样式
								ExcelFont font = new ExcelFont(workbook.createFont());
								font.setBold(true);
								cs.setFont(workbook, font.font);
							}
							/*cs.setBorderTop(BorderStyle.THIN);
							cs.setBorderRight(BorderStyle.THIN);
							cs.setBorderBottom(BorderStyle.THIN);
							cs.setBorderLeft(BorderStyle.THIN);
							cell.setCellStyle(cs.cellStyle);*/
							//在边框周围设置边框样式。
						    cs.setBorderBottom(BorderStyle.THIN);
						    cs.setBottomBorderColor(IndexedColors.BLACK.getIndex());
						    cs.setBorderLeft(BorderStyle.THIN);
						    cs.setLeftBorderColor(IndexedColors.GREEN.getIndex());
						    cs.setBorderRight(BorderStyle.THIN);
						    cs.setRightBorderColor(IndexedColors.BLUE.getIndex());
						    cs.setBorderTop(BorderStyle.MEDIUM_DASHED);
						    cs.setTopBorderColor(IndexedColors.BLACK.getIndex());
						    cell.setCellStyle(cs.cellStyle);
							//======//
							cell.setCellValue(toCellStringVal(cellValue));
							cellNo[0]++;
						});
						rowNo[0]++;
					});
					sheetNo[0]++;
				});
				os = outputWorkbook(workbook, os);
			} catch (Exception e) {
				e.printStackTrace();
				return false;
			}
		}
		return true;
	}
	
}
