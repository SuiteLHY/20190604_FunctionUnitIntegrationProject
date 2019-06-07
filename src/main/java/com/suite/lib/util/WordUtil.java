package com.suite.lib.util;

import java.io.OutputStream;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

/**
 * 使用POI处理Word文件
 * @description
 * @author Suite
 *
 */
public class WordUtil {
    //===== WORD基础操作 =====//
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
            if (data instanceof Date) {
                return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(data);
            }
        }
        return null;
    }

    //===== 读取WORD操作的实现 =====//
    //===== 生成WORD操作的实现 =====//
    @SuppressWarnings("resource")
	public static boolean createWordDocx(Map<String, List<Map<String, Object>>> data, OutputStream os) {
        if (null != data && null != os) {
            try {
                XWPFDocument docx = new XWPFDocument();
                // WORD文档页面 - 排版
                CTSectPr docxCtSectPr = docx.getDocument().getBody().isSetSectPr()
                        ? docx.getDocument().getBody().getSectPr()
                        : docx.getDocument().getBody().addNewSectPr();
                CTPageSz docxPgSz = (docxCtSectPr.isSetPgSz()) ? docxCtSectPr.getPgSz() : docxCtSectPr.addNewPgSz();
                // WORD文档页面 - 排版 - 设置横板
                docxPgSz.setW(BigInteger.valueOf(15840));
                docxPgSz.setH(BigInteger.valueOf(11907));
                docxPgSz.setOrient(STPageOrientation.LANDSCAPE);
                // 添加标题
                XWPFParagraph titleParagraph = docx.createParagraph();
                // 标题 - 内容和样式
                titleParagraph.setAlignment(ParagraphAlignment.CENTER);
                titleParagraph.setVerticalAlignment(TextAlignment.CENTER);
                XWPFRun titleParagraphRun = titleParagraph.createRun();
                titleParagraphRun.setText("xxx市xxx外出请假备案报告");
                titleParagraphRun.setColor("000000");
                titleParagraphRun.setFontFamily("宋体 (标题)");
                titleParagraphRun.setFontSize(22);
                titleParagraphRun.setBold(true);
                // 段落
                XWPFParagraph firstParagraph = docx.createParagraph();
                // 段落 - 内容和样式
                firstParagraph.setAlignment(ParagraphAlignment.LEFT);
                firstParagraph.setVerticalAlignment(TextAlignment.CENTER);
                XWPFRun firstParagraphRun = firstParagraph.createRun();
                firstParagraphRun.setText("xxx印制"
                        + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n"
                        + "报告时间：" + new SimpleDateFormat("yyyy年MM月dd日").format(new Date())
                );
                firstParagraphRun.setColor("000000");
                firstParagraphRun.setFontSize(16);
                // 段落 - 样式 - 设置中文字体
                CTRPr firstParaRunRpr = firstParagraphRun.getCTR().isSetRPr() ? firstParagraphRun.getCTR().getRPr() : firstParagraphRun.getCTR().addNewRPr() ;
                CTFonts firstParaRunFonts = firstParaRunRpr.isSetRFonts() ? firstParaRunRpr.getRFonts() : firstParaRunRpr.addNewRFonts();
                firstParaRunFonts.setAscii("仿宋");
                firstParaRunFonts.setEastAsia("仿宋");
                firstParaRunFonts.setHAnsi("仿宋");
                // 基本信息表格
                XWPFTable infoTable = docx.createTable();
                // 基本信息表格 - 列宽自动分割
                CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
                infoTableWidth.setType(STTblWidth.DXA);
                infoTableWidth.setW(BigInteger.valueOf(12069));
                // 基本信息表格 - 行数据
                int[] rowNo = new int[]{0};
                data.forEach((rowName, rowContent) -> {
                    // 创建表格行
                    XWPFTableRow row = (null != infoTable.getRow(rowNo[0])) ? infoTable.getRow(rowNo[0]) : infoTable.createRow();
                    int[] cellNo = new int[]{0};
                    rowContent.forEach((cellInfo) -> {
                        // 创建行 - 单元格
                        XWPFTableCell cell = (null != row.getCell(cellNo[0])) ? row.getCell(cellNo[0]) : row.addNewTableCell();
                        // 单元格样式
                        CTTc cttc = cell.getCTTc();
                        CTTcPr cttcPr = cttc.addNewTcPr();
                        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
                        cttcPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                        // 单元格的段落 - 内容和样式
                        XWPFParagraph cellParagraph = (null != cell.getParagraphArray(0)) ? cell.getParagraphArray(0) : cell.addParagraph();
                        cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                        cellParagraph.setVerticalAlignment(TextAlignment.CENTER);
                        XWPFRun cellParaRun = cellParagraph.createRun();
                        cellParaRun.setText(toCellStringVal(cellInfo.get("value")));
                        cellParaRun.setColor("000000");
                        cellParaRun.setFontSize(14);
                        // 单元格的段落 - 样式 - 设置中文字体
                        CTRPr paraRunRpr = cellParaRun.getCTR().isSetRPr() ? cellParaRun.getCTR().getRPr() : cellParaRun.getCTR().addNewRPr();
                        CTFonts paraRunFonts = paraRunRpr.isSetRFonts() ? paraRunRpr.getRFonts() : paraRunRpr.addNewRFonts();
                        paraRunFonts.setAscii("黑体");
                        paraRunFonts.setEastAsia("黑体");
                        paraRunFonts.setHAnsi("黑体");
                        cellNo[0]++;
                    });
                    rowNo[0]++;
                });
                // WORD文件输出
                docx.write(os);
            } catch (Exception e) {
                e.printStackTrace();
                return false;
            } finally {
                if (null != os) {
                    IOUtils.closeQuietly(os);
                }
            }
        }
        return true;
    }

    // (copy)from [简书]
	/*public void write2Docx() throws Exception {
        XWPFDocument document= new XWPFDocument();
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("G:\\Offer\\create_table.docx"));
        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText("Java PoI");
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);
        //段落
        XWPFParagraph firstParagraph = document.createParagraph();
        XWPFRun run = firstParagraph.createRun();
        run.setText("Java POI 生成word文件。");
        run.setColor("696969");
        run.setFontSize(16);
        //设置段落背景颜色
        CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
        cTShd.setVal(STShd.CLEAR);
        cTShd.setFill("97FFFF");
        //换行
        XWPFParagraph paragraph1 = document.createParagraph();
        XWPFRun paragraphRun1 = paragraph1.createRun();
        paragraphRun1.setText("\r");
        //基本信息表格
        XWPFTable infoTable = document.createTable();
        //去表格边框
        infoTable.getCTTbl().getTblPr().unsetTblBorders();
        //列宽自动分割
        CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
        infoTableWidth.setType(STTblWidth.DXA);
        infoTableWidth.setW(BigInteger.valueOf(9072));
        //表格第一行
        XWPFTableRow infoTableRowOne = infoTable.getRow(0);
        infoTableRowOne.getCell(0).setText("职位");
        infoTableRowOne.addNewTableCell().setText(": Java 开发工程师");
        //表格第二行
        XWPFTableRow infoTableRowTwo = infoTable.createRow();
        infoTableRowTwo.getCell(0).setText("姓名");
        infoTableRowTwo.getCell(1).setText(": seawater");
        //表格第三行
        XWPFTableRow infoTableRowThree = infoTable.createRow();
        infoTableRowThree.getCell(0).setText("生日");
        infoTableRowThree.getCell(1).setText(": xxx-xx-xx");
        //表格第四行
        XWPFTableRow infoTableRowFour = infoTable.createRow();
        infoTableRowFour.getCell(0).setText("性别");
        infoTableRowFour.getCell(1).setText(": 男");
        //表格第五行
        XWPFTableRow infoTableRowFive = infoTable.createRow();
        infoTableRowFive.getCell(0).setText("现居地");
        infoTableRowFive.getCell(1).setText(": xx");
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        //添加页眉
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "ctpHeader";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
        //设置为右对齐
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);
        //添加页脚
        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctFooter = ctrFooter.addNewT();
        String footerText = "ctpFooter";
        ctFooter.setStringValue(footerText);
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
        parsFooter[0] = footerParagraph;
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);
        document.write(out);
        out.close();
    }*/

}