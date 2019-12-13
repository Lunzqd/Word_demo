package com.example.utils;


import com.example.utils.CommonUtil;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Component;

import java.io.*;
import java.math.BigInteger;
import java.util.*;

@Component
public class NewCfUtils {
    /**
     * 根据模板生成word文档
     *
     * @param inputUrl 模板路径
     * @param textMap  需要替换的文本内容
     * @param dataList 需要动态生成的内容 --表格对象
     * @return
     */
    public static CustomXWPFDocument changWord(String inputUrl, Map<String, Object> textMap, List<Map<String, Object>> dataList) {
        CustomXWPFDocument document = null;
        try {
            //获取docx解析对象
            document = new CustomXWPFDocument(POIXMLDocument.openPackage(inputUrl));

            //解析替换文本段落对象
            NewCfUtils.changeText(document, textMap);

            //解析替换表格对象
            NewCfUtils.changeTable(document, textMap, dataList);


        } catch (IOException e) {
            e.printStackTrace();
        }
        return document;
    }

    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    private static void changeText(CustomXWPFDocument document, Map<String, Object> textMap) {
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    Object ob = changeValue(run.toString(), textMap);
                    System.out.println("段落：" + run.toString());
                    if (ob instanceof String) {
                        run.setText((String) ob, 0);
                    }
                }
            }
        }
    }

    /**
     * 替换表格对象方法
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     * @param dataList 需要动态生成的内容 --表格信息
     */
    private static void changeTable(CustomXWPFDocument document, Map<String, Object> textMap, List<Map<String, Object>> dataList) {

        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        //循环所有需要进行替换的文本，进行替换
        for (XWPFTable table : tables) {
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                //遍历表格,并替换模板
                eachTable(document, rows, textMap);
            }
        }

        Integer index = 0;
        Integer columns = 0;
        Integer fontSize = 10;
        Integer startLine = 1;
        List<String> singleDataList = new ArrayList<>();
        List<String[]> batchDataList = new ArrayList<>();


        if (!CommonUtil.isObjectNull(dataList) && dataList.size() > 0) {

            for (Map<String, Object> itemData : dataList) {
                index = (Integer) itemData.get("index");
                columns = (Integer) itemData.get("columns");
                fontSize = (Integer) itemData.get("fontSize");
                startLine = (Integer) itemData.get("startLine");
                if (columns == 1) {
                    // 单列表格
                    singleDataList = (List<String>) itemData.get("data");
                    XWPFTable table = tables.get(index);
                    if (null != singleDataList && 0 < singleDataList.size()) {
                        insertTable(document, table, singleDataList, null, 1, fontSize, startLine);
                    } else {
                        deleteTable(table);
                    }

                } else {
                    // 多列表格
                    batchDataList = (List<String[]>) itemData.get("data");
                    XWPFTable table = tables.get(index);
                    if (null != batchDataList && 0 < batchDataList.size()) {
                        insertTable(document, table, null, batchDataList, 2, fontSize, startLine);
//                    List<Integer[]> indexList = startEnd(nDataList);
//                    for (int c=0;c<indexList.size();c++){
//                        //合并行
//                        mergeCellVertically(table,0,indexList.get(c)[0]+1,indexList.get(c)[1]+1);
//                    }
                    } else {
                        deleteTable(table);
                    }
                }
            }
        }

        System.out.println("替换表格对象完成");

    }

    /**
     * 遍历表格
     *
     * @param rows    表格行对象
     * @param textMap 需要替换的信息集合
     */
    private static void eachTable(CustomXWPFDocument document, List<XWPFTableRow> rows, Map<String, Object> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            Object ob = changeValue(run.toString(), textMap);
                            if (ob instanceof String) {
                                run.setText((String) ob, 0);
                            } else if (ob instanceof Map) {
                                run.setText("", 0);

                                Map pic = (Map) ob;
                                int width = Integer.parseInt(pic.get("width").toString());
                                int height = Integer.parseInt(pic.get("height").toString());
                                int picType = getPictureType(pic.get("type").toString());
                                byte[] byteArray = (byte[]) pic.get("content");
                                ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                                try {
                                    int ind = document.addPicture(byteInputStream, picType);
                                    document.renderPicture(ind, width, height, paragraph, run);
                                } catch (Exception e) {
                                    e.printStackTrace();
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 为表格插入数据，行数不够添加新行
     *
     * @param table           需要插入数据的表格
     * @param singleTableData 单列表格的插入数据
     * @param batchTableData  多列表格的插入数据
     * @param type            表格类型：1-单列表格 2-多列表格
     * @param startLine       填入数据开始行
     */
    private static void insertTable(CustomXWPFDocument document, XWPFTable table, List<String> singleTableData, List<String[]> batchTableData, Integer type, Integer fontSize, Integer startLine) {

        // 设置表格宽度
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        tblPr.getTblW().setType(STTblWidth.DXA);
        tblPr.getTblW().setW(new BigInteger("10580"));
        List<XWPFTableRow> rows = table.getRows();

        for (XWPFTableRow row : rows) {
            CTTrPr trPr = row.getCtRow().addNewTrPr();
            CTHeight ht = trPr.addNewTrHeight();
            ht.setVal(BigInteger.valueOf(360));
            row.setHeight(360);
        }

        if (1 == type) {
            // 单列表格

            for (int i = 0; i < singleTableData.size(); i++) {
                XWPFTableRow row = null;

                if (startLine == 1 && i == 0) {
                    row = table.getRow(0);
                } else {
                    row = table.createRow();
                }
                // 创建列
                List<XWPFTableCell> cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    XWPFParagraph xwpfParagraph;
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    if (!CommonUtil.isObjectNull(paragraphs) && paragraphs.size() > 0) {
                        xwpfParagraph = paragraphs.get(0);
                    } else {
                        xwpfParagraph = cell.addParagraph();
                    }
                    // 处理图片
                    XWPFRun xwpfRun = xwpfParagraph.createRun();
                    xwpfRun.setFontSize(fontSize);
                    xwpfRun.setBold(false);
                    if (singleTableData.get(i).contains(".jpg") || singleTableData.get(i).contains(".png")) {
                        try {
                            String[] split = singleTableData.get(i).split("--");
                            byte[] byteArray = NewCfUtils.inputStream2ByteArray(new FileInputStream(split[1]), true);
                            ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                            int picType = getPictureType("png");
                            int ind = document.addPicture(byteInputStream, picType);
                            if (!"0".equals(split[2])) {

                                if ("1".equals(split[2])) {
                                    // 正圆
                                    document.renderPicture(ind, 100, 100, xwpfParagraph, xwpfRun);
                                } else if ("2".equals(split[2])) {
                                    // 椭圆
                                    document.renderPicture(ind, 160, 100, xwpfParagraph, xwpfRun);
                                } else if ("3".equals(split[2])) {
                                    // 顶部logo
                                    document.renderPicture(ind, 250, 30, xwpfParagraph, xwpfRun);
                                }

                            } else {
                                break;
                            }

                            xwpfRun.setText(split[0]);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    } else {
                        xwpfRun.setText(singleTableData.get(i));
                    }

                }


            }

        } else if (2 == type) {
            // 多列表格

            //创建行和创建需要的列
            for (int i = 1; i < batchTableData.size(); i++) {
                //添加一个新行
                XWPFTableRow row = table.insertNewTableRow(1);
                for (int k = 0; k < batchTableData.get(0).length; k++) {
                    row.createCell();//根据String数组第一条数据的长度动态创建列
                }
            }

            // 插入数据
            for (int i = 0; i < batchTableData.size(); i++) {
                List<XWPFTableCell> cells = table.getRow(i).getTableCells();
                for (int j = 0; j < cells.size(); j++) {
                    XWPFTableCell cell = cells.get(j);
//                    cell.setText(batchTableData.get(i)[j]);
                    // 设置单元格宽度
                    CTTcPr tcpr = cell.getCTTc().addNewTcPr();
                    CTTblWidth cellw = tcpr.addNewTcW();
                    cellw.setType(STTblWidth.DXA);
                    cellw.setW(BigInteger.valueOf(360 * 5));

                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    if (!CommonUtil.isObjectNull(paragraphs) && paragraphs.size() > 0) {
                        XWPFParagraph xwpfParagraph = paragraphs.get(0);
                        XWPFRun xwpfRun = xwpfParagraph.createRun();
                        // 处理图片
                        if (batchTableData.get(i)[j].contains(".jpg") || batchTableData.get(i)[j].contains(".png")) {
                            try {
                                String[] split = batchTableData.get(i)[j].split("--");
                                xwpfRun.setText(split[0]);
                                byte[] byteArray = NewCfUtils.inputStream2ByteArray(new FileInputStream(split[1]), true);
                                ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                                int picType = getPictureType("png");
                                int ind = document.addPicture(byteInputStream, picType);
                                if (!"0".equals(split[2])) {
                                    if ("1".equals(split[2])) {
                                        // 正圆
                                        document.renderPicture(ind, 150, 150, xwpfParagraph, xwpfRun);
                                    } else {
                                        // 椭圆
                                        document.renderPicture(ind, 400, 200, xwpfParagraph, xwpfRun);
                                    }
                                } else {
                                    break;
                                }

                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        } else {

                            xwpfRun.setText(batchTableData.get(i)[j]);
                        }

                        xwpfRun.setFontSize(fontSize);
                        xwpfRun.setBold(false);
                        //垂直居中
//                    xwpfParagraph.setVerticalAlignment(TextAlignment.CENTER);
                        //水平居中
//                    xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                        //换行
//                        xwpfParagraph.setWordWrap(true);
                    } else {
                        XWPFParagraph xwpfParagraph = cell.addParagraph();
                        XWPFRun xwpfRun = xwpfParagraph.createRun();
                        xwpfRun.setText(batchTableData.get(i)[j]);
                        xwpfRun.setFontSize(fontSize);
                        xwpfRun.setBold(false); // 是否粗体
                        //垂直居中
//                    xwpfParagraph.setVerticalAlignment(TextAlignment.CENTER);
                        //水平居中
//                    xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                        //换行
//                        xwpfParagraph.setWordWrap(true);
                    }


                }
            }
        }
    }

    /**
     * 判断文本中时候包含$
     *
     * @param text 文本
     * @return 包含返回true, 不包含返回false
     */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.contains("$")) {
            check = true;
        }
        return check;
    }

    /**
     * 匹配传入信息集合与模板
     *
     * @param value   模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static Object changeValue(String value, Map<String, Object> textMap) {
        Set<Map.Entry<String, Object>> textSets = textMap.entrySet();
        Object valu = "";
        for (Map.Entry<String, Object> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if (value.contains(key)) {
                valu = textSet.getValue();
            }
        }
        return valu;
    }

    /**
     * 将输入流中的数据写入字节数组
     *
     * @param in
     * @return
     */
    public static byte[] inputStream2ByteArray(InputStream in, boolean isClose) {
        byte[] byteArray = null;
        try {
            int total = in.available();
            byteArray = new byte[total];
            in.read(byteArray);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (isClose) {
                try {
                    in.close();
                } catch (Exception e2) {
                    System.out.println("关闭流失败");
                }
            }
        }
        return byteArray;
    }

    /**
     * 根据图片类型，取得对应的图片类型代码
     *
     * @param picType
     * @return int
     */
    private static int getPictureType(String picType) {
        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
        if (picType != null) {
            if (picType.equalsIgnoreCase("png")) {
                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {
                res = CustomXWPFDocument.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {
                res = CustomXWPFDocument.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {
                res = CustomXWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }

    /**
     * 合并行
     *
     * @param table
     * @param col     需要合并的列
     * @param fromRow 开始行
     * @param toRow   结束行
     */
    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if (rowIndex == fromRow) {
                vmerge.setVal(STMerge.RESTART);
            } else {
                vmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setVMerge(vmerge);
            } else {
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setVMerge(vmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }

    /**
     * 获取需要合并单元格的下标
     *
     * @return
     */
    public static List<Integer[]> startEnd(List<String[]> daList) {
        List<Integer[]> indexList = new ArrayList<Integer[]>();

        List<String> list = new ArrayList<String>();
        for (int i = 0; i < daList.size(); i++) {
            list.add(daList.get(i)[0]);
        }
        Map<Object, Integer> tm = new HashMap<Object, Integer>();
        for (int i = 0; i < daList.size(); i++) {
            if (!tm.containsKey(daList.get(i)[0])) {
                tm.put(daList.get(i)[0], 1);
            } else {
                int count = tm.get(daList.get(i)[0]) + 1;
                tm.put(daList.get(i)[0], count);
            }
        }
        for (Map.Entry<Object, Integer> entry : tm.entrySet()) {
            String key = entry.getKey().toString();
            String value = entry.getValue().toString();
            if (list.indexOf(key) != (-1)) {
                Integer[] index = new Integer[2];
                index[0] = list.indexOf(key);
                index[1] = list.lastIndexOf(key);
                indexList.add(index);
            }
        }
        return indexList;
    }

    /**
     * 删除表格
     *
     * @param table 表格对象
     */
    private static void deleteTable(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
//        int rowLength = rows.size();
//        for (int i = 0; i < rowLength; i++) {
//            table.removeRow(0);
//        }
        for (XWPFTableRow currentRow : rows) {
            CTTcBorders tblBorders = currentRow.getCell(0).getCTTc().getTcPr().addNewTcBorders();
            tblBorders.addNewLeft().setVal(STBorder.NIL);
            tblBorders.addNewRight().setVal(STBorder.NIL);
            tblBorders.addNewBottom().setVal(STBorder.NIL);
            tblBorders.addNewTop().setVal(STBorder.NIL);
            //隐藏这一行所有单元格的边框
            for (int i = 0; i < currentRow.getTableCells().size(); i++) {
                currentRow.getCell(i).getCTTc().getTcPr().setTcBorders(tblBorders);
            }
        }


    }

}
