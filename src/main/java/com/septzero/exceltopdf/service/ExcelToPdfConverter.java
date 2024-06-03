package com.septzero.exceltopdf.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

@Slf4j
@Service
public class ExcelToPdfConverter {
    //全局变量，用来设置的
    public static int FONT_SITE = 12;//初始字体大小
    public static float MARGIN = 20;//页边距
    public static int SPACING_CODE = 2;//间隔字符数


    public static void convertExcelToPdf(String excelFilePath, String pdfFilePath) throws IOException {
        //log.debug("开始将Excel转化为PDF---------------------------------");
        Workbook workbook = new XSSFWorkbook(new FileInputStream(new File(excelFilePath)));
        PDDocument document = new PDDocument();
        Sheet sheet = workbook.getSheetAt(0);
        int rowCount = sheet.getLastRowNum();
        int columnCount = sheet.getRow(0).getLastCellNum();
        //前面是准备工作

        //先统计出比例
        int[] ratio = new int[columnCount];
        boolean flagVertical = true;//true为竖向纸，false为横向纸
        int all = countCellProportion(sheet, ratio);//这个方法会计算出一行最多有多少字符
        int stringCount = all + SPACING_CODE * (columnCount - 1);

        //开始算字号和行高
        int fontSize = FONT_SITE;
        int cellHighUse = 240 / FONT_SITE;
        if(stringCount > 1440 / FONT_SITE){
            fontSize = 1440 / stringCount;
            cellHighUse = 2400 / stringCount;
        }
        if(stringCount > 960 / FONT_SITE){//是否需要横向纸张
            flagVertical = false;
        }
        //超长自适应分页也要靠这个方法，不分页了
        //自适应字体大小，大于120了，缩小字体，120以下是12
        //自适应行高，120，120以下是20
        //设置纸的大小(横向还是纵向，flag为true是纵向)
        float pageWidth;
        float pageHeight;
        if(flagVertical){
            pageWidth = PDRectangle.A4.getWidth();
            pageHeight = PDRectangle.A4.getHeight();
        }else {
            pageWidth = PDRectangle.A4.getHeight();
            pageHeight = PDRectangle.A4.getWidth();
        }


        float margin = MARGIN; // 边距
        float tableWidth = pageWidth - 2 * margin;
        float tableHeight = pageHeight - 2 * margin; //或许没用了
        float[] cellWidth = countCellWidth(stringCount, ratio, columnCount, tableWidth);
        float cellHeight = cellHighUse; // 每行高度为120(1440÷字号)字以下，20(2400÷字号)；以上自适应

        int currentPageIndexX = 0;//横向页数
        int currentPageIndexY = 0;//纵向页数

        int rowIndex = 0;
        //确定每一页有多少行
        int pageRowNum = (int) (tableHeight/cellHeight);
        DataFormatter dataFormatter = new DataFormatter();

        //1、先确定打印到哪一页
        //先确定一页有多少行内容，然后除法得到页数，余数得到行数
        //2、if上一页满了（flag计数），创建一页新的，然后add进去
        //3、往当前页，写内容
        //逻辑不对，应该先创建页，取每一行数据
        //打印每一页的内容
        boolean flagDoc = true;//决定是否结束写
        boolean flagPage = true;//决定是否添加一页
        //就是说，如果flag_page变false了，就出循环，然后渠道while大循环，添加一个新页，然后又来循环
        //这个i是用来定位每一页的
        int i = 0;
        PDPageContentStream contentStream = null;
        while (flagDoc){
            flagPage = true;
            //新建一页
            PDPage page = new PDPage(PDRectangle.A4);
            //设置纸张长宽
            PDRectangle rectangle = new PDRectangle(pageWidth,pageHeight);
            page.setMediaBox(rectangle);
            //添加到文档中
            document.addPage(page);
            contentStream = new PDPageContentStream(document, page);
            //这个contentStream是用来写的
            //把后面的内容封装一下，每一页的写入，就是开始坐标到结束左边，一共四个数
            //这里来确定开始坐标和结束坐标，需要两个数，一个是横向的，一个是纵向的
            //因为不做分页了，所以就把结束坐标写成死值了
            //这里调用写入方法(方法不用了)
//            writeCellPage(sheet, contentStream, 0, columnCount, rowIndex, rowIndex+pageRowNum,
//                    margin, fontSize, cellWidth, cellHeight, pageHeight, dataFormatter);
            while (flagPage){
                //这里开始给这一页写内容
                Row row = sheet.getRow(rowIndex);
                int cellCount = row.getLastCellNum();
                float wideStart = margin;
                for (int j = 0; j < cellCount; j++) {
                    //log.debug("开始写入(" + i + ", " + j + ")坐标的数据");
                    Cell cell = row.getCell(j);
                    String cellValue;

                    if (cell.getCellType() == CellType.STRING) {
                        cellValue = cell.getStringCellValue();
                    } else if (cell.getCellType() == CellType.NUMERIC) {
                        cellValue = dataFormatter.formatCellValue(cell);
                    } else {
                        cellValue = ""; // 或者抛出异常等处理方式
                    }

                    contentStream.beginText();
                    contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                    wideStart = wideStart + cellWidth[j];
                    float highStart = pageHeight - margin - (i-1) * cellHeight - 20;
                    //重点：0.0坐标是左下角的，所以是从下往上排的，得用减法，最大y最开始，后面一直减
                    contentStream.newLineAtOffset(wideStart, highStart);
                    try{
                        contentStream.showText(cellValue);
                    }catch (Exception e){
                        //如果出现中文，就重新设置字体，然后再运行
                        log.debug("开始写入中文------------------------------");
                        //Path testPath = Paths.get("").toAbsolutePath();
                        //log.debug("当前(./)目录为--------------->" + testPath);
                        Path ttfPath = Paths.get("./src/main/java/com/septzero/exceltopdf/ttf/simhei.ttf");
                        String ttfRelaPath = ttfPath.toString();
                        log.debug("当前读字体文件(.ttf)的取路径为------->" + ttfPath.toAbsolutePath());
                        File fontFile = new File(ttfRelaPath);
                        PDType0Font font = PDType0Font.load(document, fontFile);
                        contentStream.setFont(font, fontSize);
                        contentStream.showText(cellValue);
                    }

                    contentStream.endText();
                }
                i = (i+1)%pageRowNum;
                rowIndex++;
                if (i == 0){
                    flagPage = false;
                }
                if(rowIndex >= rowCount){
                    //如果写完了，就出去
                    flagDoc = false;
                    flagPage = false;
                }
            }
            contentStream.close();
        }

        document.save(new FileOutputStream(new File(pdfFilePath)));
        document.close();
        workbook.close();
    }

    private static int countCellProportion(Sheet sheet, int[] max) throws UnsupportedEncodingException {
        //计算总计字符数量
        int all = 0;
        int columnCount = sheet.getRow(0).getLastCellNum();
        int rowCount = sheet.getLastRowNum();
        DataFormatter dataFormatter = new DataFormatter();

        for(int j = 0; j < columnCount; j++){
            max[j] = 0;
            for(int i = 0; i <= rowCount; i++){
                Cell cell = sheet.getRow(i).getCell(j);
                String cellValue;

                if (cell.getCellType() == CellType.STRING) {
                    cellValue = cell.getStringCellValue();
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    cellValue = dataFormatter.formatCellValue(cell);
                } else {
                    cellValue = ""; // 或者抛出异常等处理方式
                }
                if(max[j] < cellValue.getBytes("GBK").length){
                    max[j] = cellValue.getBytes("GBK").length;
                }
            }
            all = all + max[j];
        }
        return all;
    }

    private static float[] countCellWidth(int all, int[] max, int columnCount, float tableWidth){
        //计算每一列的行宽
        //先获取有多少列，然后开始循环每一列，找到最长的一个，根据比例，分配格子
        //这段移动到了上面
        float[] out = new float[columnCount];
        //此时的max数组，拥有了最长的字段，按比例分配
        float base = tableWidth / all;
        for(int j = 0; j < columnCount-1; j++){
            out[j+1] = (max[j]+SPACING_CODE) * base;
        }
        //两个column之间，算两个字符，然后all+column-1*2的总和>80，纸张选择横向的，否则选择纵向的

        return out;
    }

    private static void writeCellPage(Sheet sheet, PDPageContentStream contentStream,
                                      int startX, int endX, int startY, int endY, float margin,
                                      int fontSize, float[] cellWidth, float cellHeight, float pageHeight,
                                      DataFormatter dataFormatter) throws IOException {
        //给每一页写入内容，入参：表，写入
        //开始X，结束X，开始Y，结束Y，页边距（坐标为excel的,X是横着的从左到右，Y是竖着的从上到下）
        //字号，每列宽度，每行高度，页面高度，其他
        for (int i = startY; i < endY; i++) {
            //这里开始给这一页写内容
            Row row = sheet.getRow(i);
            float wideStart = margin;
            for (int j = startX; j < endX; j++) {
                Cell cell = row.getCell(j);
                String cellValue;

                if (cell.getCellType() == CellType.STRING) {
                    cellValue = cell.getStringCellValue();
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    cellValue = dataFormatter.formatCellValue(cell);
                } else {
                    cellValue = ""; // 或者抛出异常等处理方式，先整个空的
                }

                contentStream.beginText();
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                wideStart = wideStart + cellWidth[j];
                if(j == startX){
                    wideStart = margin;//如果是才开始的，就给重置一下，
                    //注意：cellWidth[j]是第j-1列的宽度，cellWidth[0]=0
                }
                //重点：0.0坐标是左下角的，所以是从下往上排的，得用减法，最大y最开始，后面一直减
                contentStream.newLineAtOffset(wideStart, pageHeight-margin-i * cellHeight-20);
                contentStream.showText(cellValue);
                contentStream.endText();
            }
        }
    }

    private static int divPage(int all, int columnCount, List<Integer> pageList, int[] ratio){
        //入参：每列的max，也就是ratio，还有all
        //1、先使用all，去判断是否分页，分几页，分页后是横向还是纵向
        //当字符数量大于？？的时候就分页，然后一个取80，一个取120，距离整数最近的，
        //离80近，就选竖着，离120近，就选横着的
        //然后确定是否分页的问题，＞1，除法结果＞1就分页
        //给flagVertical赋值，所以这里的返回值是这个的数字，代表了横着分几页
        float vertical = (float) (all / 80.0);//是否竖向
        float transverse = (float) (all / 120.0);//是否横向
        //先判断分不分页
        if(all <= 80){
            return 1;
        }else if (all <= 160){
            return -1;
        }
        //需要分页了，再来这里
        //使用round()方法四舍五入，然后abs取绝对值
        int verticalInt = Math.round(vertical);
        int transverseInt = Math.round(transverse);
        float verticalAbs = Math.abs(vertical - verticalInt);//竖着的，绝对值
        float transverseAbs = Math.abs(transverse - transverseInt);//横着的，绝对值，去比大小
        boolean xyFlag = true;//默认竖向
        if(verticalAbs >= transverseAbs){
            //纸张竖着
            xyFlag = true;
        }else{
            //纸张横着
            xyFlag = false;
        }
        //然后创建一个链表，往里塞多少列一页
        //为了把数据带出去，数组放外面
        if(xyFlag){
            //竖向的纸
            divPageNum(85, columnCount, ratio, pageList);
            int out = pageList.size();
            return out;
        }else {
            //横向的纸
            divPageNum(130, columnCount, ratio, pageList);
            int out = 0-pageList.size();
            return out;
        }
    }

    public static void divPageNum(int fontNum, int columnCount, int[] ratio, List<Integer> pageList){
        //divPage用的方法，用来分页的
        int div = 0;
        for (int i = 0; i < columnCount; i++) {
            if(div != 0){
                div++;
            }
            div = div + ratio[i];
            if(div > fontNum){
                pageList.add(i);
                div = 0;
            }
        }
    }
}
