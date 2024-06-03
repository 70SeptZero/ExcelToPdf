package com.septzero.exceltopdf;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import com.septzero.exceltopdf.service.ExcelToPdfConverter;
import org.slf4j.LoggerFactory;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.IOException;

@SpringBootApplication
public class ExcelToPdfApplication {
    private static ExcelToPdfConverter converter = null;

    public ExcelToPdfApplication(ExcelToPdfConverter converter) {
        this.converter = converter;
    }

    public static void main(String[] args) {
        Logger fontLogger = (Logger) LoggerFactory.getLogger("org.apache.fontbox");
        fontLogger.setLevel(Level.WARN);
        Logger pdfLogger = (Logger) LoggerFactory.getLogger("org.apache.pdfbox");
        pdfLogger.setLevel(Level.WARN);
        String inputFolderPath = "../excel";
        String outputFolderPath = "../output";
        //读取所有文件
        try{
            File folder = new File(inputFolderPath);
            System.out.println("读取路径为：" + folder.toPath().toAbsolutePath());
            File[] inputFiles = folder.listFiles();
            if(inputFiles != null){
                for(File inputFile : inputFiles){
                    String excelFileName = inputFile.getName();
                    System.out.println("正在转化:" + excelFileName);
                    String excelFilePath = inputFolderPath + "/" + excelFileName;
                    //检查输出文件夹是否存在，不存在就创建
                    File outputPath = new File("../output");
                    if (!outputPath.exists()) {
                        outputPath.mkdirs();
                    }
                    //获取文件名(不含后缀)
                    String fileName = excelFileName.split("\\.")[0];
                    String pdfFilePath = outputFolderPath+ "/" + fileName + ".pdf";
                    converter.convertExcelToPdf(excelFilePath, pdfFilePath);
                }
            }
        }catch (IOException e){
            System.out.println("Error occurred while convert excel files: " + e.getMessage());
        }
    }

}
