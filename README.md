# ExcelToPdf
这是一个将Excel文件转化为PDF文件的项目。 This is a project to convert Excel file to Pdf.

# 语言Language

中文说明在前半段。

The English instructions are in the second half, translated by Translation tools and me.

# 中文

## 项目介绍

```
本项目用于将Excel转化为PDF文件。

由于是给项目写的一个demo，项目中的excel为接口结果生成的，所以结构较为简单，目前本项目不能支持各类格式、空行、字体、公式等
也没有支持多表、隐藏表等
未提供分页方案、最小字号限值

支持功能：
1、自适应列宽
2、自适应纸张方向
3、自适应字号大小（左右一页打完，无限向下延申）

将需要转化的Excel文件放置在excel命名文件夹，运行项目后，即可在output中获取到转化结果，输出的命名格式为：fileName.pdf。

相对路径布局如下：
├── excel                   //放需要转化的文件
├── output                  //输出文件夹
├── ExcelToPdf              //项目

未来更新企划：
把空行、公式、多表、隐藏表给支持了，
后续会提供横向过长的不同方案，如分页、自动换行
然后是一个依赖。

若有遇到什么问题，后续记录更新。
```

## 环境依赖

jdk-17.0.9

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.apache.pdfbox</groupId>
    <artifactId>pdfbox</artifactId>
    <version>2.0.24</version>
</dependency>
```

## 使用说明

1、将需要转化的Excel文件放入excel文件夹中

2、运行项目

3、在output中获取转化后的文件

### 功能说明：

目前暂时只支持常规转化功能。

不支持单文件多Sheet(只转化第一张表)、空行、公式等。

不能横向分页。

## 版本内容更新

### v0.0.97: 

```
1、实现了转化效果.
```

### v0.0.98: 

```
1、实现自适应列宽。
2、实现自适应字。
3、实现自适应页面方向。
4、实现字号大小、页边距、列间距统一调整。
```

### v0.0.99: 

```
1、支持单元格内容为中文。
2、自适应列宽对中文、英文进行适配。
```

### v1.0.0: 

```
1、上传至gitHub。
2、新增了对输出文件夹的校验。
```

# English

## Project Introduction

```
This project is designed to convert Excel files into PDF files.

As this is a demo for the project, the Excel files in the project are generated from API results, so the structure is relatively simple. Currently, the project does not support various formats, empty rows, fonts, formulas, etc., and it also does not support multiple sheets or hidden sheets. There is no pagination solution or minimum font size limit provided.

Supported functions:

Adaptive column width
Adaptive paper orientation
Adaptive font size (complete one page horizontally, extend downwards infinitely)
To use the project, place the Excel files that need to be converted into the "excel" folder, run the project, and you can find the converted results in the "output" folder. The output file naming format is: fileName.pdf.

Relative path layout:
├── excel           // target folder
├── output          // output folder
├── ExcelToPdf      // project folder

Future updates plan:
Adding support for empty rows, formulas, multiple sheets, and hidden sheets,
Will provide different solutions for horizontally long content, such as pagination and automatic line breaks
Then dependency management will be implemented

If there are any problems or suggestions, I will deal with them in the feature.
```

## Environment

jdk-17.0.9

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
<dependency>
    <groupId>org.apache.pdfbox</groupId>
    <artifactId>pdfbox</artifactId>
    <version>2.0.24</version>
</dependency>
```

## Directions for use

1、Place the Excel files to be converted into the "excel" folder.

2、Run the project.

3、The converted files will be put in the "output" folder.

### Functionality:

Currently, only basic conversion functionality is supported.

It does not support multiple sheets in a single file (only converts the first sheet), empty rows, and formulas.

Horizontal pagination is not supported.

## Version

### v0.0.97: 

```
1、Implemented convert function.
```

### v0.0.98: 

```
1、Implemented adaptive column width.
2、Implemented adaptive font size.
3、Implemented adaptive page orientation.
4、Implemented uniform adjustment of font size, page margins, and column spacing.
```

### v0.0.99: 

```
1、Supported cell content in Chinese.
2、Adapted column width to accommodate both Chinese and English content.
```

### v1.0.0: 

```
1、Uploaded to GitHub.
2、Added validation for the output folder.
```

# 