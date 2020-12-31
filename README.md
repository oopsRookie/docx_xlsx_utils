## What is docx_xlsx_utils?

docx_xlsx_utils is a code sample project for utility encapsulation of generating word(.docx only) and excel(.xlsx only) with java POJO.

And this project is based on [poi-tl](https://github.com/Sayi/poi-tl) and [easyexcel](https://github.com/alibaba/easyexcel) .

## Usage scenario

- In Spring Web project, the requirements include backend generate word or excel, and send output stream back to client browser .

  Ps: you could totally modify source code to meet your needs if you like.

## Maven

```
<dependency>
    <groupId>com.deepoove</groupId>
    <artifactId>poi-tl</artifactId>
    <version>1.8.2</version>
</dependency>

<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>easyexcel</artifactId>
    <version>2.2.7</version>
</dependency>
```

## Usage demo

open `StartController` , find these sample code below

- word generation sample

```java
DocxUtils.genWord(getClassWordVO(), "/static/office-template/class.docx",
        "班级名单class-member-list.docx", response);
```

- excel generation sample

```java
new XlsxUtils.ExcelBuilder<List<StudentVO>>(getStudentOfGroupOne(),
        "/static/office-template/class.xlsx",
        "class-member-list",
        0, "group1", response)      //sheetName is not working.ㄟ( ▔, ▔ )ㄏ
        .addKV("className", "className")
        .addKV("classNumber", "classNumber")
        .addKV("groupOneName", "group_1_Name")
        .addKV("groupOneNumber", "group_1_Number")

        .addSheet(1, "group2", getStudentOfGroupTwo())
        .addKV("groupTwoName", "group_2_Name")
        .addKV("groupTwoNumber", "group_2_Number")

        .addKV("extraKey", "extraVal")
        .build().genExcel();
```

