## docx_xlsx_utils是什么?

docx_xlsx_utils 是一个通过java POJO与模板文件生成word（仅.docx）或者excel（仅.xlsx）文件的工具封装示例项目。

此项目是基于 [poi-tl](https://github.com/Sayi/poi-tl) 和 [easyexcel](https://github.com/alibaba/easyexcel) 二次开发所成。

## 应用场景

- spring web 项目中，需要通过后端生成word或excel供给客户端下载。

  Ps: 你完全可以自己修改源码，来满足你的项目需求。

## Maven

```xml
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

## 使用示例

点开 `StartController` , 你会发现如下示例代码：

- word 生成示例

```java
DocxUtils.genWord(getClassWordVO(), "/static/office-template/class.docx",
        "班级名单class-member-list.docx", response);
```

- excel 生成示例

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

