package com.oopsRookie.controller;

import com.oopsRookie.model.vo.*;
import com.oopsRookie.util.excel.FontImage;
import com.oopsRookie.util.excel.XlsxUtils;
import com.oopsRookie.util.word.DocxUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.support.ResourcePatternResolver;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Controller
public class StartController {

    @GetMapping("genWord")
    public void genWord(HttpServletRequest request, HttpServletResponse response)
            throws IOException {

        DocxUtils.genWord(getClassWordVO(), "/static/office-template/class.docx",
                "班级名单class-member-list", response);
    }

    @GetMapping("genWordBreak")
    public void genWordBreak(HttpServletRequest request, HttpServletResponse response)
            throws IOException {
        DocxUtils.genWord(getWordBreakData(), "/static/office-template/break.docx",
                "break测试", response);
    }

    @GetMapping("genWordWithWaterMark")
    public void genWordWithWaterMark(HttpServletRequest request, HttpServletResponse response)
            throws IOException {

        DocxUtils.genWord(getClassWordVO(), "/static/office-template/class.docx",
                "班级名单class-member-list", "WaterMark水印",response);
    }

    @GetMapping("genExcel")
    public void genExcel(HttpServletRequest request, HttpServletResponse response) throws IOException, InvalidFormatException {

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
    }

    @GetMapping("genExcelWithWaterMark")
    public void genExcelWithWaterMark(HttpServletRequest request, HttpServletResponse response) throws IOException, InvalidFormatException {

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
                .withWaterMark("WaterMark水印")
                .build().genExcel();
    }



    /*-------------------------------------------- mock some data -------------------------------------*/

    private BreakVO getWordBreakData() {
        List<String> list = Arrays.asList("李XX", "李X", "李Xz", "李X", "李X",
                "李XX", "李X", "李Xz", "李X", "李X");
        BreakVO breakVO = new BreakVO();
        for (String str :
                list) {
            breakVO.addItem(new BreakItemVO(str));
        }
        return breakVO;
    }

    private List<StudentVO> getStudentOfGroupOne() {

        //students of group 1
        return new ArrayList<StudentVO>() {{
            add(new StudentVO() {{
                setName("bob");
                setNumber("g101");
            }});
            add(new StudentVO() {{
                setName("cute");
                setNumber("g102");
            }});
        }};
    }

    private List<StudentVO> getStudentOfGroupTwo() {

        //students of group 2
        return new ArrayList<StudentVO>() {{
            add(new StudentVO() {{
                setName("duck");
                setNumber("g201");
            }});
            add(new StudentVO() {{
                setName("elan");
                setNumber("g202");
            }});
        }};
    }

    //mock word data of class
    private ClassWordVO getClassWordVO() {

        //group 1
        GroupVO groupVO1 = new GroupVO() {{
            setName("group 1");
            setNumber("1");
            setStudentList(new ArrayList<StudentVO>() {{
                add(new StudentVO() {{
                    setName("bob");
                    setNumber("0");
                }});
                add(new StudentVO() {{
                    setName("cute");
                    setNumber("1");
                }});
            }});
        }};

        //group 2
        GroupVO groupVO2 = new GroupVO() {{
            setName("group 2");
            setNumber("2");
            setStudentList(new ArrayList<StudentVO>() {{
                add(new StudentVO() {{
                    setName("bob");
                    setNumber("g201");
                }});
                add(new StudentVO() {{
                    setName("club");
                    setNumber("g202");
                }});
            }});
        }};

        //class
        ClassWordVO classWordVO = new ClassWordVO() {{
            setName("class 1");
            setNumber("no.1");
        }};
        classWordVO.setGroupList(new ArrayList<GroupVO>() {{
            add(groupVO1);
            add(groupVO2);
        }});

        return classWordVO;
    }
}
