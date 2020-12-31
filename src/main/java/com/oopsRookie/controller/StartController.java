package com.oopsRookie.controller;

import com.oopsRookie.model.vo.ClassExcelVO;
import com.oopsRookie.model.vo.ClassWordVO;
import com.oopsRookie.model.vo.GroupVO;
import com.oopsRookie.model.vo.StudentVO;
import com.oopsRookie.util.excel.XlsxUtils;
import com.oopsRookie.util.word.DocxUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
public class StartController {

    @GetMapping("genWord")
    public void genWord(HttpServletRequest request, HttpServletResponse response)
            throws IOException {

        DocxUtils.genWord(getClassWordVO(), "/static/office-template/class.docx",
                "班级名单class-member-list.docx", response);
    }

    @GetMapping("genExcel")
    public void genExcel(HttpServletRequest request, HttpServletResponse response) throws IOException {
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


    /*-------------------------------------------- mock some data -------------------------------------*/

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
