package com.oopsRookie.model.vo;

import java.util.List;

public class GroupVO {
    private String name;
    private String number;
    private List<StudentVO> studentList;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public List<StudentVO> getStudentList() {
        return studentList;
    }

    public void setStudentList(List<StudentVO> studentList) {
        this.studentList = studentList;
    }

    @Override
    public String toString() {
        return "GroupVO{" +
                "name='" + name + '\'' +
                ", number='" + number + '\'' +
                ", studentList=" + studentList +
                '}';
    }
}
