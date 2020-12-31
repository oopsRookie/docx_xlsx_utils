package com.oopsRookie.model.vo;

import java.util.List;

public class ClassWordVO {
    private String name;
    private String number;
    private List<GroupVO> groupList;

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

    public List<GroupVO> getGroupList() {
        return groupList;
    }

    public void setGroupList(List<GroupVO> groupList) {
        this.groupList = groupList;
    }

    @Override
    public String toString() {
        return "ClassVO{" +
                "name='" + name + '\'' +
                ", number='" + number + '\'' +
                ", groupList=" + groupList +
                '}';
    }
}
