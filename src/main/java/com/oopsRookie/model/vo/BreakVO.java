package com.oopsRookie.model.vo;

import java.util.ArrayList;
import java.util.List;

public class BreakVO {
    private List<BreakItemVO> list = new ArrayList<>();

    public BreakVO(){}
    public BreakVO(List<BreakItemVO> list) {
        this.list = list;
    }

    public List<BreakItemVO> getList() {
        return list;
    }

    public void setList(List<BreakItemVO> list) {
        this.list = list;
    }

    public void addItem(BreakItemVO item){
        list.add(item);
    }
}
