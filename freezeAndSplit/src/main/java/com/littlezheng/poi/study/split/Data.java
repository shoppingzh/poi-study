package com.littlezheng.poi.study.split;

import java.util.ArrayList;
import java.util.List;

public class Data {

    private List<Object> datas = new ArrayList<Object>();
    
    public void add(Object o){
        datas.add(o);
    }
    
    public List<Object> get(){
        return datas;
    }

    @Override
    public String toString() {
        return datas.toString();
    }
    
}
