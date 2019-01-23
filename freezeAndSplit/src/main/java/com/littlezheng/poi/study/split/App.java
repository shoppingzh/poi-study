package com.littlezheng.poi.study.split;

import java.util.List;

public class App {

public static void main(String[] args) {
    Data d = new Data();
    for(int i=1;i<=10;i++){
        d.add(i);
    }
    
    Window v1 = new Window(d, 0, 0, 50, 100);
    Window v2 = new Window(d, 51, 0, 100, 100);
    View v = new MultipleWindow(d, v1,v2);
    v.render();
    
    System.out.println("============ 修改数据 ================");
    
    List<Object> line = d.get();
    Object o = line.set(3, "hello");
    line.remove(o);
    v.render();
}
    
}
