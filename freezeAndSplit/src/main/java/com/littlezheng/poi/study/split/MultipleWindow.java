package com.littlezheng.poi.study.split;

import java.util.ArrayList;
import java.util.List;

public class MultipleWindow extends AbstractWindow{
    
    private List<Window> wins = new ArrayList<Window>();
    
    public MultipleWindow(Data data, Window... wins) {
        super(data);
        if(wins == null){
            throw new IllegalArgumentException("必须传入至少一个Window对象！");
        }
        for(Window win : wins){
            win.data = data;
            this.wins.add(win);
        }
    }
    
    @Override
    public void render() {
        for(Window win : wins){
            win.render();
        }
    }
}
