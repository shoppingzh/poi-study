package com.littlezheng.poi.study.split;

public class Window extends AbstractWindow {

    public Window(Data data) {
        super(data);
    }

    public Window(Data data, int left, int top, int right, int bottom) {
        super(data, left, top, right, bottom);
    }

    @Override
    public void render() {
        System.out.println("在(" + left + ", " + top + ")到(" + right + ", " + bottom + ")的区域内显示：");
        System.out.println(data.toString());
    }

}
