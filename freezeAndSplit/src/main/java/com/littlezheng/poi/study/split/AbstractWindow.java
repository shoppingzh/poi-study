package com.littlezheng.poi.study.split;

public abstract class AbstractWindow implements View {

    protected Data data;
    protected int left;
    protected int top;
    protected int right;
    protected int bottom;

    public AbstractWindow(Data data) {
        this.data = data;
    }

    public AbstractWindow(Data data, int left, int top, int right, int bottom) {
        super();
        this.data = data;
        this.left = left;
        this.top = top;
        this.right = right;
        this.bottom = bottom;
    }

    public int getLeft() {
        return left;
    }

    public void setLeft(int left) {
        this.left = left;
    }

    public int getTop() {
        return top;
    }

    public void setTop(int top) {
        this.top = top;
    }

    public int getRight() {
        return right;
    }

    public void setRight(int right) {
        this.right = right;
    }

    public int getBottom() {
        return bottom;
    }

    public void setBottom(int bottom) {
        this.bottom = bottom;
    }
}
