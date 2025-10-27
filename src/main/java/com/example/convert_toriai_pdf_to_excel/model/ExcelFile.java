package com.example.convert_toriai_pdf_to_excel.model;

public class ExcelFile {
    private String name;

    private String kouSyuName;

    private final double kouzaiChouGoukei;
    private final double seiHinChouGoukei;

    public String getKouSyuName() {
        return kouSyuName;
    }

    public void setKouSyuName(String kouSyuName) {
        this.kouSyuName = kouSyuName;
    }

    public ExcelFile(String name, String kouSyuName, double kouzaiChouGoukei, double seiHinChouGoukei) {
        this.name = name;
        this.kouSyuName = kouSyuName;
        this.kouzaiChouGoukei = kouzaiChouGoukei;
        this.seiHinChouGoukei = seiHinChouGoukei;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public double getKouzaiChouGoukei() {
        return kouzaiChouGoukei;
    }

    public double getSeiHinChouGoukei() {
        return seiHinChouGoukei;
    }

    @Override
    public String toString() {
        return name;
    }
}
