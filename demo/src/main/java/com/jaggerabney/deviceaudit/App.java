package com.jaggerabney.deviceaudit;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) {
        try {
            // load first sheet in frmInventory.xlsx
            Sheet sheet = loadWorkbook("frmInventory.xlsx").getSheetAt(0);
            System.out.println(sheet.getRow(0).getCell(0));
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private static XSSFWorkbook loadWorkbook(String filename) throws IOException {
        InputStream is = App.class.getClassLoader().getResourceAsStream(filename);
        return new XSSFWorkbook(is);
    }
}
