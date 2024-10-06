/*
 * Copyright (c) 2024. All rights reserved.
 * File : App.java, Last Modified on : 06/10/24, 1:02 pm
 * Created by Vikas Tiwari.
 * Do not copy/modify without permission.
 * For any contact email at Vikastiwari199915@gmail.com
 */

package com.vikas;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class App 
{
    public static final String NAME_PREFIX = "SPLIT_EXCEL";
    private static final int CHUNK_SIZE = 10000;

    @SuppressWarnings("CallToPrintStackTrace")
    public static void main(String[] args ) {
        System.out.println( "#### EXCEL SPLITTER BY VIKAS ####");

        // 1. File Selection
        Frame frame = new Frame();
        FileDialog fileDialog = new FileDialog(frame, "Choose Large Excel file", FileDialog.LOAD);
        fileDialog.setVisible(true);
        String fileName = fileDialog.getFile();
        String directory = fileDialog.getDirectory();
        if (fileName == null || directory == null) {
            System.out.println("No file chosen.");
            System.err.println("***** EXIT *****");
            System.exit(0);
        }
        File excelFile = new File(directory, fileName);

        // 2. Choose Output Directory
        FileDialog dirDialog = new FileDialog(frame, "Choose Output Directory", FileDialog.LOAD);
        System.setProperty("apple.awt.fileDialogForDirectories", "true");
        dirDialog.setVisible(true);
        String outputDirName = dirDialog.getFile();
        String outputDirDirectory = dirDialog.getDirectory();
        if (outputDirName == null || outputDirDirectory == null) {
            System.out.println("No directory chosen.");
            System.err.println("***** EXIT *****");
            System.exit(0);
        }
        File outputDir = new File(outputDirDirectory, outputDirName);

        IOUtils.setByteArrayMaxOverride(500000000);
        System.err.println("---> STARTING SPLIT ");
        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = StreamingReader.builder().rowCacheSize(100).bufferSize(4096).open(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet

            int fileIndex = 1;
            int recordCount = 0;
            int rowIndex = 0;
            Workbook newWorkbook = WorkbookFactory.create(true);
            Workbook headerWorkbook = WorkbookFactory.create(true);
            Sheet newSheet = newWorkbook.createSheet("Sheet1");
            Row headerRow = null;
            for (Row row : sheet) {
                // Copy row to newSheet
                if(headerRow==null){
                    headerRow = headerWorkbook.createSheet("temp").createRow(0);
                    copyRow(row, headerRow);
                    System.err.println("### HEADER ROW COPIED ###");
                    continue;
                }
                if(newSheet.getPhysicalNumberOfRows()<1){
                    Row newHeaderRow = newSheet.createRow(rowIndex++);
                    copyRow(headerRow, newHeaderRow);
                    Row dataRow = newSheet.createRow(rowIndex++);
                    copyRow(row, dataRow);
                } else {
                    Row dataRow = newSheet.createRow(rowIndex++);
                    copyRow(row, dataRow);
                }
                recordCount++;
                //System.err.println("Copied : "+rowIndex);
                if (recordCount % CHUNK_SIZE == 0) {
                    // Save the current workbook and start a new one
                    rowIndex = 0;
                    try (FileOutputStream out = new FileOutputStream(new File(outputDir,NAME_PREFIX + "_" + fileIndex++ + ".xlsx"))) {
                        newWorkbook.write(out);
                        System.err.println("### SAVED :: "+NAME_PREFIX + "_" + (fileIndex-1) + ".xlsx");
                    }
                    newWorkbook.close();
                    newWorkbook = WorkbookFactory.create(true);
                    newSheet = newWorkbook.createSheet("Sheet1");
                }
            }

            // Write any remaining rows
            if (recordCount % CHUNK_SIZE != 0) {
                try (FileOutputStream out = new FileOutputStream(new File(outputDir,NAME_PREFIX + "_" + fileIndex + ".xlsx"))) {
                    newWorkbook.write(out);
                    System.err.println("### SAVED :: "+NAME_PREFIX + "_" + fileIndex + ".xlsx");
                    System.err.println("***** EXIT *****");
                }
            }
            System.exit(0);

        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("***** EXIT *****");
            System.exit(0);
        }
    }
    public static void copyRow(Row sourceRow, Row targetRow) {
        for (Cell cell : sourceRow) {
            Cell newCell = targetRow.createCell(cell.getColumnIndex(), cell.getCellType());
            switch (cell.getCellType()) {
                case STRING:
                    newCell.setCellValue(cell.getStringCellValue());
                    break;
                case NUMERIC:
                    newCell.setCellValue(cell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(cell.getCellFormula());
                    break;
                default:
                    break;
            }
        }
    }
}
