package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class App {
    public static String filePath = "C:\\poi_temp";
    public static String fileNm = "poi_reading_test.xlsx";

    public static void main(String[] args) {
        try (FileInputStream file = new FileInputStream(new File(filePath, fileNm))) {
            // 엑셀 파일로 Workbook 인스턴스를 생성한다.
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // workbook의 첫번째 sheet를 가져온다.
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 모든 행(row)들을 조회한다.
            for (Row row : sheet) {
                // 각각의 행에 존재하는 모든 열(cell)을 순회한다.
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // cell의 타입을 확인하고 값을 가져온다.
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            // getNumericCellValue 메서드는 기본적으로 double 형식을 반환한다.
                            System.out.print((int) cell.getNumericCellValue() + "\t");
                            break;

                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
