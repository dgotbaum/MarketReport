import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.hssf.record.cf.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.*;
import org.apache.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class PrintSheet {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX and XLS Spreadsheets", "xls", "xlsx");
        chooser.setFileFilter(filter);
        int returnVal = chooser.showOpenDialog(null);
        if(returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: " +
                    chooser.getSelectedFile().getName());
        }
        //Initialize Workbook and sheet
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet s = wb.createSheet("Lunches");
        //First Row/Cell for lunch choices.
		HSSFRow r = s.createRow(0);
		HSSFCell c = r.createCell((short) 1);
		c.setCellValue("Akata Sushi");
		HSSFRow ratings = s.createRow(1);
		HSSFCell akata = ratings.createCell((short)1);
		akata.setCellValue(7);
        // Creates a row of numbers to be colored
        Row numbers = s.createRow(2);
        Cell title = numbers.createCell(0);
        title.setCellValue("Color Numbers");
        for (int i = 1; i <= 10 ; i++) {
            Cell temp = numbers.createCell(i);
            temp.setCellValue(i);
        }
        // Colors the even numbers light blue
        SheetConditionalFormatting sCF = s.getSheetConditionalFormatting();
        ConditionalFormattingRule rule1 = sCF.createConditionalFormattingRule("MOD(ROW(),2)");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
        fill1.setFillPattern(org.apache.poi.hssf.record.cf.PatternFormatting.SOLID_FOREGROUND);
        CellRangeAddress[] regions =  {CellRangeAddress.valueOf("A1:Z100")};
        sCF.addConditionalFormatting(regions, rule1);


		FileOutputStream fileOut = new FileOutputStream("test.xls");
		wb.write(fileOut);
		fileOut.close();
		
		
	}

}
