package com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class CopyAndMoveWorksheet {
    private static final String TAG = CopyAndMoveWorksheet.class.getName();

    public void copyWorksheetsWithinAWorkbook() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook wb = new Workbook(filePath + File.separator + "Book1.xls");
            WorksheetCollection sheets = wb.getWorksheets();

            //Copy data to a new sheet from an existing sheet within the Workbook.
            sheets.addCopy("Sheet1");

            //Save the Excel file.
            wb.save(filePath + File.separator + "CopyWorksheetsWithinAWorkbook_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Copy Worksheets within a Workbook", e);
        }
    }

    public void copyWorksheetsBetweenWorkbooks() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a Workbook
            Workbook excelWorkbook0 = new Workbook(filePath + File.separator + "Book1.xls");

            //Create another Workbook
            Workbook excelWorkbook1 = new Workbook();

            //Copy the first sheet of the first book into second book
            excelWorkbook1.getWorksheets().get(0).copy(excelWorkbook0.getWorksheets().get(0));

            //Save the file.
            excelWorkbook1.save(filePath + File.separator + "CopyWorksheetsBetweenWorkbooks_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Copy Worksheets Between Workbook", e);
        }
    }

    public void copyAWorksheetDataFromOneWorkbookToAnother() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook excelWorkbook0 = new Workbook();

            //Get the first worksheet in the book.
            Worksheet ws0 = excelWorkbook0.getWorksheets().get(0);

            //Put some data into header rows (A1:A4)
            for (int i = 1; i < 5; i++) {
                ws0.getCells().get("A" + i).setValue("Header Row " + i);
            }

            //Put some detail data (A5:A999)
            for (int i = 5; i < 1000; i++) {
                ws0.getCells().get("A" + i).setValue("Detail Row " + i);
            }

            //Define a pagesetup object based on the first worksheet.
            PageSetup pagesetup = ws0.getPageSetup();

            //The first five rows are repeated in each page...
            //It can be seen in print preview.
            pagesetup.setPrintTitleRows("$1:$5");

            //Create another Workbook.
            Workbook excelWorkbook1 = new Workbook();

            //Get the first worksheet in the book.
            Worksheet ws1 = excelWorkbook1.getWorksheets().get(0);

            //Name the worksheet.
            ws1.setName("MySheet");

            //Copy data from the first worksheet of the first workbook into the
            //first worksheet of the second workbook.
            ws1.copy(ws0);

            //Save the Excel file.
            excelWorkbook1.save(filePath + File.separator + "CopyAWorksheetDataFromOneWorkbookToAnother_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch (Exception e) {
            Log.e(TAG, "Copy a worksheet data from one workbook to another workbook", e);
        }
    }

    public void moveWorksheetsWithinWorkbook() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook wb = new Workbook(filePath + File.separator + "Book1.xls");

            //Get the first worksheet in the book.
            Worksheet sheet = wb.getWorksheets().get(0);

            //Move the first sheet to the third position in the workbook.
            sheet.moveTo(2);

            //Save the Excel file.
            wb.save(filePath + File.separator + "MoveWorksheetsWithinWorkbook_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Moving Worksheets within Workbook", e);
        }
    }

}
