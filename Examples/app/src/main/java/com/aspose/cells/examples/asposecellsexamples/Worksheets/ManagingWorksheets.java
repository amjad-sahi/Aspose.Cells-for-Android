package com.aspose.cells.examples.asposecellsexamples.Worksheets;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;
import java.io.FileInputStream;

public class ManagingWorksheets {

    private static final String TAG = ManagingWorksheets.class.getName();

    public void createANewExcelFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            int sheetIndex = worksheets.add();
            Worksheet worksheet = worksheets.get(sheetIndex);

            //Setting the name of the newly added worksheet
            worksheet.setName("New Worksheet");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "NewExcelFile_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Create a New Excel File", e);
        }
    }

    public void addWorksheetToADesignerSpreadsheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating a file stream containing the Excel file to be opened
            FileInputStream fstream = new FileInputStream(filePath + File.separator + "Book1.xls");

            //Instantiating a Workbook object with the stream
            Workbook workbook = new Workbook(fstream);

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();
            int sheetIndex = worksheets.add();
            Worksheet worksheet = worksheets.get(sheetIndex);

            //Setting the name of the newly added worksheet
            worksheet.setName("New Worksheet");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddWorksheetToADesignerSpreadsheet_Out.xls");

            //Close the file stream to free all resources
            fstream.close();

        } catch (Exception e) {
            Log.e(TAG, "Add Worksheets to a Designer Spreadsheet", e);
        }
    }

    public void accessWorksheetUsingSheetName() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Creating a file stream containing the Excel file to be opened
            FileInputStream fstream = new FileInputStream(filePath + File.separator + "Book1.xls");

            // Instantiating a Workbook object with the stream
            Workbook workbook = new Workbook(fstream);

            // Access a worksheet using its sheet name
            Worksheet worksheet = workbook.getWorksheets().get("Sheet1");

        } catch (Exception e) {
            Log.e(TAG, "Access Worksheet using Sheet Name", e);
        }
    }

    public void removeWorksheetUsingSheetName() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Creating a file stream containing the Excel file to be opened
            FileInputStream fstream = new FileInputStream(filePath + File.separator + "Book1.xls");

            // Instantiating a Workbook object with the stream
            Workbook workbook = new Workbook(fstream);

            //Remove a worksheet using its sheet name
            workbook.getWorksheets().removeAt("Sheet1");

        } catch (Exception e) {
            Log.e(TAG, "Access Worksheet using Sheet Name", e);
        }
    }

    public void removeWorksheetUsingSheetIndex() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            // Creating a file stream containing the Excel file to be opened
            FileInputStream fstream = new FileInputStream(filePath + File.separator + "Book1.xls");

            // Instantiating a Workbook object with the stream
            Workbook workbook = new Workbook(fstream);

            //Remove a worksheet using its sheet index
            workbook.getWorksheets().removeAt(0);

        } catch (Exception e) {
            Log.e(TAG, "Remove Worksheet Using Sheet Index", e);
        }
    }

}
