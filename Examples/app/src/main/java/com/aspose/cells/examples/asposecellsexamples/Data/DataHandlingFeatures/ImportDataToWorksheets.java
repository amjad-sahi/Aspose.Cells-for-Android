package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;

public class ImportDataToWorksheets {

    private static final String TAG = ImportDataToWorksheets.class.getName();

    public void importFromArray() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet= workbook.getWorksheets().get(sheetIndex);

            //Creating an array containing names as string values
            String[] names=new String[]{"laurence chen", "roman korchagin", "kyle huang"};

            //Importing the array of names to 1st row and first column vertically
            Cells cells = worksheet.getCells();
            cells.importArray(names, 0, 0, false);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "ImportFromArray_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Import From Array", e);
        }
    }

    public void importFromMultiDimensionalArray() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook
            Workbook workbook = new Workbook();
            //Get the first worksheet (default sheet) in the Workbook
            Cells cells = workbook.getWorksheets().get("Sheet1").getCells();

            //Define a multi-dimensional array and store some data into it.
            String[][] strArray = {
                    {"A", "1A","2A" },
                    {"B", "2B", "3B"}
            };

            //Import the multi-dimensional array to the sheet
            cells.importArray(strArray, 0, 0);

            //Save the Excel file
            workbook.save(filePath + File.separator + "ImportFromMultiDimArray_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Import from Multi-dimensional Array", e);
        }
    }

    public void importFromArrayList() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook
            Workbook workbook = new Workbook();
            //Get the first worksheet (default sheet) in the Workbook
            Cells cells = workbook.getWorksheets().get("Sheet1").getCells();

            //Instantiating an ArrayList object
            ArrayList<String> list = new ArrayList();
            list.add("laurence chen");
            list.add("roman korchagin");
            list.add("kyle huang");
            list.add("tommy wang");

            //Importing the contents of ArrayList to 1st row and first column vertically
            cells.importArrayList(list, 0, 0, true);

            //Save the Excel file
            workbook.save(filePath + File.separator + "ImportFromArrayList_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Import from ArrayList", e);
        }
    }

    public void importFromResultSet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook =new Workbook();

            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
            String connectionString = "jdbc:ucanaccess://" + filePath + File.separator + "Northwind.mdb";

            // DSN-less DB connection.
            Connection connection = DriverManager.getConnection(connectionString);

            Statement statement = connection.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);

            //Get the ResultSet executing the SQL statement.
            ResultSet rs = statement.executeQuery("select EmployeeID, LastName, FirstName, Title, City from Employees");

            //Fetch the first worksheet.
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Import the ResultSet to the worksheet.
            worksheet.getCells().importResultSet(rs,"A1",true);

            //Save the excel file.
            workbook.save(filePath + File.separator + "ImportFromResultSet.xls");

        } catch (Exception e) {
            Log.e(TAG, "Import from ResultSet", e);
        }
    }
}
