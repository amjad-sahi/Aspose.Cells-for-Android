package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.aspose.cells.LookInType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FindOrSearchData {

    private static final String TAG = FindOrSearchData.class.getName();

    public void findCellsThatContainAFormula() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Finding the cell containing the specified formula
            Cells cells = worksheet.getCells();
            FindOptions findOptions = new FindOptions();
            findOptions.setLookInType(LookInType.FORMULAS);
            Cell cell = cells.find("=SUM(A2:A5)", null, findOptions);

            //Printing the name of the cell found after searching worksheet
            if(cell != null) {
                Log.v(TAG, "Name of the cell containing formula: " + cell.getName());
            }
        } catch (Exception e) {
            Log.e(TAG, "Find Cells that Contain a Formula", e);
        }
    }

    public void searchWithPartialFormula() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Finding the cell with a formula that contains an input string
            Cells cells = worksheet.getCells();
            FindOptions findOptions = new FindOptions();
            findOptions.setLookAtType(LookAtType.CONTAINS);
            Cell cell = cells.find("SUM", null, findOptions);

            //Printing the name of the cell found after searching worksheet
            if(cell != null) {
                Log.v(TAG, cell.getName());
            }
        } catch (Exception e) {
            Log.e(TAG, "Search with Partial Formula", e);
        }
    }

    public void searchForStringsThatStartWithSpecificCharacters() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Finding the cell containing a string value that starts with "Or"
            Cells cells = worksheet.getCells();
            FindOptions findOptions = new FindOptions();
            findOptions.setLookAtType(LookAtType.START_WITH);
            Cell cell = cells.find("Or", null, findOptions);

            //Printing the name of the cell found after searching worksheet
            if(cell != null) {
                Log.v(TAG, cell.getName());
            }
        } catch (Exception e) {
            Log.e(TAG, "Searching for Strings that Start with Specific Characters", e);
        }
    }

    public void searchForStringsThatEndWithSpecificCharacters() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Finding the cell containing a string value that ends with "es"
            Cells cells = worksheet.getCells();
            FindOptions findOptions = new FindOptions();
            findOptions.setLookAtType(LookAtType.END_WITH);
            Cell cell = cells.find("es", null, findOptions);

            //Printing the name of the cell found after searching worksheet
            if(cell != null) {
                Log.v(TAG, cell.getName());
            }
        } catch (Exception e) {
            Log.e(TAG, "Searching for Strings that End with Specific Characters", e);
        }
    }

    public void searchWithRegularExpressions() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            Cells cells = worksheet.getCells();
            FindOptions opt = new FindOptions();
            //Set the search key of find() method as standard RegEx
            opt.setRegexKey(true);
            opt.setLookAtType(LookAtType.ENTIRE_CONTENT);
            Cell cell = cells.find("abc[\\s]*$", null, opt);

            //Printing the name of the cell found after searching worksheet
            if(cell != null) {
                Log.v(TAG, cell.getName());
            }
        } catch (Exception e) {
            Log.e(TAG, "Searching with Regular Expressions", e);
        }
    }
}
