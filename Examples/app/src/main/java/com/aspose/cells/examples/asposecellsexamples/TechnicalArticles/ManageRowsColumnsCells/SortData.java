package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.Cells;
import com.aspose.cells.DataSorter;
import com.aspose.cells.SortOrder;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SortData {

    private static final String TAG = SortData.class.getName();

    public void sortData() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + "Book1.xls");
            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            //Get the cells collection in the sheet
            Cells cells = worksheet.getCells();

            //Obtain the DataSorter object in the workbook
            DataSorter sorter = workbook.getDataSorter();
            //Set the first order
            sorter.setOrder1(SortOrder.ASCENDING);
            //Define the first key.
            sorter.setKey1(0);
            //Set the second order
            sorter.setOrder2(SortOrder.ASCENDING);
            //Define the second key
            sorter.setKey2(1);

            //Create a cells area (range).
            CellArea ca = new CellArea();
            //Specify the start row index.
            ca.StartRow = 1;
            //Specify the start column index.
            ca.StartColumn = 0;
            //Specify the last row index.
            ca.EndRow = 9;
            //Specify the last column index.
            ca.EndColumn = 2;
            //Sort data in the specified data range (A2:C10)
            sorter.sort(cells, ca);

            //Saving the excel file
            workbook.save(filePath + "SortData_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Sorting Data", e);
        }
    }
}
