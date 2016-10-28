package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.DataSorter;
import com.aspose.cells.SortOrder;
import com.aspose.cells.Workbook;

import java.io.File;

public class DataSorting {
    private static final String TAG = DataSorting.class.getName();

    public void dataSorting() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook object.
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Get the workbook data sorter object.
            DataSorter sorter = workbook.getDataSorter();

            //Set the first order for data sorter object.
            sorter.setOrder1(SortOrder.DESCENDING);

            //Define the first key.
            sorter.setKey1(0);

            //Set the second order for data sorter object.
            sorter.setOrder2(SortOrder.ASCENDING);

            //Define the second key.
            sorter.setKey2(1);

            //Sort data in the specified data range (CellArea range: A1:B14)
            CellArea cellArea = new CellArea();
            cellArea.StartRow = 0;
            cellArea.StartColumn = 0;
            cellArea.EndRow = 13;
            cellArea.EndColumn = 1;
            sorter.sort(workbook.getWorksheets().get(0).getCells(), cellArea);

            //Save the excel file.
            workbook.save(filePath + File.separator + "DataSorting_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Data Sorting", e);
        }
    }
}
