package com.aspose.cells.examples.asposecellsexamples.Table;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ListObjectCollection;
import com.aspose.cells.TotalsCalculation;
import com.aspose.cells.Workbook;

import java.io.File;

public class CreateAListObject {
    private static final String TAG = CreateAListObject.class.getName();

    public void createAList() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a Workbook object.
            //Open a template excel file.
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Get the List objects collection in the first worksheet.
            ListObjectCollection listObjects = workbook.getWorksheets().get(0).getListObjects();

            //Add a List based on the data source range with headers on.
            listObjects.add(1, 1, 11, 5, true);

            //Show the total row for the List.
            listObjects.get(0).setShowTotals(true);

            //Calculate the total of the last (5th ) list column.
            listObjects.get(0).getListColumns().get(4).setTotalsCalculation(TotalsCalculation.SUM);

            //Save the excel file.
            workbook.save(filePath + File.separator + "CreateAList_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Create a List", e);
        }
    }

}
