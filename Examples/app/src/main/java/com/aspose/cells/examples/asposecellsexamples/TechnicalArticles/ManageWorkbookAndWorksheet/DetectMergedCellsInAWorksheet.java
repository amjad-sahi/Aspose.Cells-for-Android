package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.util.ArrayList;

public class DetectMergedCellsInAWorksheet {

    private static final String TAG = DetectMergedCellsInAWorksheet.class.getName();

    public void identifyMergedCellAreasInAWorksheetAndUnmergeThem() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new Workbook
            Workbook wkBook = new Workbook(filePath + "MergeTrial.xlsx");
            //Get a worksheet in the workbook
            Worksheet wkSheet = wkBook.getWorksheets().get("Merge Trial");
            //Clear its contents
            wkSheet.getCells().clearContents(0, 0, wkSheet.getCells().getMaxDataRow(), wkSheet.getCells().getMaxDataColumn());

            //Create an ArrayList object
            //Get the merged cells list to put it into the ArrayList object
            ArrayList<CellArea> al = wkSheet.getCells().getMergedCells();
            //Define cellarea
            CellArea ca;
            //Define some variables
            int frow, fcol, erow, ecol;
            //Loop through the ArrayList and get each cellarea
            //to unmerge it
            for (int i = al.size() - 1; i > -1; i--) {
                ca = new CellArea();
                ca = (CellArea) al.get(i);
                frow = ca.StartRow;
                fcol = ca.StartColumn;
                erow = ca.EndRow;
                ecol = ca.EndColumn;
                wkSheet.getCells().unMerge(frow, fcol, erow, ecol);
            }
            //Save the excel file
            wkBook.save(filePath + "MergeTrial_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Detect Merged Cells in a Worksheet", e);
        }
    }
}
