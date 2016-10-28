package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.CellsHelper;
import com.aspose.cells.ReferredArea;
import com.aspose.cells.ReferredAreaCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class TracePrecedentsAndDependents {

    private static final String TAG = TracePrecedentsAndDependents.class.getName();

    public void tracePrecedent() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");
            Cells cells = workbook.getWorksheets().get(0).getCells();
            Cell cell = cells.get("B7");

            //Tracing precedents of the cell B7.
            //The return array contains ranges and cells.
            ReferredAreaCollection ret = cell.getPrecedents();
            //Printing all the precedent cells' name.
            if(ret != null) {
                for(int m = 0 ; m < ret.getCount(); m++) {
                    ReferredArea area = ret.get(m);
                    StringBuilder stringBuilder = new StringBuilder();
                    if (area.isExternalLink()) {
                        stringBuilder.append("[");
                        stringBuilder.append(area.getExternalFileName());
                        stringBuilder.append("]");
                    }
                    stringBuilder.append(area.getSheetName());
                    stringBuilder.append("!");
                    stringBuilder.append(CellsHelper.cellIndexToName(area.getStartRow(), area.getStartColumn()));
                    if (area.isArea()) {
                        stringBuilder.append(":");
                        stringBuilder.append(CellsHelper.cellIndexToName(area.getEndRow(), area.getEndColumn()));
                    }
                    Log.v(TAG, stringBuilder.toString());
                }
            }

        } catch (Exception e) {
            Log.e(TAG, "Trace Precedent", e);
        }
    }

    public void traceDependents() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Open a template Excel file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");
            //Get the first worksheet (default worksheet)
            Worksheet worksheet = workbook.getWorksheets().get(0);
            //Get the A1 cell
            Cell c = worksheet.getCells().get("A1");
            //Get the all the Dependents of A1 cell
            Cell[] dependents = c.getDependents(true);
            for (int i = 0; i< dependents.length; i++) {
                Log.v(TAG, dependents[i].getWorksheet().getName() + "----" + dependents[i].getName() + ":" + dependents[i].getIntValue());
            }
        } catch (Exception e) {
            Log.e(TAG, "Trace Dependents", e);
        }
    }
}
