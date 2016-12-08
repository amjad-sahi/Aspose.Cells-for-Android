package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class CreateWorkbookAndWorksheetScopedNamedRanges {

    private static final String TAG = CreateWorkbookAndWorksheetScopedNamedRanges.class.getName();

    public void createWorkbookScopedNameRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Get Worksheets collection
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Accessing the first worksheet in the Excel file
            Worksheet sheet = worksheets.get(0);

            //Get worksheet Cells collection
            Cells cells = sheet.getCells();

            //Creating a workbook scope named range
            Range namedRange = cells.createRange("A1", "C10");
            namedRange.setName("workbookScope");

            //Saving the modified Excel file in default format
            workbook.save(filePath + "WorkbookScoped_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Create Workbook scoped name range", e);
        }
    }

    public void createWorksheetScopedNameRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Get Worksheets collection
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Accessing the first worksheet in the Excel file
            Worksheet sheet = worksheets.get(0);

            //Get worksheet Cells collection
            Cells cells = sheet.getCells();

            //Creating a workbook scope named range
            Range namedRange = cells.createRange("A1", "C10");
            namedRange.setName("Sheet1!local");

            //Saving the modified Excel file in default format
            workbook.save(filePath + "WorksheetScoped_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Create Worksheet scoped name range", e);
        }
    }

}
