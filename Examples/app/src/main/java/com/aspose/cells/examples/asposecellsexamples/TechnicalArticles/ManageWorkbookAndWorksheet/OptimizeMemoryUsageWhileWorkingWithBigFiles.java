package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.MemorySetting;
import com.aspose.cells.Workbook;

import java.io.File;

public class OptimizeMemoryUsageWhileWorkingWithBigFiles {

    private static final String TAG = OptimizeMemoryUsageWhileWorkingWithBigFiles.class.getName();

    public void readLargeExcelFiles() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Specify the LoadOptions
            LoadOptions opt = new LoadOptions();
            //Set the memory preferences
            opt.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);
            //Instantiate the Workbook
            //Load the Big Excel file having large Data set in it
            Workbook wb = new Workbook(filePath + "Book1.xlsx", opt);
        } catch (Exception e) {
            Log.e(TAG, "Read Large Excel Files", e);
        }
    }

    public void writeLargeExcelFiles() {
        try {
            Workbook wb = new Workbook();
            //Set the memory preferences
            //Note: This setting cannot take effect for the existing worksheets that are created before using the below line of code
            wb.getSettings().setMemorySetting(MemorySetting.MEMORY_PREFERENCE);

            //Note: The memory settings also would not work for the default sheet i.e., "Sheet1" etc. automatically created by the Workbook
            //To change the memory setting of existing sheets, please change memory setting for them manually:
            Cells cells = wb.getWorksheets().get(0).getCells();
            cells.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);
            //Input large dataset into the cells of the worksheet.
            //Your code goes here.

            //Get cells of the newly created Worksheet "Sheet2" whose memory setting is same with the one defined in WorkbookSettings:
            cells = wb.getWorksheets().add("Sheet2").getCells();
            //.........
            //Input large dataset into the cells of the worksheet.
            //Your code goes here.
            //.........
        } catch (Exception e) {
            Log.e(TAG, "Write Large Excel Files", e);
        }
    }

}
