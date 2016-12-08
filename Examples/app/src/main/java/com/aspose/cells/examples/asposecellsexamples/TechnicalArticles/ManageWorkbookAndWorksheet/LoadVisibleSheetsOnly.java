package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.LoadDataOption;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class LoadVisibleSheetsOnly {

    private static final String TAG = LoadVisibleSheetsOnly.class.getName();

    public void loadVisibleSheetsOnly() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            // Create a sample workbook
            // and put some data in first cell of all 3 sheets
            Workbook createWorkbook = new Workbook();
            createWorkbook.getWorksheets().get("Sheet1").getCells().get("A1").setValue("Aspose");
            createWorkbook.getWorksheets().add("Sheet2").getCells().get("A1").setValue("Aspose");
            createWorkbook.getWorksheets().add("Sheet3").getCells().get("A1").setValue("Aspose");
            // Keep Sheet3 invisible
            createWorkbook.getWorksheets().get("Sheet3").setVisible(false);
            createWorkbook.save(filePath + "VisibleSheets_Out.xlsx");

            // Load the sample workbook
            LoadDataOption loadDataOption = new LoadDataOption();
            loadDataOption.setOnlyVisibleWorksheet(true);
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.setLoadDataAndFormatting(true);
            loadOptions.setLoadDataOptions(loadDataOption);

            Workbook loadWorkbook = new Workbook(filePath + "VisibleSheets_Out.xlsx", loadOptions);

            System.out.println("Sheet1: A1: " + loadWorkbook.getWorksheets().get("Sheet1").getCells().get("A1").getValue());
            System.out.println("Sheet2: A1: " + loadWorkbook.getWorksheets().get("Sheet2").getCells().get("A1").getValue());
            System.out.println("Sheet3: A1: " + loadWorkbook.getWorksheets().get("Sheet3").getCells().get("A1").getValue());

            Log.e(TAG, "Data is not loaded from invisible sheet");
        } catch (Exception e) {
            Log.e(TAG, "Load Visible Sheets Only", e);
        }
    }

}
