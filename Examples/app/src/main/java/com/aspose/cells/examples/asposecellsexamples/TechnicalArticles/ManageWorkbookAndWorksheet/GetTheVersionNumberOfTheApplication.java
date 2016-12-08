package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class GetTheVersionNumberOfTheApplication {

    private static final String TAG = GetTheVersionNumberOfTheApplication.class.getName();

    public void getTheVersionNumberOfTheApplicationThatCreatedTheExcelDocument() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook reference
            Workbook workbook = null;

            //Print the version number of Excel 2003 XLS file
            workbook = new Workbook(filePath + "Excel2003.xls");
            System.out.println("Excel 2003 XLS Version: " + workbook.getBuiltInDocumentProperties().getVersion());

            //Print the version number of Excel 2010 XLSX file
            workbook = new Workbook(filePath + "Excel2010.xlsx");
            System.out.println("Excel 2010 XLS Version: " + workbook.getBuiltInDocumentProperties().getVersion());
        } catch (Exception e) {
            Log.e(TAG, "Get the Version Number of the Application that Created the Excel Document", e);
        }
    }
}
