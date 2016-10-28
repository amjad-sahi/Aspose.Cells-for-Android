package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ConvertExcelFilesToXPS {

    private static final String TAG = ConvertExcelFilesToXPS.class.getName();

    public void convertSingleWorksheetToXPS() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            ///Get the first worksheet
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Apply different Image and Print options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            //Set the format to XPS
            options.setSaveFormat(SaveFormat.XPS);

            //Render the sheet with respect to specified printing options
            SheetRender render = new SheetRender(sheet, options);
            render.toImage(0, filePath + File.separator + "ConvertWorksheetToXPS_Out.xps");
        } catch (Exception e) {
            Log.e(TAG, "Convert Excel to HTML", e);
        }
    }

    public void quickExcelToXPSConversion() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a workbook object from the template file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Save in XPS format
            workbook.save(filePath + File.separator + "QuickExcelToXPSConversion_Out.xps", SaveFormat.XPS);
        } catch (Exception e) {
            Log.e(TAG, "Convert Excel to HTML", e);
        }
    }


}
