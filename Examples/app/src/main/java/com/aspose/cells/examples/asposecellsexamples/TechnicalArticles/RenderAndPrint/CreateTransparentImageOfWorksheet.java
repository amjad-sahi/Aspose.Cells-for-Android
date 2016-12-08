package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;

import java.io.File;

public class CreateTransparentImageOfWorksheet {

    private static final String TAG = CreateTransparentImageOfWorksheet.class.getName();

    public void generateATransparentImage() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source file
            Workbook wb = new Workbook(filePath + "Book1.xls");

            //Apply different image or print options
            ImageOrPrintOptions imgOption = new ImageOrPrintOptions();
            imgOption.setImageFormat(ImageFormat.getPng());
            imgOption.setHorizontalResolution(200);
            imgOption.setVerticalResolution(200);
            imgOption.setOnePagePerSheet(true);

            //Apply transparency to the output image
            imgOption.setTransparent(true);

            //Create image after apply image or print options
            SheetRender sr = new SheetRender(wb.getWorksheets().get(0), imgOption);
            sr.toImage(0, filePath + "GenerateATransparentImage_Out.png");
        } catch (Exception e) {
            Log.e(TAG, "Create Transparent Image of Excel Worksheet", e);
        }
    }
}