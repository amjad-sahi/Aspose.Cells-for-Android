package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExportWorksheetToImageByPage {

    private static final String TAG = ExportWorksheetToImageByPage.class.getName();

    public void renderFirstPageOfWorksheetToJPEGFormat() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load an existing spreadsheet
            Workbook book = new Workbook(filePath + "Book1.xlsx");
            //Access the first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Define ImageOrPrintOptions
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            //Specify the image format
            imgOptions.setImageFormat(ImageFormat.getJpeg());

            //Render the sheet with respect to specified image/print options
            SheetRender render = new SheetRender(sheet, imgOptions);
            //Render the first page of the Worksheet to image
            render.toImage(0, filePath + "SheetImage_Out.jpg");
        } catch (Exception e) {
            Log.e(TAG, "Render the first page of worksheet to JPEG format", e);
        }
    }

    public void renderAllWorksheetsPrintingPagesToSeparateImages() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load an existing spreadsheet
            Workbook book = new Workbook(filePath + "Book1.xlsx");
            //Access the first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Define ImageOrPrintOptions
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            //Specify the image format
            imgOptions.setImageFormat(ImageFormat.getJpeg());

            //Render the sheet with respect to specified image/print options
            SheetRender render = new SheetRender(sheet, imgOptions);
            //Iterate over all available worksheet pages
            for (int j = 0; j < render.getPageCount(); j++) {
                //Export each page to separate image
                render.toImage(j, filePath + sheet.getName() + " Page" + (j + 1) + "_Out.jpg");
            }
        } catch (Exception e) {
            Log.e(TAG, "Render all worksheets printing pages to separate images", e);
        }
    }
}
