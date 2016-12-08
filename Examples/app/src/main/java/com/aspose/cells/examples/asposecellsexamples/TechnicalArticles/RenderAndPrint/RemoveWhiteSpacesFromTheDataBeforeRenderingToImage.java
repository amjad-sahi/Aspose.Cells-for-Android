package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.PrintingPageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class RemoveWhiteSpacesFromTheDataBeforeRenderingToImage {

    private static final String TAG = RemoveWhiteSpacesFromTheDataBeforeRenderingToImage.class.getName();

    public void removeWhiteSpacesFromTheDataBeforeRenderingToImage() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load an existing spreadsheet
            Workbook book = new Workbook(filePath + "sample.xlsx");

            //Get the first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Specify print area if required
            //sheet.PageSetup.PrintArea = "A1:H8";

            //Set the margins to 0 so that there is no white space around the image
            sheet.getPageSetup().setLeftMargin(0);
            sheet.getPageSetup().setRightMargin(0);
            sheet.getPageSetup().setTopMargin(0);
            sheet.getPageSetup().setBottomMargin(0);

            //Define ImageOrPrintOptions
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            imgOptions.setImageFormat(ImageFormat.getEmf());

            //Set OnePagePerSheet to true
            imgOptions.setOnePagePerSheet(true);
            //Set to ignore blank cells
            imgOptions.setPrintingPage(PrintingPageType.IGNORE_BLANK);

            //Create the SheetRender object based on the sheet with its
            //ImageOrPrintOptions attributes
            SheetRender render = new SheetRender(sheet, imgOptions);
            //Convert the worksheet to image
            render.toImage(0, filePath + "RemoveWhiteSpaces_Out.emf");

        } catch (Exception e) {
            Log.e(TAG, "Remove white spaces from the Data before Rendering to Image", e);
        }
    }
}
