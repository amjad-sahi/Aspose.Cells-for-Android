package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SetDefaultFontWhileRenderingSpreadsheetToImage {
    private static final String TAG = SetDefaultFontWhileRenderingSpreadsheetToImage.class.getName();

    public void setDefaultFontWhileRenderingSpreadsheetToImages() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Set default font of the workbook to none
            Style style = book.getDefaultStyle();
            style.getFont().setName("");
            book.setDefaultStyle(style);

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Access cell A4 and add some text inside it
            Cell cell = sheet.getCells().get("A4");
            cell.putValue("This text has some unknown or invalid font which does not exist.");

            //Set the font of cell A4 which is unknown
            style = cell.getStyle();
            style.getFont().setName("UnknownNotExist");
            style.getFont().setSize(20);
            style.setTextWrapped(true);
            cell.setStyle(style);

            //Set first column width and fourth column height
            sheet.getCells().setColumnWidth(0, 80);
            sheet.getCells().setRowHeight(3, 60);

            //Create an instance of ImageOrPrintOptions
            ImageOrPrintOptions opts = new ImageOrPrintOptions();
            opts.setOnePagePerSheet(true);
            opts.setImageFormat(ImageFormat.getPng());

            //Render worksheet image with Courier New as default font
            opts.setDefaultFont("Courier New");
            SheetRender sr = new SheetRender(sheet, opts);
            sr.toImage(0, filePath + "out_courier_new.png");

            //Render worksheet image again with Times New Roman as default font
            opts.setDefaultFont("Times New Roman");
            sr = new SheetRender(sheet, opts);
            sr.toImage(0, filePath + "out_times_new_roman.png");
        } catch (Exception e) {
            Log.e(TAG, "Set Default Font while Rendering Spreadsheet to Images", e);
        }
    }
}
