package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageConditionalFormattingAndIcons;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.DataBar;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.io.FileOutputStream;

public class GenerateConditionalFormattingDataBarsImages {

    private static final String TAG = GenerateConditionalFormattingDataBarsImages.class.getName();

    public void generateConditionalFormattingDataBarsImages() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the cell which contains conditional formatting databar
            Cell cell = worksheet.getCells().get("C1");

            //Get the conditional formatting of the cell
            FormatConditionCollection[] fcc = cell.getFormatConditions();

            //Access the conditional formatting databar
            DataBar dbar = fcc[0].get(0).getDataBar();

            //Create image or print options
            ImageOrPrintOptions opts = new ImageOrPrintOptions();
            opts.setImageFormat(ImageFormat.getPng());

            //Get the image bytes of the databar
            byte[] imgBytes = dbar.toImage(cell, opts);

            //Write image bytes on the disk
            FileOutputStream out = new FileOutputStream(filePath + "school.jpg");
            out.write(imgBytes);
            out.close();
        } catch (Exception e) {
            Log.e(TAG, "Generate Conditional Formatting DataBars Images", e);
        }
    }
}
