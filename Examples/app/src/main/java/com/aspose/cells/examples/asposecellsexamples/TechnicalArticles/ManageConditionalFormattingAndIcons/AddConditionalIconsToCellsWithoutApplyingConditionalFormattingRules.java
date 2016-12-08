package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageConditionalFormattingAndIcons;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.ConditionalFormattingIcon;
import com.aspose.cells.IconSetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.ByteArrayInputStream;
import java.io.File;

public class AddConditionalIconsToCellsWithoutApplyingConditionalFormattingRules {

    private static final String TAG = AddConditionalIconsToCellsWithoutApplyingConditionalFormattingRules.class.getName();

    public void addConditionalIconsToCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate an instance of Workbook
            Workbook workbook = new Workbook();

            //Get the first worksheet (default worksheet) in the workbook
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the cells
            Cells cells = worksheet.getCells();

            //Set the columns widths (A, B and C)
            worksheet.getCells().setColumnWidth(0, 24);
            worksheet.getCells().setColumnWidth(1, 24);
            worksheet.getCells().setColumnWidth(2, 24);

            //Input date into the cells
            cells.get("A1").setValue("KPIs");
            cells.get("A2").setValue("Total Turnover (Sales at List)");
            cells.get("A3").setValue("Total Gross Margin %");
            cells.get("A4").setValue("Total Net Margin %");
            cells.get("B1").setValue("UA Contract Size Group 4");
            cells.get("B2").setValue(19551794);
            cells.get("B3").setValue(11.8070745566204);
            cells.get("B4").setValue(11.858589818569);
            cells.get("C1").setValue("UA Contract Size Group 3");
            cells.get("C2").setValue(8150131.66666667);
            cells.get("C3").setValue(10.3168384396244);
            cells.get("C4").setValue(11.3326931937091);

            //Get the conditional icon's image data
            byte[] imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.TRAFFIC_LIGHTS_31, 0);
            //Create a stream based on the image data
            ByteArrayInputStream stream = new ByteArrayInputStream(imagedata);
            //Add the picture to the cell based on the stream
            worksheet.getPictures().add(1, 1, stream);

            imagedata = null;

            //Get the conditional icon's image data
            imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.ARROWS_3, 2);
            //Create a stream based on the image data
            stream = new ByteArrayInputStream(imagedata);
            //Add the picture to the cell based on the stream
            worksheet.getPictures().add(1, 2, stream);

            imagedata = null;

            //Get the conditional icon's image data
            imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.SYMBOLS_3, 0);
            //Create a stream based on the image data
            stream = new ByteArrayInputStream(imagedata);
            //Add the picture to the cell based on the stream
            worksheet.getPictures().add(2, 1, stream);

            imagedata = null;

            //Get the conditional icon's image data
            imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.STARS_3, 0);
            //Create a stream based on the image data
            stream = new ByteArrayInputStream(imagedata);
            //Add the picture to the cell based on the stream
            worksheet.getPictures().add(2, 2, stream);

            imagedata = null;

            //Get the conditional icon's image data
            imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.BOXES_5, 1);
            //Create a stream based on the image data
            stream = new ByteArrayInputStream(imagedata);
            //Add the picture to the cell based on the stream
            worksheet.getPictures().add(3, 1, stream);

            imagedata = null;

            //Get the conditional icon's image data
            imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.FLAGS_3, 1);
            //Create a stream based on the image data
            stream = new ByteArrayInputStream(imagedata);
            //Add the picture to the cell based on the stream
            worksheet.getPictures().add(3, 2, stream);

            imagedata = null;

            //Save the Excel file
            workbook.save(filePath + "AddConditionalIconsToCells.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Add Conditional Icons to Cells without Applying Conditional Formatting Rules", e);
        }
    }
}
