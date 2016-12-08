package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;

import java.io.File;

public class ModifyAnExistingStyle {

    private static final String TAG = ModifyAnExistingStyle.class.getName();

    public void createAndModifyAStyle() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook.
            Workbook workbook = new Workbook();

            //Create a new style object.
            Style style = workbook.createStyle();

            //Set the number format.
            style.setNumber(14);

            //Set the font color to red color.
            style.getFont().setColor(Color.getRed());

            //Name the style.
            style.setName("Date1");

            //Get the first worksheet cells.
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Specify the style (described above) to A1 cell.
            cells.get("A1").setStyle(style);

            //Create a range (B1:D1).
            Range range = cells.createRange("B1", "D1");

            //Initialize styleflag object.
            StyleFlag flag = new StyleFlag();

            //Set all formatting attributes on.
            flag.setAll(true);

            //Apply the style (described above)to the range.
            range.applyStyle(style, flag);

            //Modify the style (described above) and change the font color from red to black.
            style.getFont().setColor(Color.getBlack());

            //Done! Since the named style (described above) has been set to a cell and range,
            //the change would be Reflected(new modification is implemented) to cell(A1) and //range (B1:D1).
            style.update();

            //Save the Excel file.
            workbook.save(filePath + "BookStyles_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Create and Modify a Style", e);
        }
    }

    public void modifyAStyleInATemplateFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook.
            //Open a template file.
            //In the book1.xls file, we have applied Ms Excel's
            //Named style i.e., "Percent" to the range "A1:C8".
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //We get the Percent style and create a style object.
            Style style = workbook.getNamedStyle("Percent");

            //Change the number format to "0.00%".
            style.setNumber(10);

            //Set the font color.
            style.getFont().setColor(Color.getRed());

            //Update the style. so, the style of range "A1:C8" will be changed too.
            style.update();

            //Save the Excel file.
            workbook.save(filePath + "ModifyAStyleInATemplateFile_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Modify a Style in a Template File", e);
        }
    }
}
