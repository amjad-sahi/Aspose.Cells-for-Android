package com.aspose.cells.examples.asposecellsexamples.Data.AddOnFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class MergeAndUnmergeCells {

    private static final String TAG = MergeAndUnmergeCells.class.getName();

    public void mergeCellsInAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a Workbook.
            Workbook wbk = new Workbook();

            //Create a Worksheet and get the first sheet.
            Worksheet worksheet = wbk.getWorksheets().get(0);

            //Create a Cells object to fetch all the cells.
            Cells cells = worksheet.getCells();

            //Merge some Cells (C6:E7) into a single C6 Cell.
            cells.merge(5, 2, 7, 4);

            //Input data into C6 Cell.
            worksheet.getCells().get("C6").setValue("This is my value");

            //Create a Style object to fetch the Style of C6 Cell.
            Style style = worksheet.getCells().get("C6").getStyle();

            //Create a Font object
            Font font = style.getFont();

            //Set the name.
            font.setName("Times New Roman");

            //Set the font size.
            font.setSize(18);

            //Set the font color
            font.setColor(Color.getBlue());

            //Bold the text
            font.setBold(true);

            //Make it italic
            font.setItalic(true);

            //Set the backgrond color of C6 Cell to Red
            style.setBackgroundColor(Color.getRed());
            style.setPattern(BackgroundType.SOLID);

            //Apply the Style to C6 Cell.
            cells.get("C6").setStyle(style);

            //Save the Workbook.
            wbk.save(filePath + File.separator + "MergeCells_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Merge Cells in a Worksheet", e);
        }
    }

    public void unmergeMergedCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a Workbook.
            Workbook wbk = new Workbook(filePath + File.separator + "mergingcells.xls");

            //Create a Worksheet and get the first sheet.
            Worksheet worksheet = wbk.getWorksheets().get(0);

            //Create a Cells object to fetch all the cells.
            Cells cells = worksheet.getCells();

            //Unmerge the cells.
            cells.unMerge(5, 2, 7, 4);

            //Save the file.
            wbk.save(filePath + File.separator + "UnmergeCells.xls");
        } catch (Exception e) {
            Log.e(TAG, "Unmerge Merged Cells", e);
        }
    }
}
