package com.aspose.cells.examples.asposecellsexamples.DataFormatting.AdvanceFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Column;
import com.aspose.cells.Font;
import com.aspose.cells.Row;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FormatRowsAndColumns {

    private static final String TAG = FormatRowsAndColumns.class.getName();

    /**
     * The Cells collection provides a Rows collection.
     * Each item in the Rows collection represents a Row object.
     * The Row object offers an applyStyle method that is used to set formatting on a complete row.
     */
    public void formatARow() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Accessing the newly added Style to the Excel object
            Style style = workbook.createStyle();

            //Setting the vertical alignment of the text in the cell
            style.setVerticalAlignment(TextAlignmentType.CENTER);

            //Setting the horizontal alignment of the text in the cell
            style.setHorizontalAlignment(TextAlignmentType.CENTER);

            //Setting the font color of the text in the cell
            Font font = style.getFont();
            font.setColor(Color.getGreen());

            //Shrinking the text to fit in the cell
            style.setShrinkToFit(true);

            //Setting the bottom border of the cell
            style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

            //Creating StyleFlag
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.setHorizontalAlignment(true);
            styleFlag.setVerticalAlignment(true);
            styleFlag.setShrinkToFit(true);
            styleFlag.setBottomBorder(true);
            styleFlag.setFontColor(true);

            //Accessing a row from the Rows collection
            Row row = cells.getRows().get(0);

            //Assigning the Style object to the Style property of the row
            row.applyStyle(style, styleFlag);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "FormatARow_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Format a Row", e);
        }
    }

    /**
     * The Cells collection also provides a Columns collection.
     * Each item in the Columns collection represents a Column object.
     * Similar to the Row object, the Column object also offers the applyStyle method that is used to format a column.
     * Use the Column object's applyStyle method to format a column in the same way as for a row.
     */
    public void formatAColumn() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Accessing the newly added Style to the Excel object
            Style style = workbook.createStyle();

            //Setting the vertical alignment of the text in the cell
            style.setVerticalAlignment(TextAlignmentType.CENTER);

            //Setting the horizontal alignment of the text in the cell
            style.setHorizontalAlignment(TextAlignmentType.CENTER);

            //Setting the font color of the text in the cell
            Font font = style.getFont();
            font.setColor(Color.getGreen());

            //Shrinking the text to fit in the cell
            style.setShrinkToFit(true);

            //Setting the bottom border of the cell
            style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

            //Creating StyleFlag
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.setHorizontalAlignment(true);
            styleFlag.setVerticalAlignment(true);
            styleFlag.setShrinkToFit(true);
            styleFlag.setBottomBorder(true);
            styleFlag.setFontColor(true);

            //Accessing a column from the Columns collection
            Column column = cells.getColumns().get(0);

            //Applying the style to the column
            column.applyStyle(style, styleFlag);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "FormatAColumn_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Format a Column", e);
        }
    }

    /**
     * Setting Display Format of Numbers & Dates for Rows & Columns
     */
    public void setDisplayFormatOfNumbersAndDatesForRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first (default) worksheet by passing its sheet index
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Adding a new Style to the styles collection of the Workbook object
            Style style = workbook.createStyle();

            //Setting the Number property to 4 which corresponds to the pattern #,##0.00
            style.setNumber(4);

            //Creating an object of StyleFlag
            StyleFlag flag = new StyleFlag();

            //Setting NumberFormat property to true so that only this aspect takes effect from Style object
            flag.setNumberFormat(true);

            //Applying style to the first row of the worksheet
            worksheet.getCells().getRows().get(0).applyStyle(style, flag);

            //Re-initializing the Style object
            style = workbook.createStyle();

            //Setting the Custom property to the pattern d-mmm-yy
            style.setCustom("d-mmm-yy");

            //Applying style to the first column of the worksheet
            worksheet.getCells().getColumns().get(0).applyStyle(style, flag);

            //Saving spreadsheet on disc
            workbook.save(filePath + File.separator + "DisplayFormat_Out.xlsx");
        } catch(Exception e) {
            Log.e(TAG, "Setting Display Format of Numbers and Dates for Rows and Columns", e);
        }
    }
}
