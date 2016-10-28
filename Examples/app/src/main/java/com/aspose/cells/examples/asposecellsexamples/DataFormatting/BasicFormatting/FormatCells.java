package com.aspose.cells.examples.asposecellsexamples.DataFormatting.BasicFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FormatCells {

    private static final String TAG = FormatCells.class.getName();

    /**
     * To apply different formatting to different cells, use the Cell class' setStyle method.
     * This example shows how to use it to apply formatting to a cell.
     */
    public void usingTheSetStyleMethod() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Accessing the "A1" cell from the worksheet
            Cell cell = cells.get("A1");

            //Adding some value to the "A1" cell
            cell.setValue("Hello Aspose!");

            Style style = cell.getStyle();

            //Setting the vertical alignment of the text in the "A1" cell
            style.setVerticalAlignment(TextAlignmentType.CENTER);

            //Setting the horizontal alignment of the text in the "A1" cell
            style.setHorizontalAlignment(TextAlignmentType.CENTER);

            //Setting the font color of the text in the "A1" cell
            Font font = style.getFont();
            font.setColor(Color.getGreen());

            //Setting the cell to shrink according to the text contained in it
            style.setShrinkToFit(true);

            //Setting the bottom border
            style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

            //Saved style
            cell.setStyle(style);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "StyleMethod_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Using the setStyle Method", e);
        }
    }

    /**
     * To apply the same formatting to different cells, use the Style object.
     * This approach can greatly improve the efficiency of the application and saves memory too.
     */
    public void usingTheStyleObject() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Accessing the "A1" cell from the worksheet
            Cell cell = cells.get("A1");

            //Adding some value to the "A1" cell
            cell.setValue("Hello Aspose!");

            //Adding a new Style to the styles collection of the Excel object
            Style style = workbook.createStyle();

            //Setting the vertical alignment of the text in the "A1" cell
            style.setVerticalAlignment(TextAlignmentType.CENTER);

            //Setting the horizontal alignment of the text in the "A1" cell
            style.setHorizontalAlignment(TextAlignmentType.CENTER);

            //Setting the font color of the text in the "A1" cell
            Font font = style.getFont();
            font.setColor (Color.getGreen());

            //Setting the cell to shrink according to the text contained in it
            style.setShrinkToFit(true);

            //Setting the bottom border
            style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

            //Saved style
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "StyleObject_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Using the Style Object", e);
        }
    }
}
