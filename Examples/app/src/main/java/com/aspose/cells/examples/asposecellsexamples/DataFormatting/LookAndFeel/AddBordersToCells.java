package com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AddBordersToCells {

    private static final String TAG = AddBordersToCells.class.getName();

    /**
     * Add borders to a cell by using the Style object's setBorder method.
     * Pass the border type as a parameter.
     * The border types are pre-defined in the BorderType enumeration.
     */
    public void addBordersToACell() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Accessing the "A1" cell from the worksheet
            Cell cell = cells.get("A1");

            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");
            Style style = cell.getStyle();

            //Setting the line of the top border
            style.setBorder(BorderType.TOP_BORDER, CellBorderType.THICK, Color.getBlack());

            //Setting the line of the bottom border
            style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.THICK, Color.getBlack());

            //Setting the line of the left border
            style.setBorder(BorderType.LEFT_BORDER, CellBorderType.THICK, Color.getBlack());

            //Setting the line of the right border
            style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.THICK, Color.getBlack());

            //Saving the modified style to the "A1" cell.
            cell.setStyle(style);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddBordersToACell_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Add Borders to a Cell", e);
        }
    }

    /**
     * Sometimes, developers may need to add borders to a range of cells rather than just a single cell.
     * To do so, developers would first need to create a range of cells by calling the createRange method of the Cells collection.
     */
    public void addBordersToARangeOfCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Workbook object
            //Obtaining the reference of the newly added worksheet
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Accessing the "A1" cell from the worksheet
            Cell cell = worksheet.getCells().get("A1");

            //Adding some value to the "A1" cell
            cell.setValue("Hello World From Aspose");

            //Creating a range of cells starting from "A1" cell to 3rd column in a row
            Range range = worksheet.getCells().createRange(0, 0, 0, 2);
            range.setName("MyRange");

            //Adding a thick outline border with the blue line
            range.setOutlineBorders(CellBorderType.THICK, Color.getBlue());

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddBordersToARangeOfCells_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Borders to a Range of Cells", e);
        }
    }
}
