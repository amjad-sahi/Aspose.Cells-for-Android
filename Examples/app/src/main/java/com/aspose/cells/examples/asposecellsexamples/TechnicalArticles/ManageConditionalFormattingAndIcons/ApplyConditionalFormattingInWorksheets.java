package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageConditionalFormattingAndIcons;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.Color;
import com.aspose.cells.ConditionalFormattingCollection;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ApplyConditionalFormattingInWorksheets {

    private static final String TAG = ApplyConditionalFormattingInWorksheets.class.getName();

    public void applyConditionalFormattingBasedOnCellValue() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            // Instantiating a Workbook object
            Workbook workbook = new Workbook();
            // Retrieve the first Worksheet of Workbook
            Worksheet sheet = workbook.getWorksheets().get(0);

            // Adds an empty conditional formatting
            int index = sheet.getConditionalFormattings().add();

            // Retrieve the instance of newly added condition
            FormatConditionCollection fcs = sheet.getConditionalFormattings().get(index);

            // Sets the conditional format range.
            CellArea ca = new CellArea();
            ca.StartRow = 0;
            ca.EndRow = 0;
            ca.StartColumn = 0;
            ca.EndColumn = 0;

            // Add the range to FormatConditionCollection
            fcs.addArea(ca);

            // Sets condition formula
            int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "50", "100");

            // Retrieve the FormatCondition
            FormatCondition fc = fcs.get(conditionIndex);

            // Set color for the FormatCondition
            fc.getStyle().setBackgroundColor(Color.getRed());

            // Save result in spreadsheet format
            workbook.save(filePath + "ConditionalFormattingBasedOnCellValue_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Apply Conditional Formatting Based on Cell Value", e);
        }
    }

    public void applyConditionalFormattingBasedOnAFormula() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Retrieve the first Worksheet of Workbook
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Retrieve the collection of conditional formatting rules of worksheet
            ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();

            //Add a new formatting condition
            int index = cfs.add();

            //Retrieve newly added format condition collection
            FormatConditionCollection fcs = cfs.get(index);

            //Set the conditional format range
            CellArea ca = new CellArea();
            ca = new CellArea();
            ca.StartRow = 2;
            ca.EndRow = 2;
            ca.StartColumn = 1;
            ca.EndColumn = 1;

            //Add the CellArea to FormatConditionCollection
            fcs.addArea(ca);

            //Sets condition formula
            int conditionIndex = fcs.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "", "");

            FormatCondition fc = fcs.get(conditionIndex);

            //Set formula for the condition
            fc.setFormula1("=IF(SUM(B1:B2)>100,TRUE,FALSE)");

            //Set the style for conditional formatting rule
            fc.getStyle().setBackgroundColor(Color.getRed());

            //Set formula for the cell on which condition has been applied
            sheet.getCells().get("B3").setFormula("=SUM(B1:B2)");

            //Display tip
            sheet.getCells().get("C4").setValue("If Sum of B1:B2 is greater than 100, B3 will have RED background");

            //Save the result in spreadsheet format
            workbook.save(filePath + "ConditionalFormattingBasedOnFormula_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Apply Conditional Formatting Based on Cell Value", e);
        }
    }
}
