package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageConditionalFormattingAndIcons;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.CellArea;
import com.aspose.cells.Color;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ApplyShadingToAlternateRowsAndColumnsUsingConditionalFormatting {

    private static final String TAG = ApplyShadingToAlternateRowsAndColumnsUsingConditionalFormatting.class.getName();

    public void applyShadingToAlternateRowsAndColumnsUsingConditionalFormatting() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            //Optionally load an existing spreadsheet by passing its stream or path to Workbook constructor
            Workbook book = new Workbook();

            //Access the Worksheet on which desired rule has to be applied
            Worksheet sheet = book.getWorksheets().get(0);

            //Add FormatConditions to the instance of Worksheet
            int index = sheet.getConditionalFormattings().add();

            //Access the newly added FormatConditions via its index
            FormatConditionCollection conditionCollection = sheet.getConditionalFormattings().get(index);

            //Define a CellsArea on which conditional formatting will be applicable
            CellArea area = CellArea.createCellArea("A1", "I20");

            //Add area to the instance of FormatConditions
            conditionCollection.addArea(area);

            //Add a condition to the instance of FormatConditions
            //For this case, the condition type is expression, which is based on some formula
            index = conditionCollection.addCondition(FormatConditionType.EXPRESSION);

            //Access the newly added FormatCondition via its index
            FormatCondition formatCondirion = conditionCollection.get(index);

            //Set the formula for the FormatCondition
            //Formula uses the Excel's built-in functions as discussed earlier in this article
            formatCondirion.setFormula1("=MOD(ROW(),2)=0");

            //Set the background color and patter for the FormatCondition's Style
            formatCondirion.getStyle().setBackgroundColor(Color.getBlue());
            formatCondirion.getStyle().setPattern(BackgroundType.SOLID);

            //Save the result on disk
            book.save("ApplyShadingToAlternateRowsAndColumns_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Apply Shading to Alternate Rows and Columns with Conditional Formatting", e);
        }
    }
}
