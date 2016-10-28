package com.aspose.cells.examples.asposecellsexamples.DataFormatting.AdvanceFormatting;

import android.util.Log;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.CellArea;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.ConditionalFormattingCollection;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ConditionalFormatting {

    private static final String TAG = ConditionalFormatting.class.getName();

    /**
     * Add and Delete Conditional Formatting.
     */
    public void applyConditionalFormatting() {
        try {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.getWorksheets().get(0);
            ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();

            // Adds an empty conditional formatting
            int index = cfs.add();
            FormatConditionCollection fcs = cfs.get(index);

            //Sets the conditional format range.
            CellArea ca1 = new CellArea();
            ca1.StartRow = 0;
            ca1.StartColumn = 0;
            ca1.EndRow = 0;
            ca1.EndColumn = 0;

            CellArea ca2 = new CellArea();
            ca2.StartRow = 0;
            ca2.StartColumn = 0;
            ca2.EndRow = 0;
            ca2.EndColumn = 0;

            CellArea ca3 = new CellArea();
            ca3.StartRow = 0;
            ca3.StartColumn = 0;
            ca3.EndRow = 0;
            ca3.EndColumn = 0;

            fcs.addArea(ca1);
            fcs.addArea(ca2);
            fcs.addArea(ca3);

            //Sets condition formulas.
            int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN,"=A2","100");

            FormatCondition fc = fcs.get(conditionIndex);

            int conditionIndex2 = fcs.addCondition(FormatConditionType.CELL_VALUE,OperatorType.BETWEEN,"50","100");
        } catch(Exception e) {
            Log.e(TAG, "Apply Conditional Formatting", e);
        }
    }

    /**
     * Set Font
     */
    public void setFont() {
        try {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.getWorksheets().get(0);
            ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();

            // Adds an empty conditional formatting
            int index = cfs.add();
            FormatConditionCollection fcs = cfs.get(index);

            //Sets condition formulas.
            int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN,"=A2","100");

            FormatCondition fc = fcs.get(conditionIndex);

            Style style = fc.getStyle();
            Font font = style.getFont();
            font.setItalic(true);
            font.setBold(true);
            font.setStrikeout(true);
            font.setUnderline(FontUnderlineType.DOUBLE);
            font.setColor(Color.getBlack());
            fc.setStyle(style);

        } catch(Exception e) {
            Log.e(TAG, "Set Font", e);
        }
    }

    /**
     * Set Border
     */
    public void setBorder() {
        try {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.getWorksheets().get(0);
            ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();

            // Adds an empty conditional formatting
            int index = cfs.add();
            FormatConditionCollection fcs = cfs.get(index);

            //Sets condition formulas.
            int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN,"=A2","100");

            FormatCondition fc = fcs.get(conditionIndex);

            Style style = fc.getStyle();
            style.setBorder(BorderType.LEFT_BORDER, CellBorderType.DASHED,Color.fromArgb(0, 255, 255));
            style.setBorder(BorderType.TOP_BORDER,CellBorderType.DASHED,Color.fromArgb(0, 255, 255));
            style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.DASHED,Color.fromArgb(0, 255, 255));
            style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.DASHED,Color.fromArgb(255, 255, 0));
            fc.setStyle(style);

        } catch(Exception e) {
            Log.e(TAG, "Set Border", e);
        }
    }

    /**
     * Set Pattern
     */
    public void setPattern() {
        try {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.getWorksheets().get(0);
            ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();

            // Adds an empty conditional formatting
            int index = cfs.add();
            FormatConditionCollection fcs = cfs.get(index);

            //Sets condition formulas.
            int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN,"=A2","100");

            FormatCondition fc = fcs.get(conditionIndex);

            Style style = fc.getStyle();
            style.setPattern(BackgroundType.REVERSE_DIAGONAL_STRIPE);
            style.setForegroundColor(Color.fromArgb(255,255,0));
            style.setBackgroundColor(Color.fromArgb(0,255,255));
            fc.setStyle(style);

        } catch(Exception e) {
            Log.e(TAG, "Set Pattern", e);
        }
    }
}
