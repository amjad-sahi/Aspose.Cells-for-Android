package com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Color;
import com.aspose.cells.PivotField;
import com.aspose.cells.PivotFieldCollection;
import com.aspose.cells.PivotFieldDataDisplayFormat;
import com.aspose.cells.PivotFieldSubtotalType;
import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotItemPosition;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableAutoFormatType;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.PivotTableStyleType;
import com.aspose.cells.Style;
import com.aspose.cells.TableStyle;
import com.aspose.cells.TableStyleElement;
import com.aspose.cells.TableStyleElementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CustomizeTheAppearanceOfPivotTable {

    private static final String TAG = CustomizeTheAppearanceOfPivotTable.class.getName();

    /**
     * This example illustrates how to set the auto format type and the pivot table style using the AutoFormat and PivotTableStyle properties.
     */
    public void setTheAutoFormatAndPivotTableStyleTypes(PivotTable pivotTable) {
        try {
            //Setting the PivotTable report is automatically formatted for Excel 2003 formats
            pivotTable.setAutoFormat(true);
            //Setting the PivotTable atuoformat type.
            pivotTable.setAutoFormatType(PivotTableAutoFormatType.CLASSIC);

            //Setting the PivotTable's Styles for Excel 2007/2010 formats e.g XLSX.
            pivotTable.setPivotTableStyleType(PivotTableStyleType.PIVOT_TABLE_STYLE_LIGHT_1);
        } catch (Exception e) {
            Log.e(TAG, "Set the AutoFormat and PivotTableStyle Types", e);
        }
    }

    /**
     * This example shows how to access row fields, access a particular row, set subtotals, apply automatic sorting, and using the autoShow option.
     */
    public void setRowColumnAndPageFieldFormat(PivotTable pivotTable) {
        try {
            //Accessing the row fields.
            PivotFieldCollection pivotFields = pivotTable.getRowFields();

            //Accessing the first row field in the row fields.
            PivotField pivotField = pivotFields.get(0);

            //Setting Subtotals.
            pivotField.setSubtotals(PivotFieldSubtotalType.SUM, true);
            pivotField.setSubtotals(PivotFieldSubtotalType.COUNT, true);

            //Setting autosort options.
            //Setting the field auto sort.
            pivotField.setAutoSort(true);

            //Setting the field auto sort ascend.
            pivotField.setAscendSort(true);

            //Setting the field auto sort using the field itself.
            pivotField.setAutoSortField(-1);

            //Setting autoShow options.
            //Setting the field auto show.
            pivotField.setAutoShow(true);

            //Setting the field auto show ascend.
            pivotField.setAscendShow(false);

            //Setting the auto show using field(data field).
            pivotField.setAutoShowField(0);
        } catch (Exception e) {
            Log.e(TAG, "Set Row, Column and Page Fields Format", e);
        }
    }

    /**
     * The example illustrate how to format data fields.
     */
    public void setDataFieldsFormat(PivotTable pivotTable) {
        try {
            //Accessing the data fields.
            PivotFieldCollection pivotFields = pivotTable.getDataFields();

            //Accessing the first data field in the data fields.
            PivotField pivotField = pivotFields.get(0);

            //Setting data display format
            pivotField.setDataDisplayFormat(PivotFieldDataDisplayFormat.PERCENTAGE_OF);

            //Setting the base field.
            pivotField.setBaseField(1);

            //Setting the base item.
            pivotField.setBaseItem(PivotItemPosition.NEXT);

            //Setting number format
            pivotField.setNumber(10);
        } catch (Exception e) {
            Log.e(TAG, "Set Data Fields Format", e);
        }
    }

    /**
     * This examples shows how to modify the quick style applied to a pivot table.
     */
    public void modifyAPivotTableQuickStyle(PivotTable pivotTable) {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Open the template file containing the pivot table.
            Workbook wb = new Workbook(filePath + File.separator + "Template.xlsx");

            //Add pivot table style
            Style style1 = wb.createStyle();
            com.aspose.cells.Font font1 = style1.getFont();
            font1.setColor(Color.getRed());
            Style style2 = wb.createStyle();
            com.aspose.cells.Font font2 = style2.getFont();
            font2.setColor(Color.getBlue());
            int i = wb.getWorksheets().getTableStyles().addPivotTableStyle("tt");

            //Get and Set the table style for different categories
            TableStyle ts = wb.getWorksheets().getTableStyles().get(i);
            int index = ts.getTableStyleElements().add(TableStyleElementType.FIRST_COLUMN);
            TableStyleElement e = ts.getTableStyleElements().get(index);
            e.setElementStyle(style1);
            index = ts.getTableStyleElements().add(TableStyleElementType.GRAND_TOTAL_ROW);
            e = ts.getTableStyleElements().get(index);
            e.setElementStyle(style2);

            //Set Pivot Table style name
            PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);
            pt.setPivotTableStyleName("tt");

            //Save the file.
            wb.save(filePath + File.separator + "ModifyAPivotTableQuickStyle_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Modify a Pivot Table Quick Style", e);
        }
    }

    /**
     * PivotFieldCollection has a method named clear() for the task.
     * When you want to clear all the PivotFields in the areas e.g., page, column, row or data, you can use it.
     * This example shows how to clear all the PivotFields in data area.
     */
    public PivotTable clearPivotFields() {
        try {
            File sdDir = Environment.getExternalStorageDirectory();
            String sdPath = sdDir.getCanonicalPath();

            //Open the template file containing the pivot table.
            Workbook workbook = new Workbook(sdPath + "pivot.xlsm");

            //Get the first worksheet
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Get the pivot tables in the sheet
            PivotTableCollection pivotTables = sheet.getPivotTables();

            //Get the first PivotTable
            PivotTable pivotTable = pivotTables.get(0);

            //Clear all the data fields
            pivotTable.getDataFields().clear();

            //Add new data field
            pivotTable.addFieldToArea(PivotFieldType.DATA, "Betrag Netto FW");

            //Set the refresh data flag on
            pivotTable.setRefreshDataFlag(false);

            //Refresh and calculate the pivot table data
            pivotTable.refreshData();
            pivotTable.calculateData();

            //Save the Excel file
            workbook.save(sdPath + "ClearPivotFields_Out.xlsx");

            return pivotTable;
        } catch (Exception e) {
            Log.e(TAG, "Clear PivotFields", e);
            return null;
        }
    }
}
