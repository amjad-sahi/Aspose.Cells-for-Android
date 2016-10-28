package com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableAutoFormatType;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.SheetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

import static android.content.ContentValues.TAG;

public class CreatePivotTablesAndPivotCharts {

    public void createPivotTablesAndPivotCharts() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating an Workbook object
            Workbook workbook = new Workbook();
            //Obtaining the reference of the first worksheet
            Worksheet sheet = workbook.getWorksheets().get(0);
            //Name the sheet
            sheet.setName("Data");
            Cells cells = sheet.getCells();

            //Setting the values to the cells
            Cell cell = cells.get("A1");
            cell.setValue("Employee");
            cell = cells.get("B1");
            cell.setValue("Quarter");
            cell = cells.get("C1");
            cell.setValue("Product");
            cell = cells.get("D1");
            cell.setValue("Continent");
            cell = cells.get("E1");
            cell.setValue("Country");
            cell = cells.get("F1");
            cell.setValue("Sale");

            cell = cells.get("A2");
            cell.setValue("David");
            cell = cells.get("A3");
            cell.setValue("David");
            cell = cells.get("A4");
            cell.setValue("David");
            cell = cells.get("A5");
            cell.setValue("David");
            cell = cells.get("A6");
            cell.setValue("James");
            cell = cells.get("A7");
            cell.setValue("James");
            cell = cells.get("A8");
            cell.setValue("James");
            cell = cells.get("A9");
            cell.setValue("James");
            cell = cells.get("A10");
            cell.setValue("James");
            cell = cells.get("A11");
            cell.setValue("Miya");
            cell = cells.get("A12");
            cell.setValue("Miya");
            cell = cells.get("A13");
            cell.setValue("Miya");
            cell = cells.get("A14");
            cell.setValue("Miya");
            cell = cells.get("A15");
            cell.setValue("Miya");
            cell = cells.get("A16");
            cell.setValue("Miya");
            cell = cells.get("A17");
            cell.setValue("Miya");
            cell = cells.get("A18");
            cell.setValue("Elvis");
            cell = cells.get("A19");
            cell.setValue("Elvis");
            cell = cells.get("A20");
            cell.setValue("Elvis");
            cell = cells.get("A21");
            cell.setValue("Elvis");
            cell = cells.get("A22");
            cell.setValue("Elvis");
            cell = cells.get("A23");
            cell.setValue("Elvis");
            cell = cells.get("A24");
            cell.setValue("Elvis");
            cell = cells.get("A25");
            cell.setValue("Jean");
            cell = cells.get("A26");
            cell.setValue("Jean");
            cell = cells.get("A27");
            cell.setValue("Jean");
            cell = cells.get("A28");
            cell.setValue("Ada");
            cell = cells.get("A29");
            cell.setValue("Ada");
            cell = cells.get("A30");
            cell.setValue("Ada");

            cell = cells.get("B2");
            cell.setValue("1");
            cell = cells.get("B3");
            cell.setValue("2");
            cell = cells.get("B4");
            cell.setValue("3");
            cell = cells.get("B5");
            cell.setValue("4");
            cell = cells.get("B6");
            cell.setValue("1");
            cell = cells.get("B7");
            cell.setValue("2");
            cell = cells.get("B8");
            cell.setValue("3");
            cell = cells.get("B9");
            cell.setValue("4");
            cell = cells.get("B10");
            cell.setValue("4");
            cell = cells.get("B11");
            cell.setValue("1");
            cell = cells.get("B12");
            cell.setValue("1");
            cell = cells.get("B13");
            cell.setValue("2");
            cell = cells.get("B14");
            cell.setValue("2");
            cell = cells.get("B15");
            cell.setValue("3");
            cell = cells.get("B16");
            cell.setValue("4");
            cell = cells.get("B17");
            cell.setValue("4");
            cell = cells.get("B18");
            cell.setValue("1");
            cell = cells.get("B19");
            cell.setValue("1");
            cell = cells.get("B20");
            cell.setValue("2");
            cell = cells.get("B21");
            cell.setValue("3");
            cell = cells.get("B22");
            cell.setValue("3");
            cell = cells.get("B23");
            cell.setValue("4");
            cell = cells.get("B24");
            cell.setValue("4");
            cell = cells.get("B25");
            cell.setValue("1");
            cell = cells.get("B26");
            cell.setValue("2");
            cell = cells.get("B27");
            cell.setValue("3");
            cell = cells.get("B28");
            cell.setValue("1");
            cell = cells.get("B29");
            cell.setValue("2");
            cell = cells.get("B30");
            cell.setValue("3");

            cell = cells.get("C2");
            cell.setValue("Maxilaku");
            cell = cells.get("C3");
            cell.setValue("Maxilaku");
            cell = cells.get("C4");
            cell.setValue("Chai");
            cell = cells.get("C5");
            cell.setValue("Maxilaku");
            cell = cells.get("C6");
            cell.setValue("Chang");
            cell = cells.get("C7");
            cell.setValue("Chang");
            cell = cells.get("C8");
            cell.setValue("Chang");
            cell = cells.get("C9");
            cell.setValue("Chang");
            cell = cells.get("C10");
            cell.setValue("Chang");
            cell = cells.get("C11");
            cell.setValue("Geitost");
            cell = cells.get("C12");
            cell.setValue("Chai");
            cell = cells.get("C13");
            cell.setValue("Geitost");
            cell = cells.get("C14");
            cell.setValue("Geitost");
            cell = cells.get("C15");
            cell.setValue("Maxilaku");
            cell = cells.get("C16");
            cell.setValue("Geitost");
            cell = cells.get("C17");
            cell.setValue("Geitost");
            cell = cells.get("C18");
            cell.setValue("Ikuru");
            cell = cells.get("C19");
            cell.setValue("Ikuru");
            cell = cells.get("C20");
            cell.setValue("Ikuru");
            cell = cells.get("C21");
            cell.setValue("Ikuru");
            cell = cells.get("C22");
            cell.setValue("Ipoh Coffee");
            cell = cells.get("C23");
            cell.setValue("Ipoh Coffee");
            cell = cells.get("C24");
            cell.setValue("Ipoh Coffee");
            cell = cells.get("C25");
            cell.setValue("Chocolade");
            cell = cells.get("C26");
            cell.setValue("Chocolade");
            cell = cells.get("C27");
            cell.setValue("Chocolade");
            cell = cells.get("C28");
            cell.setValue("Chocolade");
            cell = cells.get("C29");
            cell.setValue("Chocolade");
            cell = cells.get("C30");
            cell.setValue("Chocolade");

            cell = cells.get("D2");
            cell.setValue("Asia");
            cell = cells.get("D3");
            cell.setValue("Asia");
            cell = cells.get("D4");
            cell.setValue("Asia");
            cell = cells.get("D5");
            cell.setValue("Asia");
            cell = cells.get("D6");
            cell.setValue("Europe");
            cell = cells.get("D7");
            cell.setValue("Europe");
            cell = cells.get("D8");
            cell.setValue("Europe");
            cell = cells.get("D9");
            cell.setValue("Europe");
            cell = cells.get("D10");
            cell.setValue("Europe");
            cell = cells.get("D11");
            cell.setValue("America");
            cell = cells.get("D12");
            cell.setValue("America");
            cell = cells.get("D13");
            cell.setValue("America");
            cell = cells.get("D14");
            cell.setValue("America");
            cell = cells.get("D15");
            cell.setValue("America");
            cell = cells.get("D16");
            cell.setValue("America");
            cell = cells.get("D17");
            cell.setValue("America");
            cell = cells.get("D18");
            cell.setValue("Europe");
            cell = cells.get("D19");
            cell.setValue("Europe");
            cell = cells.get("D20");
            cell.setValue("Europe");
            cell = cells.get("D21");
            cell.setValue("Oceania");
            cell = cells.get("D22");
            cell.setValue("Oceania");
            cell = cells.get("D23");
            cell.setValue("Oceania");
            cell = cells.get("D24");
            cell.setValue("Oceania");
            cell = cells.get("D25");
            cell.setValue("Africa");
            cell = cells.get("D26");
            cell.setValue("Africa");
            cell = cells.get("D27");
            cell.setValue("Africa");
            cell = cells.get("D28");
            cell.setValue("Africa");
            cell = cells.get("D29");
            cell.setValue("Africa");
            cell = cells.get("D30");
            cell.setValue("Africa");

            cell = cells.get("E2");
            cell.setValue("China");
            cell = cells.get("E3");
            cell.setValue("India");
            cell = cells.get("E4");
            cell.setValue("Korea");
            cell = cells.get("E5");
            cell.setValue("India");
            cell = cells.get("E6");
            cell.setValue("France");
            cell = cells.get("E7");
            cell.setValue("France");
            cell = cells.get("E8");
            cell.setValue("Germany");
            cell = cells.get("E9");
            cell.setValue("Italy");
            cell = cells.get("E10");
            cell.setValue("France");
            cell = cells.get("E11");
            cell.setValue("U.S.");
            cell = cells.get("E12");
            cell.setValue("U.S.");
            cell = cells.get("E13");
            cell.setValue("Brazil");
            cell = cells.get("E14");
            cell.setValue("U.S.");
            cell = cells.get("E15");
            cell.setValue("U.S.");
            cell = cells.get("E16");
            cell.setValue("Canada");
            cell = cells.get("E17");
            cell.setValue("U.S.");
            cell = cells.get("E18");
            cell.setValue("Italy");
            cell = cells.get("E19");
            cell.setValue("France");
            cell = cells.get("E20");
            cell.setValue("Italy");
            cell = cells.get("E21");
            cell.setValue("New Zealand");
            cell = cells.get("E22");
            cell.setValue("Australia");
            cell = cells.get("E23");
            cell.setValue("Australia");
            cell = cells.get("E24");
            cell.setValue("New Zealand");
            cell = cells.get("E25");
            cell.setValue("S.Africa");
            cell = cells.get("E26");
            cell.setValue("S.Africa");
            cell = cells.get("E27");
            cell.setValue("S.Africa");
            cell = cells.get("E28");
            cell.setValue("Egypt");
            cell = cells.get("E29");
            cell.setValue("Egypt");
            cell = cells.get("E30");
            cell.setValue("Egypt");

            cell = cells.get("F2");
            cell.setValue(2000);
            cell = cells.get("F3");
            cell.setValue(500);
            cell = cells.get("F4");
            cell.setValue(1200);
            cell = cells.get("F5");
            cell.setValue(1500);
            cell = cells.get("F6");
            cell.setValue(500);
            cell = cells.get("F7");
            cell.setValue(1500);
            cell = cells.get("F8");
            cell.setValue(800);
            cell = cells.get("F9");
            cell.setValue(900);
            cell = cells.get("F10");
            cell.setValue(500);
            cell = cells.get("F11");
            cell.setValue(1600);
            cell = cells.get("F12");
            cell.setValue(600);
            cell = cells.get("F13");
            cell.setValue(2000);
            cell = cells.get("F14");
            cell.setValue(500);
            cell = cells.get("F15");
            cell.setValue(900);
            cell = cells.get("F16");
            cell.setValue(700);
            cell = cells.get("F17");
            cell.setValue(1400);
            cell = cells.get("F18");
            cell.setValue(1350);
            cell = cells.get("F19");
            cell.setValue(300);
            cell = cells.get("F20");
            cell.setValue(500);
            cell = cells.get("F21");
            cell.setValue(1000);
            cell = cells.get("F22");
            cell.setValue(1500);
            cell = cells.get("F23");
            cell.setValue(1500);
            cell = cells.get("F24");
            cell.setValue(1600);
            cell = cells.get("F25");
            cell.setValue(1000);
            cell = cells.get("F26");
            cell.setValue(1200);
            cell = cells.get("F27");
            cell.setValue(1300);
            cell = cells.get("F28");
            cell.setValue(1500);
            cell = cells.get("F29");
            cell.setValue(1400);
            cell = cells.get("F30");
            cell.setValue(1000);



            // Creating Pivot Table
            //Adding a new sheet
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet sheet2 = workbook.getWorksheets().get(sheetIndex);
            //Naming the sheet
            sheet2.setName("PivotTable");
            //Getting the pivottables collection in the sheet
            PivotTableCollection pivotTables = sheet2.getPivotTables();
            //Adding a PivotTable to the worksheet
            int index = pivotTables.add("=Data!A1:F30", "B3", "PivotTable1");
            //Accessing the instance of the newly added PivotTable
            PivotTable pivotTable = pivotTables.get(index);
            //Showing the grand totals
            pivotTable.setRowGrand(true);
            pivotTable.setColumnGrand(true);
            //Setting the PivotTable report is automatically formatted
            pivotTable.setAutoFormat(true);
            //Setting the PivotTable autoformat type.
            pivotTable.setAutoFormatType(PivotTableAutoFormatType.REPORT_6);
            //Draging the first field to the row area.
            pivotTable.addFieldToArea(PivotFieldType.ROW, 0);
            //Draging the third field to the row area.
            pivotTable.addFieldToArea(PivotFieldType.ROW, 2);
            //Draging the second field to the row area.
            pivotTable.addFieldToArea(PivotFieldType.ROW, 1);
            //Draging the fourth field to the column area.
            pivotTable.addFieldToArea(PivotFieldType.COLUMN, 3);
            //Draging the fifth field to the data area.
            pivotTable.addFieldToArea(PivotFieldType.DATA, 5);
            //Setting the number format of the first data field
            pivotTable.getDataFields().get(0).setNumber(7);




            // Creating a Pivot Chart based on the Pivot Table
            //Adding a new sheet
            sheetIndex = workbook.getWorksheets().add(SheetType.CHART);
            Worksheet sheet3 = workbook.getWorksheets().get(sheetIndex);
            //Naming the sheet
            sheet3.setName("PivotChart");
            //Adding a column chart
            int chartIndex = sheet3.getCharts().add(ChartType.COLUMN, 0, 5, 28, 16);
            Chart chart = sheet3.getCharts().get(chartIndex);
            //Setting the pivot chart data source
            chart.setPivotSource("PivotTable!PivotTable1");
            chart.setHidePivotFieldButtons(false);

            //Saving the Excel file
            workbook.save(filePath + "CreatePivotTablesAndPivotCharts_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Create a Spreadsheet with Data", e);
        }
    }

}
