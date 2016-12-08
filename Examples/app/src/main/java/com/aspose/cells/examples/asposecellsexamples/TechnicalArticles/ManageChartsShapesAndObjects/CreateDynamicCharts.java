package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.ComboBox;
import com.aspose.cells.ListObjectCollection;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CreateDynamicCharts {

    private static final String TAG = CreateDynamicCharts.class.getName();

    public void usingExcelTables() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook();
            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);
            //Access cells collection of the first worksheet
            Cells cells = sheet.getCells();

            //Insert data column wise
            cells.get("A1").putValue("Category");
            cells.get("A2").putValue("Fruit");
            cells.get("A3").putValue("Fruit");
            cells.get("A4").putValue("Fruit");
            cells.get("A5").putValue("Fruit");
            cells.get("A6").putValue("Vegetables");
            cells.get("A7").putValue("Vegetables");
            cells.get("A8").putValue("Vegetables");
            cells.get("A9").putValue("Vegetables");
            cells.get("A10").putValue("Beverages");
            cells.get("A11").putValue("Beverages");
            cells.get("A12").putValue("Beverages");

            cells.get("B1").putValue("Food");
            cells.get("B2").putValue("Apple");
            cells.get("B3").putValue("Banana");
            cells.get("B4").putValue("Apricot");
            cells.get("B5").putValue("Grapes");
            cells.get("B6").putValue("Carrot");
            cells.get("B7").putValue("Onion");
            cells.get("B8").putValue("Cabage");
            cells.get("B9").putValue("Potatoe");
            cells.get("B10").putValue("Coke");
            cells.get("B11").putValue("Coladas");
            cells.get("B12").putValue("Fizz");

            cells.get("C1").putValue("Cost");
            cells.get("C2").putValue(2.2);
            cells.get("C3").putValue(3.1);
            cells.get("C4").putValue(4.1);
            cells.get("C5").putValue(5.1);
            cells.get("C6").putValue(4.4);
            cells.get("C7").putValue(5.4);
            cells.get("C8").putValue(6.5);
            cells.get("C9").putValue(5.3);
            cells.get("C10").putValue(3.2);
            cells.get("C11").putValue(3.6);
            cells.get("C12").putValue(5.2);

            cells.get("D1").putValue("Profit");
            cells.get("D2").putValue(0.1);
            cells.get("D3").putValue(0.4);
            cells.get("D4").putValue(0.5);
            cells.get("D5").putValue(0.6);
            cells.get("D6").putValue(0.7);
            cells.get("D7").putValue(1.3);
            cells.get("D8").putValue(0.8);
            cells.get("D9").putValue(1.3);
            cells.get("D10").putValue(2.2);
            cells.get("D11").putValue(2.4);
            cells.get("D12").putValue(3.3);

            //Create ListObject
            //Get the List objects collection in the first worksheet
            ListObjectCollection listObjects = sheet.getListObjects();

            //Add a List based on the data source range with headers on
            int index = listObjects.add(0, 0, 11, 3, true);

            sheet.autoFitColumns();

            //Create chart based on ListObject
            index = sheet.getCharts().add(ChartType.COLUMN, 21, 1, 35, 18);
            Chart chart = sheet.getCharts().get(index);
            chart.setChartDataRange("A1:D12", true);
            chart.getNSeries().setCategoryData("A2:B12");

            //Calculate chart
            chart.calculate();

            //Save spreadsheet
            book.save(filePath + "ChartBasedOnExcelTables_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Using Excel Tables", e);
        }
    }

    public void usingDynamicFormulas() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook object
            Workbook workbook = new Workbook();

            //Get the first worksheet
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Access cells collection of first worksheet
            Cells cells = sheet.getCells();

            //Create a range in the second worksheet
            Range range = cells.createRange("C21", "C24");

            //Name the range
            range.setName("MyRange");

            //Fill different cells with data in the range
            range.get(0, 0).putValue("North");
            range.get(1, 0).putValue("South");
            range.get(2, 0).putValue("East");
            range.get(3, 0).putValue("West");

            ComboBox comboBox = (ComboBox) sheet.getShapes().addShape(MsoDrawingType.COMBO_BOX, 15, 0, 2, 0, 17, 64);
            comboBox.setInputRange("=MyRange");
            comboBox.setLinkedCell("=B16");
            comboBox.setSelectedIndex(0);
            Cell cell = cells.get("B16");
            Style style = cell.getStyle();
            style.getFont().setColor(Color.getWhite());
            cell.setStyle(style);

            cells.get("C16").setFormula("=INDEX(Sheet1!$C$21:$C$24,$B$16,1)");

            //Put some data for chart source
            //Data Headers
            cells.get("D15").putValue("Jan");
            cells.get("D20").putValue("Jan");

            cells.get("E15").putValue("Feb");
            cells.get("E20").putValue("Feb");

            cells.get("F15").putValue("Mar");
            cells.get("F20").putValue("Mar");

            cells.get("G15").putValue("Apr");
            cells.get("G20").putValue("Apr");

            cells.get("H15").putValue("May");
            cells.get("H20").putValue("May");

            cells.get("I15").putValue("Jun");
            cells.get("I20").putValue("Jun");

            //Data
            cells.get("D21").putValue(304);
            cells.get("D22").putValue(402);
            cells.get("D23").putValue(321);
            cells.get("D24").putValue(123);

            cells.get("E21").putValue(300);
            cells.get("E22").putValue(500);
            cells.get("E23").putValue(219);
            cells.get("E24").putValue(422);

            cells.get("F21").putValue(222);
            cells.get("F22").putValue(331);
            cells.get("F23").putValue(112);
            cells.get("F24").putValue(350);

            cells.get("G21").putValue(100);
            cells.get("G22").putValue(200);
            cells.get("G23").putValue(300);
            cells.get("G24").putValue(400);

            cells.get("H21").putValue(200);
            cells.get("H22").putValue(300);
            cells.get("H23").putValue(400);
            cells.get("H24").putValue(500);

            cells.get("I21").putValue(400);
            cells.get("I22").putValue(200);
            cells.get("I23").putValue(200);
            cells.get("I24").putValue(100);

            //Dynamically load data on selection of Dropdown value
            cells.get("D16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,2,FALSE),0)");
            cells.get("E16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,3,FALSE),0)");
            cells.get("F16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,4,FALSE),0)");
            cells.get("G16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,5,FALSE),0)");
            cells.get("H16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,6,FALSE),0)");
            cells.get("I16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,7,FALSE),0)");

            //Create Chart
            int index = sheet.getCharts().add(ChartType.COLUMN, 0, 3, 12, 9);
            Chart chart = sheet.getCharts().get(index);
            chart.getNSeries().add("='Sheet1'!$D$16:$I$16", false);
            chart.getNSeries().get(0).setName("=C16");
            chart.getNSeries().setCategoryData("=$D$15:$I$15");

            //Save result on disc
            workbook.save(filePath + "ChartUsingDynamicFormulas_Out.xlsx");

        } catch (Exception e) {
            Log.e(TAG, "Using Dynamic Formulas", e);
        }
    }
}
