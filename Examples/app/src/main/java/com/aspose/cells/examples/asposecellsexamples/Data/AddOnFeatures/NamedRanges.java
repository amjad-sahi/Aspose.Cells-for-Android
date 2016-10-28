package com.aspose.cells.examples.asposecellsexamples.Data.AddOnFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Name;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;
import java.util.Iterator;

public class NamedRanges {

    private static final String TAG = NamedRanges.class.getName();

    public void createANamedRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            WorksheetCollection worksheets = workbook.getWorksheets();

            //Accessing the first worksheet in the Excel file
            Worksheet sheet = worksheets.get(0);
            Cells cells = sheet.getCells();

            //Creating a named range
            Range namedRange = cells.createRange("B4", "G14");
            namedRange.setName("TestRange");

            //Saving the modified Excel file in default (that is Excel 2000) format
            workbook.save(filePath + File.separator + "CreateANamedRange_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Create a Named Range", e);
        }
    }

    public void accessAllNamedRangesInAFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            WorksheetCollection worksheets = workbook.getWorksheets();

            //Getting all named ranges
            Range[] namedRanges = worksheets.getNamedRanges();

        } catch (Exception e) {
            Log.e(TAG, "Accessing All Named Ranges in a File", e);
        }
    }

    public void accessASpecificNamedRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            WorksheetCollection worksheets = workbook.getWorksheets();

            //Getting the specified named range
            Range namedRange = worksheets.getRangeByName("TestRange");

        } catch (Exception e) {
            Log.e(TAG, "Accessing a Specific Named Range", e);
        }
    }

    public void inputDataIntoANamedRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the workbook.
            Worksheet worksheet1 = workbook.getWorksheets().get(0);

            //Create a range of cells and specify its name based on H1:J4.
            Range range = worksheet1.getCells().createRange(0, 7, 3, 9);
            range.setName("MyRange");

            //Input some data into cells in the range.
            range.get(0, 0).setValue("USA");
            range.get(0, 1).setValue("SA");
            range.get(0, 2).setValue("Israel");
            range.get(1, 0).setValue("UK");
            range.get(1, 1).setValue("AUS");
            range.get(1, 2).setValue("Canada");
            range.get(2, 0).setValue("France");
            range.get(2, 1).setValue("India");
            range.get(2, 2).setValue("Egypt");
            range.get(3, 0).setValue("China");
            range.get(3, 1).setValue("Philipine");
            range.get(3, 2).setValue("Brazil");

            //Identify range cells.
            int firstrow = range.getFirstRow();
            int firstcol = range.getFirstColumn();
            int trows = range.getRowCount();
            int tcols = range.getColumnCount();

            //Save the excel file.
            workbook.save(filePath + File.separator + "RangeCells_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Input Data into a Named Range", e);
        }
    }

    public void setBackgroundColorAndFontAttributes() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the book.
            Worksheet WS = workbook.getWorksheets().get(0);

            //Create a named range of cells.
            com.aspose.cells.Range range = WS.getCells().createRange(1, 1, 1, 17);
            range.setName("MyRange");

            //Declare a style object.
            Style stl;

            //Create the style object with respect to the style of a cell.
            stl = WS.getCells().getCell(1, 1).getStyle();

            //Specify some Font settings.
            stl.getFont().setName("Arial");
            stl.getFont().setBold(true);

            //Set the font text color
            stl.getFont().setColor(Color.getRed());

            //To Set the fill color of the range, you may use ForegroundColor with
            //solid Pattern setting.
            stl.setBackgroundColor(Color.getYellow());
            stl.setPattern(BackgroundType.SOLID);

            //Apply the style to the range.
            for (int r = 1; r < 2; r++) {
                for (int c = 1; c < 18; c++) {
                    WS.getCells().getCell(r, c).setStyle(stl);
                }

            }

            //Save the Excel file.
            workbook.save(filePath + File.separator + "BackgroundColorAndFontAttributes_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Set Background Color and Font Attributes", e);
        }
    }

    public void addBordersToANamedRange() {
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
            Range range = worksheet.getCells().createRange(0,0,0,2);
            range.setName("MyRange");

            //Adding a thick outline border with the blue line
            range.setOutlineBorders(CellBorderType.THICK, Color.getBlue());

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddBordersToANamedRange_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Add Borders to a Named Range", e);
        }
    }

    public void convertCellsAddressToRangeOrCellArea() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the newly added worksheet
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Accessing the "A1" cell from the worksheet
            Cell cell = worksheet.getCells().get("A1");

            //Adding some value to the "A1" cell
            cell.setValue("Hello World!");

            //Creating a range of cells based on cells Address.
            Range range = worksheet.getCells().createRange("A1:F10");

            //Specify a Style object for borders.
            Style style = cell.getStyle();
            //Setting the line style of the top border

            style.setBorder(BorderType.TOP_BORDER, CellBorderType.THICK, Color.getBlack());

            style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK, Color.getBlack());

            style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK, Color.getBlack());

            style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack());

            Iterator cellArray = range.iterator();
            while(cellArray.hasNext())
            {
                Cell temp = (Cell)cellArray.next();
                //Saving the modified style to the cell.
                temp.setStyle(style);
            }

            //Saving the Excel file
            workbook.save(filePath + File.separator + "ConvertCellsAddressToRangeOrCellArea_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Convert Cells Address to Range or CellArea", e);
        }
    }

    public void setASimpleFormulaForNamedRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Get the WorksheetCollection
            WorksheetCollection worksheets = book.getWorksheets();

            //Add a new Named Range with name "myName"
            int index = worksheets.getNames().add("myName");

            //Access the newly created Named Range
            Name name = worksheets.getNames().get(index);

            //Set RefersTo property of the Named Range to a formula
            //Formula references another cell in the same worksheet
            name.setRefersTo("=Sheet1!$A$3");

            //Set the formula in the cell A1 to the newly created Named Range
            worksheets.get(0).getCells().get("A1").setFormula("myName");

            //Insert the value in cell A3 which is being referenced in the Named Range
            worksheets.get(0).getCells().get("A3").putValue("This is the value of A3");

            //Calculate formulas
            book.calculateFormula();

            //Save the result in XLSX format
            book.save(filePath + File.separator + "SetASimpleFormulaForNamedRange_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Set a Simple Formula for Named Range", e);
        }
    }

    public void setAComplexFormulaForNamedRange() {
        try {
            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Get the WorksheetCollection
            WorksheetCollection worksheets = book.getWorksheets();

            //Add a new Named Range with name "data"
            int index = worksheets.getNames().add("data");

            //Access the newly created Named Range from the collection
            Name data = worksheets.getNames().get(index);

            //Set RefersTo property of the Named Range to a cell range in same worksheet
            data.setRefersTo("=Sheet1!$A$1:$A$10");

            //Add another Named Range with name "range"
            index = worksheets.getNames().add("range");

            //Access the newly created Named Range from the collection
            Name range = worksheets.getNames().get(index);

            //Set RefersTo property to a formula using the Named Range data
            range.setRefersTo("=INDEX(data,Sheet1!$A$1,1):INDEX(data,Sheet1!$A$1,9)");
        } catch (Exception e) {
            Log.e(TAG, "Set a Complex Formula for Named Range", e);
        }
    }

    public void useANamedRangeToSumValuesFrom2CellsInDifferentWorksheets() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Get the WorksheetCollection
            WorksheetCollection worksheets = book.getWorksheets();

            //Insert some data in cell A1 of Sheet1
            worksheets.get("Sheet1").getCells().get("A1").putValue(10);

            //Add a new Worksheet and insert a value to cell A1
            worksheets.get(worksheets.add()).getCells().get("A1").putValue(10);

            //Add a new Named Range with name "range"
            int index = worksheets.getNames().add("range");

            //Access the newly created Named Range from the collection
            Name range = worksheets.getNames().get(index);

            //Set RefersTo property of the Named Range to a SUM function
            range.setRefersTo("=SUM(Sheet1!$A$1,Sheet2!$A$1)");

            //Insert the Named Range as formula to 3rd worksheet
            worksheets.get(worksheets.add()).getCells().get("A1").setFormula("range");

            //Calculate formulas
            book.calculateFormula();

            //Save the result in XLSX format
            book.save(filePath + File.separator + "SumValueFrom2Cells_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Use a named range to sum values from 2 cells in different worksheets", e);
        }
    }

}
