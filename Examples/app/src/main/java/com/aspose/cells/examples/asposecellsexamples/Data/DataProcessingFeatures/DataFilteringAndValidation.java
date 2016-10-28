package com.aspose.cells.examples.asposecellsexamples.Data.DataProcessingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.AutoFilter;
import com.aspose.cells.CellArea;
import com.aspose.cells.Cells;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Range;
import com.aspose.cells.Validation;
import com.aspose.cells.ValidationAlertType;
import com.aspose.cells.ValidationCollection;
import com.aspose.cells.ValidationType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class DataFilteringAndValidation {

    private static final String TAG = DataFilteringAndValidation.class.getName();

    public void autofilter() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Creating AutoFilter by giving the cells range
            AutoFilter autoFilter = worksheet.getAutoFilter();
            CellArea area = new CellArea();
            autoFilter.setRange("A1:B4");

            //Saving the modified Excel file
            workbook.save(filePath + File.separator + "Autofilter_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Auto Filter", e);
        }
    }

    public void  filterColumnsWithSpecifiedValues() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Creating AutoFilter by giving the cells range
            AutoFilter autoFilter = worksheet.getAutoFilter();
            autoFilter.setRange("A1:B4");

            //Filtering columns with specified values
            autoFilter.filter(1, "Bananas");

            //Saving the modified Excel file
            workbook.save(filePath + File.separator + "FilterColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Filter Columns with Specified Values", e);
        }
    }

    public void autoFilterOptions() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Obtaining the auto-filters in the worksheet
            AutoFilter autoFilter = worksheet.getAutoFilter();
            autoFilter.addFilter(0, "6");
            autoFilter.addFilter(0, "7");
            autoFilter.addFilter(0, "10");

            //Saving the modified Excel file
            workbook.save(filePath + File.separator + "AutoFilterOptions_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Advanced Auto-Filter Options", e);
        }
    }

    public void wholeNumberDataValidation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Accessing the Validations collection of the worksheet
            Worksheet worksheet = worksheets.get(0);
            ValidationCollection validations = worksheet.getValidations();

            //Creating the CellArea on which validation has to be applied
            CellArea area = new CellArea();
            area.StartRow = 0;
            area.StartColumn = 0;
            area.EndRow = 1;
            area.EndColumn = 1;

            //Creating a Validation object
            int index = validations.add(area);
            Validation validation = validations.get(index);

            //Setting the validation type to whole number
            validation.setType(ValidationType.WHOLE_NUMBER);

            //Setting the operator for validation to Between
            validation.setOperator(OperatorType.BETWEEN);

            //Setting the minimum value for the validation
            validation.setFormula1("10");

            //Setting the maximum value for the validation
            validation.setFormula2 ("1000");

            //Adding the cell area to Validation
            validation.addArea(area);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "WholeNumberDataValidation_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Whole Number Data Validation", e);
        }
    }

    public void decimalDataValidation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an instance of Workbook
            Workbook workbook = new Workbook();

            //Accessing the first worksheet in the collection
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Accessing the worksheet's validation collection
            ValidationCollection validations = worksheet.getValidations();

            //Creating the CellArea structure on which validation has to be applied
            CellArea area= new CellArea();
            area.StartRow = 0;
            area.EndRow = 9;
            area.StartColumn = 0;
            area.EndColumn = 0;

            //Creating a validation object by adding it to the collection
            int index = validations.add(area);
            Validation validation = validations.get(index);

            //Setting the validation type to Decimal
            validation.setType(ValidationType.DECIMAL);

            //Specifying the operator type
            validation.setOperator(OperatorType.BETWEEN);

            //Setting the lower & upper limits
            validation.setFormula1(new Double(Double.MIN_VALUE).toString());
            validation.setFormula2(new Double(Double.MAX_VALUE).toString());

            //Setting the error message
            validation.setErrorMessage("Please enter a valid integer or decimal number");

            //Setting the number format to 2 decimal places for the validation area
            for (int i = 0; i < 10; i++) {
                worksheet.getCells().get(i, 0).getStyle().setCustom("0.00");
            }

            workbook.save(filePath + File.separator + "DecimalDataValidation_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Decimal Data Validation", e);
        }
    }

    public void listDataValidation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an instance of Workbook
            Workbook workbook = new Workbook();

            //Accessing the first worksheet from collection
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Creating a range in the worksheet
            Range range = worksheet.getCells().createRange("E1", "E4");
            range.setName("MyRange");

            //Filling different cells with data in the range
            range.get(0,0).setValue("Blue");
            range.get(1,0).setValue("Red");
            range.get(2,0).setValue("Green");
            range.get(3,0).setValue("Yellow");

            //Getting the validations collection
            ValidationCollection validations = worksheet.getValidations();

            //Specifying the validation area
            CellArea area = new CellArea();
            area.StartRow = 0;
            area.StartColumn = 0;
            area.EndRow = 4;
            area.EndColumn = 0;

            //Creating a new validation to the validation collection
            int index = validations.add(area);
            Validation validation = validations.get(index);

            //Setting the validation type
            validation.setType(ValidationType.LIST);

            //Setting the in cell drop down
            validation.setInCellDropDown(true);

            //Setting the formula
            validation.setFormula1("=MyRange");

            //Enabling it to show error
            validation.setShowError(true);

            //Setting the alert type severity level
            validation.setAlertStyle(ValidationAlertType.STOP);

            //Setting the error title
            validation.setErrorTitle("Error");

            //Setting the error message
            validation.setErrorMessage("Please select a color from the list");

            workbook.save(filePath + File.separator + "ListDataValidation_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "List Data Validation", e);
        }
    }

    public void dateDataValidation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an instance of Workbook
            Workbook workbook = new Workbook();

            //Accessing the cells of the first worksheet
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Putting a string value into the A1 cell
            cells.get("A1").putValue("Please enter Date between 1/1/1970 and 12/31/1999");

            //Wrapping the text
            cells.get("A1").getStyle().setTextWrapped(true);

            //Setting row height and column width for the cells
            cells.setRowHeight(0, 31);
            cells.setColumnWidth(0, 35);

            //Getting the validations collection
            ValidationCollection validations = workbook.getWorksheets().get(0).getValidations();

            //Setting the CellArea on which the data validation rule will be applied
            CellArea area = CellArea.createCellArea(0, 1, 0, 1);

            //Adding a new validation to the collection
            int index = validations.add(area);
            Validation validation = validations.get(index);

            //Setting the data validation type
            validation.setType(ValidationType.DATE);

            //Setting the operator for the data validation
            validation.setOperator(OperatorType.BETWEEN);

            //Setting the value or expression associated with the data validation
            validation.setFormula1("1/1/1970");

            //Setting value or expression associated with the second part of the data validation
            validation.setFormula2("12/31/1999");

            //Enabling the error
            validation.setShowError(true);

            //Setting the validation alert style
            validation.setAlertStyle(ValidationAlertType.STOP);

            //Setting the title of the data validation error dialog box
            validation.setErrorTitle("Date Error");

            //Setting the data validation error message
            validation.setErrorMessage("Enter a Valid Date");

            //Setting and enabling the data validation input message
            validation.setInputMessage("Date Validation Type");
            validation.setIgnoreBlank(true);
            validation.setShowInput(true);

            workbook.save(filePath + File.separator + "DateDataValidation_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Date Data Validation", e);
        }
    }

    public void timeDataValidation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an instance of Workbook
            Workbook workbook = new Workbook();

            //Accessing the cells of the first worksheet
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Putting a string value into the A1 cell
            cells.get("A1").setValue("Please enter Time b/w 09:00 and 11:30 'o Clock");

            //Wrapping the text
            cells.get("A1").getStyle().setTextWrapped(true);

            //Setting row height and column width for the cells
            cells.setRowHeight(0, 31);
            cells.setColumnWidth(0, 35);

            //Getting the validations collection
            ValidationCollection validations = workbook.getWorksheets().get(0).getValidations();

            //Setting the CellArea on which the data validation rule will be applied
            CellArea area = CellArea.createCellArea(0, 1, 0, 1);

            //Adding a new validation to the collection
            int index = validations.add(area);
            Validation validation = validations.get(index);

            //Setting the data validation type
            validation.setType(ValidationType.TIME);

            //Setting the operator for the data validation
            validation.setOperator(OperatorType.BETWEEN);

            //Setting the value or expression associated with the data validation
            validation.setFormula1("09:00");

            //Setting value or expression associated with the second part of the data validation
            validation.setFormula2("11:30");

            //Enabling the error
            validation.setShowError(true);

            //Setting the validation alert style
            validation.setAlertStyle(ValidationAlertType.STOP);

            //Setting the title of the data validation error dialog box
            validation.setErrorTitle("Time Error");

            //Setting the data validation error message
            validation.setErrorMessage("Enter a Valid Time");

            //Setting and enabling the data validation input message
            validation.setInputMessage("Time Validation Type");
            validation.setIgnoreBlank(true);
            validation.setShowInput(true);

            workbook.save(filePath + File.separator + "TimeDataValidation_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Time Data Validation", e);
        }
    }

    public void textLengthDataValidation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an instance of Workbook
            Workbook workbook = new Workbook();

            //Accessing the cells of default worksheet
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Inserting the help text in cell A1
            cells.get("A1").putValue("Please enter a string not more than 5 chars");

            //Wrapping the text of cell A1
            cells.get("A1").getStyle().setTextWrapped(true);

            //Setting row height & column width
            cells.setRowHeight(0, 31);
            cells.setColumnWidth(0, 35);

            //Accessing the validation collection of first worksheet
            ValidationCollection validations = workbook.getWorksheets().get(0).getValidations();

            //Creating cell area on which validation has to be applied
            CellArea area = CellArea.createCellArea(0, 1, 0, 1);

            //Adding a new validation to the collection
            int index = validations.add(area);
            Validation validation = validations.get(index);

            //Setting the data validation type
            validation.setType(ValidationType.TEXT_LENGTH);

            //Setting the operator for the data validation
            validation.setOperator(OperatorType.LESS_OR_EQUAL);

            //Setting the value or expression associated with the data validation
            validation.setFormula1("5");

            //Enabling the error
            validation.setShowError(true);

            //Setting the validation alert style
            validation.setAlertStyle(ValidationAlertType.WARNING);

            //Setting the title of the data-validation error dialog box
            validation.setErrorTitle("Text Length Error");

            //Setting the data validation error message
            validation.setErrorMessage(" Enter a Valid String");

            //Setting and enabling the data validation input message
            validation.setInputMessage("TextLength Validation Type");
            validation.setIgnoreBlank(true);
            validation.setShowInput(true);

            workbook.save(filePath + File.separator + "TextLengthDataValidation_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Text Length Data Validation", e);
        }
    }
}
