package com.example.datavalidation;

import java.io.File;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void dataValidation() throws Exception {

		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		Workbook wb = new Workbook();
	  	WorksheetCollection worksheets = wb.getWorksheets();
        Worksheet worksheet = worksheets.get(0);
        Cells cells = worksheet.getCells();
        ValidationCollection validations = worksheet.getValidations();

        Cell cell;
        cell = cells.get("A1");
        cell.setValue("Enter Number:");
        cell = cells.get("B1");
        cell.setValue(10);

        //Create a validation object
        int index = validations.add();
        Validation validation = validations.get(index);
        
        //Set the validation type to whole number
        validation.setType(ValidationType.WHOLE_NUMBER);
        
        //Set the operator for validation to between
        validation.setOperator(OperatorType.BETWEEN);
        
        //Set the minimum value for the validation
        validation.setFormula1("0");
        
        //Set the maximum value for the validation
        validation.setFormula2("10");
        validation.setShowError(true);
        validation./* setAlertType */setAlertStyle(ValidationAlertType.INFORMATION);
        validation.setErrorTitle("Error");
        validation.setErrorMessage("Enter value between 0 to 10");
        validation.setInputMessage("Data Validation using Condition for Numbers");
        validation.setIgnoreBlank(true);
        validation.setShowInput(true);
        validation.setShowError(true);

        //Apply the validation to a range of cells from B1 to B1 using the CellArea structure
        CellArea cellArea = CellArea.createCellArea(0, 1, 0, 1);

        //Add the cell area to Validation
        validation.addArea(cellArea);

        wb.save(sdPath + "DataValidation.xlsx", SaveFormat.XLSX);
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			dataValidation();
			tx.setText("Data Validation created successfully. Please check the root of SD path.");
		}
		catch(Exception ex)
		{
			tx.setText("Error during document processing: " + ex.getMessage());
		}
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

}
