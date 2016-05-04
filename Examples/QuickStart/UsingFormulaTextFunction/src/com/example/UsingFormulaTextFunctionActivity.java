package com.example;

import com.aspose.cells.*;

import android.os.Bundle;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;


public class UsingFormulaTextFunctionActivity extends Activity {

	public void usingFormulaTextFunction() throws Exception
	{
		//Create a workbook object
		Workbook workbook = new Workbook();

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Put some formula in cell A1
		Cell cellA1 = worksheet.getCells().get("A1");
		cellA1.setFormula("=Sum(B1:B10)");

		//Get the text of the formula in cell A2 using FORMULATEXT function
		Cell cellA2 = worksheet.getCells().get("A2");
		cellA2.setFormula("=FormulaText(A1)");

		//Calculate the workbook
		workbook.calculateFormula();

		//Print the results of A2
		//It will now print the text of the formula inside cell A1
		System.out.println(cellA2.getStringValue());
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_using_formula_text_function);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			System.out.println("---------------Executing--------------------------");
			usingFormulaTextFunction();
			tx.setText("Successfully Done. Please see the console output of the code in the LogCat.");
			System.out.println("---------------Done--------------------------");
		}
		catch(Exception ex)
		{
			tx.setText("Error during document processing: " + ex.getMessage());
		}

	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.using_formula_text_function, menu);
		return true;
	}

}
