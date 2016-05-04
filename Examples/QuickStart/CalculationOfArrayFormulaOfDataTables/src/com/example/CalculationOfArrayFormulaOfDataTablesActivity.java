package com.example;

import java.io.File;
import java.io.InputStream;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.content.res.AssetManager;
import android.view.Menu;
import android.widget.TextView;

public class CalculationOfArrayFormulaOfDataTablesActivity extends Activity {

	public void calculationOfArrayFormulaOfDataTables() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		//Load workbook object found in assets	
		AssetManager am = getApplicationContext().getAssets();
		InputStream is = am.open("DataTable.xlsx");	
		Workbook workbook = new Workbook(is);

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//When you will put 100 in B1, then all Data Table values
		//formatted as Yellow will become 120
		worksheet.getCells().get("B1").putValue(100);

		//Calculate formula, now it also calculates Data Table array formula
		workbook.calculateFormula();

		//Save the workbook in pdf format
		workbook.save(sdPath + "output.pdf");
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_calculation_of_array_formula_of_data_tables);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			System.out.println("---------------Executing--------------------------");
			calculationOfArrayFormulaOfDataTables();
			tx.setText("Document created successfully. Please check the root of SD path.");
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
		getMenuInflater().inflate(
				R.menu.calculation_of_array_formula_of_data_tables, menu);
		return true;
	}

}
