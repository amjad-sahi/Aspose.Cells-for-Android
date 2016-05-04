package com.example;


import com.aspose.cells.*;
import java.io.File;
import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;



public class SearchDataUsingOriginalValuesActivity extends Activity {

	public void searchDataUsingOriginalValues() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		//Create workbook object
		Workbook workbook = new Workbook();

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Add 10 in cell A1 and A2
		worksheet.getCells().get("A1").putValue(10);
		worksheet.getCells().get("A2").putValue(10);

		//Add Sum formula in cell D4 but customize it as ---
		Cell cell = worksheet.getCells().get("D4");

		Style style = cell.getStyle();
		style.setCustom("---");
		cell.setStyle(style);

		//The result of formula will be 20
		//but 20 will not be visible because
		//the cell is formated as ---
		cell.setFormula("=Sum(A1:A2)");

		//Calculate the workbook
		workbook.calculateFormula();

		//Create find options, we will search 20 using
		//original values otherwise 20 will never be found
		//because it is formatted as ---
		FindOptions options = new FindOptions();
		options.setLookInType(LookInType.ORIGINAL_VALUES);
		options.setLookAtType(LookAtType.ENTIRE_CONTENT);

		Cell foundCell = null;
		Object obj = 20;

		//Find 20 which is Sum(A1:A2) and formatted as ---
		foundCell = worksheet.getCells().find(obj, foundCell, options);

		//Print the found cell
		System.out.println(foundCell);

		//Save the workbook
		workbook.save(sdPath + "output.xlsx");

	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_search_data_using_original_values);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			System.out.println("---------------Executing--------------------------");
			searchDataUsingOriginalValues();
			tx.setText("Documents created successfully. Please check the root of SD path.");
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
		getMenuInflater().inflate(R.menu.search_data_using_original_values,
				menu);
		return true;
	}

}
