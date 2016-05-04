package com.example;

import android.os.Bundle;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;
import com.aspose.cells.*;

public class DetectCellValueStartsWithSingleQuoteActivity extends Activity {

	public void detectCellValueStartsWithSingleQuote() throws Exception
	{
		//Create an instance of workbook
		Workbook workbook = new Workbook();

		//Access first worksheet from the collection
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Access cells A1 and A2
		Cell a1 = worksheet.getCells().get("A1");
		Cell a2 = worksheet.getCells().get("A2");

		//Add simple text to cell A1 and text with quote prefix to cell A2
		a1.putValue("sample");
		a2.putValue("'sample");

		//Print their string values, A1 and A2 both are same
		System.out.println("String value of A1: " + a1.getStringValue());
		System.out.println("String value of A2: " + a2.getStringValue());

		//Access styles of cells A1 and A2
		Style s1 = a1.getStyle();
		Style s2 = a2.getStyle();

		System.out.println();

		//Check if A1 and A2 has a quote prefix
		System.out.println("A1 has a quote prefix: " + s1.getQuotePrefix());
		System.out.println("A2 has a quote prefix: " + s2.getQuotePrefix());

	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_detect_cell_value_starts_with_single_quote);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
				
		try
		{
			System.out.println("---------------Executing--------------------------");
			detectCellValueStartsWithSingleQuote();
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
		getMenuInflater().inflate(
				R.menu.detect_cell_value_starts_with_single_quote, menu);
		return true;
	}

}
