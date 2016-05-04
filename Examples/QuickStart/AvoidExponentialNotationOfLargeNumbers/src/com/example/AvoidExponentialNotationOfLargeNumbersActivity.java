package com.example;

import java.io.File;
import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class AvoidExponentialNotationOfLargeNumbersActivity extends Activity {


	public void AvoidExponentialNotationOfLargeNumbersWhileImportingFromHtml() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		//Sample Html containing large number with digits greater than 15
		String html = "<html>"
						+ "<body>"
							+ "<p>1234567890123456</p>"
						+ "</body>"
					+ "</html>";

		//Convert Html to byte array
		byte[] byteArray = html.getBytes();

		//Set Html load options and keep precision true
		HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
		loadOptions.setKeepPrecision(true);

		//Convert byte array into stream
		java.io.ByteArrayInputStream stream = new java.io.ByteArrayInputStream(byteArray);

		//Create workbook from stream with Html load options
		Workbook workbook = new Workbook(stream, loadOptions);

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Auto fit the sheet columns
		worksheet.autoFitColumns();

		//Save the workbook
		workbook.save(sdPath + "output.xlsx", SaveFormat.XLSX);

	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
				
		try
		{
			System.out.println("---------------Executing--------------------------");
			AvoidExponentialNotationOfLargeNumbersWhileImportingFromHtml();
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
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

}
