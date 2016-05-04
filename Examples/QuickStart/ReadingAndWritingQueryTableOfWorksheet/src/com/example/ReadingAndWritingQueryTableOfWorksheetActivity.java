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

public class ReadingAndWritingQueryTableOfWorksheetActivity extends Activity {

	public void readingAndWritingQueryTableOfWorksheet() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;
	
		//Load workbook object found in assets	
		AssetManager am = getApplicationContext().getAssets();
		InputStream is = am.open("Sample.xlsx");	
		Workbook workbook = new Workbook(is);

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Access first Query Table
		QueryTable qt = worksheet.getQueryTables().get(0);

		//Print Query Table Data
		System.out.println("Adjust Column Width: " + qt.getAdjustColumnWidth());
		System.out.println("Preserve Formatting: " + qt.getPreserveFormatting());

		//Now set Preserve Formatting to true
		qt.setPreserveFormatting(true);

		//Save the workbook
		workbook.save(sdPath + "output.xlsx");
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_reading_and_writing_query_table_of_worksheet);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);

		try
		{
			System.out.println("---------------Executing--------------------------");
			readingAndWritingQueryTableOfWorksheet();
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
		getMenuInflater().inflate(
				R.menu.reading_and_writing_query_table_of_worksheet, menu);
		return true;
	}

}
