package com.example.helloworld;

import java.io.File;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	public void GenerateWorkbook() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;
		
		//Create a workbook
		Workbook workbook = new Workbook();
		
		//Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);
		
		//Access cell A1 and put Hello World message
		Cell cell = worksheet.getCells().get("A1");
		cell.putValue("Hello World!");
		
		//Save the workbook in different formats
		workbook.save(sdPath + "HelloWorld.xls", SaveFormat.EXCEL_97_TO_2003);
		workbook.save(sdPath + "HelloWorld.xlsx", SaveFormat.XLSX);
		workbook.save(sdPath + "HelloWorld.ods", SaveFormat.ODS);
		workbook.save(sdPath + "HelloWorld.pdf", SaveFormat.PDF);
		
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
				
		try
		{
			GenerateWorkbook();
			tx.setText("Documents created successfully. Please check the root of SD path.");
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
