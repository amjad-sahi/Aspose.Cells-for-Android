package com.example;

import java.io.File;
import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class CustomPropertiesOfDocumentInformationPanelActivity extends
		Activity {

	public void customPropertiesOfDocumentInformationPanel() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		//Create workbook object
		Workbook workbook = new Workbook(FileFormatType.XLSX);

		//Add simple property without any type
		workbook.getContentTypeProperties().add("MK31", "Simple Data");

		//Add date time property with type
		workbook.getContentTypeProperties().add("MK32", "04-Mar-2015", "DateTime");

		//Save the workbook
		workbook.save(sdPath + "output.xlsx");
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_custom_properties_of_document_information_panel);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);

		try
		{
			System.out.println("---------------Executing--------------------------");
			customPropertiesOfDocumentInformationPanel();
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
				R.menu.custom_properties_of_document_information_panel, menu);
		return true;
	}

}
