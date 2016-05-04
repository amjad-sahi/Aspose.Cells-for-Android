package com.example;

import java.io.File;
import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class UsingCustomXmlPartsActivity extends Activity {

	public void usingCustomXmlParts() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		String booksXML= "<catalog><book><title>Complete C#</title><price>44</price></book><book><title>Complete Java</title><price>76</price></book><book><title>Complete SharePoint</title><price>55</price></book><book><title>Complete PHP</title><price>63</price></book><book><title>Complete VB.NET</title><price>72</price></book></catalog>";

		Workbook workbook = new Workbook();
		workbook.getContentTypeProperties().add("BookStore", booksXML);
		workbook.save(sdPath + "output.xlsx");
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_using_custom_xml_parts);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);

		try
		{
			System.out.println("---------------Executing--------------------------");
			usingCustomXmlParts();
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
		getMenuInflater().inflate(R.menu.using_custom_xml_parts, menu);
		return true;
	}

}
