package com.example;

import java.io.*;
import java.util.*;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.content.res.AssetManager;
import android.view.Menu;
import android.widget.TextView;

public class SmartMarkerGroupingImageActivity extends Activity {
	
	public class SmartMarkerGroupingImage 
	{
		public void Execute(String dataDir) throws Exception {

			//Get the first image from assets
			AssetManager am = getApplicationContext().getAssets();
			InputStream is = am.open("moon.png");		
			byte[] photo1 = new byte[is.available()];
			is.read(photo1);
			is.close();
			
			//Get the second image from assets			
			am = getApplicationContext().getAssets();
			is = am.open("moon2.png");		
			byte[] photo2 = new byte[is.available()];
			is.read(photo2);
			is.close();
	
			//Create a new workbook and access its worksheet
			Workbook workbook = new Workbook();
			Worksheet worksheet = workbook.getWorksheets().get(0);
			
			//Set the standard row height to 35
			worksheet.getCells().setStandardHeight(35);

			//Set column widhts of D, E and F
			worksheet.getCells().setColumnWidth(3, 20);
			worksheet.getCells().setColumnWidth(4, 20);
			worksheet.getCells().setColumnWidth(5, 40);
			
			//Add the headings in columns D, E and F
			worksheet.getCells().get("D1").putValue("Name");
			Style st = worksheet.getCells().get("D1").getStyle();
			st.getFont().setBold(true);
			worksheet.getCells().get("D1").setStyle(st);

			worksheet.getCells().get("E1").putValue("City");
			worksheet.getCells().get("E1").setStyle(st);

			worksheet.getCells().get("F1").putValue("Photo");
			worksheet.getCells().get("F1").setStyle(st);
		
			//Add smart marker tags in columns D, E, F
			worksheet.getCells().get("D2").putValue("&=Person.Name(group:normal,skip:1)");
			worksheet.getCells().get("E2").putValue("&=Person.City");
			worksheet.getCells().get("F2").putValue("&=Person.Photo(Picture:FitToCell)");
		
			//Create Persons objects with photos
			ArrayList<Person> persons = new ArrayList<Person>();       
			persons.add(new Person("George", "New York", photo1));
			persons.add(new Person("George", "New York", photo2));
			persons.add(new Person("George", "New York", photo1));
			persons.add(new Person("George", "New York", photo2));
			persons.add(new Person("Johnson", "London", photo2));
			persons.add(new Person("Johnson", "London", photo1));
			persons.add(new Person("Johnson", "London", photo2));
			persons.add(new Person("Simon", "Paris", photo1));
			persons.add(new Person("Simon", "Paris", photo2));
			persons.add(new Person("Simon", "Paris", photo1));
			persons.add(new Person("Henry", "Sydney", photo2));
			persons.add(new Person("Henry", "Sydney", photo1));
			persons.add(new Person("Henry", "Sydney", photo2));

			//Create a workbook designer
			WorkbookDesigner designer = new WorkbookDesigner(workbook);

			//Set the data source and process smart marker tags
			designer.setDataSource("Person", persons);
			designer.process();

			//Save the workbook
			workbook.save(dataDir + "output.xlsx", SaveFormat.XLSX);

		}
		
		public class Person
		{
			//Create Name, City and Photo properties
			private String m_Name;
			private String m_City;
			private byte[] m_Photo;

			public Person(String name, String city, byte[] photo)
			{
				m_Name = name;
				m_City = city;
				m_Photo = photo;
			}

			public String getName() { return this.m_Name; }
			public void setName(String name) { this.m_Name = name; }

			public String getCity() { return this.m_City; }
			public void setCity(String address) { this.m_City = address; }

			public byte[] getPhoto() { return this.m_Photo; }
			public void setPhoto(byte[] photo) { this.m_Photo = photo; }
		}    

	}

	
	public void smartMarkerGroupingImage() throws Exception
	{
		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		SmartMarkerGroupingImage grouping = new SmartMarkerGroupingImage();
		grouping.Execute(sdPath);
	}

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_smart_marker_grouping_image);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			System.out.println("---------------Executing--------------------------");
			smartMarkerGroupingImage();
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
		getMenuInflater().inflate(R.menu.smart_marker_grouping_image, menu);
		return true;
	}

}
