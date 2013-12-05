package com.example.chart;

import java.io.File;

import com.aspose.cells.*;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	void createAreaChart() throws Exception {

		//Get the SD card path
		String sdPath = Environment.getExternalStorageDirectory().getPath() + File.separator;

		String filePath = sdPath + "AreaTemplate.xlsx";
				
		Workbook workbook = new Workbook(filePath);

		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);
		// Set the name of worksheet
		worksheet.setName("Area");

		// Create chart
		ChartCollection charts = worksheet.getCharts();
		Chart chart = charts.get(charts.add(ChartType.AREA, 5, 1, 29, 10));

		// Set properties of nseries
		SeriesCollection nSeries = chart.getNSeries();
		nSeries.add("B4:F4", false);
		nSeries.add("B3:F3", false);
		nSeries.add("B2:F2", false);
		nSeries.setCategoryData("B1:F1");

		Cells cells = worksheet.getCells();
		String temp = "";
		for (int i = 0; i < chart.getNSeries().getCount(); i++) {
			temp = cells.get(i + 1, 0).getStringValue();
			nSeries.get(i).setName(temp);
			nSeries.get(i).setColorVaried(true);
		}

		// Set properties of Legend
		chart.getLegend().setPosition(LegendPositionType.TOP);

		// Set properties of chart title
		Title title = chart.getTitle();
		title.setText("Sales By Region");
		Font font1 = title.getFont();
		font1.setColor(Color.getBlack());
		font1.setBold(true);
		font1.setSize(12);

		// Set properties of categoryaxis title
		Axis categoryAxis = chart.getCategoryAxis();
		title = categoryAxis.getTitle();
		title.setText("Year(2002-2006)");
		Font font2 = title.getFont();
		font2.setColor(Color.getBlack());
		font2.setBold(true);
		font2.setSize(10);
		categoryAxis.setAxisBetweenCategories(false);

		workbook.save(filePath + ".out.xlsx", SaveFormat.XLSX);
	}

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView)findViewById(R.id.myTextBox);
		
		try
		{
			createAreaChart();
			tx.setText("Area Chart created successfully. Please check the root of SD path.");
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
