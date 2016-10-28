package com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.Color;
import com.aspose.cells.Label;
import com.aspose.cells.MsoFillFormat;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.MsoTextFrame;
import com.aspose.cells.Picture;
import com.aspose.cells.PlacementType;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.io.FileInputStream;

public class InsertControlsInExcelCharts {

    private static final String TAG = InsertControlsInExcelCharts.class.getName();

    public void addLabelControlToTheChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook
            //Open the existing file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the designer chart in the first sheet
            Worksheet sheet = workbook.getWorksheets().get(0);
            Chart chart = sheet.getCharts().get(0);

            //Add a new label to the chart
            Label label = chart.getShapes().addLabelInChart(100, 100, 250, 830);

            //Set the caption of the label
            label.setText("A Label In Chart");

            //Set the Placement Type, the way the
            //label is attached to the cells
            label.setPlacement(PlacementType.FREE_FLOATING);

            //Set the fill color of the label
            label.getFillFormat().setForeColor(Color.getAzure());

            //Save the excel file
            workbook.save(filePath + File.separator + "LabelControlToChart_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Add Label Control to the Chart", e);
        }
    }

    public void addTextBoxControlToTheChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the designer chart in the first sheet
            Worksheet sheet = workbook.getWorksheets().get(0);
            Chart chart = sheet.getCharts().get(0);

            //Add a new TextBox to the chart
            TextBox textbox = chart.getShapes().addTextBoxInChart(0, 0, 350, 2550);

            //Fill the text
            textbox.setText("Sales By Region");

            //Get the TextBox'x text frame
            MsoTextFrame textframe = textbox.getTextFrame();

            //Set the TextBox to adjust it according to its contents
            textframe.setAutoSize(true);

            //Set the font color
            textbox.getFont().setColor(Color.getMaroon());

            //Set the font to bold
            textbox.getFont().setBold(true);

            //Set the font size
            textbox.getFont().setSize(14);

            //Set font attribute to italic
            textbox.getFont().setItalic(true);

            //Get the FillFormat of the TextBox
            MsoFillFormat fillformat = textbox.getFillFormat();

            //Set the ForeColor
            fillformat.setForeColor(Color.getSilver());

            //Get the LineFormat type of the TextBox
            MsoLineFormat lineformat = textbox.getLineFormat();

            //Set the line style
            lineformat.setStyle(MsoLineStyle.THIN_THICK);

            //Set the line weight
            lineformat.setWeight(2);

            //Set the dash style to solid
            lineformat.setDashStyle(MsoLineDashStyle.SOLID);

            //Save the excel file
            workbook.save(filePath + File.separator + "AddTextBoxControl_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Add TextBox Control to the Chart", e);
        }
    }

    public void addPictureToTheChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the designer chart in the first sheet
            Worksheet sheet = workbook.getWorksheets().get(0);
            Chart chart = sheet.getCharts().get(0);

            FileInputStream stream = new FileInputStream(filePath + File.separator + "school.jpg");

            //Add a new picture to the chart
            Picture picture = chart.getShapes().addPictureInChart(50, 50, stream, 40, 40);

            //Get the LineFormat type of the picture
            MsoLineFormat lineformat = picture.getLineFormat();

            //Set the line color
            lineformat.setForeColor(Color.getRed());

            //Set the line style
            lineformat.setStyle(MsoLineStyle.THIN_THICK);

            //Set the line weight
            lineformat.setWeight(2);

            //Set the dash style to solid
            lineformat.setDashStyle(MsoLineDashStyle.SOLID);

            //Save the excel file
            workbook.save(filePath + File.separator + "AddPicture_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Add Picture to the Chart", e);
        }
    }
}
