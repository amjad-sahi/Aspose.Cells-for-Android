package com.aspose.cells.examples.asposecellsexamples.DrawingObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ArcShape;
import com.aspose.cells.Button;
import com.aspose.cells.Cells;
import com.aspose.cells.CheckBox;
import com.aspose.cells.Color;
import com.aspose.cells.ComboBox;
import com.aspose.cells.LineShape;
import com.aspose.cells.MsoArrowheadLength;
import com.aspose.cells.MsoArrowheadStyle;
import com.aspose.cells.MsoArrowheadWidth;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoFillFormat;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.Oval;
import com.aspose.cells.PlacementType;
import com.aspose.cells.RadioButton;
import com.aspose.cells.SelectionType;
import com.aspose.cells.Style;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class Controls {

    private static final String TAG = Controls.class.getName();

    /**
     * Adding a Text Box Control to the Worksheet
     */
    public void addATextBoxControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the book.
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the textbox object.
            int textboxIndex = worksheet.getTextBoxes().add(2, 1, 160, 200);
            TextBox textbox0 = worksheet.getTextBoxes().get(textboxIndex);

            //Fill the text.
            textbox0.setText("ASPOSE______The .NET & JAVA Component Publisher!");

            //Set the placement.
            textbox0.setPlacement(PlacementType.FREE_FLOATING);

            //Set the font color.
            textbox0.getFont().setColor(Color.getBlue());

            //Set the font to bold.
            textbox0.getFont().setBold(true);

            //Set the font size.
            textbox0.getFont().setSize(14);

            //Set font attribute to italic.
            textbox0.getFont().setItalic(true);

            //Add a hyperlink to the textbox.
            textbox0.addHyperlink("http://www.aspose.com/");

            //Get the filformat of the textbox.
            MsoFillFormat fillformat = textbox0.getFillFormat();

            //Set the fillcolor.
            fillformat.setForeColor(Color.getSilver());

            //Get the lineformat type of the textbox.
            MsoLineFormat lineformat = textbox0.getLineFormat();

            //Set the line style.
            lineformat.setStyle(MsoLineStyle.THIN_THICK);

            //Set the line weight.
            lineformat.setWeight(6);

            //Set the dash style to square_dot.
            lineformat.setDashStyle(MsoLineDashStyle.SQUARE_DOT);


            //Get the second textbox.
            TextBox textbox1 = (com.aspose.cells.TextBox) worksheet.getShapes().addShape(MsoDrawingType.TEXT_BOX, 15, 0, 4, 0, 85, 120);

            //Input some text to it.
            textbox1.setText("This is another simple text box");

            //Set the placement type as the textbox will move and
            //resize with cells.
            textbox1.setPlacement(PlacementType.MOVE_AND_SIZE);

            //Save the Excel file.
            workbook.save(filePath + File.separator + "AddATextBoxControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a Text Box Control", e);
        }
    }

    /**
     * Manipulating TextBox Controls in Designer Spreadsheets
     */
    public void manipulateTextBoxControlsInDesignerSpreadsheets() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "TextBoxes.xls");

            //Get the first worksheet in the book.
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the first textbox object.
            TextBox textbox0 = worksheet.getTextBoxes().get(0);

            //Obtain the text in the first textbox.
            String text0 = textbox0.getText();
            System.out.println(text0);

            //Get the second textbox object.
            TextBox textbox1 = worksheet.getTextBoxes().get(1);

            //Obtain the text in the second textbox.
            String text1 = textbox1.getText();

            //Change the text of the second textbox.
            textbox1.setText("This is an alternative text");

            //Save the excel file.
            workbook.save(filePath + File.separator + "ManipulateTextBox_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Manipulate TextBox Controls in Designer Spreadsheet", e);
        }
    }

    /**
     * Adding Checkbox Control to a Worksheet
     */
    public void addCheckboxControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the book.
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Add a checkbox to the first worksheet in the workbook.
            int checkBoxIndex = worksheet.getCheckBoxes().add(5, 5, 100, 120);
            CheckBox checkBox = worksheet.getCheckBoxes().get(checkBoxIndex);

            //Set its text string.
            checkBox.setText("Check it!");

            //Put a value into B1 cell.
            worksheet.getCells().get("B1").setValue("LnkCell");

            //Set B1 cell as a linked cell for the checkbox.
            checkBox.setLinkedCell("=B1");

            //Save the excel file.
            workbook.save(filePath + File.separator + "AddCheckboxControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Checkbox Control to a Worksheet", e);
        }
    }

    /**
     * Adding a Radio Button Control to a Worksheet
     */
    public void addARadioButton() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Get the worksheet cells collection.
            Cells cells = sheet.getCells();

            //Insert a value.
            cells.get("C2").setValue("Age Groups");

            Style style = cells.get("B3").getStyle();
            style.getFont().setBold(true);
            //Set it bold.
            cells.get("C2").setStyle(style);

            //Add a radio button to the first sheet.
            RadioButton radio1 = (RadioButton) sheet.getShapes().addShape(MsoDrawingType.RADIO_BUTTON, 3, 0, 1, 0, 20, 100);

            //Set its text string.
            radio1.setText("20-29");

            //Set A1 cell as a linked cell for the radio button.
            radio1.setLinkedCell("A1");

            //Make the radio button 3-D.
            radio1.setShadow(true);

            //Set the foreground color of the radio button.
            radio1.getFillFormat().setForeColor(Color.getGreen());

            // set the line style of the radio button.
            radio1.getLineFormat().setStyle(MsoLineStyle.THICK_THIN);

            //Set the weight of the radio button.
            radio1.getLineFormat().setWeight(4);

            //Set the line color of the radio button.
            radio1.getLineFormat().setForeColor(Color.getBlue());

            //Set the dash style of the radio button.
            radio1.getLineFormat().setDashStyle(MsoLineDashStyle.SOLID);

            //Make the line format visible.
            radio1.getLineFormat().setVisible(true);

            //Make the fill format visible.
            radio1.getFillFormat().setVisible(true);

            //Add another radio button to the first sheet.
            com.aspose.cells.RadioButton radio2 = (com.aspose.cells.RadioButton) sheet.getShapes().addShape(MsoDrawingType.RADIO_BUTTON, 6, 0, 1, 0, 20, 100);

            //Set its text string.
            radio2.setText("30-39");

            //Set A1 cell as a linked cell for the radio button.
            radio2.setLinkedCell("A1");

            //Make the radio button 3-D.
            radio2.setShadow(true);

            //Set the foreground color of the radio button.
            radio2.getFillFormat().setForeColor(Color.getGreen());

            // set the line style of the radio button.
            radio2.getLineFormat().setStyle(MsoLineStyle.THICK_THIN);

            //Set the weight of the radio button.
            radio2.getLineFormat().setWeight(4);

            //Set the line color of the radio button.
            radio2.getLineFormat().setForeColor(Color.getBlue());

            //Set the dash style of the radio button.
            radio2.getLineFormat().setDashStyle(MsoLineDashStyle.SOLID);

            //Make the line format visible.
            radio2.getLineFormat().setVisible(true);

            //Make the fill format visible.
            radio2.getFillFormat().setVisible(true);

            //Add another radio button to the first sheet.
            RadioButton radio3 = (RadioButton) sheet.getShapes().addShape(MsoDrawingType.RADIO_BUTTON, 9, 0, 1, 0, 20, 100);

            //Set its text string.
            radio3.setText("40-49");

            //Set A1 cell as a linked cell for the radio button.
            radio3.setLinkedCell("A1");

            //Make the radio button 3-D.
            radio3.setShadow(true);

            //Set the foreground color of the radio button.
            radio3.getFillFormat().setForeColor(Color.getGreen());

            // set the line style of the radio button.
            radio3.getLineFormat().setStyle(MsoLineStyle.THICK_THIN);

            //Set the weight of the radio button.
            radio3.getLineFormat().setWeight(4);

            //Set the line color of the radio button.
            radio3.getLineFormat().setForeColor(Color.getBlue());

            //Set the dash style of the radio button.
            radio3.getLineFormat().setDashStyle(MsoLineDashStyle.SOLID);

            //Make the line format visible.
            radio3.getLineFormat().setVisible(true);

            //Make the fill format visible.
            radio3.getFillFormat().setVisible(true);

            //Save the Excel file.
            workbook.save(filePath + File.separator + "AddRadioButton_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a Radio Button Control", e);
        }
    }

    /**
     * Adding ComboBox Control to the Worksheet
     */
    public void addComboBoxControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Get the worksheet cells collection.
            Cells cells = sheet.getCells();

            //Input a value.
            cells.get("B3").setValue("Employee:");

            Style style = cells.get("B3").getStyle();
            style.getFont().setBold(true);
            //Set it bold.
            cells.get("B3").setStyle(style);

            //Input some values that denote the input range for the combo box.
            cells.get("A2").setValue("Emp001");
            cells.get("A3").setValue("Emp002");
            cells.get("A4").setValue("Emp003");
            cells.get("A5").setValue("Emp004");
            cells.get("A6").setValue("Emp005");
            cells.get("A7").setValue("Emp006");

            //Add a new combo box.
            ComboBox comboBox = (ComboBox) sheet.getShapes().addShape(MsoDrawingType.COMBO_BOX, 3, 0, 1, 0, 20, 100);

            //Set the linked cell;
            comboBox.setLinkedCell("A1");

            //Set the input range.
            comboBox.setInputRange("=A2:A7");

            //Set no. of list lines displayed in the combo box's list portion.
            comboBox.setDropDownLines(5);

            //Set the combo box with 3-D shading.
            comboBox.setShadow(true);

            //AutoFit Columns
            sheet.autoFitColumns();

            //Saves the file.
            workbook.save(filePath + File.separator + "AddAComboBox_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a ComboBox Control", e);
        }
    }

    /**
     * Adding Label Control to the Worksheet
     */
    public void addLabelControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Add a new label to the worksheet.
            com.aspose.cells.Label label = (com.aspose.cells.Label) sheet.getShapes().addShape(MsoDrawingType.LABEL, 2, 2, 2, 0, 60, 120);

            //Set the caption of the label.
            label.setText("This is a Label");

            //Set the Placement Type, the way the label is attached to the cells.
            label.setPlacement(PlacementType.FREE_FLOATING);

            //Set the fill color of the label.
            label.getFillFormat().setForeColor(Color.getYellow());

            //Saves the file.
            workbook.save(filePath + File.separator + "AddLabelControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a Label Control", e);
        }
    }

    /**
     * Adding ListBox Control to the Worksheet
     */
    public void addListBoxControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Get the worksheet cells collection.
            Cells cells = sheet.getCells();

            //Input a value.
            cells.get("B3").setValue("Choose Dept:");

            Style style = cells.get("B3").getStyle();
            style.getFont().setBold(true);
            //Set it bold.
            cells.get("B3").setStyle(style);

            //Input some values that denote the input range for the combo box.
            cells.get("A2").setValue("Sales");
            cells.get("A3").setValue("Finance");
            cells.get("A4").setValue("MIS");
            cells.get("A5").setValue("R&D");
            cells.get("A6").setValue("Marketing");
            cells.get("A7").setValue("HRA");

            //Add a new list box.
            com.aspose.cells.ListBox listBox = (com.aspose.cells.ListBox) sheet.getShapes().addShape(MsoDrawingType.LIST_BOX, 3, 3, 1, 0, 100, 122);

            //Set the linked cell;
            listBox.setLinkedCell("A1");

            //Set the input range.
            listBox.setInputRange("=A2:A7");

            //Set the Placement Type, the way the list box is attached to the cells.
            listBox.setPlacement(PlacementType.FREE_FLOATING);

            //Set the list box with 3-D shading.
            listBox.setShadow(true);

            //Set the selection type.
            listBox.setSelectionType(SelectionType.SINGLE);

            //AutoFit Columns
            sheet.autoFitColumns();

            //Saves the file.
            workbook.save(filePath + File.separator + "AddListBox_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add ListBox Control", e);
        }
    }

    /**
     * Adding Button Control to the Worksheet
     */
    public void addButtonControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Add a new button to the worksheet.
            Button button = (Button) sheet.getShapes().addShape(MsoDrawingType.BUTTON, 2, 2, 2, 0, 20, 80);

            //Set the caption of the button.
            button.setText("Aspose");

            //Set the Placement Type, the way the button is attached to the cells.
            button.setPlacement(PlacementType.FREE_FLOATING);

            //Set the font name.
            button.getFont().setName("Tahoma");

            //Set the caption string bold.
            button.getFont().setBold(true);

            //Set the color to blue.
            button.getFont().setColor(Color.getBlue());

            //Set the hyperlink for the button.
            button.addHyperlink("http://www.aspose.com/");

            //Saves the file.
            workbook.save(filePath + File.separator + "AddButtonControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Button Control", e);
        }
    }

    /**
     * Adding Line Control to the Worksheet
     */
    public void addLineControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the book.
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Add a new line to the worksheet.
            LineShape line1 = (LineShape) worksheet.getShapes().addShape(MsoDrawingType.LINE, 5, 1, 0, 0, 0, 250);

            //Set the line dash style
            MsoLineFormat shapeline = line1.getLineFormat();
            shapeline.setDashStyle(MsoLineDashStyle.SOLID);

            //Set the placement.
            line1.setPlacement(PlacementType.FREE_FLOATING);

            //Add another line to the worksheet.
            LineShape line2 = (LineShape) worksheet.getShapes().addShape(MsoDrawingType.LINE, 7, 1, 0, 0, 85, 250);

            //Set the line dash style.
            shapeline = line2.getLineFormat();
            shapeline.setDashStyle(MsoLineDashStyle.DASH_LONG_DASH);

            //Set the weight of the line.
            MsoLineFormat lineformat = line2.getLineFormat();
            lineformat.setWeight(4);

            //Set the placement.
            line2.setPlacement(PlacementType.FREE_FLOATING);

            //Add the third line to the worksheet.
            LineShape line3 = (LineShape) worksheet.getShapes().addShape(MsoDrawingType.LINE, 13, 1, 0, 0, 0, 250);

            //Set the line dash style
            shapeline = line1.getLineFormat();
            shapeline.setDashStyle(MsoLineDashStyle.SOLID);

            //Set the placement.
            line3.setPlacement(PlacementType.FREE_FLOATING);

            //Make the gridlines invisible in the first worksheet.
            workbook.getWorksheets().get(0).setGridlinesVisible(false);

            //Save the excel file.
            workbook.save(filePath + File.separator + "AddLineControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Line Control", e);
        }
    }

    /**
     * Adding an ArrowHead to the Line
     */
    public void addAnArrowHead() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet in the book.
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Add a line to the worksheet.
            com.aspose.cells.LineShape line2 = (com.aspose.cells.LineShape) worksheet.getShapes().addShape(MsoDrawingType.LINE, 7, 1, 1, 0, 85, 250);
            MsoLineFormat lineformat = line2.getLineFormat();

            //Set the line color
            lineformat.setForeColor(Color.getBlue());

            //Set the line style.
            lineformat.setDashStyle(MsoLineDashStyle.SOLID);

            //Set the weight of the line.
            lineformat.setWeight(3);

            //Set the placement.
            line2.setPlacement(PlacementType.FREE_FLOATING);

            //Set the line arrows.
            line2.setEndArrowheadWidth(MsoArrowheadWidth.MEDIUM);
            line2.setEndArrowheadStyle(MsoArrowheadStyle.ARROW);
            line2.setEndArrowheadLength(MsoArrowheadLength.MEDIUM);

            line2.setBeginArrowheadWidth(MsoArrowheadWidth.NARROW);
            line2.setBeginArrowheadStyle(MsoArrowheadStyle.ARROW_DIAMOND);
            line2.setBeginArrowheadLength(MsoArrowheadLength.MEDIUM);

            //Make the gridlines invisible in the first worksheet.
            workbook.getWorksheets().get(0).setGridlinesVisible(false);

            //Save the excel file.
            workbook.save(filePath + File.separator + "AddAnArrowHead_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add an ArrowHead", e);
        }
    }

    /**
     * Adding Rectangle Control to the Worksheet
     */
    public void addRectangleControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook excelbook = new Workbook();

            //Add a rectangle control.
            com.aspose.cells.RectangleShape rectangle = (com.aspose.cells.RectangleShape) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.RECTANGLE, 3, 2, 0, 0, 70, 130);

            //Set the placement of the rectangle.
            rectangle.setPlacement(PlacementType.FREE_FLOATING);

            //Set the fill format.
            MsoFillFormat fillformat = rectangle.getFillFormat();
            fillformat.setForeColor(Color.getOlive());

            //Set the line style.
            MsoLineFormat linestyle = rectangle.getLineFormat();
            linestyle.setStyle(MsoLineStyle.THICK_THIN);

            //Set the line weight.
            linestyle.setWeight(4);

            //Set the color of the line.
            linestyle.setForeColor(Color.getBlue());

            //Set the dash style of the rectangle.
            linestyle.setDashStyle(MsoLineDashStyle.SOLID);

            //Save the excel file.
            excelbook.save(filePath + File.separator + "AddRectangle_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Rectangle Control", e);
        }
    }

    /**
     * Adding Arc Control to the Worksheet
     */
    public void addArcControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook excelbook = new Workbook();
            //Add an arc shape.
            ArcShape arc1 = (ArcShape) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.ARC, 2, 2, 0, 0, 130, 130);
            //Set the placement of the arc.
            arc1.setPlacement(PlacementType.FREE_FLOATING);
            //Set the fill format.
            MsoFillFormat fillformat = arc1.getFillFormat();
            fillformat.setForeColor(Color.getBlue());
            //Set the line style.
            MsoLineFormat lineformat = arc1.getLineFormat();
            lineformat.setStyle(MsoLineStyle.SINGLE);
            //Set the line weight.
            lineformat.setWeight(1);
            //Set the color of the arc line.
            lineformat.setForeColor(Color.getBlue());
            //Set the dash style of the arc.
            lineformat.setDashStyle(MsoLineDashStyle.SOLID);
            //Add another arc shape.
            ArcShape arc2 = (ArcShape) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.ARC, 9, 2, 0, 0, 130, 130);
            //Set the placement of the arc.
            arc2.setPlacement(PlacementType.FREE_FLOATING);
            //Set the line style.
            MsoLineFormat lineformat1 = arc2.getLineFormat();
            lineformat1.setStyle(MsoLineStyle.SINGLE);
            //Set the line weight.
            lineformat1.setWeight(1);
            //Set the color of the arc line.
            lineformat1.setForeColor(Color.getBlue());
            //Set the dash style of the arc.
            lineformat1.setDashStyle(MsoLineDashStyle.SOLID);
            //Save the excel file.
            excelbook.save(filePath + File.separator + "AddArcControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Arc Control", e);
        }
    }

    /**
     * Adding Oval Control to the Worksheet
     */
    public void addOvalControl() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook.
            Workbook excelbook = new Workbook();

            //Add an oval shape.
            Oval oval1 = (Oval) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.OVAL, 2, 2, 0, 0, 130, 130);

            //Set the placement of the oval.
            oval1.setPlacement(PlacementType.FREE_FLOATING);

            //Set the fill format.
            MsoFillFormat fillformat = oval1.getFillFormat();
            fillformat.setForeColor(Color.getNavy());

            //Set the line style.
            MsoLineFormat lineformat = oval1.getLineFormat();
            lineformat.setStyle(MsoLineStyle.SINGLE);

            //Set the line weight.
            lineformat.setWeight(1);

            //Set the color of the oval line.
            lineformat.setForeColor(Color.getGreen());

            //Set the dash style of the oval.
            lineformat.setDashStyle(MsoLineDashStyle.SOLID);

            //Add another arc shape.
            Oval oval2 = (Oval) excelbook.getWorksheets().get(0).getShapes().addShape(MsoDrawingType.OVAL, 9, 2, 0, 0, 130, 130);

            //Set the placement of the oval.
            oval2.setPlacement(PlacementType.FREE_FLOATING);

            //Set the line style.
            MsoLineFormat lineformat1 = oval2.getLineFormat();
            lineformat1.setStyle(MsoLineStyle.SINGLE);

            //Set the line weight.
            lineformat1.setWeight(1);

            //Set the color of the oval line.
            lineformat1.setForeColor(Color.getBlue());

            //Set the dash style of the oval.
            lineformat1.setDashStyle(MsoLineDashStyle.SOLID);

            //Save the excel file.
            excelbook.save(filePath + File.separator + "AddOvalControl_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Oval Control", e);
        }
    }
}
