package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CustomDocumentPropertyCollection;
import com.aspose.cells.DocumentProperty;
import com.aspose.cells.PropertyType;
import com.aspose.cells.Workbook;

import java.io.File;

public class DocumentProperties {

    private static final String TAG = DocumentProperties.class.getName();

    public void getPropertyUsingNameOrIndex() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Retrieve a list of all custom document properties of the Excel file
            CustomDocumentPropertyCollection customProperties = workbook.getWorksheets().getCustomDocumentProperties();

            //Accessing a custom document property by using the property index
            DocumentProperty customProperty1 = customProperties.get(3);

            //Accessing a custom document property by using the property name
            DocumentProperty customProperty2 = customProperties.get("Owner");
        } catch(Exception e) {
            Log.e(TAG, "Get Property Using Name or Index", e);
        }
    }

    public void retrieveNameValueAndTypeOfDocumentProperty() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Retrieve a list of all custom document properties of the Excel file
            CustomDocumentPropertyCollection customProperties = workbook.getWorksheets().getCustomDocumentProperties();

            //Access a custom document property
            DocumentProperty customProperty1 = customProperties.get(0);

            //Store the value of the document property as an object
            Object objectValue = customProperty1.getValue();

            //Access a custom document property
            DocumentProperty customProperty2 = customProperties.get(1);

            //Checking the type of the document property and then storing the value of the
            //document property according to that type
            if (customProperty2.getType() == PropertyType.NUMBER) {
                int intValue = customProperty2.toInt();
            }
        } catch(Exception e) {
            Log.e(TAG, "Retrieve Name, Value And Type of Document Property", e);
        }
    }

    public void addCustomProperty() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            // Retrieve a list of all custom document properties of the Excel file
            CustomDocumentPropertyCollection customProperties = workbook.getWorksheets().getCustomDocumentProperties();

            // Add a custom document property to the Excel file
            DocumentProperty publisher = customProperties.add("Publisher", "Aspose");
        } catch(Exception e) {
            Log.e(TAG, "Add Custom Properties", e);
        }
    }

    public void removeCustomProperty() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            // Retrieve a list of all custom document properties of the Excel file
            CustomDocumentPropertyCollection customProperties = workbook.getWorksheets().getCustomDocumentProperties();

            //Remove a custom document property
            customProperties.remove("Publisher");
        } catch(Exception e) {
            Log.e(TAG, "Remove Custom Properties", e);
        }
    }
}
