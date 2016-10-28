package com.aspose.cells.examples.asposecellsexamples.DrawingObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.OleFileType;
import com.aspose.cells.OleObject;
import com.aspose.cells.OleObjectCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ManageOLEObjects {

    private static final String TAG = ManageOLEObjects.class.getName();

    public void insertOLEObject() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Get the image file.
            File file = new File(filePath + File.separator + "school.jpg");

            //Get the picture into the streams.
            byte[] img = new byte[(int) file.length()];
            FileInputStream fis = new FileInputStream(file);
            fis.read(img);

            //Get the Excel file into the streams.
            file = new File(filePath + File.separator + "Book1.xls");
            byte[] data = new byte[(int) file.length()];
            fis = new FileInputStream(file);
            fis.read(data);

            //Instantiate a new Workbook.
            Workbook wb = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = wb.getWorksheets().get(0);

            //Add an Ole object into the worksheet with the image
            //shown in MS Excel.
            int oleObjIndex = sheet.getOleObjects().add(14, 3, 200, 220, img);
            OleObject oleObj = sheet.getOleObjects().get(oleObjIndex);

            //Set embedded OLE object data.
            oleObj.setObjectData(data);

            //Save the Excel file
            wb.save(filePath + File.separator + "InsertOLEObject_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Insert OLE Object", e);
        }
    }

    public void extractOLEObject() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the OleObject Collection in the first worksheet.
            OleObjectCollection oles = workbook.getWorksheets().get(0).getOleObjects();

            //Loop through all the OLE objects and extract each object in the worksheet
            for (int i = 0; i < oles.getCount(); i++) {
                if (oles.get(i).getMsoDrawingType() == MsoDrawingType.OLE_OBJECT) {
                    OleObject ole = (OleObject)oles.get(i);
                    //Specify the output filename.
                    String fileName = filePath + File.separator + "ExtractOLEObject_Out" + i + ".";
                    //Specify each file format based on the oleformattype.
                    switch (ole.getFileType())
                    {
                        case OleFileType.DOC:
                            fileName += "doc";
                            break;
                        case OleFileType.XLS:
                            fileName += "Xls";
                            break;
                        case OleFileType.PPT:
                            fileName += "Ppt";
                            break;
                        case OleFileType.PDF:
                            fileName += "Pdf";
                            break;
                        case OleFileType.UNKNOWN:
                            fileName += "Jpg";
                            break;
                        default:
                            fileName += "data";
                            break;
                    }

                    FileOutputStream fos = new FileOutputStream(fileName);
                    byte[] data = ole.getObjectData();
                    fos.write(data);
                    fos.close();
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Extract OLE Object", e);
        }
    }
}
