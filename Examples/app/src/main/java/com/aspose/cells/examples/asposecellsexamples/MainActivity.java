package com.aspose.cells.examples.asposecellsexamples;

import android.Manifest;
import android.app.ProgressDialog;
import android.content.pm.PackageManager;
import android.os.AsyncTask;
import android.os.Bundle;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.ListView;
import android.widget.Toast;

import com.aspose.cells.PivotTable;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables.ApplyConsolidationFunctionToDataFieldsOfPivotTable;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables.ChangeAPivotTableSourceData;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables.CreateAPivotTableReport;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables.CreatePivotTablesAndPivotCharts;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.PivotTables.CustomizeTheAppearanceOfPivotTable;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.SmartMarkersAndFormulaCalculation.FormulaCalculationEngine;
import com.aspose.cells.examples.asposecellsexamples.AdvancedTopics.SmartMarkersAndFormulaCalculation.SmartMarkers;
import com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods.DetectFileFormat;
import com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods.GetCellNameFromRowAndColumnIndices;
import com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods.GetRowAndColumnIndicesFromCellName;
import com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods.MergeFiles;
import com.aspose.cells.examples.asposecellsexamples.Chart.AdvancedFeatures.CustomChart;
import com.aspose.cells.examples.asposecellsexamples.Chart.AdvancedFeatures.ManipulateDesignerCharts;
import com.aspose.cells.examples.asposecellsexamples.Chart.AdvancedFeatures.UsingSparklines;
import com.aspose.cells.examples.asposecellsexamples.Chart.ChangeChartPositionAndSize;
import com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures.Apply3DFormatToChart;
import com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures.ChartAppearance;
import com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures.InsertControlsInExcelCharts;
import com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures.SetChartData;
import com.aspose.cells.examples.asposecellsexamples.Chart.CreateASimpleChart;
import com.aspose.cells.examples.asposecellsexamples.Data.AddOnFeatures.AddHyperlinks;
import com.aspose.cells.examples.asposecellsexamples.Data.AddOnFeatures.MergeAndUnmergeCells;
import com.aspose.cells.examples.asposecellsexamples.Data.AddOnFeatures.NamedRanges;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.AccessAWorksheetMaximumDisplayRange;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.AccessCells;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.AddDataToCells;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.DataSorting;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.ExportDataFromWorksheets;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.FindOrSearchData;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.ImportDataToWorksheets;
import com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures.TracePrecedentsAndDependents;
import com.aspose.cells.examples.asposecellsexamples.Data.DataProcessingFeatures.CreatingSubtotals;
import com.aspose.cells.examples.asposecellsexamples.Data.DataProcessingFeatures.DataFilteringAndValidation;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.AdvanceFormatting.ActivateSheetsAndMakeAnActiveCell;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.AdvanceFormatting.ConditionalFormatting;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.AdvanceFormatting.FormatRowsAndColumns;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.BasicFormatting.FormatCells;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.BasicFormatting.SetDisplayFormats;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel.AddBordersToCells;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel.ColorAndPalette;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel.ColorsAndBackgroundPatterns;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel.ConfigureAlignmentSettings;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel.FontSettings;
import com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel.FormatSelectedCharacters;
import com.aspose.cells.examples.asposecellsexamples.DrawingObjects.Comments;
import com.aspose.cells.examples.asposecellsexamples.DrawingObjects.Controls;
import com.aspose.cells.examples.asposecellsexamples.DrawingObjects.ManageOLEObjects;
import com.aspose.cells.examples.asposecellsexamples.DrawingObjects.Pictures;
import com.aspose.cells.examples.asposecellsexamples.Files.OpenFile;
import com.aspose.cells.examples.asposecellsexamples.Files.SaveFile;
import com.aspose.cells.examples.asposecellsexamples.Formulas.CalculateFormulas;
import com.aspose.cells.examples.asposecellsexamples.Formulas.CalculateFormulasDirectly;
import com.aspose.cells.examples.asposecellsexamples.Formulas.CalculateFormulasOnceOnly;
import com.aspose.cells.examples.asposecellsexamples.Formulas.UsingFormulasToProcessData;
import com.aspose.cells.examples.asposecellsexamples.RowsAndColumns.AdjustRowHeightAndColumnWidth;
import com.aspose.cells.examples.asposecellsexamples.RowsAndColumns.AutoFitRowsAndColumns;
import com.aspose.cells.examples.asposecellsexamples.RowsAndColumns.CopyRowsAndColumns;
import com.aspose.cells.examples.asposecellsexamples.RowsAndColumns.GroupUngroupRowsAndColumns;
import com.aspose.cells.examples.asposecellsexamples.RowsAndColumns.HideAndShowRowsAndColumns;
import com.aspose.cells.examples.asposecellsexamples.RowsAndColumns.ManagingRowsAndColumns;
import com.aspose.cells.examples.asposecellsexamples.Table.ConvertTableToRangeOfData;
import com.aspose.cells.examples.asposecellsexamples.Table.CreateAListObject;
import com.aspose.cells.examples.asposecellsexamples.Table.FormatAListObject;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertChartToImage;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertChartToPDF;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertExcelFilesToHTML;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertExcelFilesToXPS;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertExcelToPDF;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertToMHTML;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertWorksheetToImage;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.ConvertWorksheetToSVG;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.DocumentProperties;
import com.aspose.cells.examples.asposecellsexamples.UtilityFeatures.EncryptFile;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.FreezePanes;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.HideOrShowAWorksheet;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.HideOrShowRowColumnHeaders;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.HideOrShowScrollBars;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.HideOrShowTabs;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.PageBreakPreview;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.SplitPanes;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures.ZoomFactor;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.ManagingWorksheets;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup.CopyAndMoveWorksheet;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup.HeadersAndFooters;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup.ManagePageBreaks;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup.PageOptions;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup.PrintOptions;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup.SetMargins;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.SecurityFeatures.AdvancedProtectionSettingsSinceExcelXP;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.SecurityFeatures.ProtectWorksheet;
import com.aspose.cells.examples.asposecellsexamples.Worksheets.SecurityFeatures.UnprotectWorksheet;

import java.util.ArrayList;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = MainActivity.class.getName();
    private static final int REQUEST_WRITE_EXTERNAL_STORAGE = 1;
    private ListView listView;

    private AdapterView.OnItemClickListener sectionsListener = new AdapterView.OnItemClickListener() {
        @Override
        public void onItemClick(AdapterView<?> parent, View view, int position, long id) {

            if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                Toast.makeText(MainActivity.this, getString(R.string.write_to_external_storage_permission), Toast.LENGTH_SHORT).show();
                return;
            }

            switch (position) {
                case 0:
                    // Working with Files
                    runWorkingWithFilesExamples();
                    break;
                case 1:
                    // Working with Worksheets
                    runWorkingWithWorksheetsExamples();
                    break;
                case 2:
                    // Working with Rows and Columns
                    runWorkingWithRowsAndColumnsExamples();
                    break;
                case 3:
                    // Working with Data
                    runWorkingWithDataExamples();
                    break;
                case 4:
                    // Working with Data Formatting examples
                    runWorkingWithDataFormattingExamples();
                    break;
                case 5:
                    // Creating Charts examples
                    runCreatingChartsExamples();
                    break;
                case 6:
                    // Working with Other Drawing Objects
                    runWorkingWithOtherDrawingObjectsExamples();
                    break;
                case 7:
                    // Advanced Topics
                    runAdvancedTopicsExamples();
                    break;
                case 8:
                    // Working with Tables
                    runWorkingWithTablesExamples();
                    break;
                case 9:
                    // Working with Formulas
                    runWorkingWithFormulasExamples();
                    break;
                case 10:
                    // Working with CellsHelper Methods
                    runWorkingWithCellsHelperMethodsExamples();
                    break;
                default:
                    break;
            }
        }
    };

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        listView = (ListView)findViewById(R.id.list_view);
        ArrayList<String> sectionsNames = new ArrayList<String>();
        sectionsNames.add("Working with Files");
        sectionsNames.add("Working with Worksheets");
        sectionsNames.add("Working with Rows and Columns");
        sectionsNames.add("Working with Data");
        sectionsNames.add("Working with Data Formatting");
        sectionsNames.add("Creating Charts");
        sectionsNames.add("Working with Other Drawing Objects");
        sectionsNames.add("Advanced Topics");
        sectionsNames.add("Working with Tables");
        sectionsNames.add("Working with Formulas");
        sectionsNames.add("Working with CellsHelper Methods");

        ArrayAdapter<String> arrayAdapter = new ArrayAdapter<String>(this, android.R.layout.simple_list_item_1, sectionsNames);
        listView.setAdapter(arrayAdapter);

        listView.setOnItemClickListener(sectionsListener);

        // Request write to external storage permission.
        requestWriteToExternalStoragePermission();
    }

    public void runWorkingWithFilesExamples() {
        // Open File
        OpenFile openFile = new OpenFile();
        openFile.openThroughPath();
        openFile.openThroughStream(this);
        openFile.openMicrosoftExcel97File();
        openFile.openMicrosoftExcel2007XLSXFile();
        openFile.openCSVFile();
        openFile.openTabDelimitedFile();
        openFile.openEncryptedExcelFile();

        // Save File
        SaveFile saveFile = new SaveFile();
        saveFile.saveToALocation();
        saveFile.saveToAStream(this);
        saveFile.saveEntireWorkbookIntoTextOrCSVFormat();

        // Document Properties
        DocumentProperties documentProperties = new DocumentProperties();
        documentProperties.getPropertyUsingNameOrIndex();
        documentProperties.retrieveNameValueAndTypeOfDocumentProperty();
        documentProperties.addCustomProperty();
        documentProperties.removeCustomProperty();

        // Encrypt File
        EncryptFile encryptFile = new EncryptFile();
        encryptFile.encryptFile();

        // Convert to MHTML
        ConvertToMHTML convertToMHTMLFiles = new ConvertToMHTML();
        convertToMHTMLFiles.convertToMHTML();
        convertToMHTMLFiles.ConvertFromMHTML();

        // Convert to PDF
        ConvertExcelToPDF convertExcelToPDF = new ConvertExcelToPDF();
        convertExcelToPDF.convertExcelToPDF();
        convertExcelToPDF.pdfAConversion();
        convertExcelToPDF.setCreationTimeForOutputPDF();

        // Convert Chart to Image
        ConvertChartToImage convertChartToImage = new ConvertChartToImage();
        convertChartToImage.convertChartToImage();

        // Convert Worksheet to Image
        ConvertWorksheetToImage convertWorksheetToImage = new ConvertWorksheetToImage();
        convertWorksheetToImage.convertToImage();

        // Convert Worksheet to SVG
        ConvertWorksheetToSVG convertWorksheetToSVG = new ConvertWorksheetToSVG();
        convertWorksheetToSVG.convertWorksheetToSVG();

        // Convert Excel Files to HTML
        ConvertExcelFilesToHTML convertExcelToHTML = new ConvertExcelFilesToHTML();
        convertExcelToHTML.convertExcelToHTML();
        convertExcelToHTML.setImagePreferencesForHTML();

        // Convert Excel Files to XPS
        ConvertExcelFilesToXPS convertExcelToXPS = new ConvertExcelFilesToXPS();
        convertExcelToXPS.convertSingleWorksheetToXPS();
        convertExcelToXPS.quickExcelToXPSConversion();

        // Convert Chart to PDF
        ConvertChartToPDF convertChartToPDF = new ConvertChartToPDF();
        convertChartToPDF.convertChartToPDF();
        convertChartToPDF.saveChartPDFInByteArrayOutputStream();
    }

    public void runWorkingWithWorksheetsExamples() {
        // Manage Worksheets
        ManagingWorksheets managingWorksheets = new ManagingWorksheets();
        managingWorksheets.createANewExcelFile();
        managingWorksheets.addWorksheetToADesignerSpreadsheet();
        managingWorksheets.accessWorksheetUsingSheetName();
        managingWorksheets.removeWorksheetUsingSheetName();
        managingWorksheets.removeWorksheetUsingSheetIndex();

        // Hide or Show a Worksheet
        HideOrShowAWorksheet hideOrShowAWorksheet = new HideOrShowAWorksheet();
        hideOrShowAWorksheet.hideAWorksheet();
        hideOrShowAWorksheet.makeAWorksheetVisible();
        hideOrShowAWorksheet.setVisibilityType();

        // Hide or Show tabs
        HideOrShowTabs hideOrShowTabs = new HideOrShowTabs();
        hideOrShowTabs.hideTabs();
        hideOrShowTabs.makeTabsVisible();

        // Hide or Show ScrollBars
        HideOrShowScrollBars hideOrShowScrollBars = new HideOrShowScrollBars();
        hideOrShowScrollBars.hideScrollBars();
        hideOrShowScrollBars.makeScrollBarsVisible();

        // Hide or Show Row Column Headers
        HideOrShowRowColumnHeaders hideOrShowRowColumnHeaders = new HideOrShowRowColumnHeaders();
        hideOrShowRowColumnHeaders.hideRowAndColumnHeaders();
        hideOrShowRowColumnHeaders.showRowAndColumnHeaders();

        // Page Break Preview
        PageBreakPreview pageBreakPreview = new PageBreakPreview();
        pageBreakPreview.enableNormalView();
        pageBreakPreview.enablePageBreakPreview();

        // Zoom Factor
        ZoomFactor zoomFactor = new ZoomFactor();
        zoomFactor.controllingTheZoomFactor();

        // Freeze Panes
        FreezePanes freezePanes = new FreezePanes();
        freezePanes.setFreezePanes();

        // Split Panes
        SplitPanes splitPanes = new SplitPanes();
        splitPanes.splitPanes();
        splitPanes.removePanes();

        // Protect Worksheets
        ProtectWorksheet protectWorksheet = new ProtectWorksheet();
        protectWorksheet.protectAWorksheet();
        protectWorksheet.protectCells();
        protectWorksheet.protectARow();
        protectWorksheet.protectAColumn();

        // Advanced Protection Settings since Excel XP
        AdvancedProtectionSettingsSinceExcelXP protectionSettings = new AdvancedProtectionSettingsSinceExcelXP();
        protectionSettings.advanceProtectionSettings();
        protectionSettings.lockCells();

        // Unprotect a Worksheet
        UnprotectWorksheet unprotectWorksheet = new UnprotectWorksheet();
        unprotectWorksheet.unprotectASimplyProtectedWorksheet();
        unprotectWorksheet.unprotectAPasswordProtectedWorksheet();

        // Setting Page Options
        PageOptions pageOptions = new PageOptions();
        pageOptions.pageOrientation();
        pageOptions.scalingFactor();
        pageOptions.fitToPageOptions();
        pageOptions.paperSize();
        pageOptions.printQuality();
        pageOptions.firstPageNumber();

        // Page Margins
        SetMargins margins = new SetMargins();
        margins.pageMargins();
        margins.centerOnPage();
        margins.headerAndFooterMargins();

        // Set Headers and Footers
        HeadersAndFooters headersAndFooters = new HeadersAndFooters();
        headersAndFooters.setHeadersAndFooters();
        headersAndFooters.insertAGraphicInAHeaderOrFooter();

        // Set Print Options
        PrintOptions printOptions = new PrintOptions();
        printOptions.setPrintArea();
        printOptions.setPrintTitles();
        printOptions.setOtherPrintOptions();
        printOptions.setPageOrder();

        // Manage Page Breaks
        ManagePageBreaks pageBreaks = new ManagePageBreaks();
        pageBreaks.addPageBreaks();
        pageBreaks.clearAllPageBreaks();
        pageBreaks.removeSpecificPageBreak();

        CopyAndMoveWorksheet copyAndMoveWorksheet = new CopyAndMoveWorksheet();
        copyAndMoveWorksheet.copyWorksheetsWithinAWorkbook();
        copyAndMoveWorksheet.copyWorksheetsBetweenWorkbooks();
        copyAndMoveWorksheet.copyAWorksheetDataFromOneWorkbookToAnother();
        copyAndMoveWorksheet.moveWorksheetsWithinWorkbook();
    }

    public void runWorkingWithRowsAndColumnsExamples() {
        // Inserting Deleting Rows and Columns
        ManagingRowsAndColumns rowsAndColumns = new ManagingRowsAndColumns();
        rowsAndColumns.insertRows();
        rowsAndColumns.deleteRows();
        rowsAndColumns.insertColumns();
        rowsAndColumns.deleteColumns();

        // Hiding and Showing Rows and Columns
        HideAndShowRowsAndColumns hideAndShowRowsAndColumns = new HideAndShowRowsAndColumns();
        hideAndShowRowsAndColumns.hideRowsAndColumns();
        hideAndShowRowsAndColumns.showRowsAndColumns();

        // Grouping Ungrouping Rows and Columns
        GroupUngroupRowsAndColumns groupUngroupRowsAndColumns = new GroupUngroupRowsAndColumns();
        groupUngroupRowsAndColumns.groupRowsAndColumns();
        groupUngroupRowsAndColumns.ungroupRowsAndColumns();
        groupUngroupRowsAndColumns.summaryRowsBelowDetail();

        // Adjusting Row Height and Column Width
        AdjustRowHeightAndColumnWidth adjustRowHeightAndColumnWidth = new AdjustRowHeightAndColumnWidth();
        adjustRowHeightAndColumnWidth.setRowHeight();
        adjustRowHeightAndColumnWidth.setRowHeightForAllRows();
        adjustRowHeightAndColumnWidth.setColumnWidth();
        adjustRowHeightAndColumnWidth.setWidthOfAllColumns();

        // AutoFit Rows and Columns
        AutoFitRowsAndColumns autoFitRowsAndColumns = new AutoFitRowsAndColumns();
        autoFitRowsAndColumns.autoFitRow();
        autoFitRowsAndColumns.autoFitRowInARangeOfCells();
        autoFitRowsAndColumns.autoFitColumn();
        autoFitRowsAndColumns.autoFitColumnInARangeOfCells();

        // Copying Rows and Columns
        CopyRowsAndColumns copyRowsAndColumns = new CopyRowsAndColumns();
        copyRowsAndColumns.copyRows();
        copyRowsAndColumns.copyColumns();
    }

    public void runWorkingWithDataExamples() {
        // Accessing Worksheet Cells
        AccessCells accessCells = new AccessCells();
        accessCells.accessUsingCellName();

        // Adding Data to Cells
        AddDataToCells addDataToCells = new AddDataToCells();
        addDataToCells.addDataToCells();

        // Find or Search Data
        FindOrSearchData findOrSearchData = new FindOrSearchData();
        // Finding Cells that Contain a Formula
        findOrSearchData.findCellsThatContainAFormula();
        // Searching with Partial Formula
        findOrSearchData.searchWithPartialFormula();
        // Searching for Strings that Start with Specific Characters
        findOrSearchData.searchForStringsThatStartWithSpecificCharacters();
        // Searching for Strings that End with Specific Characters
        findOrSearchData.searchForStringsThatEndWithSpecificCharacters();
        // Searching with Regular Expressions: the RegEx Feature
        findOrSearchData.searchWithRegularExpressions();

        // Data Sorting
        DataSorting dataSorting = new DataSorting();
        dataSorting.dataSorting();

        // Import Data to Worksheets
        ImportDataToWorksheets importDataToWorksheets = new ImportDataToWorksheets();
        // Importing from Array
        importDataToWorksheets.importFromArray();
        // Importing from Multi-dimensional Arrays
        importDataToWorksheets.importFromMultiDimensionalArray();
        // Importing from an ArrayList
        importDataToWorksheets.importFromArrayList();

        // Importing from ResultSet
        // Before calling importFromResultSet(), save Resource file to External Storage
        //Utils.saveFileToExternalStorage(MainActivity.this, "Northwind.mdb");
        //importDataToWorksheets.importFromResultSet();

        // Exporting Data from Worksheets
        ExportDataFromWorksheets exportDataFromWorksheets = new ExportDataFromWorksheets();
        exportDataFromWorksheets.exportDataToArray();

        // Tracing Precedents and Dependents
        TracePrecedentsAndDependents tracePrecedentsAndDependents = new TracePrecedentsAndDependents();
        tracePrecedentsAndDependents.tracePrecedent();
        tracePrecedentsAndDependents.traceDependents();

        // Accessing a Worksheet's Maximum Display Range
        AccessAWorksheetMaximumDisplayRange accessAWorksheetMaximumDisplayRange = new AccessAWorksheetMaximumDisplayRange();
        accessAWorksheetMaximumDisplayRange.accessAWorksheetMaximumDisplayRange();

        // Accessing a Worksheet's Maximum Display Range
        DataFilteringAndValidation dataFilteringAndValidation = new DataFilteringAndValidation();
        // Autofilter Data
        dataFilteringAndValidation.autofilter();
        // To filter columns with specified values, developers may also call the AutoFilter class' Filter method.
        dataFilteringAndValidation.filterColumnsWithSpecifiedValues();
        // Advanced Auto-Filter Options
        dataFilteringAndValidation.autoFilterOptions();
        // Whole Number Data Validation
        dataFilteringAndValidation.wholeNumberDataValidation();
        // Decimal Data Validation
        dataFilteringAndValidation.decimalDataValidation();
        // List Data Validation
        dataFilteringAndValidation.listDataValidation();
        // Date Data Validation
        dataFilteringAndValidation.dateDataValidation();
        // Time Data Validation
        dataFilteringAndValidation.timeDataValidation();
        // Text Length Data Validation
        dataFilteringAndValidation.textLengthDataValidation();

        // Creating Subtotals
        CreatingSubtotals creatingSubtotals = new CreatingSubtotals();
        creatingSubtotals.creatingSubtotals();

        // Add on Features
        // Adding Hyperlinks to Link Data
        AddHyperlinks addHyperlinks = new AddHyperlinks();
        // Adding a URL Link
        addHyperlinks.addAURLLink();
        addHyperlinks.applyFormattingToLookLikeHyperlink();
        // Adding a Link to Another Cell in the Same File
        addHyperlinks.addALinkToAnotherCellInTheSameFile();
        // Adding a Link to an External File
        addHyperlinks.addALinkToAnExternalFile();

        // Merging and Unmerging (Splitting) Cells
        MergeAndUnmergeCells mergeAndUnmergeCells = new MergeAndUnmergeCells();
        // Merging Cells in a Worksheet
        mergeAndUnmergeCells.mergeCellsInAWorksheet();
        // Unmerging (Splitting) Merged Cells
        mergeAndUnmergeCells.unmergeMergedCells();

        // Named Ranges
        NamedRanges namedRanges = new NamedRanges();
        // Create a Named Range
        namedRanges.createANamedRange();
        // Access All Named Ranges in a File
        namedRanges.accessAllNamedRangesInAFile();
        // Accessing a Specific Named Range
        namedRanges.accessASpecificNamedRange();
        // Inputting Data into a Named Range
        namedRanges.inputDataIntoANamedRange();
        // Setting Background Color and Font Attributes
        namedRanges.setBackgroundColorAndFontAttributes();
        // Adding Borders to a Named Range
        namedRanges.addBordersToANamedRange();
        // Converting Cells Address to Range or CellArea
        namedRanges.convertCellsAddressToRangeOrCellArea();
        // Setting a Simple Formula for Named Range
        namedRanges.setASimpleFormulaForNamedRange();
        // Setting a Complex Formula for Named Range
        namedRanges.setAComplexFormulaForNamedRange();
        namedRanges.useANamedRangeToSumValuesFrom2CellsInDifferentWorksheets();
    }

    public void runWorkingWithDataFormattingExamples() {
        // Basic Formatting
        // Formatting Data in Cells
        FormatCells formatCells = new FormatCells();
        // Using the setStyle Method
        formatCells.usingTheSetStyleMethod();
        // Using the Style Object
        formatCells.usingTheStyleObject();

        // Setting Display Formats for Numbers and Dates
        SetDisplayFormats setDisplayFormats = new SetDisplayFormats();
        // Using Built-in Number Formats
        setDisplayFormats.usingBuiltInNumberFormats();
        // Using Custom Number Formats
        setDisplayFormats.usingCustomNumberFormats();

        // Look and Feel
        // Configuring Alignment Settings
        ConfigureAlignmentSettings configureAlignmentSettings = new ConfigureAlignmentSettings();
        // Text Alignment - Horizontal
        configureAlignmentSettings.horizontalTextAlignment();
        // Text Alignment - Vertical
        configureAlignmentSettings.verticalTextAlignment();
        // Indentation
        configureAlignmentSettings.indentation();
        // Orientation
        configureAlignmentSettings.orientation();
        // Text Controls
        // Wrap Text
        configureAlignmentSettings.wrapText();
        // Shrink to Fit
        configureAlignmentSettings.shrinkToFit();
        // Merge Cells
        configureAlignmentSettings.mergeCells();
        // Text Direction
        configureAlignmentSettings.textDirection();

        // Dealing with Font Settings
        FontSettings fontSettings = new FontSettings();
        // Setting Font Name
        fontSettings.setFontName();
        // Setting Font Style to Bold
        fontSettings.setFontStyleToBold();
        // Setting Font Size
        fontSettings.setFontSize();
        // Setting Font Underline Type
        fontSettings.setFontUnderline();
        // Setting Font Color
        fontSettings.setFontColor();
        // Setting Strike Out Effect on Font
        fontSettings.setStrikeOutEffectOnFont();
        // Setting SubScript Effect on Font
        fontSettings.setSubScriptEffectOnFont();
        // Setting SuperScript Effect on Font
        fontSettings.setSuperScriptEffectOnFont();

        // Colors and Palette
        ColorAndPalette colorAndPalette = new ColorAndPalette();
        // Adding Custom Colors to Palette
        colorAndPalette.addCustomColorsToPalette();

        // Formatting Selected Characters in a Cell
        FormatSelectedCharacters formatSelectedCharacters = new FormatSelectedCharacters();
        formatSelectedCharacters.formatSelectedCharacters();

        // Adding Borders to Cells
        AddBordersToCells addBordersToCells = new AddBordersToCells();
        // Adding Borders to a Cell
        addBordersToCells.addBordersToACell();
        // Adding Borders to a Range of Cells
        addBordersToCells.addBordersToARangeOfCells();

        // Colors and Background Patterns
        ColorsAndBackgroundPatterns colorsAndBackgroundPatterns = new ColorsAndBackgroundPatterns();
        colorsAndBackgroundPatterns.setColorsAndBackgroundPatterns();

        // Advanced Formatting
        // Activating Sheets and Making an Active Cell in the Worksheet
        ActivateSheetsAndMakeAnActiveCell activateSheets = new ActivateSheetsAndMakeAnActiveCell();
        activateSheets.activateSheetAndMakeAnActiveCell();

        // Formatting Rows and Columns
        FormatRowsAndColumns formatRowsAndColumns = new FormatRowsAndColumns();
        // Formatting a Row
        formatRowsAndColumns.formatARow();
        // Formatting a Column
        formatRowsAndColumns.formatAColumn();
        // Setting Display Format of Numbers & Dates for Rows & Columns
        formatRowsAndColumns.setDisplayFormatOfNumbersAndDatesForRowsAndColumns();

        // Conditional Formatting
        ConditionalFormatting conditionalFormatting = new ConditionalFormatting();
        // Add and Delete Conditional Formatting
        conditionalFormatting.applyConditionalFormatting();
        // Set Font
        conditionalFormatting.setFont();
        // Set Border
        conditionalFormatting.setBorder();
        // Set Pattern
        conditionalFormatting.setPattern();
    }

    public void runCreatingChartsExamples() {
        // Getting Started with Charts
        CreateASimpleChart createChart = new CreateASimpleChart();
        // Creating a Simple Chart
        createChart.createAChart();

        // Change the Chart's Position and Size
        ChangeChartPositionAndSize changeChartPositionAndSize = new ChangeChartPositionAndSize();
        changeChartPositionAndSize.changeChartPositionAndSize();

        // Set Chart Appearance
        ChartAppearance chartAppearance = new ChartAppearance();
        // Setting Chart Area
        chartAppearance.setChartArea();
        // Setting Chart Lines
        chartAppearance.setChartLines();
        // Applying Microsoft Excel 2007/2010 Themes to Charts
        chartAppearance.applyMicrosoftExcel20072010ThemesToCharts();
        // Setting the Titles of Charts or Axes
        chartAppearance.setTitlesOfChartsOrAxes();
        // Setting Major Gridlines
        // Hiding Major Gridlines
        chartAppearance.hidingMajorGridlines();
        // Changing Major Gridlines Settings
        chartAppearance.changingMajorGridlinesSettings();
        // Setting Borders for Back and Side Walls
        chartAppearance.setBordersForBackAndSideWalls();

        // Setting Chart Data
        SetChartData chartData = new SetChartData();
        // Chart Data
        chartData.chartData();
        // Category Data
        chartData.categoryData();
        // Adds a column chart to the worksheet.
        chartData.setChartAndCategoryData();

        // Inserting Controls in Excel Charts
        InsertControlsInExcelCharts insertControls = new InsertControlsInExcelCharts();
        // Adding Label Control to the Chart
        insertControls.addLabelControlToTheChart();
        // Adding TextBox Control to the Chart
        insertControls.addTextBoxControlToTheChart();
        // Adding Picture to the Chart
        insertControls.addPictureToTheChart();

        // Applying 3D Format to Chart
        Apply3DFormatToChart apply3DFormatToChart = new Apply3DFormatToChart();
        apply3DFormatToChart.set3DFormatToChart();

        // Advanced Features
        ManipulateDesignerCharts designerCharts = new ManipulateDesignerCharts();
        // Creating a Chart
        designerCharts.createAChart();
        // Manipulating the Chart
        designerCharts.manipulateTheChart();
        // Manipulating a Line Chart in the Designer Template
        designerCharts.manipulateALineChartInTheDesignerTemplate();
        // Applying Microsoft Excel 2007/2010 Themes to Charts
        designerCharts.applyMicrosoftExcel20072010ThemesToCharts();

        // Creating Custom Charts
        CustomChart customChart = new CustomChart();
        customChart.createCustomChart();

        // Using Sparklines
        UsingSparklines usingSparklines = new UsingSparklines();
        usingSparklines.usingSparklines();
    }

    public void runWorkingWithOtherDrawingObjectsExamples() {
        // Adding Pictures
        Pictures pictures = new Pictures();
        // Adding Pictures
        pictures.addPicture();
        // Positioning Pictures
        pictures.positionPicture();

        // Adding Comments
        Comments comments = new Comments();
        // Adding a Comment
        comments.addComment();
        // Formatting Comments
        comments.formatComment();

        // Working with Controls
        Controls controls = new Controls();
        // Adding a Text Box Control to the Worksheet
        controls.addATextBoxControl();
        // Manipulating TextBox Controls in Designer Spreadsheets
        controls.manipulateTextBoxControlsInDesignerSpreadsheets();
        // Adding Checkbox Control to a Worksheet
        controls.addCheckboxControl();
        // Adding a Radio Button Control to a Worksheet
        controls.addARadioButton();
        // Adding ComboBox Control to the Worksheet
        controls.addComboBoxControl();
        // Adding Label Control to the Worksheet
        controls.addLabelControl();
        // Adding ListBox Control to the Worksheet
        controls.addListBoxControl();
        // Adding Button Control to the Worksheet
        controls.addButtonControl();
        // Adding Line Control to the Worksheet
        controls.addLineControl();
        // Adding an ArrowHead to the Line
        controls.addAnArrowHead();
        // Adding Rectangle Control to the Worksheet
        controls.addRectangleControl();
        // Adding Arc Control to the Worksheet
        controls.addArcControl();
        // Adding Oval Control to the Worksheet
        controls.addOvalControl();

        // Managing OLE Objects
        ManageOLEObjects oleObjects = new ManageOLEObjects();
        // Inserting OLE Objects into a Worksheet
        oleObjects.insertOLEObject();
        // Extracting OLE Objects in the Workbook
        oleObjects.extractOLEObject();
    }

    public void runAdvancedTopicsExamples() {
        // Designer Spreadsheet & Smart Markers
        SmartMarkers smartMarkers = new SmartMarkers();
        smartMarkers.smartMarkers();

        // Formula Calculation Engine
        FormulaCalculationEngine calculationEngine = new FormulaCalculationEngine();
        // Adding Formulas & Calculating Results
        calculationEngine.addFormulasAndCalculateResults();

        // Changing a Pivot Table's Source Data
        ChangeAPivotTableSourceData pivotTableSourceData = new ChangeAPivotTableSourceData();
        pivotTableSourceData.changeAPivotTableSourceData();

        // Create a Pivot Table Report
        CreateAPivotTableReport pivotTableReport = new CreateAPivotTableReport();
        pivotTableReport.createASimplePivotTable();

        // Customizing the Appearance of Pivot Table Reports
        CustomizeTheAppearanceOfPivotTable appearanceOfPivotTable = new CustomizeTheAppearanceOfPivotTable();
        // Clearing PivotFields
        PivotTable pivotTable = appearanceOfPivotTable.clearPivotFields();
        // Setting the AutoFormat and PivotTableStyle Types
        appearanceOfPivotTable.setTheAutoFormatAndPivotTableStyleTypes(pivotTable);
        // Setting Row, Column, and Page Fields Format
        appearanceOfPivotTable.setRowColumnAndPageFieldFormat(pivotTable);
        // Setting Data Fields Format
        appearanceOfPivotTable.setDataFieldsFormat(pivotTable);
        // Modify a Pivot Table Quick Style
        appearanceOfPivotTable.modifyAPivotTableQuickStyle(pivotTable);

        // Applying ConsolidationFunction to Data Fields of Pivot Table
        ApplyConsolidationFunctionToDataFieldsOfPivotTable consolidationFunctionToDataFieldsOfPivotTable = new ApplyConsolidationFunctionToDataFieldsOfPivotTable();
        consolidationFunctionToDataFieldsOfPivotTable.applyConsolidationFunctionToDataFieldsOfPivotTable();

        // Create Pivot Tables and Pivot Charts
        CreatePivotTablesAndPivotCharts pivotTablesAndPivotCharts = new CreatePivotTablesAndPivotCharts();
        pivotTablesAndPivotCharts.createPivotTablesAndPivotCharts();
    }

    public void runWorkingWithTablesExamples() {
        // Convert an Excel Table to a Range of Data
        ConvertTableToRangeOfData convertTableToRangeOfData = new ConvertTableToRangeOfData();
        convertTableToRangeOfData.convertTableToRange();

        // Create a List Object
        CreateAListObject listObject = new CreateAListObject();
        listObject.createAList();

        // Formatting a List Object
        FormatAListObject formatAListObject = new FormatAListObject();
        formatAListObject.formatAListObject();
    }

    public void runWorkingWithFormulasExamples() {
        // Calculating Formulas
        CalculateFormulas calculateFormulas = new CalculateFormulas();
        // Adding Formulas & Calculating Results
        calculateFormulas.addFormulasAndCalculateResults();

        // Using Formulas or Functions to Process Data
        UsingFormulasToProcessData formulasToProcessData = new UsingFormulasToProcessData();
        // Using Built-in Functions
        formulasToProcessData.usingBuiltInFunctions();
        // Using Add-in Functions
        formulasToProcessData.usingAddInFunctions();
        // Using Array Formula
        formulasToProcessData.usingArrayFormula();
        // Using R1C1 Formula
        formulasToProcessData.usingR1C1Formula();

        // Calculating Formulas without Adding them to a Worksheet
        CalculateFormulasDirectly calculateFormulasDirectly = new CalculateFormulasDirectly();
        calculateFormulasDirectly.calculateFormulasDirectly();

        // Calculating Formulas Once Only
        CalculateFormulasOnceOnly calculateFormulasOnceOnly = new CalculateFormulasOnceOnly();
        calculateFormulasOnceOnly.calculateFormulasOnceOnly();
    }

    public void runWorkingWithCellsHelperMethodsExamples() {
        // Detect File Format
        DetectFileFormat detectFileFormat = new DetectFileFormat();
        detectFileFormat.detectFileFormat();

        // Merging Files
        MergeFiles mergeFiles = new MergeFiles();
        mergeFiles.mergeFiles();

        // Getting Cell Name from Row and Column Indices
        GetCellNameFromRowAndColumnIndices rowAndColumnIndices = new GetCellNameFromRowAndColumnIndices();
        rowAndColumnIndices.getCellNameFromRowAndColumnIndices();

        // Get Row and Column Indices from Cell Name
        GetRowAndColumnIndicesFromCellName rowAndColumnIndicesFromCellName = new GetRowAndColumnIndicesFromCellName();
        rowAndColumnIndicesFromCellName.getRowAndColumnIndicesFromCellName();
    }

    public void saveAssetsFilesToExternalStorage() {
        new SaveFilesToExternalStorageTask().execute();
    }

    private class SaveFilesToExternalStorageTask extends AsyncTask<Void, Integer, Void> {
        private ProgressDialog progressDialog = new ProgressDialog(MainActivity.this);

        @Override
        protected void onPreExecute() {
            super.onPreExecute();
            progressDialog.setMessage(getString(R.string.saving_resource_files_external_storage));
            progressDialog.setProgressStyle(ProgressDialog.STYLE_HORIZONTAL);
            progressDialog.setIndeterminate(false);
            progressDialog.setMax(13);
            progressDialog.setCancelable(false);
            progressDialog.show();
        }

        @Override
        protected Void doInBackground(Void... voids) {
            // Save Resource Files To External Storage
            Utils.saveFileToExternalStorage(MainActivity.this, "Book1.xlsx");
            publishProgress(Integer.valueOf(1));
            Utils.saveFileToExternalStorage(MainActivity.this, "Book1.xls");
            publishProgress(Integer.valueOf(2));
            Utils.saveFileToExternalStorage(MainActivity.this, "Book1.csv");
            publishProgress(Integer.valueOf(3));
            Utils.saveFileToExternalStorage(MainActivity.this, "Book1.txt");
            publishProgress(Integer.valueOf(4));
            Utils.saveFileToExternalStorage(MainActivity.this, "source.xlsx");
            publishProgress(Integer.valueOf(5));
            Utils.saveFileToExternalStorage(MainActivity.this, "Source.html");
            publishProgress(Integer.valueOf(6));
            Utils.saveFileToExternalStorage(MainActivity.this, "sample.xlsx");
            publishProgress(Integer.valueOf(7));
            Utils.saveFileToExternalStorage(MainActivity.this, "footer.jpg");
            publishProgress(Integer.valueOf(8));
            Utils.saveFileToExternalStorage(MainActivity.this, "school.jpg");
            publishProgress(Integer.valueOf(9));
            Utils.saveFileToExternalStorage(MainActivity.this, "mergingcells.xls");
            publishProgress(Integer.valueOf(10));
            Utils.saveFileToExternalStorage(MainActivity.this, "pivot.xlsm");
            publishProgress(Integer.valueOf(11));
            Utils.saveFileToExternalStorage(MainActivity.this, "Template.xlsx");
            publishProgress(Integer.valueOf(12));
            Utils.saveFileToExternalStorage(MainActivity.this, "PageBreaks.xls");
            publishProgress(Integer.valueOf(13));

            return null;
        }

        @Override
        protected void onProgressUpdate(Integer... values) {
            super.onProgressUpdate(values);
            progressDialog.setProgress(values[0]);
        }

        @Override
        protected void onPostExecute(Void aVoid) {
            progressDialog.dismiss();
        }
    }

    public void requestWriteToExternalStoragePermission() {

        if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                != PackageManager.PERMISSION_GRANTED) {

            if (ActivityCompat.shouldShowRequestPermissionRationale(this,
                    Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
                Toast.makeText(this, getString(R.string.write_to_external_storage_permission), Toast.LENGTH_SHORT).show();
            } else {
                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                        REQUEST_WRITE_EXTERNAL_STORAGE);
            }
        } else {
            saveAssetsFilesToExternalStorage();
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String permissions[], int[] grantResults) {
        switch (requestCode) {
            case REQUEST_WRITE_EXTERNAL_STORAGE: {

                if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                    saveAssetsFilesToExternalStorage();
                } else {
                    // permission denied
                    Toast.makeText(this, getString(R.string.permission_denied_external_storage), Toast.LENGTH_SHORT).show();
                }
                return;
            }
        }
    }
}
