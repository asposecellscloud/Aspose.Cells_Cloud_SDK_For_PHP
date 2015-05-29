<?php

use Aspose\Cells\CellsApi;

class CellsApiTest extends PHPUnit_Framework_TestCase {
    
    protected $cells;

    protected function setUp()
    {        
        $this->cells = new CellsApi();
    }
    
    public function testDeleteDecryptDocument()
    {
        $body = array("Password" => "123456");
        $result = $this->cells->DeleteDecryptDocument($name="test_cells.xlsx", $storage = null, $folder = null, $body);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteDocumentProperties()
    {
        $result = $this->cells->DeleteDocumentProperties($name="test_convert_cell.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteDocumentProperty()
    {
        $result = $this->cells->DeleteDocumentProperty($name="test_convert_cell.xlsx", $propertyName="Author", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteDocumentUnProtectFromChanges()
    {
        $result = $this->cells->DeleteDocumentUnProtectFromChanges($name="test_convert_cell.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteUnProtectDocument()
    {
        $body = array("Password" => "123456");
        $result = $this->cells->DeleteUnProtectDocument($name="test_convert_cell.xlsx", $storage = null, $folder = null, $body);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteUnprotectWorksheet()
    {
        $body = array("Password" => "123456");
        $result = $this->cells->DeleteUnprotectWorksheet($name="test_convert_cell.xlsx", $sheetName="Sheet1", $storage = null, $folder = null, $body);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorkSheetBackground()
    {
        $result = $this->cells->DeleteWorkSheetBackground($name="test_convert_cell.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorkSheetComment()
    {
        $result = $this->cells->DeleteWorkSheetComment($name="test_cells.xlsx", $sheetName="Sheet1", $cellName="A2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorkSheetHyperlink()
    {
        $result = $this->cells->DeleteWorkSheetHyperlink($name="test_cells.xlsx", $sheetName="Sheet3", $hyperlinkIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorkSheetHyperlinks()
    {
        $result = $this->cells->DeleteWorkSheetHyperlinks($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorkSheetPictures()
    {
        $result = $this->cells->DeleteWorkSheetPictures($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorkSheetValidation()
    {
        $result = $this->cells->DeleteWorkSheetValidation($name="test_cells.xlsx", $sheetName="Sheet3", $validationIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheet()
    {
        $result = $this->cells->DeleteWorksheet($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetChartLegend()
    {
        $result = $this->cells->DeleteWorksheetChartLegend($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetChartTitle()
    {
        $result = $this->cells->DeleteWorksheetChartTitle($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetClearCharts()
    {
        $result = $this->cells->DeleteWorksheetClearCharts($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    /*public function testDeleteWorksheetColumns()
    {
        $result = $this->cells->DeleteWorksheetColumns($name="test_cells.xlsx", $sheetName="Sheet1", $columnIndex="0", $columns="0", $updateReference=true, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }*/
    
    public function testDeleteWorksheetDeleteChart()
    {
        $result = $this->cells->DeleteWorksheetDeleteChart($name="test_cells.xlsx", $sheetName="Sheet2", $chartIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetFreezePanes()
    {
        $result = $this->cells->DeleteWorksheetFreezePanes($name="test_cells.xlsx", $sheetName="Sheet3", $row=1, $column=1, $freezedRows=1, $freezedColumns=1, $folder = null, $storage = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetOleObject()
    {
        $result = $this->cells->DeleteWorksheetOleObject($name="test_cells.xlsx", $sheetName="Sheet2", $oleObjectIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetOleObjects()
    {
        $result = $this->cells->DeleteWorksheetOleObjects($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetPicture()
    {
        $result = $this->cells->DeleteWorksheetPicture($name="test_cells.xlsx", $sheetName="Sheet2", $pictureIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetPivotTable()
    {
        $result = $this->cells->DeleteWorksheetPivotTable($name="test_cells.xlsx", $sheetName="Sheet4", $pivotTableIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetPivotTables()
    {
        $result = $this->cells->DeleteWorksheetPivotTables($name="test_cells.xlsx", $sheetName="Sheet4", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetRow()
    {
        $result = $this->cells->DeleteWorksheetRow($name="test_cells.xlsx", $sheetName="Sheet3", $rowIndex=1, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testDeleteWorksheetRows()
    {
        $result = $this->cells->DeleteWorksheetRows($name="test_cells.xlsx", $sheetName="Sheet3", $startrow=1, $totalRows=10, $updateReference = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetChartArea()
    {
        $result = $this->cells->GetChartArea($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetChartAreaBorder()
    {
        $result = $this->cells->GetChartAreaBorder($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetChartAreaFillFormat()
    {
        $result = $this->cells->GetChartAreaFillFormat($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetDocumentProperties()
    {
        $result = $this->cells->GetDocumentProperties($name="test_cells.xlsx", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetDocumentProperty()
    {
        $result = $this->cells->GetDocumentProperty($name="test_cells.xlsx", $propertyName="Author", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetExtractBarcodes()
    {
        $result = $this->cells->GetExtractBarcodes($name="test_cells.xlsx", $sheetName="Sheet1", $pictureNumber="0", $storage = null, $folder = null);        
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBook()
    {
        $result = $this->cells->GetWorkBook($name="test_cells.xlsx", $password = null, $isAutoFit = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookDefaultStyle()
    {
        $result = $this->cells->GetWorkBookDefaultStyle($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookName()
    {
        $result = $this->cells->GetWorkBookName($name="test_cells.xlsx", $nameName="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookNames()
    {
        $result = $this->cells->GetWorkBookNames($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookTextItems()
    {
        $result = $this->cells->GetWorkBookTextItems($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkBookWithFormat()
    {
        $result = $this->cells->GetWorkBookWithFormat($name="test_cells.xlsx", $format="pdf", $password = null, $isAutoFit = null, $storage = null, $folder = null, $outPath = null);
        $fh = fopen(getcwd(). '/Data/Output/Workbook.pdf', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Workbook.pdf');
    }
    
    public function testGetWorkSheet()
    {
        $result = $this->cells->GetWorkSheet($name="test_cells.xlsx", $sheetName="Sheet1", $verticalResolution = null, $horizontalResolution = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetCalculateFormula()
    {
        $result = $this->cells->GetWorkSheetCalculateFormula($name="test_cells.xlsx", $sheetName="Sheet3", $formula="SUM(A3,A4)", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetComment()
    {
        $result = $this->cells->GetWorkSheetComment($name="test_cells.xlsx", $sheetName="Sheet1", $cellName="A2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetComments()
    {
        $result = $this->cells->GetWorkSheetComments($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetHyperlink()
    {
        $result = $this->cells->GetWorkSheetHyperlink($name="test_cells.xlsx", $sheetName="Sheet3", $hyperlinkIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetHyperlinks()
    {
        $result = $this->cells->GetWorkSheetHyperlinks($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetMergedCell()
    {
        $result = $this->cells->GetWorkSheetMergedCell($name="test_cells.xlsx", $sheetName="Sheet3", $mergedCellIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetMergedCells()
    {
        $result = $this->cells->GetWorkSheetMergedCells($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetTextItems()
    {
        $result = $this->cells->GetWorkSheetTextItems($name="test_cells.xlsx", $sheetName="Sheet3", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetValidation()
    {
        $result = $this->cells->GetWorkSheetValidation($name="test_cells.xlsx", $sheetName="Sheet1", $validationIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetValidations()
    {
        $result = $this->cells->GetWorkSheetValidations($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorkSheetWithFormat()
    {
        $result = $this->cells->GetWorkSheetWithFormat($name="test_cells.xlsx", $sheetName="Sheet1", $format="png", $verticalResolution = null, $horizontalResolution = null, $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Sheet1.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Sheet1.png');
    }
    
    public function testGetWorkSheets()
    {
        $result = $this->cells->GetWorkSheets($name="test_cells.xlsx", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetAutoshape()
    {
        $result = $this->cells->GetWorksheetAutoshape($name="test_cells.xlsx", $sheetName="Sheet2", $autoshapeNumber=2, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetAutoshapeWithFormat()
    {
        $result = $this->cells->GetWorksheetAutoshapeWithFormat($name="test_cells.xlsx", $sheetName="Sheet2", $autoshapeNumber=2, $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Autoshape.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Autoshape.png');
    }
    
    public function testGetWorksheetAutoshapes()
    {
        $result = $this->cells->GetWorksheetAutoshapes($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetCell()
    {
        $result = $this->cells->GetWorksheetCell($name="test_cells.xlsx", $sheetName="Sheet1", $cellOrMethodName="A1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetCellStyle()
    {
        $result = $this->cells->GetWorksheetCellStyle($name="test_cells.xlsx", $sheetName="Sheet1", $cellName="A1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetCells()
    {
        $result = $this->cells->GetWorksheetCells($name="test_cells.xlsx", $sheetName="Sheet1", $offest = null, $count = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetChart()
    {
        $result = $this->cells->GetWorksheetChart($name="test_cells.xlsx", $sheetName="Sheet1", $chartNumber="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetChartLegend()
    {
        $result = $this->cells->GetWorksheetChartLegend($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetChartWithFormat()
    {
        $result = $this->cells->GetWorksheetChartWithFormat($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Chart.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Chart.png');
    }
    
    public function testGetWorksheetCharts()
    {
        $result = $this->cells->GetWorksheetCharts($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetColumn()
    {
        $result = $this->cells->GetWorksheetColumn($name="test_cells.xlsx", $sheetName="Sheet1", $columnIndex=1, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetColumns()
    {
        $result = $this->cells->GetWorksheetColumns($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetOleObject()
    {
        $result = $this->cells->GetWorksheetOleObject($name="test_cells.xlsx", $sheetName="Sheet2", $objectNumber="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetOleObjectWithFormat()
    {
        $result = $this->cells->GetWorksheetOleObjectWithFormat($name="test_cells.xlsx", $sheetName="Sheet2", $objectNumber="0", $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Ole.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Ole.png');
    }
    
    public function testGetWorksheetOleObjects()
    {
        $result = $this->cells->GetWorksheetOleObjects($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPicture()
    {
        $result = $this->cells->GetWorksheetPicture($name="test_cells.xlsx", $sheetName="Sheet2", $pictureNumber="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPictureWithFormat()
    {
        $result = $this->cells->GetWorksheetPictureWithFormat($name="test_cells.xlsx", $sheetName="Sheet2", $pictureNumber="0", $format="png", $storage = null, $folder = null);
        $fh = fopen(getcwd(). '/Data/Output/Picture.png', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/Picture.png');
    }
    
    public function testGetWorksheetPictures()
    {
        $result = $this->cells->GetWorksheetPictures($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPivotTable()
    {
        $result = $this->cells->GetWorksheetPivotTable($name="test_cells.xlsx", $sheetName="Sheet1", $pivottableIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetPivotTables()
    {
        $result = $this->cells->GetWorksheetPivotTables($name="test_cells.xlsx", $sheetName="Sheet4", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetRow()
    {
        $result = $this->cells->GetWorksheetRow($name="test_cells.xlsx", $sheetName="Sheet2", $rowIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testGetWorksheetRows()
    {
        $result = $this->cells->GetWorksheetRows($name="test_cells.xlsx", $sheetName="Sheet2", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutAddNewWorksheet()
    {
        $result = $this->cells->PutAddNewWorksheet($name="test_cells.xlsx", $sheetName="NewSheet", $storage = null, $folder = null);
        $this->assertEquals('Created', $result->Status);
    }
    
    public function testPutChangeVisibilityWorksheet()
    {
        $result = $this->cells->PutChangeVisibilityWorksheet($name="test_cells.xlsx", $sheetName="Sheet1", $isVisible="true", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutConvertWorkBook()
    {
        $file = getcwd() . '/Data/Input/test_convert_cell.xlsx';
        $result = $this->cells->PutConvertWorkBook($format = "pdf", $password = null, $outPath = null, $file);
        $fh = fopen(getcwd(). '/Data/Output/test_convert_cell.pdf', 'w');
        fwrite($fh, $result);
        fclose($fh);
        $this->assertFileExists(getcwd(). '/Data/Output/test_convert_cell.pdf');
    }
    
    public function testPutDocumentProperty()
    {
        $body = array("Value"=>"Aspose");
        $result = $this->cells->PutDocumentProperty($name="test_cells.xlsx", $propertyName="Title", $storage = null, $folder = null, $body);
        $this->assertEquals('Created', $result->Status);
    }
    
    public function testPutDocumentProtectFromChanges()
    {
        $body = array("ProtectionType"=>"all", "Password"=>"123456");
        $result = $this->cells->PutDocumentProtectFromChanges($name="test_cells.xlsx", $storage = null, $folder = null, $body);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutInsertWorksheetColumns()
    {
        $result = $this->cells->PutInsertWorksheetColumns($name="test_cells.xlsx", $sheetName="Sheet1", $columnIndex=1, $columns=5, $updateReference = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutInsertWorksheetRow()
    {
        $result = $this->cells->PutInsertWorksheetRow($name="test_cells.xlsx", $sheetName="Sheet1", $rowIndex=1, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutInsertWorksheetRows()
    {
        $result = $this->cells->PutInsertWorksheetRows($name="test_cells.xlsx", $sheetName="Sheet1", $startrow=1, $totalRows=10, $updateReference = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorkSheetBackground()
    {
        $file = getcwd() . '/Data/Input/watermark.png';
        $result = $this->cells->PutWorkSheetBackground($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null, $file);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorkSheetComment()
    {
        $body = '{
                    "CellName": "A1",
                    "Author": "Masood",
                    "HtmlNote": "",
                    "Note": "Add comments",
                    "AutoSize": true,
                    "IsVisible": true,
                    "Width": 0,
                    "Height": 0,
                    "TextHorizontalAlignment": "",
                    "TextOrientationType": "",
                    "TextVerticalAlignment": "",
                    "link": {
                      "Href": "",
                      "Rel": "",
                      "Type": "",
                      "Title": ""
                    }
                  }';
        $result = $this->cells->PutWorkSheetComment($name="test_cells.xlsx", $sheetName="Sheet1", $cellName="A1", $storage = null, $folder = null, $body);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorkSheetHyperlink()
    {
        $result = $this->cells->PutWorkSheetHyperlink($name="test_cells.xlsx", $sheetName="Sheet1", $firstRow=1, $firstColumn=1, $totalRows=1, $totalColumns=1, $address="www.aspose.com", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorkSheetValidation()
    {
        $result = $this->cells->PutWorkSheetValidation($name="test_cells.xlsx", $sheetName="Sheet1", $range="A1:A5", $storage = null, $folder = null, $file = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorkbookCreate()
    {
        $file = getcwd() . '/Data/Input/test_convert_cell.xlsx';
        $result = $this->cells->PutWorkbookCreate($name="new_workbook.xlsx", $templateFile = null, $dataFile = null, $storage = null, $folder = null, $file);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorksheetAddChart()
    {
        $result = $this->cells->PutWorksheetAddChart($name="test_cells.xlsx", $sheetName="Sheet1", $chartType="bar", $upperLeftRow = 12, $upperLeftColumn = 12, $lowerRightRow = 20, $lowerRightColumn = 20, $area = null, $isVertical = null, $categoryData = null, $isAutoGetSerialName = null, $title = null, $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorksheetAddPicture()
    {
        $file = getcwd() . '/Data/Input/watermark.png';
        $result = $this->cells->PutWorksheetAddPicture($name="test_cells.xlsx", $sheetName="Sheet3", $upperLeftRow = 12, $upperLeftColumn = 12, $lowerRightRow = 20, $lowerRightColumn = 20, $picturePath = "watermark.png", $storage = null, $folder = null, $file);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorksheetChartLegend()
    {
        $result = $this->cells->PutWorksheetChartLegend($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorksheetChartTitle()
    {
        $body = array("Title"=>array("Text"=>"Sales Report"));
        $result = $this->cells->PutWorksheetChartTitle($name="test_cells.xlsx", $sheetName="Sheet1", $chartIndex="0", $storage = null, $folder = null, $body);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorksheetFreezePanes()
    {
        $result = $this->cells->PutWorksheetFreezePanes($name="test_cells.xlsx", $sheetName="Sheet1", $row=1, $column=1, $freezedRows=5, $freezedColumns=5, $folder = null, $storage = null);
        $this->assertEquals(200, $result->Code);
    }
    
    public function testPutWorksheetOleObject()
    {
        $body = '{
                    "DisplayAsIcon": true,
                    "FileFormatType": "",
                    "ImageSourceFullName": "watermark.png",
                    "IsAutoSize": true,
                    "IsLink": true,
                    "ProgID": "",
                    "SourceFullName": "ole.docx",
                    "Name": "OLE",
                    "MsoDrawingType": "",
                    "AutoShapeType": "",
                    "Placement": "",
                    "UpperLeftRow": 20,
                    "Top": 20,
                    "UpperLeftColumn": 20,
                    "Left": 20,
                    "LowerRightRow": 10,
                    "Bottom": 10,
                    "LowerRightColumn": 10,
                    "Right": 20,
                    "Width": 200,
                    "Height": 100,
                    "X": 0,
                    "Y": 0,
                    "RotationAngle": 0,
                    "HtmlText": "",
                    "Text": "ole object",
                    "AlternativeText": "no alternative text",
                    "TextHorizontalAlignment": "",
                    "TextHorizontalOverflow": "",
                    "TextOrientationType": "string",
                    "TextVerticalAlignment": "string",
                    "TextVerticalOverflow": "string",
                    "IsGroup": true,
                    "IsHidden": true,
                    "IsLockAspectRatio": true,
                    "IsLocked": true,
                    "IsPrintable": true,
                    "IsTextWrapped": true,
                    "IsWordArt": true,
                    "LinkedCell": "string",
                    "ZOrderPosition": 0
                  }';
        $result = $this->cells->PutWorksheetOleObject($name="test_cells.xlsx", $sheetName="Sheet3", $upperLeftRow = null, $upperLeftColumn = null, $height = null, $width = null, $oleFile = null, $imageFile = null, $storage = null, $folder = null, $body);
        $this->assertEquals(200, $result->Code);
    }
    
    /*public function testPutWorksheetPivotTable()
    {
        $body = '';
        $result = $this->cells->PutWorksheetPivotTable($name="test_cells.xlsx", $sheetName="Sheet1", $storage = null, $folder = null, $sourceData = null, $destCellName = null, $tableName = null, $useSameSource = null, $body);
        $this->assertEquals(200, $result->Code);
    }*/
                           
}    