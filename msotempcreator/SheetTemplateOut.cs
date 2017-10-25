using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Text.RegularExpressions;


//For Open XML - (s)pread(sheet) operations 
namespace msotempcreator
{
    class SheetTemplateOut : TemplateOut
    {
        private SpreadsheetDocument templateAndOutputFile;
        public SheetTemplateOut(string templateFilePath, string outputFilePath)
            : base(templateFilePath, outputFilePath)
        {
        }
        MemoryStream templateStream = new MemoryStream();

        override protected void defineDocument()
        {            
            //byte[] byteArray = File.ReadAllBytes(GeneralOptions.templatePath);            
            byte[] byteArray = File.ReadAllBytes(_templateFilePath);
            templateStream.Write(byteArray, 0, (int)byteArray.Length);
            templateAndOutputFile = SpreadsheetDocument.Open(templateStream, true);

        }

        

        protected override void replaceTempsWithDatasOnTables(DataGridView sourceView)
        {
            int sourceColCount = 0;
            while (sourceColCount < sourceView.ColumnCount)
            {
                Application.DoEvents();
                foreach (DataGridViewRow filteredRow in sourceView.Rows)
                {
                    DataRowCollection scanRows = scanTablesinFile("__" + sourceView.Columns[sourceColCount].Name).Rows;
                    foreach (DataRow scanResult in scanRows)
                    {
                        WorksheetPart currentWorkSheetPart = ((WorksheetPart)templateAndOutputFile.WorkbookPart.GetPartById(scanResult["sheetID"].ToString()));
                        addValuesToCells(sourceView.Columns[sourceColCount].Name, scanResult["columnName"].ToString(), Convert.ToUInt32(scanResult["row"]), currentWorkSheetPart,sourceView);

                    }
                }
                sourceColCount++;
            }

        }

        protected override void replaceFunctionsWithResultsOnTables(DataGridView sourceView)
        {
            runRowNumberFunction();


        }


        private void runRowNumberFunction()
        {
            DataRowCollection scanRows = scanTablesinFile("__ROWNUMBER__").Rows;

            foreach (DataRow scanResult in scanRows)
            {
                //make a worksheet Part according to sheet which is find at location by scanTablesinFile
                WorksheetPart currentWorkSheetPart = ((WorksheetPart)templateAndOutputFile.WorkbookPart.GetPartById(scanResult["sheetID"].ToString()));

                int counter = 1;
                uint rowIndex = (Convert.ToUInt32(scanResult["row"]));

                // starts to count and put row numbers while next cell is not empty
                do
                {
                    insertTextInWorksheet(counter.ToString(), scanResult["columnName"].ToString(), rowIndex, currentWorkSheetPart);
                    rowIndex++;
                    counter++;

                } while (getOffsetCellTextInWorksheet(scanResult["columnName"].ToString(), rowIndex, currentWorkSheetPart, 1) != "");

            }
        }

        private void addValuesToCells(string temp, string columnName, uint rowIndex, WorksheetPart worksheetPart,DataGridView sourceView)
        {            
            for (uint q = 0; q < sourceView.Rows.Count; q++)
            {
                insertTextInWorksheet("", columnName, rowIndex + q, worksheetPart);
                if (sourceView[temp, (int)q].Value != null && sourceView[temp, (int)q].GetType()!= typeof(DataGridViewButtonCell)) insertTextInWorksheet(sourceView[temp, (int)q].Value.ToString(), columnName, rowIndex + q, worksheetPart);
            }

        }

        

        // write text to a given location and convert cell data type according to text
        private void insertTextInWorksheet(string newText, string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            string cellReference = columnName + rowIndex;            
            WorksheetWriter writer = new WorksheetWriter(templateAndOutputFile, worksheetPart);
            //writer.PasteText(cellReference, newText);
            writer.PasteValue(cellReference,  general.ConvertDoubleStr(newText), ConvertTextToCellValuesType(newText));          
        }

        // write a formula to a given location and force file to calculate formula
        private void insertFormulaInWorksheet(string newFormula, string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            string cellReference = columnName + rowIndex;
            WorksheetWriter writer = new WorksheetWriter(templateAndOutputFile, worksheetPart);            

            Cell formulaCell = writer.FindCell(cellReference);
           
            formulaCell.CellFormula = new CellFormula(newFormula);            
            templateAndOutputFile.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            templateAndOutputFile.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;          

        }



        public string getCellTextInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            string cellReference = columnName + rowIndex;
            WorksheetWriter writer = new WorksheetWriter(templateAndOutputFile, worksheetPart);        
            Cell resultCell = writer.FindCell(cellReference);            
            return resultCell.CellValue == null ? "" : convertAnyTypeCellValueToString(resultCell);           

        }
        
        //gets cell referenced at first sheet of workbook
        public string getCellTextInWorksheet(string columnName, uint rowIndex)
        {           
            Sheet firstSheet = templateAndOutputFile.WorkbookPart.Workbook.Descendants<Sheet>().First();
            WorksheetPart firstWorksheetPart = ((WorksheetPart)templateAndOutputFile.WorkbookPart.GetPartById(firstSheet.Id));                        
            return getCellTextInWorksheet(columnName, rowIndex, firstWorksheetPart);

        }

        //gets cell referenced at given sheet of workbook (if sheet doesnt exceeds, throw exception
        public string getCellTextInWorksheet(string columnName, uint rowIndex,string sheetName)
        {
            Sheet sourceSheet = templateAndOutputFile.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (sourceSheet == null) throw new ArgumentException("No sheet"); 
            WorksheetPart firstWorksheetPart = ((WorksheetPart)templateAndOutputFile.WorkbookPart.GetPartById(sourceSheet.Id));
            return getCellTextInWorksheet(columnName, rowIndex, firstWorksheetPart);

        }        

        // get text of offseted cell of a given source location, example if offset count = 1 then it will take value of right next cell
        private string getOffsetCellTextInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart, int offsetCount)
        {
            
            WorksheetWriter writer = new WorksheetWriter(templateAndOutputFile, worksheetPart);
            string newColoumnName = convertToColumnName(convertToColumnIndex(columnName) + offsetCount);
            
            string cellReference = newColoumnName + rowIndex;
            Cell resultCell = writer.FindCell(cellReference);

            return resultCell.CellValue == null ? "" : resultCell.CellValue.Text;
        }

        protected override void replaceAllTemps(DataGridView sourceView)
        {
            //not needed because all platform in spreadsheet is table
        }

        protected override void saveDocument()
        {
            templateAndOutputFile.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            templateAndOutputFile.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true; 
            // before closing(and saving) templateAndOutputFile, firstly try if outputFile open or not. templateAndOutputFile can be necessary for second attempt of saving.
            File.WriteAllBytes(_outputFilePath, templateStream.ToArray());            
            templateAndOutputFile.Close();
            File.WriteAllBytes(_outputFilePath, templateStream.ToArray());            
            templateStream.Close();
            
        }



        private DataTable scanTablesinFile(string findText)
        {
            DataTable scanTableResult = new DataTable();
            scanTableResult.Columns.Clear();
            scanTableResult.Columns.Add("id", typeof(Int32));
            scanTableResult.Columns.Add("text", typeof(string));            
            scanTableResult.Columns.Add("sheetID", typeof(string)); //sheet id
            scanTableResult.Columns.Add("row", typeof(Int32));
            scanTableResult.Columns.Add("cell", typeof(Int32));
            scanTableResult.Columns.Add("columnName", typeof(string));
            int resultId = 0;
            
            foreach (Sheet currentSheet in templateAndOutputFile.WorkbookPart.Workbook.Descendants<Sheet>())
            {
                Worksheet currentWorkSheet = ((WorksheetPart)templateAndOutputFile.WorkbookPart.GetPartById(currentSheet.Id)).Worksheet;
                IEnumerable<Row> allRows = currentWorkSheet.GetFirstChild<SheetData>().Descendants<Row>();
           
               
                int w=1; //row counter
                foreach (Row currentRow in allRows)
                {
                    IEnumerable<Cell> allCells = currentRow.Descendants<Cell>();
                    int e =1; // cell counter 

                    // rows start with first fill cell so when it counts as first, row counter can be wrong. because of that w starts from first cell referance row index
                    if (allCells.FirstOrDefault() == null)
                    {
                        continue;                       
                    }
                    w = (int)GetRowIndex((allCells.FirstOrDefault()).CellReference);
                    foreach (Cell currentCell in allCells)
                    {
                        CellValue currentCellValue = currentCell.GetFirstChild<CellValue>();
                            if (currentCellValue == null)
                                currentCellValue = new CellValue("");

                        string data = currentCellValue.Text;
                        if (currentCell.DataType != null)
                        {
                            if (currentCell.DataType == CellValues.SharedString) // cell has a cell value that is a string, thus, stored else where
                            {
                                // gets text according to element SharedString, it is neccessary for merged cells.
                                data = templateAndOutputFile.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault().SharedStringTable.ElementAt(int.Parse(currentCellValue.Text)).InnerText;
                            }
                        }    
                           // if text which is seek, is found then
                            if (data.Contains(findText))
                            {
                                //.. put it is location (as sheet id,row no,cell no,cell coloumn) to the scanTableResult.
                                if (data == findText)
                                {
                                    scanTableResult.Rows.Add(resultId, findText, currentSheet.Id, w, e, GetColumnName(currentCell.CellReference.Value));
                                    resultId++;
                                }
                            }
                        e++;
                    }
                    w++;
                }                
            }        

            return scanTableResult;
        }        

        

        // return cell values type according to text. for example "1" will CellTypes.Number but "D1" will be CellValues.SharedString
        private static CellValues ConvertTextToCellValuesType(string cellRef)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellRef);
            // text like phone number must be sharedstring
            if (match.Value == "" && cellRef.Trim().Contains(" ") == false)
            {
                return CellValues.Number;
            }            
            //return match.Value == "" ? CellValues.Number : CellValues.SharedString;
            return CellValues.SharedString;
        }


        // give the Coloumn Alpha(name) from a given cell referance(adress like A11) -> result is A
        public static string GetColumnName(string cellRef)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellRef);

            return match.Value;
        }


        // give the Row index from a given cell referance(adress like A11) -> result is 11
        public static uint GetRowIndex(string cellRef)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellRef);

            return uint.Parse(match.Value);
        }

        // converts column Alpha to index number . ex: A->0 B -> 1
        public static int convertToColumnIndex(String columnName)
        {
            columnName = columnName.ToUpper();
            int value = 0;
            for (int i = 0, k = columnName.Length - 1; i < columnName.Length; i++, k--)
            {
                int alpabetIndex = Convert.ToInt32(Convert.ToChar(columnName.Substring(i,1))) - 64;
                int delta = 0;
                // last column simply add it
                if (k == 0)
                {
                    delta = alpabetIndex - 1;
                }
                else
                { // aggregate
                    if (alpabetIndex == 0)
                        delta = (26 * k);
                    else
                        delta = (alpabetIndex * 26 * k);
                }
                value += delta;
            }
            return value;
        }

        // converts column index number to name(alpha)  ex: 0-> A , 1 -> B
        public static string convertToColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

        public static string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that Sheet
                // object to retrieve a reference to the appropriate worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                if (theSheet == null)
                {
                    theSheet = wbPart.Workbook.Descendants<Sheet>().First();
                            
                    //throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part, and then use its 
                // Worksheet property to get a reference to the cell whose 
                // address matches the address you supplied:
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string:
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents a numeric value, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and Booleans
                    // individually. For shared strings, the code looks up the 
                    // corresponding value in the shared string table. For Booleans, 
                    // the code converts the value into the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                // For shared strings, look up the value in the shared 
                                // strings table.
                                var stringTable = wbPart.
                                  GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                // If the shared string table is missing, something is 
                                // wrong. Return the index that you found in the cell.
                                // Otherwise, look up the correct text in the table.
                                if (stringTable != null)
                                {
                                    value = stringTable.SharedStringTable.
                                      ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            return value;
        }

    private string convertAnyTypeCellValueToString(Cell theCell)
    {
        string value = "";
        if (theCell != null)
        {
            value = theCell.InnerText;

            // If the cell represents a numeric value, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and Booleans
            // individually. For shared strings, the code looks up the 
            // corresponding value in the shared string table. For Booleans, 
            // the code converts the value into the words TRUE or FALSE.
            if (theCell.DataType != null)
            {
                switch (theCell.DataType.Value)
                {
                    case CellValues.SharedString:
                        // For shared strings, look up the value in the shared 
                        // strings table.
                        var stringTable = templateAndOutputFile.WorkbookPart.
                          GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        // If the shared string table is missing, something is 
                        // wrong. Return the index that you found in the cell.
                        // Otherwise, look up the correct text in the table.
                        if (stringTable != null)
                        {
                            value = stringTable.SharedStringTable.
                              ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
        }
        return value;
    }

    public static fResult sendViewToXLS(ref DataGridView sourceView)
    {
       //  string xlsFile = SpreadsheetDocument.Open(, true);
        string xlsFile;
        general.getSaveDialog(language.XSTR("Save Output File"), Globals.extentionTypes[".xlsx"].ToString(),ExeArguments.DWGPREFIXFolder, language.XSTR("generalOut") + ".xlsx",out xlsFile);
        return sendViewToXLS(xlsFile, ref sourceView);
    }


    class BlankExcelDoc
    {
        public string _filePath;
        public SpreadsheetDocument _outDoc;
        public WorkbookPart _outPart;
        public WorksheetPart _outSheetPart;
        public BlankExcelDoc(string filePath)
        {
            _filePath=filePath;
            string errorFilePath = Path.GetPathRoot(_filePath) + Path.GetFileNameWithoutExtension(_filePath)+DateTime.Now.ToString("yyyyMMddHHmmss")+Path.GetExtension(_filePath);

            try
            {
                _outDoc = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook, true);
            }
            catch (IOException )
            {
                if (general.closeProccess(Path.GetFileName(_filePath)) == (int)fResult.NORMAL)
                {
                    try                    
                    {
                        _outDoc = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook, true);
                    }
                    catch (IOException )
                    {
                        _filePath = errorFilePath;
                        _outDoc = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook, true);
                        MessageBox.Show(language.XSTR("Target is not accessible")+" "+language.XSTR("File saved as ") + _filePath);
                    }
                    catch (Exception ext)
                    {
                        general.ShowException(ext);
                        throw ext;
                    }
                    
                }
                else
                {
                    try
                    {
                        _filePath = errorFilePath;
                        _outDoc = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook, true);
                        MessageBox.Show(language.XSTR("Target is not accessible") + " " + language.XSTR("File saved as ") + _filePath);
                    }
                    catch (System.Exception ex)
                    {
                        general.ShowException(ex);
                        throw ex;
                    }
                    
                }
            }
            _outPart = _outDoc.AddWorkbookPart();
            _outPart.Workbook = new Workbook();
            _outPart.Workbook.Save();

            SharedStringTablePart _outSharedPart = _outPart.AddNewPart<SharedStringTablePart>();
            _outSharedPart.SharedStringTable = new SharedStringTable();
            _outSharedPart.SharedStringTable.Save();
         

            _outSheetPart = _outPart.AddNewPart<WorksheetPart>();
            _outSheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = _outPart.Workbook.AppendChild<Sheets>(new Sheets());
            _outDoc.WorkbookPart.Workbook.Save();

            Sheet sheet = new Sheet()            
            {
                Id = _outPart.GetIdOfPart(_outSheetPart),
                SheetId = 1,
                Name = "1"
            };

            sheets.Append(sheet);
            _outDoc.WorkbookPart.Workbook.Save();
        }

        //saves and close document
        public fResult endDoc()
        {
            _outPart.Workbook.Save();
            _outDoc.Close();
            return fResult.NORMAL;

        }
    }
    public static fResult sendViewToXLS(string filePath,ref DataGridView sourceView)
    {
        BlankExcelDoc newExcelOut = new BlankExcelDoc(filePath);
        string refadress ="";

        int col = 1;
        foreach (DataGridViewColumn tempCol in sourceView.Columns)
        {
            WorksheetWriter writer = new WorksheetWriter(newExcelOut._outDoc, newExcelOut._outSheetPart);
            writer.PasteValue(general.convertNumericToAlpha(col, true)+"1", general.ConvertDoubleStr(general.convertNulltoEmpty(tempCol.HeaderText)), ConvertTextToCellValuesType(general.convertNulltoEmpty(general.convertNulltoEmpty(tempCol.HeaderText))));
            ++col;
        }

        int row = 2;
        foreach (DataGridViewRow tempRow in sourceView.Rows)
        {
            col = 1;
            foreach (DataGridViewCell tempCell in tempRow.Cells)
            {                
                WorksheetWriter writer = new WorksheetWriter(newExcelOut._outDoc, newExcelOut._outSheetPart);
                refadress = general.convertNumericToAlpha(col,true) + row;

                writer.PasteValue(refadress, general.ConvertDoubleStr(general.convertNulltoEmpty(tempCell.Value)), ConvertTextToCellValuesType(general.convertNulltoEmpty(tempCell.Value)));  
                //writer.PasteValue("A12", "AAA", ConvertTextToCellValuesType("AAAS"));  
                ++col;
            }
            ++row;
        }

        newExcelOut.endDoc();
        return fResult.NORMAL;
    }


    }   

}
