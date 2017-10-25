using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;



//For Open XML - (doc)ument operations 
// odt-> docx converting wanted cli_* dll needed
// doc -> docx converting wanted Interop.Word & excel needed

namespace msotempcreator
{
    class DocTemplateOut : TemplateOut
    {

        private WordprocessingDocument templateAndOutputFile;
        MemoryStream templateStream = new MemoryStream();
        public DocTemplateOut(string m_templateFilePath, string m_outputFilePath)
            : base(m_templateFilePath, m_outputFilePath)
        {
        }
                

        override protected void defineDocument()
        {            
            //byte[] byteArray = File.ReadAllBytes(GeneralOptions.templatePath);
            byte[] byteArray = File.ReadAllBytes(_templateFilePath);
            templateStream.Write(byteArray, 0, (int)byteArray.Length);
            templateAndOutputFile = WordprocessingDocument.Open(templateStream, true);            
        }       

        private DataTable scanTablesinFile(string findText)
        {
            DataTable scanTableResult = new DataTable();
            scanTableResult.Columns.Clear();
            scanTableResult.Columns.Add("id", typeof(Int32));
            scanTableResult.Columns.Add("text", typeof(string));            
            scanTableResult.Columns.Add("table", typeof(Int32));
            scanTableResult.Columns.Add("row", typeof(Int32));
            scanTableResult.Columns.Add("cell", typeof(Int32));            
            int tableCount = templateAndOutputFile.MainDocumentPart.Document.Body.Elements<Table>().Count();
            int resultId = 0;
            
            // gets all tables in document
            for (int q = 0 ; q < tableCount ; q++)
            {
                Table currentTable = templateAndOutputFile.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(q);
                int tableRowCount = currentTable.Elements<TableRow>().Count();
                // gets all rows in current table
                for (int w = 0; w < tableRowCount; w++)
                {
                    TableRow currentRow = currentTable.Elements<TableRow>().ElementAt(w);
                    int tableCellCount = currentRow.Elements<TableCell>().Count();

                    // gets all cells in current row
                    for (int e = 0; e < tableCellCount; e++)
                    {   
                        //scan Cell to get if cell text equals to seeked cell , if it is true then..
                        if (scanCell(currentRow.Elements<TableCell>().ElementAt(e), findText) == true)
                        {
                            //.. put it is location (as table no,row no,cell no) to the scanTableResult.
                            scanTableResult.Rows.Add(resultId, findText, q, w,e);
                            resultId++;
                        }
                    }                        
                }
            }           

            return scanTableResult;
        }        
                        
        private void writeCell(object node, string newText)
        {
            if (node is TableCell)
                foreach (object nodeCell in ((TableCell)node).ChildElements)
                    writeCell(nodeCell, newText);
            if (node is Paragraph)
                foreach (object nodePar in ((Paragraph)node).ChildElements)
                    writeCell(nodePar, newText);
            if (node is Run)
                foreach (object nodeRun in ((Run)node).ChildElements)
                    writeCell(nodeRun, newText);
            else
                if (node is Text)
                {
                    //((Text)node).Parent.Elements<Text>().FirstOrDefault().Text = newText;
                    
                    ((Text)node).Text = newText;
                   
                }
        }

               
        // removes paragraphs in a cell except first, removes runs in a paragraph except first, texts in a run except first
        private void removeCellChildItemsExceptFirst(object node)
        {
            if (node is TableCell)
            {
                while (((TableCell)node).Elements<Paragraph>().Count() > 1)
                {
                    ((TableCell)node).Elements<Paragraph>().ElementAt(1).Remove();
                }
                foreach (object nodeCell in ((TableCell)node).ChildElements)
                    removeCellChildItemsExceptFirst(nodeCell);                    
            }

            if (node is Paragraph)
            {
                while (((Paragraph)node).Elements<Run>().Count() > 1)
                {
                    ((Paragraph)node).Elements<Run>().ElementAt(1).Remove();
                }
                foreach (object nodePar in ((Paragraph)node).ChildElements)
                    removeCellChildItemsExceptFirst(nodePar);
            }
                
            if (node is Run)
            {
                while (((Run)node).Elements<Text>().Count() > 1)
                {
                    ((Run)node).Elements<Text>().ElementAt(1).Remove();
                }                
            }              
            
        }
        
        //replace all text in a node, node can be BODY if all things are wanted to change.
        private void replaceText(object node,string oldText, string newText)
        {
            if (node is Body)
                goto GoIn;            
            if (node is Table)
                goto GoIn;
            if (node is TableRow)
                goto GoIn;
            if (node is TableCell)
                goto GoIn;
            if (node is Paragraph)
                goto GoIn;
            if (node is Run)
                goto GoIn;
            if (node is Text)
            {
                if (((Text)node).Text == oldText)
                {
                    ((Text)node).Text = newText;
                }         
            }

            return;

            GoIn:
            foreach (object nodesNode in ((OpenXmlCompositeElement)node).ChildElements)
                replaceText(nodesNode, oldText, newText);
            
        }
           
        //scan cell for a text, but not suitable for parent object. for example if row or table is given , it will return false.
        private bool scanCell(object node, string newText)
        {
            if (node is TableCell)
            {
                if (((TableCell)node).InnerText.Trim() == newText)
                {
                    return true;
                }
                return false;
            }
                return false;               

        }

        // get a spesific cell value , also not suitable for parents.
        private string getCellValue(object node)
        {
            if (node is TableCell)
            {
                return ((TableCell)node).InnerText;
            }
            return "";

        }      

        // puts datas at given position, and also creates extra rows for datas if necessary.
        private void addValuesToCells(Table currentTable, Int32 rowIndex, Int32 cellIndex, string temp, DataTable sourceTable)
        {            
            //DataRow[] filteredRows = sourceTable.Select("temp = '" + temp + "'");
            // instead of filteredRows, maybe directly table could be use but first structure is based on a spesific coloumn(above). So converted like below.

            DataRow[] filteredRows = sourceTable.Select();         
           
            int dataRowCount = 0;
            for (int q = rowIndex; q < rowIndex + filteredRows.Length; q++)
            {                
                removeCellChildItemsExceptFirst(currentTable.Elements<TableRow>().ElementAt(q).Elements<TableCell>().ElementAt(cellIndex));
                // insert values to the position that scantableinFile has found.
                try
                {                   
                    writeCell(currentTable.Elements<TableRow>().ElementAt(q).Elements<TableCell>().ElementAt(cellIndex), filteredRows[dataRowCount][temp].ToString());
                }
                catch(Exception ex)
                {
                   // general.ShowException(ex);
                
                }
                dataRowCount++;
                
                // controls if "q"th(record) is the last or not, if not then ...
                if ((q + 1) != rowIndex + filteredRows.Length)
                {
                    // ... controls if there is an empty cell for next record . // and also controls table has enough rows for control // if there is no empty cell, add rows with empty cells                    
                    // 1.condition: firstly controls if there is enough row for next record , if not ,add row )code doesnt control 2. and 3. condition so there is no error while controlin)
                    // 2.condition: if current cell numbers and next cell number is different , then add row , so 3.condition will be pass away again.
                    // 3.condition: if  next rows cell which is at the same coloumn index with current row cell has "_empty", if not then add row.
                    
                    // command out for to use contain property of string
                    //while ((q + 1) >= currentTable.Elements<TableRow>().Count() || scanCell(currentTable.Elements<TableRow>().ElementAt(q + 1).Elements<TableCell>().ElementAt(cellIndex), "_empty_") == false )                    
                    while ((q + 1) >= currentTable.Elements<TableRow>().Count() || 
                            currentTable.Elements<TableRow>().ElementAt(q).Elements<TableCell>().Count() != currentTable.Elements<TableRow>().ElementAt(q+1).Elements<TableCell>().Count() || 
                            (getCellValue(currentTable.Elements<TableRow>().ElementAt(q + 1).Elements<TableCell>().ElementAt(cellIndex))).Contains("_empty_") == false)                    
                    {
                        //creating an empty row by cloning
                        TableRow newEmptyRow = (TableRow)currentTable.Elements<TableRow>().ElementAt(q).Clone();
                        //cleans clone rows cells ( "" causes problem while scaning at created extra rows, so "_empty_" is used)
                        writeAllCellsInRow(newEmptyRow, "_empty_");

                        currentTable.InsertAfter(newEmptyRow, currentTable.Elements<TableRow>().ElementAt(q));
                        
                        
                    }
                  
                    //tablonun sayısı 
                    
                }
            }
        }

        // find template data labels and give their positions to "addValuesToCells"
        protected override void replaceTempsWithDatasOnTables(DataTable sourceView)
        {
            int sourceColCount = 0;
            while (sourceColCount < sourceView.Columns.Count)
            {
                foreach (TableRow filteredRow in sourceView.Rows)
                {                    
                    DataRowCollection scanRows = scanTablesinFile("__" + sourceView.Columns[sourceColCount].ColumnName).Rows;
                    foreach (DataRow scanResult in scanRows)
                    {                        
                        addValuesToCells(templateAndOutputFile.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(Convert.ToInt32(scanResult["table"])), Convert.ToInt32(scanResult["row"]), Convert.ToInt32(scanResult["cell"]), sourceView.Columns[sourceColCount].Name,sourceView);                      

                    }
                }
                sourceColCount++;
            }
        }

        // cleans Extra Rows Label Texts.  (Extra Rows Label: "" causes problem while scaning extra rows, so "_empty_" text is using when extra rows created. so Extra Row Label Text is "empty")
        private void cleanExtraRowsCellLabel()
        {
            replaceText(templateAndOutputFile.MainDocumentPart.Document.Body, "_empty_", "");
        }
        
        //run row number function where field is defined.
        private void executeRowNumber(Table currentTable, Int32 rowIndex, Int32 cellIndex)
        {
            int rowCount = 0;
            // add row numbers if the next row cell is empty and the next row is same with current one. to understand it , controls number of rows are equal or not.
            do
            {                
                writeCell(currentTable.Elements<TableRow>().ElementAt(rowIndex + rowCount).Elements<TableCell>().ElementAt(cellIndex), (rowCount + 1).ToString());
                rowCount++;
                if ((rowIndex + rowCount) >= currentTable.Elements<TableRow>().Count())
                    break;
            }
            while (scanCell(currentTable.Elements<TableRow>().ElementAt(rowIndex + rowCount).Elements<TableCell>().ElementAt(cellIndex), "_empty_") == true && (currentTable.Elements<TableRow>().ElementAt(rowIndex + rowCount - 1).Elements<TableCell>().Count() == currentTable.Elements<TableRow>().ElementAt(rowIndex + rowCount).Elements<TableCell>().Count()));
            
        }

        // it returns cell on destination at same alignment with source row cell.
        private int findCellIndexAtSamePositionBetweenRows(TableRow sourceRow, int sourceCellIndex, TableRow destinationRow)
        {
            int sourcePosition = findCellPositionAtRow(sourceRow,sourceCellIndex);

            int cumulativeCellWidths = 0;            
            for (int i = 0; i < destinationRow.Elements<TableCell>().Count(); i++)
            {
                cumulativeCellWidths = cumulativeCellWidths + Convert.ToInt32(destinationRow.Elements<TableCell>().ElementAt(i).GetFirstChild<TableCellProperties>().GetFirstChild<TableCellWidth>().Width);
                
                if (cumulativeCellWidths > sourcePosition )
                {
                    return i;
                }
            }
            return -1;
        }

        // it returns LEFT position of cell in row. 
        private int findCellPositionAtRow(TableRow sourceRow, int sourceCellIndex)
        {
            int cumulativeCellWidths = 0;

            if (sourceCellIndex < sourceRow.Elements<TableCell>().Count())
            {
                for (int i = 0; i < sourceCellIndex; i++)
                {                    
                    cumulativeCellWidths = cumulativeCellWidths + Convert.ToInt32(sourceRow.Elements<TableCell>().ElementAt(i).GetFirstChild<TableCellProperties>().GetFirstChild<TableCellWidth>().Width);
                }
            }
            return cumulativeCellWidths;
            
        }

        protected override void saveDocument()
        {
            // before closing(and saving) templateAndOutputFile, firstly try if outputFile open or not. templateAndOutputFile can be necessary for second attempt of saving.
            File.WriteAllBytes(_outputFilePath, templateStream.ToArray());

            templateAndOutputFile.Close();            
            File.WriteAllBytes(_outputFilePath, templateStream.ToArray());
            templateStream.Close();

        }

        //add text to document independently
        private void addText(string textAdd)
        {
            Paragraph para = templateAndOutputFile.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(textAdd));
        }

        //replace template text with all related datas unified with seperator Mark
        override protected void replaceAllTemps(DataGridView sourceView)
        {
            int sourceColCount = 0;
            
            while (sourceColCount < sourceView.ColumnCount)
            {
                string result = "";
                foreach (DataGridViewRow filteredRow in sourceView.Rows)
                {
                    result = result + GeneralOptions.docOutSeperatorMark + filteredRow.Cells[sourceColCount].Value;             
                }
                result = result.Trim(GeneralOptions.docOutSeperatorMark.ToCharArray());
                replaceText(templateAndOutputFile.MainDocumentPart.Document.Body, "__" + sourceView.Columns[sourceColCount].Name, result);
                sourceColCount++;
            }
            
        }

        protected override void replaceFunctionsWithResultsOnTables(DataTable sourceView)
        {
            // only for excel for now
            return;
        }

        private void writeAllCellsInRow(TableRow sourceRow,string newText)
        {
             for (int g = 0; g < sourceRow.Elements<TableCell>().Count(); g++)
            {
                writeCell(sourceRow.Elements<TableCell>().ElementAt(g), newText);      

            }
        }

    }
}

