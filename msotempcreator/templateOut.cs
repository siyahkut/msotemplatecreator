using System;
using System.IO;
using Defs


namespace msotempcreator
{
    // derived classes : ssheetTemplateOut.cs , docTemplateOut.cs

    abstract class TemplateOut
    {
        public string _templateFilePath = GeneralOptions.generalOptionsTable.Rows[0][GeneralOptions.templatePathOptionName].ToString() ;
        public string _outputFilePath = "";
        public bool tempFileExists = false;
        public TemplateOut(string templateFilePath, string outputFilePath)
        {
            _templateFilePath = templateFilePath;
            _outputFilePath = outputFilePath;            
            
            try
            {                               
                defineDocument();
                tempFileExists = true;
            }
            catch (IOException)
            {                
                DialogResult dialogResult = MessageBox.Show(language.XSTR("Template file is not accessible. Do you want toolbox try to close it if it is open as a solution?"), definitions.MsgBoxCaptions.sureWindow.ToString(), MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (general.closeProccess(Path.GetFileName(_templateFilePath)) == (int)fResult.NORMAL)
                    {
                        try
                        {
                           defineDocument();
                           tempFileExists = true;
                        }
                        catch (Exception ex)
                        {
                            general.ShowException(language.XSTR("Template file is not accessible. As solution you can try to close it manually if it is open or select a new file as a template. "),ex);
                            tempFileExists = false;                            
                        }
                    }
                }
            }
            catch (Exception ext2)
            {
                general.ShowException(ext2);
            }

        }

        public fResult convertTempToOut(DataGridView sourceView)
        {
            string temp = "";
            return convertTempToOut(sourceView, true,ref temp);            
        }

        public fResult convertTempToOut(DataGridView sourceView, bool saveFile, ref string saveFilePath)
        {
            string filePath;
            if (saveFilePath == null)
            {
                filePath = general.selectOutputFileName();
                saveFilePath = filePath;
            }
            else
            {
                filePath = saveFilePath;
            }
            if (filePath != "" && filePath != null)
            {
                _outputFilePath = filePath;
                replaceTempsWithDatasOnTables(sourceView);
                replaceAllTemps(sourceView);
                replaceFunctionsWithResultsOnTables(sourceView);
                if (saveFile) saveOutputFile();
                return fResult.NORMAL;
            }
            else
            {
                return fResult.ERROR;
            }
        }        


        protected void saveOutputFile()
        {
            
            try
            {               
                saveDocument();
            }
            catch (IOException ext)
            {
                DialogResult dialogResult = MessageBox.Show(language.XSTR("Output file is open. Do you want toolbox try to close it for new exporting?"), definitions.MsgBoxCaptions.sureWindow.ToString(), MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (general.closeProccess(Path.GetFileName(_outputFilePath)) == (int)fResult.NORMAL)
                    {
                        try
                        {                           
                            saveDocument();
                        }
                        catch (Exception ex)
                        {
                            general.ShowException(language.XSTR("Output file must be close before exporting.  Please close it manually.")+" "+ext, ex);
                        }
                    }
                }
            }
            catch (Exception ext2)
            {
                general.ShowException(ext2);
            }
        }

        

        abstract protected void defineDocument();
        abstract protected void saveDocument();
               
        abstract protected void replaceTempsWithDatasOnTables(DataGridView sourceView);
        abstract protected void replaceAllTemps(DataGridView sourceView);
        abstract protected void replaceFunctionsWithResultsOnTables(DataGridView sourceView);
        

        

    }
}
