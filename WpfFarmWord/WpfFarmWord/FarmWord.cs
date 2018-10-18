using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

// for using Open XML SDK
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

//for OutPut method & some types in ReadingWordTables required. 
using System.Data; // for using DataTable objects
using System.IO;

namespace WpfFarmWord
{
    class DocumentType
    {
        StringBuilder _name;
        public StringBuilder Name { get { return _name; } set { _name = value; } }

        StringBuilder _cipher;
        public StringBuilder Cipher { get { return _cipher; } set { _cipher = value; } }
    }

    class Standard : DocumentType
    {
        StringBuilder _trlink;
        public StringBuilder Trlink { get { return _trlink; } set { _trlink = value; } }

        StringBuilder _comment;
        public StringBuilder Comment { get { return _comment; } set { _comment = value; } }

        StringBuilder _remark;
        public StringBuilder Remark { get { return _remark; } set { _remark = value; } }
    }

    class FarmWord
    {
        public FarmWord(string filepath)

        {
            // 1. validation of file before read-write procedure.
            // This code has been taken from Open XML SDK:
            // https://docs.microsoft.com/en-us/office/open-xml/how-to-validate-a-word-processing-document
            ValidateWordDocument(filepath);

            // 2. Bibliography section
            // Search and replace text 
            // https://docs.microsoft.com/en-us/office/open-xml/how-to-search-and-replace-text-in-a-document-part
            // Change text in cell (Word)
            // https://docs.microsoft.com/en-us/office/open-xml/how-to-change-text-in-a-table-in-a-word-processing-document#change-text-in-a-cell-in-a-table
            // Vertical cells
            // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff951689(v=office.14)
             
            DataTable TableNeedsPolish = ReadWordTables(filepath);

            ////////////////////Testing!////////////////////////
            OutPut(TableNeedsPolish);
            ///////An illusion! What are you hiding?///////////

        }
        #region  Validation method
        private static void ValidateWordDocument(string filepath)
        {
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filepath, true))
            {
                StringBuilder @string = new StringBuilder();
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    int count = 0;
                    foreach (ValidationErrorInfo error in
                        validator.Validate(wordprocessingDocument))
                    {
                        count++;
                        @string.Append("Error " + count);
                        @string.Append("Description: " + error.Description+"\n");
                        @string.Append("ErrorType: " + error.ErrorType + "\n");
                        @string.Append("Node: " + error.Node + "\n");
                        @string.Append("Path: " + error.Path.XPath + "\n");
                        @string.Append("Part: " + error.Part.Uri + "\n");
                    }

                    @string.Append("Total count =" + count);
                    if (count != 0)
                    {
                        MessageBox.Show(@string.ToString());
                    }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                wordprocessingDocument.Close();
            }
        }
        #endregion

        /// <summary>
        /// Read text from all cells from all tables in Word document as they populated by author of file. 
        /// </summary>
        /// <param name="filepath">path to source file with data, that populated into tables with hard structure.</param>
        /// <returns>DataTable object for usage as source for anyone user</returns>
        private static DataTable ReadWordTables(string filepath)
        {
            DataTable tableWithMess = null;
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, isEditable: false))
                {
                    List<Table> tables =
                        doc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                    List<List<string>> totalRows = new List<List<string>>();

                    foreach (Table table in tables)
                    {
                        int maxCol = 0;
                        foreach (TableRow row in table.Elements<TableRow>())
                        {
                            List<string> tempRowValues = new List<string>();
                            foreach (TableCell cell in row.Elements<TableCell>())
                            {
                                tempRowValues.Add(cell.InnerText);
                            }
                            maxCol = ProcessList(tempRowValues, totalRows, maxCol);
                        }
                        tableWithMess = ConvertListListStringToDataTable(totalRows, maxCol);
                    }
                    return tableWithMess;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }

        }
        
        /// <summary>
        /// Add each row to the totalRows.
        /// </summary>
        /// <param name="tempRows"></param>
        /// <param name="totalRows"></param>
        /// <param name="MaxCol">the max column number in rows of the totalRows</param>
        /// <returns></returns>
        private static int ProcessList(List<string> tempRows, List<List<string>> totalRows, int MaxCol)
        {
            if (tempRows.Count > MaxCol)
            {
                MaxCol = tempRows.Count;
            }

            totalRows.Add(tempRows);
            return MaxCol;
        }

        /// <summary>
        /// This method converts list data to a data table
        /// </summary>
        /// <param name="totalRows"></param>
        /// <param name="maxCol"></param>
        /// <returns>returns datatable object</returns>
        private static DataTable ConvertListListStringToDataTable(List<List<string>> totalRows, int maxCol)
        {
            DataTable table = new DataTable();
            for (int i = 0; i < maxCol; i++)
            {
                table.Columns.Add();
            }
            foreach (List<string> row in totalRows)
            {
                while (row.Count < maxCol)
                {
                    row.Add("");
                }
                table.Rows.Add(row.ToArray());
            }
            return table;
        }

        /// <summary>
        /// This method prints dataTable to TXT file
        /// </summary>
        /// <param name="dataTable"></param>
        private void OutPut(DataTable dataTable)
        {
            StreamWriter swExtLogFile = new StreamWriter(@"D:\Users\Dreamwalker\source\repos\Accreditation scope generator\WpfFarmWord\WpfFarmWord\bin\Debug\log.txt", true);

            int i;
            swExtLogFile.Write(Environment.NewLine);
            foreach (DataRow row in dataTable.Rows)
            {
                object[] array = row.ItemArray;
                for (i = 0; i < array.Length - 1; i++)
                {
                    swExtLogFile.Write(array[i].ToString() + " | ");
                }
                swExtLogFile.WriteLine(array[i].ToString());
            }
            swExtLogFile.Write("*****END OF DATA****" + DateTime.Now.ToString());
            swExtLogFile.Flush();
            swExtLogFile.Close();
        }
        
    }
}
