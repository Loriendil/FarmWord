using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

// for using Open XML SDK
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

//for OutPut method & some types in ReadingWordTables required. 
using System.Data; // for using DataTable objects
using System.IO;
// for Regex
using System.Text;
using System.Text.RegularExpressions;

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
        public FarmWord(string filename, string path)

        {

            string filepath = path + filename;
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

            var tableFulled = ReadWordTables(filepath);
            DataTable TableAfterPolish = CleantableFromMess(tableFulled.Key, tableFulled.Value);
            OutPut(TableAfterPolish, path);
        }

        /// <summary>
        /// Validate document before using.
        /// </summary>
        /// <param name="filepath">path to source file with data, that populated into tables with hard structure.</param>
        private static void ValidateWordDocument(string filepath)
        {
            try
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
                            @string.Append("Description: " + error.Description + "\n");
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
            catch (FileFormatException ex)
            {
                MessageBox.Show(ex.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                Application.Current.Shutdown();
            }
        }

        /// <summary>
        /// Read text from all cells from all tables in Word document as they populated by author of file. 
        /// </summary>
        /// <param name="filepath">path to source file with data, that populated into tables with hard structure.</param>
        /// <returns>Tuple<> object for usage as source for anyone user. Tuples returns 2 valiable!</returns>
        private static KeyValuePair<List<List<string>>, int> ReadWordTables(string filepath)
        {

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, isEditable: false))
                {
                    List<Table> tables =
                        doc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                    List<List<string>> totalRows = new List<List<string>>();
                    int oMax = 0;

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
                            oMax = maxCol;
                        }
                    }                    
                    return new KeyValuePair<List<List<string>>, int>(totalRows, oMax);
                }
            }
            catch (Exception ex)
            {
                string error = ex.ToString();
                MessageBox.Show(error);
                List<string> expr = new List<string>();
                expr.Add(error);
                List<List<string>> express = new List<List<string>>();
                express.Add(expr);
                return new KeyValuePair<List<List<string>>, int>(express, -1);
            }
        }

        
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
            
            table = table.Rows
                    .Cast<DataRow>()
                    .Where(row => !row.ItemArray.All(field => field is DBNull ||
                                    string.IsNullOrWhiteSpace(field as string)))
                    .CopyToDataTable();
            return table;
        }

        /// <summary>
        /// This method prints dataTable to TXT file
        /// </summary>
        /// <param name="dataTable"></param>
        private void OutPut(DataTable dataTable, string path)
        {
            string logName = "\\log.txt";
            string pathToFile = path + logName;
            StreamWriter swExtLogFile = new StreamWriter(pathToFile, true);

            int i;
            swExtLogFile.Write(Environment.NewLine);
            foreach (DataRow row in dataTable.Rows)
            {
                object[] array = row.ItemArray;
                for (i = 0; i < array.Length - 1; i++)
                {
                    swExtLogFile.Write(array[i].ToString() + "|");
                }
                swExtLogFile.WriteLine(array[i].ToString());
            }
            swExtLogFile.Write("*****END OF DATA****" + DateTime.Now.ToString());
            swExtLogFile.Flush();
            swExtLogFile.Close();
        }

        /// <summary>
        /// Usage Regex for clean string from field code HYPERLINK 
        /// </summary>
        /// <param name="target">Source for clean in type List<List<string>></param>
        /// <returns>Cleaned table</returns>
        private static DataTable CleantableFromMess(List<List<string>> targets, int maxCol)
        {
            DataTable outsource = new DataTable();
            string markword = "Примечание ";            
            string cleanedArrow = string.Empty;
            foreach (List<string> target in targets)
            {
                maxCol = target.Count;
                for (int i = 0; i < maxCol; i++)
                {
                    string temp = target[i].ToString();
                    cleanedArrow = CleanFromHyperlinks(temp);
                    target[i] = cleanedArrow;
                }
            }
           targets = DeleteHeaderAndSubHeaderStrings(targets, markword);
           outsource = ConvertListListStringToDataTable(targets, maxCol);
           return outsource;
        }

        /// <summary>
        /// Method cleans from hyperlinks hidden by text string
        /// </summary>
        /// <param name="str">input string with hyperlink</param>
        /// <returns> Same string without hyperlink</returns>
        private static string CleanFromHyperlinks(string str)
        {
            string temp = str;
            string replacement = "";
            do
            {
                if (temp.Contains("HYPERLINK"))
                {
                    string start = "HYPERLINK\\s*\"(.*?)\"";
                    string sNext = @"\\o\s*";
                    string s = "\"(.*?)\"";
                    string pattern = start + sNext + s;
                    Regex regex = new Regex(pattern);
                    temp = regex.Replace(temp, replacement);
                }
                if (temp.Contains("+"))
                {
                    string ss = @"\s*\+\s*";
                    string s1 = "\"(.*?)\"";
                    string pattern = ss + s1;
                    Regex regex = new Regex(pattern);
                    temp = regex.Replace(temp, replacement);
                }
            }
            while (temp.Contains("+"));
            return temp;
        }

        /// <summary>
        /// Method deletes from list every row contains mark word
        /// </summary>
        /// <param name="listOfTargets">list of arrows of strings that reprecent a table from Word document</param>
        /// <param name="variable">Variable represents a mark of row</param>
        /// <returns>List of arrows of strings without row with mark word.</returns>
        private static List<List<string>> DeleteHeaderAndSubHeaderStrings(List<List<string>> listOfTargets, string variable)
        {
            foreach (List<string> targets in listOfTargets.ToList())
            {
                foreach (string target in targets.ToList())
                {
                    int numOfColomns = targets.Count();
                    int indexOfTarElem = targets.FindIndex(x => x == variable);
                    if (indexOfTarElem != -1)
                    {
                        targets.RemoveRange(0, numOfColomns);
                    }
                    if (IsItNumber(target) == true)
                    {
                        int count = 0;
                        List<int> indexs = new List<int>();
                        indexs.Add(targets.IndexOf(target));
                        if (count <= numOfColomns)
                        {
                            foreach (int index in indexs)
                            {
                                targets.RemoveAt(index);
                            }
                        }
                        count++;
                    }
                    if (string.IsNullOrWhiteSpace(target))
                    {
                        int count = 0;
                        List<int> indexs = new List<int>();
                        indexs.Add(targets.IndexOf(target));
                        if (count <= numOfColomns)
                        {
                            foreach (int index in indexs)
                            {
                                targets.RemoveAt(index);
                            }
                        }
                        count++;
                    }
                }
            }
            return listOfTargets;
        }
 
        /// <summary>
        /// Methods determine input string is it integer or not
        /// </summary>
        /// <param name="target"> input string</param>
        /// <returns>bool value: true, input string is an integer, false: input string is not. 
        /// Optional, int.TryParse() returns integer that stored as string variable.</returns>
        private static bool IsItNumber(string target)
        {
            if (!String.IsNullOrEmpty(target))
            {
                int intValue;
                bool myValue = int.TryParse(target, out intValue);
                if (myValue)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
                return false;
        }
    }
}