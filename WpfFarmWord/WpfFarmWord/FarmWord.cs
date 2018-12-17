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
            var TableAfterPolish = CleantableFromMess(tableFulled.Key, tableFulled.Value);
            // OutPut(TableAfterPolish, path);
            OutputAnother(TableAfterPolish, path);
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
        //private static DataTable ConvertListListStringToDataTable(List<List<string>> totalRows, int maxCol)
        //{
        //    DataTable table = new DataTable();
        //    for (int i = 0; i < maxCol; i++)
        //    {
        //        table.Columns.Add();
        //    }
        //    foreach (List<string> row in totalRows)
        //    {
        //        while (row.Count < maxCol)
        //        {
        //            row.Add("");
        //        }
        //        table.Rows.Add(row.ToArray());
        //    }

        //    table = table.Rows
        //            .Cast<DataRow>()
        //            .Where(row => !row.ItemArray.All(field => field is DBNull ||
        //                            string.IsNullOrWhiteSpace(field as string)))
        //            .CopyToDataTable();
        //    return table;
        //}

        /// <summary>
        /// This method prints dataTable to TXT file
        /// </summary>
        /// <param name="dataTable"></param>
        //private void OutPut(DataTable dataTable, string path)
        //{
        //    string logName = "\\log.txt";
        //    string pathToFile = path + logName;
        //    StreamWriter swExtLogFile = new StreamWriter(pathToFile, true);

        //    int i;
        //    //swExtLogFile.Write(Environment.NewLine); // 1st \n at the begining of file
        //    swExtLogFile.Write("*****START OF DATA****" + DateTime.Now.ToString()+ Environment.NewLine);
        //    foreach (DataRow row in dataTable.Rows)
        //    {
        //        object[] array = row.ItemArray;
        //        for (i = 0; i < array.Length - 1; i++)
        //        {
        //            swExtLogFile.Write(array[i].ToString() + "|");
        //        }
        //        swExtLogFile.WriteLine(array[i].ToString());
        //    }
        //    swExtLogFile.Write("*****END OF DATA****" + DateTime.Now.ToString());
        //    swExtLogFile.Flush();
        //    swExtLogFile.Close();
        //}

        private void OutputAnother(List<List<string>> dataTable, string path)
        {
            string logName = "\\log.txt";
            string pathToFile = path + logName;
            string sum = string.Empty;
            using (StreamWriter swExtLogFile = new StreamWriter(pathToFile, true))
            {
                swExtLogFile.Write("*****START OF DATA****" + DateTime.Now.ToString() + Environment.NewLine);

                foreach (var line in dataTable)
                {
                    foreach (var item in line)
                    {
                        sum = sum + "|" + item;
                        
                    }
                    swExtLogFile.WriteLine(sum+ "|\n");
                    sum = string.Empty;
                }
                swExtLogFile.Write("*****END OF DATA****" + DateTime.Now.ToString());
            }
        }

        /// <summary>
        /// Usage Regex for clean string from field code HYPERLINK 
        /// </summary>
        /// <param name="target">Source for clean in type List<List<string>></param>
        /// <returns>Cleaned table</returns>
        private static List<List<string>> CleantableFromMess(List<List<string>> targets, int maxCol)
        {
            targets = DeleteEmptyStrings(targets);
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
            
            targets = MarkStandardsByListType(targets, markword);
            targets = DeleteHeaderAndSubHeaderStrings(targets, markword);
            targets = InsertCorrectionsFromDBPublisher(targets, maxCol, "*");
            targets = TRsClausePopulate(targets);
            targets = targets.Where(x => x.Count != 1).ToList(); // need this code, because method InsertCorrectionsFromDBPublisher()
                                                                 // does not remove cell after inserting correction.
            targets = SplitStringToCoupleStrings(targets);
            targets = AddGroupAndGroupType2List4TR4And20(targets);
            targets = OriginPublisher(targets);
            return targets;
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
                    //if (string.IsNullOrWhiteSpace(target))
                    //{
                    //    int count = 0;
                    //    List<int> indexs = new List<int>();
                    //    indexs.Add(targets.IndexOf(target));
                    //    if (count <= numOfColomns)
                    //    {
                    //        foreach (int index in indexs)
                    //        {
                    //            targets.RemoveAt(index);
                    //        }
                    //    }
                    //    count++;
                    //}
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
        private static List<List<string>> InsertCorrectionsFromDBPublisher(List<List<string>> listOfTargets, int maxCol, string separator)
        {
            string singleString = string.Empty;
            int indexIntoString = 0;
            int indexIntoList = 0;
            int indexOfSeparator = 0;
            int indexIntoListOfLists = 0;
            string output = string.Empty;
            string focus = string.Empty;
            string stdname = "ГОСТ";
            //-------------------------------------------------------------------
            List<List<string>> end = new List<List<string>>(listOfTargets.Count);

            for (int i = 0; i < listOfTargets.Count; i++)
            {
                List<string> temp = new List<string>();
                temp = listOfTargets[i];

                for (int k = 0; k < temp.Count; k++)
                {
                    if (temp[k].Contains(separator))
                    {
                        if (WhereStarIs(temp[k].ToString(), separator))
                        {
                            indexIntoList = k;
                            indexIntoListOfLists = i;
                            focus = temp[k].ToString();
                            indexIntoString = focus.IndexOf(stdname);
                            indexOfSeparator = focus.IndexOf(separator);
                        }
                        else
                        {
                            List<string> item = end[indexIntoListOfLists];
                            IEnumerable<string> result = GetSubStrings(temp[k].ToString(), "\"", "\"");
                            singleString = string.Join(",", result);
                            output = CorrectStringByAnother(item[indexIntoList].ToString(), singleString, indexIntoString, indexOfSeparator);
                            Console.WriteLine(output);
                            item[indexIntoList] = output;
                            indexIntoList = 0;
                        }
                    }
                }
                end.Add(temp);
            }

            // copy with out specific keyword
            for (int i = 0; i < end.Count; i++)
            {
                var temp = end[i];
                for (int j = 0; j < temp.Count; j++)
                {
                    if (temp[j].Contains(separator))
                    {
                        temp.RemoveAt(j);
                    }
                }
                end[i] = temp;
            }
            // clean from empty strings
            List<List<string>> final = new List<List<string>>(end.Count);
            foreach (List<string> list in end)
            {
                if (IsListEmpty(list))
                {
                    final.Add(list);
                }
            }

            return final;
        }

        /// <summary>
        ///  Method determine empty string. COunts string.Empty, whitespaces and null values.
        /// </summary>
        /// <param name="targets"> List of string for analysis. </param>
        /// <returns></returns>
        private static bool IsListEmpty(List<string> targets)
        {
            int countOfSpaces = 0;
            foreach (string target in targets)
            {
                if (string.IsNullOrEmpty(target) || string.IsNullOrWhiteSpace(target))
                {
                    countOfSpaces++;
                }
            }
            if (countOfSpaces != targets.Count)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 
        ///      find index of string
        ///      find index of separator( = "*")
        ///      if(index of separator<index of string)
        ///      {
        ///           star before string; // it's a remark
        ///      }
        ///      else
        ///      {
        ///      star after string; // it's string with corection
        ///      }
        /// </summary>
        /// <param name="target"> Target, where method will serach a separator</param>
        /// <param name="separator">Any sign that could be a separator</param>
        /// <returns>True: separator after keyword (Here i need to fix a hard written sequence. Here is "ГОСТ") False: if separator is before keyword.</returns>
        private static bool WhereStarIs(string target, string separator)
        {
            // get index of  separator
            int sepIndex = target.IndexOf(separator);
            // get index of keyword
            int keywordIndex = target.IndexOf("ГОСТ");
            if (sepIndex < keywordIndex)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Gets all sub strings from start to end expresstion
        /// </summary>
        /// <param name="input"> String, which will be under analysis</param>
        /// <param name="start"> Escape string that marks a start of coping.</param>
        /// <param name="end"> Escape string that marks an end of coping. This string will be excluded from result! </param>
        /// <returns> IEnumerable<string> type. For convertion to string you need to use this: "string.Join(",", result);" </returns>
        /// <remark>If you want to toogle on start string into output string you need to write "yield return match.Groups[0].Value;",
        ///  if you want not - "yield return match.Groups[1].Value;" </remark>
        private static IEnumerable<string> GetSubStrings(string input, string start, string end)
        {
            Regex r = new Regex(Regex.Escape(start) + "(.*?)" + Regex.Escape(end));
            MatchCollection matches = r.Matches(input);
            foreach (Match match in matches)
                yield return match.Groups[1].Value;
        }

        private static string CorrectStringByAnother(string target, string erratum, int indexIntoString, int indexOfSeparator)
        {
            StringBuilder outputValueBuilder = new StringBuilder(target);
            int Range = indexOfSeparator - indexIntoString + 1; // rewrite this shit magic!
            outputValueBuilder.Remove(indexIntoString, Range);
            outputValueBuilder.Insert(indexIntoString, erratum);
            target = outputValueBuilder.ToString();
            return target;
        }

        private static List<List<string>> TRsClausePopulate(List<List<string>> targets) // For TR CU 4, 20 only! 
        {
            List<List<string>> output = new List<List<string>>();
            foreach (var line in targets)
            {
                if (line.ToString() == null)
                { continue; }
                else
                {
                    List<string> newLine = new List<string>();
                    newLine = line.GetRange(1, line.Count - 1);
                    newLine.Insert(0, "статьи 4,5 ТР ТС 004/2011");
                    output.Add(newLine);
                }
            }
            return output;
        }

        private static List<List<string>> MarkStandardsByListType(List<List<string>> table, string markword)
        {
            List<List<string>> output = new List<List<string>>();
            int count = 0;
            foreach (var line in table)
            {
                if (line.ToString() == null)
                {
                  table.RemoveAt(table.IndexOf(line));
                }
                else
                {
                    if (line.Contains(markword))
                    {
                        count++;
                    }
                    if (count == 1)
                    {
                            line.Add("Free");
                            output.Add(line);
                    }
                    if (count == 2)
                    {
                        line.Add("Test");
                        output.Add(line);
                    }
                }
            }
            return output;
        }

        private static List<List<string>> DeleteEmptyStrings(List<List<string>> table)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (var line in table)
            {
                string buff = string.Empty;
                foreach (var item in line)
                {
                    buff += item;
                }
                if (string.IsNullOrEmpty(buff))
                {
                    continue;
                }
                else
                {
                    output.Add(line);
                }
            }
            return output;
        }

        private static List<List<string>> SplitStringToCoupleStrings(List<List<string>> table)
        {
            List<List<string>> output = new List<List<string>>();
            foreach (var line in table)
            {
                List<string> buffer = new List<string>();
                foreach (var item in line)
                {
                    
                    if (line.IndexOf(item) == 1)
                    {
                        if (item.Contains("раздел"))
                        {
                            string[] buff1 = item.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                            string sum = string.Empty;
                            int index = 0;
                            foreach (var element in buff1)
                            {
                                if (element.Contains("ГОСТ"))
                                {
                                    break;
                                }
                                if (element.Contains("СТБ"))
                                {
                                    break;
                                }
                                if (element.Contains("СТ"))
                                {
                                    break;
                                }
                                else
                                {
                                    sum += element + " ";
                                    index++;
                                }
                            }
                            buffer.Add(sum);
                            sum = string.Empty;
                            for (int k = index; k < buff1.Length; k++)
                            {
                                sum += buff1[k] + " ";
                            }
                            buffer.Add(sum);
                            continue;
                        }
                        else
                        {
                            buffer.Add(string.Empty);
                            buffer.Add(item);

                        }
                    }
                    else
                    {
                        buffer.Add(item);
                    }
                }
                output.Add(buffer);
            }
            return output;
        }

        private static List<List<string>> AddGroupAndGroupType2List4TR4And20(List<List<string>> table) // For TR CU 4,20 only!
        {
            List<List<string>> output = new List<List<string>>();
            foreach (var line in table)
            {
                List<string> buffer = line;
                buffer.AddRange(new List<string> { string.Empty, string.Empty});
                output.Add(buffer);
            }
            return table;
        }

        private static List<List<string>> OriginPublisher(List<List<string>> table)
        {
            List<List<string>> output =new List<List<string>>();
            
            foreach (var line in table)
            {
                List<string> buffer = line;
                for (int i=0; i<line.Count; i++)
                {
                    if (line[i].Contains("ГОСТ"))
                    {
                        buffer.Add("RU");
                    }
                    if (line[i].Contains("СТБ"))
                    {
                        buffer.Add("BEL");
                    }
                    if (line[i].Contains("СТ РК"))
                    {
                        buffer.Add("KZ");
                    }
                }
                output.Add(buffer);
            }
            return output; 
        }
    }

}