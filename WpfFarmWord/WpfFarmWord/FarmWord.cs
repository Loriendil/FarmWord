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

            // 2. retrive text from cell[2,1] 
            // Search and replace text 
            // https://docs.microsoft.com/en-us/office/open-xml/how-to-search-and-replace-text-in-a-document-part
            // Change text in cell (Word)
            // https://docs.microsoft.com/en-us/office/open-xml/how-to-change-text-in-a-table-in-a-word-processing-document#change-text-in-a-cell-in-a-table
            // Vertical cells
            // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff951689(v=office.14)
            // 3. if-block for understand is it a new table or a next part from already excisting table
            // 4. make a remark
            GetAllTables(filepath);
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

        #region method for retrive lists of tables
        private static void GetAllTables(string filepath)
        {
            int rows = 10000;
            int colomns = 5;
            string[,] stds = new string[rows, colomns];
           

            using (WordprocessingDocument doc =
           WordprocessingDocument.Open(filepath, isEditable: false))
            {
                List<Table> tables =
                    doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                MessageBox.Show(tables.Count().ToString());

                //List<Table> tables =
                //    doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                // MessageBox.Show(tables.Count().ToString());

                foreach (Table table in tables)
                {
                    int i = 0;
                    int j = 0;
                    while (i < rows)
                    {
                        foreach (var row in table.Descendants<TableRow>())
                        {
                            StringBuilder textBuilder = new StringBuilder();
                            while (j < colomns)
                            {
                                
                                foreach (var cell in row.Descendants<TableCell>())
                                {
                                
                                
                                    foreach (var para in cell.Descendants<Paragraph>())
                                    {
                                        foreach (var run in para.Descendants<Run>())
                                        {
                                            foreach (var text in run.Descendants<Text>())
                                            {
                                                textBuilder.Append(text.InnerText);

                                            }
                                        } 
                                    }
                                    stds[i, j] = textBuilder.ToString();
                                    textBuilder.Clear();
                                    j++;
                                }
                                i++;
                            }
                        }
                        
                    }
                }
                
                MessageBox.Show(stds[2,4]);
            }    
        }
        #endregion 
        #region method for clause 2 - disabled 
        private static void CheckIsTablePart(Table table, List<Standard> stds)
        {
            string txt = "Элементы ";
            // Find the second row in the table.
            TableRow rowInd = table.Elements<TableRow>().ElementAt(1);

            // Find the second cell in the row.
            TableCell cellInd = rowInd.Elements<TableCell>().ElementAt(1);

            // Find the first paragraph in the table cell.
            Paragraph pInd = cellInd.Elements<Paragraph>().First();

            // Find the first run in the paragraph.
            Run rInd = pInd.Elements<Run>().First();

            // Set the text for the run.
            Text tInd = rInd.Elements<Text>().First();
            if (IsTablePart(tInd, txt) == true)
            {
                // populate new List<> with new remark
                MessageBox.Show(string.Format("This after call CheckIsTablePart, IsTablePart is true"));
            }
            else
            {
                StringBuilder textBuilder = new StringBuilder();
                // continue populate List<> with old remark
                MessageBox.Show(string.Format("This after call CheckIsTablePart, IsTablePart is false"));
                //ProcessTable(table, textBuilder, stds);
            }
            
        }
        #endregion

        #region method for clause 3 - disabled
        private static bool IsTablePart(Text tInd, string txt)
        {
            
                if (tInd.Text == txt)
                {
                    return true;
                }
                else
                {
                    return false;
                }
        }
        #endregion

        #region method for clause 4 - disabled
        private static void ProcessTable(Table node, StringBuilder textBuilder, List<Standard> standd)
        {
            foreach (var row in node.Descendants<TableRow>())
            {
                foreach (var cell in row.Descendants<TableCell>())
                {
                    for (int i = 0; i < node.Elements<TableGrid>().First().Elements<GridColumn>().Count(); i++) //<- iterate instances
                    {
                        foreach (var para in cell.Descendants<Paragraph>())
                        {
                            standd[i].Name = ProcessParagraph(para, textBuilder); // <- all Names in all stadds fulfilled. Anothers not! Incorrect! 
                        }
                        // http://www.cyberforum.ru/csharp-beginners/thread1104040.html
                    }
                }
            }
        }

        private static StringBuilder ProcessParagraph(Paragraph node, StringBuilder textBuilder)
        {
            foreach (var text in node.Descendants<Text>())
            {
                textBuilder.Append(text.InnerText);
            }
            return textBuilder;
        }
        #endregion
    }
}
