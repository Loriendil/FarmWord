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
         
            using (WordprocessingDocument doc =
           WordprocessingDocument.Open(filepath, isEditable: false))
            {
                List<Table> tables =
                    doc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                //MessageBox.Show(tables.Count().ToString());

                foreach (Table table in tables)
                {
                    
                }
                
            }    
        }
        #endregion 


    }
}
