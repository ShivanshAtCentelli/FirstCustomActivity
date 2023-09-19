using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel; // Excel
using System.Runtime.InteropServices; // Marshal
using System.Reflection; //Missing.Values

namespace StarterPackage.FirstActivityBuild
{
    public class ConcatenateStrings: CodeActivity
    {
        [Category("Input")]
        [DisplayName("First String")]
        [Description("First string to be concatenated with another string")]
        public InArgument<string> istrFirstString { get; set; }

        [Category("Input")]
        [DisplayName("Second String")]
        [Description("Secong string to be concatenated with another string")]
        public InArgument<string> istrSecondString { get; set; }

        [Category("Output")]
        [DisplayName("Concatenated String")]
        [Description("Concatenated string")]
        public OutArgument<string> ostrOutputString { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            string strFirstString = istrFirstString.Get(context);
            string strSecondString = istrSecondString.Get(context);
            string strConcatenatedString = strFirstString + " " + strSecondString;
            ostrOutputString.Set(context, strConcatenatedString);

        }
    }

    public class CreateWorkbook: CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Workbook Path")]
        [Description("Full path where the workbook is to be created")]
        public InArgument<string> istrFilePath { get; set; }

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("Workbook")]
        [Description("Workbook object")]
        public OutArgument<Workbook> owrbkOutputWorkbook { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            string strExcelFile = istrFilePath.Get(context);

            // Initialize excel application
            Application xlApplication = new Application();
            Workbook xlWorkbook = xlApplication.Workbooks.Add(Missing.Value);
            xlWorkbook.SaveAs(strExcelFile);

            // Close all application
            xlWorkbook.Close(true, Missing.Value, Missing.Value);
            xlApplication.Quit();

            // Output workbook object
            owrbkOutputWorkbook.Set(context, xlWorkbook);

            // Release COM objects
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApplication);
        }

        
    }

    public class ReadRange : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Workbook Path")]
        [Description("Please enter the full path of the where workbook is to be created")]
        public InArgument<string> istrFilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Sheet Name")]
        [Description("Please enter the sheet name")]
        public InArgument<string> istrSheetName { get; set; } = "Sheet1";

        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Range")]
        [Description("Please enter the range")]
        public InArgument<string> istrRange { get; set; } = "A1:B2";

        [Category("Output")]
        [RequiredArgument]
        [DisplayName("DataTable")]
        [Description("Use a DataTable variable")]
        public OutArgument<System.Data.DataTable> odtOutput { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                string strFileName = istrFilePath.Get(context);
                string strSheetName = istrSheetName.Get(context);
                string strRange = istrRange.Get(context);

                // Insert data into DataTable
                System.Data.OleDb.OleDbConnection oleDbConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + strFileName + "';Extended Properties=Excel 8.0;");
                System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM  [" + strSheetName + "$" + strRange + "]", oleDbConnection);
                oleDbDataAdapter.TableMappings.Add("Table", "TestTable");
                System.Data.DataSet dataSet = new System.Data.DataSet();
                oleDbDataAdapter.Fill(dataSet);
                System.Data.DataTable dataTable = dataSet.Tables[0];

                // Close all the connections
                oleDbConnection.Close();

                // Return output
                odtOutput.Set(context, dataTable);
            }
            catch(System.Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message + Environment.NewLine + istrRange.Get(context) + Environment.NewLine + istrSheetName.Get(context) + Environment.NewLine + istrFilePath.Get(context));
            }
        }
    }
}
