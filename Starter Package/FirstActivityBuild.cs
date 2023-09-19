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
}
