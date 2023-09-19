using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;

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
}
