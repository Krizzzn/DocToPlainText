using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace DocToPlainText
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) {
                PrintHelp();
                return;
            }
            if (!File.Exists(args[0])) {
                Console.WriteLine(args[0] + " file does not exist");
                PrintHelp();
                return;
            }

            Microsoft.Office.Interop.Word.ApplicationClass wordObject = new ApplicationClass();
            object filename = Path.GetTempFileName();

            
            try {
                object file = args[0];
                object nullobject = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document docs = wordObject.Documents.Open
                    (ref file, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                try {
                    object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDOSText;
                    object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                    docs.SaveAs(ref filename, ref format, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref encoding, ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                }
                finally {
                    docs.Close();
                }
            }
            finally {
                wordObject.Quit();

            }

            Console.Write(File.ReadAllText(filename.ToString()));

            if (args.Length >= 2)
                File.Copy(filename.ToString(), args[1], true);

            File.Delete(filename.ToString());
        }

        private static void PrintHelp()
        {
            Console.WriteLine("\nSample call:\nDocToPlainText.exe \"d:\\mywordfile.docx\"");
        }
    }
}
