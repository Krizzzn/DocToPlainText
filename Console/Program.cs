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

            var arguments = ReadArgs(args);

            Microsoft.Office.Interop.Word.ApplicationClass wordObject = new ApplicationClass();
            object filename = Path.GetTempFileName();

            try {
                object file = Path.GetFullPath(arguments["in"]);
                object nullobject = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document docs = wordObject.Documents.Open
                    (ref file, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                try {
                    object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDOSText;
                    object encoding = ReadEncoding(arguments["enc"]);
                    docs.SaveAs(ref filename, ref format, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref encoding, ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                } finally {
                    docs.Close();
                }
            } finally {
                wordObject.Quit();
            }

            Console.Write(File.ReadAllText(filename.ToString()));

            if (arguments["out"] != null)
                File.Copy(filename.ToString(), args[1], true);

            File.Delete(filename.ToString());
        }

        private static Dictionary<string, string> ReadArgs(string[] args)
        {
            var dict = new Dictionary<string, string>();
            new[] { "in", "out", "enc" }.ToList().ForEach(m => dict.Add(m, null));

            dict["in"] = args[0];
            for (int i = 1; i < args.Length; i++) {
                if (args[i].ToLower().StartsWith("enc:"))
                    dict["enc"] = args[i].Replace("enc:","");
                else if (dict["out"] == null)
                    dict["out"] = args[i];
            }
            return dict;
        }

        private static object ReadEncoding(string argument)
        {
            object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
            if (argument != null) {
                var enCode = 0;
                if (int.TryParse(argument, out enCode)) {
                    if (Enum.IsDefined(typeof(Microsoft.Office.Core.MsoEncoding), enCode)) 
                        encoding = (Microsoft.Office.Core.MsoEncoding)enCode;
                }
            }
            return encoding;
        }

        private static void PrintHelp()
        {
            Console.WriteLine("\nSample call:\nDocToPlainText.exe \"d:\\mywordfile.docx\"");
            Console.WriteLine("\nSpecify the encoding with the enc: parameter.\nSample call:\nDocToPlainText.exe \"d:\\mywordfile.docx\" enc:874");
        }

        public static void GetEncodings()
        {
            var names = Enum.GetNames(typeof(Microsoft.Office.Core.MsoEncoding));
            foreach (var name in names) {
                Console.WriteLine("{1} = {0}   ", name.Replace("msoEncoding", ""), ((int)Enum.Parse(typeof(Microsoft.Office.Core.MsoEncoding), name, true)).ToString().PadLeft(5));
            }
        }
    }
}
