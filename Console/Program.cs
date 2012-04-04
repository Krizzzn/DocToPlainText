﻿using System;
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
                Console.WriteLine("Sample call:\nDocToPlainText.exe \"d:\\mywordfile.docx\"");
                return;
            }
            if (!File.Exists(args[0])) {
                Console.WriteLine(args[0] + " file does not exist");
                Console.WriteLine("Sample call:\nDocToPlainText.exe \"d:\\mywordfile.docx\"");
                return;
            }

            Microsoft.Office.Interop.Word.ApplicationClass wordObject = new ApplicationClass();

            try {
                object file = args[0];
                object nullobject = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document docs = wordObject.Documents.Open
                    (ref file, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                    ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                try {
                    object filename = Path.GetTempFileName();
                    object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDOSText;
                    object encoding = "1252";
                    docs.SaveAs(ref filename, ref format, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref encoding, ref nullobject, ref nullobject, ref nullobject, ref nullobject);

                    Console.Write(File.ReadAllText(filename.ToString()));

                    if (args.Length >= 2)
                        File.Copy(filename.ToString(), args[1], true);

                    File.Delete(filename.ToString());
                }
                finally {
                    docs.Close();
                }
            }
            finally {
                wordObject.Quit();

            }
        }
    }
}
