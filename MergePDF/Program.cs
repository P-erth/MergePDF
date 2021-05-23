using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Drawing;
using ZXing;

namespace MergePDF
{
    class Program
    {

        static void Main(string[] args)
        {
            string ruta= "sin ini";
            foreach (string value in args) ruta = value;
            Console.WriteLine(ruta);
            Console.ReadLine();
        }
    }
}