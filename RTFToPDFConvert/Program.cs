using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using LibreOfficeLibrary;

namespace RTFToPDFConvert
{
    class Program
    {
		//LibreOfficeLibrary.DocumentConverter convert = new DocumentConverter();
		static void Main(string[] args)
        {
			string[] pdfFiles = Directory.GetFiles(@"C:\Users\jona1993\Documents\FR", "*")
									 .Select(Path.GetFileName)
									 .ToArray();

			LibreOfficeConverter converter = new LibreOfficeConverter();

			foreach (string f in pdfFiles)
			{
				converter.ConvertToPDF(f);
			}
			

            Console.WriteLine("Hello World!");

			Console.Read();

			//var conv = new DocumentConverter();

			//conv.ConvertToPdf(@"C:\Users\jona1993\Documents\From.rtf", @"C:\Users\jona1993\Documents\to.pdf"); // Fonctionne pas (Mais d'apres le code source, fait la même chose)

		}

		
	}


}
