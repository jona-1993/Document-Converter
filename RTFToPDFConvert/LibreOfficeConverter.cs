using System;
using System.Collections.Generic;
using System.Text;

namespace RTFToPDFConvert
{
    class LibreOfficeConverter
    {
		public string LibreOfficeDirectory { get; set; }
		public string SourceDirectory { get; set; }
		public string DestinationDirectory { get; set; }

		public LibreOfficeConverter()
		{
			SourceDirectory = @"C:\Users\jona1993\Documents\FR\";
			DestinationDirectory = @"C:\Users\jona1993\Documents\FR_PDF";
			LibreOfficeDirectory = @"C:\Program Files\LibreOffice";
		}

		public LibreOfficeConverter(string sourcedir, string destinationdir, string lodir)
		{
			SourceDirectory = sourcedir;
			DestinationDirectory = destinationdir;
			LibreOfficeDirectory = lodir;
		}

		public async void ConvertToPDF(string filename)
		{
			System.Diagnostics.Process process = new System.Diagnostics.Process();

			System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();

			startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;

			startInfo.FileName = LibreOfficeDirectory + @"\program\soffice"; // Il faut avoir installé Libre Office

			startInfo.Arguments = @"--headless --norestore --convert-to pdf """ + SourceDirectory + filename + "\"" + @" --outdir " + DestinationDirectory;

			process.StartInfo = startInfo;

			process.Start();
		}

		public async void ConvertToDOCX(string filename)
		{
			System.Diagnostics.Process process = new System.Diagnostics.Process();

			System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();

			startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;

			startInfo.FileName = LibreOfficeDirectory + @"\program\soffice";

			startInfo.Arguments = @"--headless --norestore --convert-to docx """ + SourceDirectory + filename + "\"" + @" --outdir " + DestinationDirectory;

			process.StartInfo = startInfo;

			process.Start();
		}

	}

}
