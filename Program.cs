using System;
using System.Configuration;
//using System.Collections.Specialized;
using System.IO;

namespace MonitorDeArchivos
{
    class Program
    {
        static void Main(string[] args)
        {
            FileSystemWatcher PDFspool = new FileSystemWatcher(@"D:\RICOH\Projects\PrintMultisize\PDFspool");
            PDFspool.NotifyFilter = (NotifyFilters.FileName);
            PDFspool.Created += ProcessFile;
            PDFspool.EnableRaisingEvents = true;
            Console.WriteLine("Presione <ENTER> tecla para detener: ");
            Console.ReadLine();
        }
        private static void ProcessFile(object source, FileSystemEventArgs e)
        {
            WatcherChangeTypes tipoDeCambio = e.ChangeType;
            string QtyPaper = ConfigurationManager.AppSettings.Get("QtyPaper");
            string[] TXvector = new string[Convert.ToInt32(QtyPaper)];
            string[] TYvector = new string[Convert.ToInt32(QtyPaper)];
            string[] TSvector = new string[Convert.ToInt32(QtyPaper)];
            for (int i = 0; i < (Convert.ToInt32(QtyPaper)); i++)
            {
                TXvector[i] = ConfigurationManager.AppSettings.Get("TX" + (i + 1));
                TYvector[i] = ConfigurationManager.AppSettings.Get("TY" + (i + 1));
                TSvector[i] = ConfigurationManager.AppSettings.Get("S" + (i + 1));
            }
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo scriptProc = new System.Diagnostics.ProcessStartInfo();
            scriptProc.RedirectStandardOutput = true;
            scriptProc.UseShellExecute = false;
            scriptProc.FileName = @"D:\RICOH\Projects\PrintMultisize\pdfinfo.exe";
            scriptProc.Arguments = "\""+e.FullPath+"\"";
            process.StartInfo = scriptProc;
            process.Start();
            string breakText;
            string sizeX, sizeY;
            do
            {
                breakText = process.StandardOutput.ReadLine();
                if (breakText != null && breakText.Substring(0, 10) == "Page size:")
                {
                    sizeX = breakText.Substring(16, breakText.IndexOf("x") - 17);
                    sizeY = breakText.Substring(breakText.IndexOf("x") + 2, breakText.IndexOf("pts") - breakText.IndexOf("x") - 3);
                    for (int i = 0; i < (Convert.ToInt32(QtyPaper)); i++)
                    {
                        if (TXvector[i] == sizeX && TYvector[i] == sizeY)
                        {
                            string printer = ConfigurationManager.AppSettings.Get(e.Name.Substring(0, 2)) + TSvector[i];
                            Console.WriteLine(">>Imprimir en: " + printer);
                            PrintFile(printer,e.FullPath);
                            //Console.WriteLine(sizeX + "," + sizeY);
                        }
                    }
                }
            } while (breakText.Substring(0,10) != "Page size:" || breakText is null);
            process.WaitForExit();
            process.Close();
        }
        private static void PrintFile(string printer, string file)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo scriptProc = new System.Diagnostics.ProcessStartInfo();
            scriptProc.RedirectStandardOutput = true;
            scriptProc.UseShellExecute = false;
            scriptProc.FileName = @"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe";
            scriptProc.Arguments = "/t \"" + file + "\" \"" + printer + "\" \"PCL6 Driver for Universal Print\" \"\"";
            process.StartInfo = scriptProc;
            process.Start();
            process.WaitForExit();
            process.Close();
        }
    }
}