using System;
using System.Configuration;
using System.IO;
using RawPrint;

namespace MonitorDeArchivos
{
    class Program
    {
        static void Main(string[] args)
        {
            string PDFspoolXML = ConfigurationManager.AppSettings.Get("PDFspool");
            FileSystemWatcher PDFspool = new FileSystemWatcher(PDFspoolXML);
            PDFspool.NotifyFilter = (NotifyFilters.FileName);
            PDFspool.Created += ProcessFile;
            PDFspool.EnableRaisingEvents = true;
            Console.WriteLine("Esta ventana permite el redireccionamiento de impresiones, por favor no lo cierre");
            Console.ReadLine();
        }
        private static void ProcessFile(object source, FileSystemEventArgs e)
        {
            WatcherChangeTypes tipoDeCambio = e.ChangeType;
            string PDFrejectXML = ConfigurationManager.AppSettings.Get("PDFreject");
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
            string PDFinfoXML = ConfigurationManager.AppSettings.Get("PDFinfo");
            scriptProc.FileName = PDFinfoXML;
            scriptProc.Arguments = "\""+e.FullPath+"\"";
            process.StartInfo = scriptProc;
            process.Start();
            string breakText;
            string sizeX, sizeY;
            int noFound;
            do
            {
                breakText = process.StandardOutput.ReadLine();
                if (breakText != null && breakText.Substring(0, 10) == "Page size:")
                {
                    sizeX = breakText.Substring(16, breakText.IndexOf("x") - 17);
                    sizeY = breakText.Substring(breakText.IndexOf("x") + 2, breakText.IndexOf("pts") - breakText.IndexOf("x") - 3);
                    noFound = 0;
                    for (int i = 0; i < (Convert.ToInt32(QtyPaper)); i++)
                    {
                        if (TXvector[i] == sizeX && TYvector[i] == sizeY)
                        {
                            string printer = ConfigurationManager.AppSettings.Get(e.Name.Substring(0, 2)) + TSvector[i];
                            printPDF(printer, e.FullPath, e.Name);
                            File.Delete(e.FullPath);
                            noFound = 1;
                        }
                    }
                    if (noFound == 0)
                    {
                        if (File.Exists(PDFrejectXML + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + e.Name))
                        {
                            File.Delete(PDFrejectXML + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + e.Name);
                        }
                        File.Move(e.FullPath, PDFrejectXML + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + e.Name);
                        Console.WriteLine("No se pudo imprimir el siguiente archivo: {0} > {1}", e.Name, DateTime.Now.ToString("yyyyMMddHHmmss"));
                    }
                }
            } while (breakText.Substring(0,10) != "Page size:" || breakText is null);
            process.WaitForExit();
            process.Close();
        }
        private static void printPDF(string printerrelease, string filepath, string filename)
        {
            IPrinter printer = new Printer();
            printer.PrintRawFile(printerrelease, filepath, filename);
        }

    }
}