using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
namespace WebApplication2
{
    public class Util
    {
        private Util() { }
        public static bool WordToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Word.WdExportFormat exportFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
            Microsoft.Office.Interop.Word.Application application = null;

            Microsoft.Office.Interop.Word.Document document = null;
            object unknow = System.Type.Missing;
            application = new Microsoft.Office.Interop.Word.Application();
            application.Visible = false;
            document = application.Documents.Open(sourcePath);
            document.SaveAs();
            document.ExportAsFixedFormat(targetPath, exportFormat, false);
            //document.ExportAsFixedFormat(targetPath, exportFormat);
            result = true;

            //application.Documents.Close(ref unknow, ref unknow, ref unknow);
            document.Close(ref unknow, ref unknow, ref unknow);
            document = null;
            application.Quit();
            //application.Quit(ref unknow, ref unknow, ref unknow);
            application = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return result;
        }

        public static bool PDFToSWF(string toolPah, string sourcePath, string targetPath)
        {
            Process pc = new Process();
            bool returnValue = true;

            string cmd = toolPah;
            string args = " -t " + sourcePath + " -s flashversion=9 -o " + targetPath;
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo(cmd, args);
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                pc.StartInfo = psi;
                pc.Start();
                pc.WaitForExit();
            }
            catch (Exception ex)
            {
                returnValue = false;
                throw new Exception(ex.Message);
            }
            finally
            {
                pc.Close();
                pc.Dispose();
            }
            return returnValue;
        }

    }
}