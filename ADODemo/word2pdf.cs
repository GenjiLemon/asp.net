using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace ADODemo
{
    class word2pdf
    {
        public bool WordToPdf(object sourcePath, string targetPath)
        {
            bool result = false;
            WdExportFormat wdExportFormatPDF = WdExportFormat.wdExportFormatPDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Word.ApplicationClass applicationClass = null;
            Microsoft.Office.Interop.Word.Document document = null;
            try
            {
                applicationClass = new Microsoft.Office.Interop.Word.ApplicationClass();
                document = applicationClass.Documents.Open(ref sourcePath, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                if (document != null)
                {
                    document.ExportAsFixedFormat(targetPath, wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 0, 0, WdExportItem.wdExportDocumentContent, true, true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false, ref missing);
                }
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (document != null)
                {
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit(ref missing, ref missing, ref missing);
                    applicationClass = null;
                }
            }
            return result;
        }

        /// <summary>
        /// 打开pdf文件方法
        /// </summary>
        /// <param name="p"></param>
        /// <param name="inFilePath">文件路径及文件名</param>
        //public static void Priview(System.Web.UI.Page p, string inFilePath)
        //{
        //    p.Response.ContentType = "Application/pdf";

        //    string fileName = inFilePath.Substring(inFilePath.LastIndexOf('\\') + 1);
        //    p.Response.AddHeader("content-disposition", "filename=" + fileName);
        //    p.Response.WriteFile(inFilePath);
        //    p.Response.End();
        //}

    }
}
