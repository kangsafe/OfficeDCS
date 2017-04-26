using System;
using System.IO;
using System.Text;

namespace XDPI
{
    class OfficeUtils
    {
        public OfficeUtils()
        { }
        /// <summary>把Word文件转换成为PDF格式文件</summary>   
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns>
        public static bool WordToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Word.WdExportFormat exportFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
            Microsoft.Office.Interop.Word.ApplicationClass application = null;
            Microsoft.Office.Interop.Word.Document document = null; try
            {
                application = new Microsoft.Office.Interop.Word.ApplicationClass();
                application.Visible = false;
                document = application.Documents.Open(sourcePath);
                document.SaveAs2();
                document.ExportAsFixedFormat(targetPath, exportFormat);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                result = false;
            }
            finally
            {
                if (document != null)
                {
                    document.Close();
                    document = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        /// <summary>把Word文件转换成为PDF格式文件</summary>   
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns>
        public static bool WordToHtml(string sourcePath, string targetPath)
        {
            bool result = false;

            if (File.Exists(sourcePath))
            {
                Console.WriteLine("正在生成html，请稍候...");
                object oMissing = System.Reflection.Missing.Value;
                object oTrue = true;
                object oFalse = false;

                Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word._Document oWordDoc = new Microsoft.Office.Interop.Word.Document();

                oWord.Visible = false;
                object openFormat = Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto;
                object openName = sourcePath;

                try
                {
                    oWordDoc = oWord.Documents.Open(ref openName, ref oMissing, ref oTrue, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref openFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                }
                catch (Exception e)
                {
                    Console.WriteLine("读取Word文档时发生异常");
                    oWord.Quit(ref oTrue, ref oMissing, ref oMissing);
                    return false;
                }

                object saveFileName = targetPath;
                object saveFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML;

                oWordDoc.SaveAs(ref saveFileName, ref saveFormat, ref oMissing, ref oMissing, ref oFalse, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                oWordDoc.Close(ref oTrue, ref oMissing, ref oMissing);
                oWord.Quit(ref oTrue, ref oMissing, ref oMissing);

                Encoding enc = Encoding.GetEncoding("GB2312");
                string s = File.ReadAllText(targetPath, enc);
                s = s.Replace("position:absolute;", "");
                File.WriteAllText(targetPath, s, enc);

                Console.WriteLine("Word文档已转换为html格式");
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return true;
            }
            return false;
        }


        /// <summary>把Microsoft.Office.Interop.Excel文件转换成PDF格式文件</summary>   
                /// <param name="sourcePath">源文件路径</param> 
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns> 
        public static bool ExcelToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Excel.XlFixedFormatType targetType =
            Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.ApplicationClass application = null; Microsoft.Office.Interop.Excel.Workbook workBook = null;
            try
            {
                application = new Microsoft.Office.Interop.Excel.ApplicationClass();
                application.Visible = false;
                workBook = application.Workbooks.Open(sourcePath);
                workBook.SaveAs();
                workBook.ExportAsFixedFormat(targetType, targetPath);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        /// <summary> 把PowerPoint文件转换成PDF格式文件</summary>   
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns>   
        public static bool PowerPointToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            //Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF;
            //object missing = Type.Missing;
            //Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null; Microsoft.Office.Interop.PowerPoint.Presentation persentation = null;
            //try
            //{
            //    application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            //    application.Visible = MsoTriState.msoFalse;
            //    persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            //    persentation.SaveAs(targetPath, targetFileType, MsoTriState.msoTrue);
            //    result = true;
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //    result = false;
            //}
            //finally
            //{
            //    if (persentation != null)
            //    {
            //        persentation.Close();
            //        persentation = null;
            //    }
            //    if (application != null)
            //    {
            //        application.Quit();
            //        application = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //}
            return result;
        }
        /// <summary> 把Visio文件转换成PDF格式文件  </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns>   
        public static bool VisioToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            //Microsoft.Office.Interop.Visio.VisFixedFormatTypes targetType =
            //Microsoft.Office.Interop.Visio.VisFixedFormatTypes.visFixedFormatPDF;
            //object missing = Type.Missing;
            //Microsoft.Office.Interop.Visio.ApplicationClass application = null; Microsoft.Office.Interop.Visio.Document document = null;
            //try
            //{
            //    application = new Microsoft.Office.Interop.Visio.ApplicationClass();
            //    application.Visible = false;
            //    document = application.Documents.Open(sourcePath);
            //    document.Save();
            //    document.ExportAsFixedFormat(targetType, targetPath, Microsoft.Office.Interop.Visio.VisDocExIntent.visDocExIntentScreen, Microsoft.Office.Interop.Visio.VisPrintOutRange.visPrintAll);
            //    result = true;
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //    result = false;
            //}
            //finally
            //{
            //    if (application != null)
            //    {
            //        application.Quit();
            //        application = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //}
            return result;
        }
        /// <summary>把Project文件转换成PDF格式文件</summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns>   
        public static bool ProjectToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            object missing = Type.Missing;
            //Microsoft.Office.Interop.MSProject.ApplicationClass application = null;
            //try
            //{
            //    application = new Microsoft.Office.Interop.MSProject.ApplicationClass();
            //    application.Visible = false;
            //    application.FileOpenEx(sourcePath);
            //    application.DocumentExport(targetPath, Microsoft.Office.Interop.MSProject.PjDocExportType.pjPDF);
            //    result = true;
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message); result = false;
            //}
            //finally
            //{
            //    if (application != null)
            //    {
            //        application.DocClose();
            //        application.Quit();
            //        application = null;
            //    }
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //}
            return result;
        }
    }
}