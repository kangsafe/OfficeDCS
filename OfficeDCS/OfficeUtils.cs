using Microsoft.Office.Core;
using System;
using System.IO;
using System.Text;

namespace XDPI
{
    class OfficeUtils
    {
        public OfficeUtils()
        { }

        /// <summary>
        /// GB2312转换成UTF8
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string gb2312_utf8(string text)
        {
            //声明字符集   
            System.Text.Encoding utf8, gb2312;
            //gb2312   
            gb2312 = System.Text.Encoding.GetEncoding("gb2312");
            //utf8   
            utf8 = System.Text.Encoding.GetEncoding("utf-8");
            byte[] gb;
            gb = gb2312.GetBytes(text);
            gb = System.Text.Encoding.Convert(gb2312, utf8, gb);
            //返回转换后的字符   
            return utf8.GetString(gb);
        }

        /// <summary>
        /// UTF8转换成GB2312
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string utf8_gb2312(string text)
        {
            //声明字符集   
            System.Text.Encoding utf8, gb2312;
            //utf8   
            utf8 = System.Text.Encoding.GetEncoding("utf-8");
            //gb2312   
            gb2312 = System.Text.Encoding.GetEncoding("gb2312");
            byte[] utf;
            utf = utf8.GetBytes(text);
            utf = System.Text.Encoding.Convert(utf8, gb2312, utf);
            //返回转换后的字符   
            return gb2312.GetString(utf);
        }

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

        /// <summary>把Microsoft.Office.Interop.Excel文件转换成PDF格式文件</summary>   
                /// <param name="sourcePath">源文件路径</param> 
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns> 
        public static bool ExcelToHtml(string sourcePath, string targetPath)
        {
            bool result = false;
            try
            {
                //实例化Excel  
                Microsoft.Office.Interop.Excel.Application repExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = null;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                //打开文件，n.FullPath是文件路径  
                workbook = repExcel.Application.Workbooks.Open(sourcePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                object savefilename = targetPath;
                object ofmt = Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml;
                //进行另存为操作    
                workbook.SaveAs(savefilename, ofmt, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                object osave = false;
                //逐步关闭所有使用的对象  
                workbook.Close(osave, Type.Missing, Type.Missing);
                repExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                worksheet = null;
                //垃圾回收  
                GC.Collect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                GC.Collect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(repExcel.Application.Workbooks);
                GC.Collect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(repExcel);
                repExcel = null;
                GC.Collect();
                //依据时间杀灭进程  
                System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process p in process)
                {
                    if (DateTime.Now.Second - p.StartTime.Second > 0 && DateTime.Now.Second - p.StartTime.Second < 5)
                    {
                        p.Kill();
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return false;
            }

        }

        /// <summary> 把PowerPoint文件转换成PDF格式文件</summary>   
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns>true=转换成功</returns>   
        public static bool PowerPointToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null; Microsoft.Office.Interop.PowerPoint.Presentation persentation = null;
            try
            {
                application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                //application.Visible = MsoTriState.msoFalse;
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, MsoTriState.msoTrue);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
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
        public static bool PowerPointToPNG(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPNG;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null; Microsoft.Office.Interop.PowerPoint.Presentation persentation = null;
            try
            {
                application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                //application.Visible = MsoTriState.msoFalse;
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, MsoTriState.msoTrue);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
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
        public static bool PowerPointToJPG(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsJPG;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null; Microsoft.Office.Interop.PowerPoint.Presentation persentation = null;
            try
            {
                application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                //application.Visible = MsoTriState.msoFalse;
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, MsoTriState.msoTrue);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
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
        public static bool PowerPointToGIF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null; Microsoft.Office.Interop.PowerPoint.Presentation persentation = null;
            try
            {
                application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                //application.Visible = MsoTriState.msoFalse;
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, MsoTriState.msoTrue);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
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
        public static bool PowerPointToBMP(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsBMP;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null; Microsoft.Office.Interop.PowerPoint.Presentation persentation = null;
            try
            {
                application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                //application.Visible = MsoTriState.msoFalse;
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, MsoTriState.msoTrue);
                result = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
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
        public static bool PowerPointToHtml(string sourcePath, string targetPath)
        {
            bool result = false;
            try
            {
                //Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType targetFileType = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.;
                Microsoft.Office.Interop.PowerPoint.Application ppt = new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Core.MsoTriState m1 = new MsoTriState();
                Microsoft.Office.Core.MsoTriState m2 = new MsoTriState();
                Microsoft.Office.Core.MsoTriState m3 = new MsoTriState();
                Microsoft.Office.Interop.PowerPoint.Presentation pp = ppt.Presentations.Open(sourcePath, m1, m2, m3);
                pp.SaveAs(targetPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsHTML, Microsoft.Office.Core.MsoTriState.msoTriStateMixed);
                pp.Close();
                ppt.Quit();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return false;
            }

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