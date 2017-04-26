using System;
using System.Text;
using System.IO;

namespace OfficeDCS
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("说明：本程序用于将word文档转换为html格式文档，支持.doc和.docx格式");
                Console.WriteLine("用法：word2html.exe <待转换的word文档>");
                Console.WriteLine("Copyleft(C)2015 Solomon");
                Console.ReadLine();
                return;
            }

            string srcInputName = args[0]; // 打开文件的位置
            string ext = Path.GetExtension(srcInputName);

            string current_cmd = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            string current_dir = Path.GetDirectoryName(current_cmd);
            //    Console.WriteLine("正在生成html，请稍候...");

            string inputName = current_dir + "\\" + srcInputName;
            string outputName = inputName.Replace(ext, ".html"); // 同路径保存
            XDPI.OfficeUtils.WordToHtml(inputName, outputName);
            //    if (File.Exists(inputName))
            //    {

            //        object oMissing = System.Reflection.Missing.Value;
            //        object oTrue = true;
            //        object oFalse = false;

            //        Word._Application oWord = new Word.Application();
            //        Word._Document oWordDoc = new Word.Document();

            //        oWord.Visible = false;
            //        object openFormat = Word.WdOpenFormat.wdOpenFormatAuto;
            //        object openName = inputName;

            //        try
            //        {
            //            oWordDoc = oWord.Documents.Open(ref openName, ref oMissing, ref oTrue, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref openFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //        }
            //        catch (Exception e)
            //        {
            //            Console.WriteLine("读取Word文档时发生异常");
            //            oWord.Quit(ref oTrue, ref oMissing, ref oMissing);
            //            return;
            //        }

            //        object saveFileName = outputName;
            //        object saveFormat = Word.WdSaveFormat.wdFormatFilteredHTML;

            //        oWordDoc.SaveAs(ref saveFileName, ref saveFormat, ref oMissing, ref oMissing, ref oFalse, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            //        oWordDoc.Close(ref oTrue, ref oMissing, ref oMissing);
            //        oWord.Quit(ref oTrue, ref oMissing, ref oMissing);

            //        Encoding enc = Encoding.GetEncoding("GB2312");
            //        string s = File.ReadAllText(outputName, enc);
            //        s = s.Replace("position:absolute;", "");
            //        File.WriteAllText(outputName, s, enc);

            //        Console.WriteLine("Word文档已转换为html格式");
            //    }
            //}
        }
    }
}
