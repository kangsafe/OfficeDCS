using System;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace OfficeDCS
{
    class Program
    {
        //匹配编码格式
        static string reg_e = "\\s+-e\\s+(gb2312|utf-8|gbk)";
        //匹配输出格式
        static string reg_t = "\\s+-t\\s+(html|pdf|txt|png|jpg|bmp|gif)";
        //匹配输入文件
        static string reg_i = "\\s+-i\\s+\\S*";
        //匹配输出文件
        static string reg_o = "\\s+-o\\s+\\S*";
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("说明：本程序用于将word文档转换为html格式文档，支持.doc和.docx格式");
                Console.WriteLine("用法：word2html.exe <待转换的word文档>");
                Console.WriteLine(
                    "参数:-e gb2312 编码\n" +
                    "     -t [html|pdf|txt|png|jpg|bmp|gif] 输出格式\n" +
                    "     -i 源文件路径\n" +
                    "     -o 输出文件路径");
                Console.WriteLine("Copyleft(C)2015 Solomon");
                Console.ReadLine();
                return;
            }
            //string paramStr = "OfficeDCS.exe test.doc -i test.doc -o 1.html";
            string paramStr = getStr(args);
            Console.WriteLine(paramStr);
            //string srcInputName = args[0]; // 打开文件的位置


            string current_cmd = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            string current_dir = Path.GetDirectoryName(current_cmd);



            string inPath = getParam(paramStr, reg_i);
            inPath = inPath.Replace("-i", "").Trim();
            if (!inPath.Contains("\\"))
            {
                inPath = current_dir + "\\" + inPath;
            }
            Console.WriteLine(inPath);
            string outPath = getParam(paramStr, reg_o);
            outPath = outPath.Replace("-o", "").Trim();
            if (!outPath.Contains("\\"))
            {
                outPath = current_dir + "\\" + outPath; // 同路径保存
            }
            Console.WriteLine(outPath);
            string ext = Path.GetExtension(inPath);
            string type = getParam(paramStr, reg_t);
            type = type.Replace("-t", "").Trim();
            if (ext.Contains("doc"))
            {
                if (type == "pdf")
                {
                    Console.WriteLine("正在生成pdf，请稍候...");
                    XDPI.OfficeUtils.WordToPDF(inPath, outPath);
                    Console.WriteLine("Word文档已转换为pdf格式");
                }
                else if (type == "html")
                {
                    Console.WriteLine("正在生成html，请稍候...");
                    XDPI.OfficeUtils.WordToHtml(inPath, outPath);
                    Encoding enc = Encoding.GetEncoding("GB2312");
                    string s = File.ReadAllText(outPath,enc);
                    s = s.Replace("charset=gb2312", "charset=utf-8");
                    s = XDPI.OfficeUtils.gb2312_utf8(s);
                    File.WriteAllText(outPath, s, Encoding.UTF8);
                    Console.WriteLine("Word文档已转换为html格式");
                }
            }
            else if (ext.Contains("ppt"))
            {
                if (type == "pdf")
                {
                    Console.WriteLine("正在生成pdf，请稍候...");
                    XDPI.OfficeUtils.PowerPointToPDF(inPath, outPath);
                    Console.WriteLine("PowerPoint文档已转换为pdf格式");
                }
                else if (type == "html")
                {
                    Console.WriteLine("正在生成html，请稍候...");
                    XDPI.OfficeUtils.PowerPointToHtml(inPath, outPath);
                    Encoding enc = Encoding.GetEncoding("GB2312");
                    string s = File.ReadAllText(outPath, enc);
                    s = s.Replace("charset=gb2312", "charset=utf-8");
                    s = XDPI.OfficeUtils.gb2312_utf8(s);
                    File.WriteAllText(outPath, s, Encoding.UTF8);
                    Console.WriteLine("PowerPoint文档已转换为html格式");
                }else if (type == "png")
                {
                    Console.WriteLine("正在生成png，请稍候...");
                    XDPI.OfficeUtils.PowerPointToPNG(inPath, outPath);
                    Console.WriteLine("PowerPoint文档已转换为png格式");
                }
                else if (type == "jpg")
                {
                    Console.WriteLine("正在生成jpg，请稍候...");
                    XDPI.OfficeUtils.PowerPointToJPG(inPath, outPath);
                    Console.WriteLine("PowerPoint文档已转换为jpg格式");
                }
                else if (type == "bmp")
                {
                    Console.WriteLine("正在生成bmp，请稍候...");
                    XDPI.OfficeUtils.PowerPointToBMP(inPath, outPath);
                    Console.WriteLine("PowerPoint文档已转换为bmp格式");
                }
                else if (type == "gif")
                {
                    Console.WriteLine("正在生成gif，请稍候...");
                    XDPI.OfficeUtils.PowerPointToGIF(inPath, outPath);
                    Console.WriteLine("PowerPoint文档已转换为gif格式");
                }
            }
            else if (ext.Contains("xls"))
            {
                if (type == "pdf")
                {
                    Console.WriteLine("正在生成pdf，请稍候...");
                    XDPI.OfficeUtils.ExcelToPDF(inPath, outPath);
                    Console.WriteLine("Excel文档已转换为pdf格式");
                }
                else if (type == "html")
                {
                    Console.WriteLine("正在生成html，请稍候...");
                    XDPI.OfficeUtils.ExcelToHtml(inPath, outPath);
                    Encoding enc = Encoding.GetEncoding("GB2312");
                    string s = File.ReadAllText(outPath, enc);
                    s = s.Replace("charset=gb2312", "charset=utf-8");
                    s = XDPI.OfficeUtils.gb2312_utf8(s);
                    File.WriteAllText(outPath, s, Encoding.UTF8);
                    Console.WriteLine("Excel文档已转换为html格式");
                }
            }
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
        static string getParam(string str, string reg)
        {
            try
            {
                Regex r = new Regex(reg, RegexOptions.IgnoreCase);
                Match m = r.Match(str);
                return m.Groups[0].Value.Trim();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return "";
            }
        }

        static string getStr(string[] args)
        {
            string s = "";
            foreach (string str in args)
            {
                s += " " + str;
            }
            return s;
        }
    }


}
