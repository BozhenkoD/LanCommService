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

namespace WinServices
{

    public class Search
    {
        private Packet Packets;
        
        public Search(Packet pak)
        {
            Packets = pak;
        }

        double a = 0, b = 0;

        private void FindInDir(DirectoryInfo dir, string logFileName, int CardNumbReg, int FilesTypesReg, int mainReg, long files, long currentFiles)
        {
            foreach (FileInfo file in dir.GetFiles())
            {
                if (mainReg == 1)
                {
                    if (FilesTypesReg == 1)
                    {
                        if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".xls" || file.Extension.ToLower() == ".xlsx")
                        {
                            files++;
                            continue;
                        }
                        continue;
                    }
                    else
                    {
                        files++;
                        continue;
                    }
                }
                currentFiles++;
                string Path;
                a = currentFiles;
                b = files;

                Packets.Progress = ((int)((a / b) * 100));
               

                if (file.Extension.ToLower() == ".exe" || file.Extension.ToLower() == ".dll")
                    continue;
                if (file.Length > 120000000)
                    continue;

                if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx" || file.Extension.ToLower() == ".xls" || file.Extension.ToLower() == ".xlsx")
                {
                    if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".docx")
                    {
                        if (MSWordMatches(file.FullName, CardNumbReg) == true)
                        {
                            Path = file.FullName;
                            //bgW.ReportProgress(-1, Path);
                           
                           // WriteLog(logFileName);
                        }
                    }
                    else
                    {
                        if (MSExcelMatches(file.FullName, CardNumbReg) == true)
                        {
                            Path = file.FullName;
                            //bgW.ReportProgress(-1, Path);
                            //WriteLog(logFileName);
                        }
                    }
                    continue;
                }

                if ((FindMatches(file.FullName, CardNumbReg) == true) && FilesTypesReg != 1)
                {
                    Path = file.FullName;
                }

                WriteLog(logFileName);
            }

            foreach (DirectoryInfo subdir in dir.GetDirectories())
            {
                try
                {
                    this.FindInDir(subdir, logFileName, CardNumbReg, FilesTypesReg, mainReg, ref files, ref currentFiles);
                }
                catch (Exception)
                {

                }
            }

        }
        public void FindFiles(Packet Pak, int mainReg)
        {
            string LogFile = (DateTime.Now.ToString() + ".txt").Replace(':', '-');

            try
            {
                this.FindInDir(new DirectoryInfo(Pak.Directory), LogFile, Convert.ToInt32(Pak.CVV), Convert.ToInt32(Pak.MSOffice), mainReg, Pak.CountFiles, Pak.CurrentFile);
            }
            catch (Exception)
            {
                return;
            }
        }

        public bool FindMatches(string fPath, int reg)
        {

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

           

            if (matches1.Count != 0)
            {
                foreach (Match match in regex1.Matches(fl))
                {
                    int i = match.Index;
                    if (CalcLune(fl.Substring(i, 16)))
                    {
                        Packets.FileInfo.Add(fPath);
                        Packets.FileInfo.Add(fl.Substring(i, 16));
                        Packets.FileInfo.Add(i.ToString());
                        
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
                        Packets.FileInfo.Add(fPath);
                        Packets.FileInfo.Add(fl.Substring(i, 23));
                        Packets.FileInfo.Add(i.ToString());
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
                        Packets.FileInfo.Add(fPath);
                        Packets.FileInfo.Add(fl.Substring(i, 19));
                        Packets.FileInfo.Add(i.ToString());
                    }
                }
            }

            return true;

        }

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

        public bool MSWordMatches(string fPath, int reg)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = false;
            Document doc = app.Documents.Open(fPath);

            //Get all words
            string allWords = doc.Content.Text;
            doc.Close();
            app.Quit();

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

            if (matches1.Count != 0)
            {
                foreach (Match match in regex1.Matches(allWords))
                {
                    int i = match.Index;
                    if (CalcLune(allWords.Substring(i, 16)))
                        return true;

                }
            }
            if (matches2.Count != 0)
            {
                

                foreach (Match match in regex2.Matches(allWords))
                {
                    int i = match.Index;
                    if (CalcLune(allWords.Substring(i, 23).Replace("\r", "").Replace("\n", "")))
                    {
                        

                        return true;
                    }

                }
            }
            if (matches3.Count != 0)
            {
                foreach (Match match in regex3.Matches(allWords))
                {
                    int i = match.Index;
                    if (CalcLune(allWords.Substring(i, 19).Replace(" ", "").Replace("\r", "").Replace("\n", "")))
                        return true;

                }
            }

            return false;

        }
        public bool MSExcelMatches(string fPath, int reg)
        {

            Microsoft.Office.Interop.Excel.Application excelapp;
            Microsoft.Office.Interop.Excel.Workbook excelappworkbook;

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



            string cellValue = "";

            for (int sheet = 1; sheet <= CountSheet; sheet++)
            {
                for (int i = 1; i <= usedRowsNum; i++)
                {
                    for (int y = 1; y <= usedColumnsNum; y++)
                    {
                        Microsoft.Office.Interop.Excel.Range cellRange = (Microsoft.Office.Interop.Excel.Range)excelworksheet.Cells[i, y];

                        if (cellRange.Value != null)
                        {
                            cellValue += cellRange.Value.ToString();
                        }
                    }
                }
                cellValue += "\r\n";
            }
            excelappworkbook.Save();
            excelappworkbook.Close();

            excelapp.Quit();


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
            MatchCollection matches1 = regex1.Matches(cellValue);
            MatchCollection matches2 = regex2.Matches(cellValue);
            MatchCollection matches3 = regex3.Matches(cellValue);

            if (matches1.Count != 0)
            {
                foreach (Match match in regex1.Matches(cellValue))
                {
                    int i = match.Index;
                    if (CalcLune(cellValue.Substring(i, 16)))
                    {

                        return true;
                    }

                }
            }
            if (matches2.Count != 0)
            {
                foreach (Match match in regex2.Matches(cellValue))
                {
                    int i = match.Index;
                    if (CalcLune(cellValue.Substring(i, 23).Replace("\r", "").Replace("\n", "")))
                    {
                        return true;
                    }

                }
            }
            if (matches3.Count != 0)
            {
                foreach (Match match in regex3.Matches(cellValue))
                {
                    int i = match.Index;
                    if (CalcLune(cellValue.Substring(i, 19).Replace(" ", "").Replace("\r", "").Replace("\n", "")))
                    {

                        return true;
                    }


                }
            }



            return false;

        }


        private void WriteLog(string logFileName)
        {
            try
            {
                StreamWriter sw = new StreamWriter(logFileName, true);

                foreach (var item in Packets.FileInfo)
                {
                    sw.WriteLine(item);
                }

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
