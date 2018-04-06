using HtmlAgilityPack;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace HeartSongs
{
    public class DataStore
    {
        public string Date { get; set; }
        public string Time { get; set; }
        public string Artist { get; set; }
        public string Song { get; set; }

        public DataStore()
        {
            Date = "";
            Time = "";
            Artist = "";
            Song = "";
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var html = @"http://www.heart.co.uk/radio/last-played-songs/";
            HtmlWeb web = new HtmlWeb();
            var htmlDoc = web.Load(html);

            var node = htmlDoc.DocumentNode.SelectSingleNode("//body");

            string artistName = "";
            string songName = "";
            string timePlayed = "";

            bool songFound = false;
            bool artistFound = false;

            List<DataStore> dataList = new List<DataStore>();

            foreach (var nNodeSpan in node.Descendants("span"))
            {
                int pFrom = 0;
                int pTo = 0;

                if (nNodeSpan.OuterHtml.Contains(@"itemprop=""name""") && nNodeSpan.OuterHtml.Contains(@"class=""track"""))
                {
                    pFrom = nNodeSpan.OuterHtml.IndexOf(">") + ">".Length;
                    pTo = nNodeSpan.OuterHtml.LastIndexOf("<");

                    songName = nNodeSpan.OuterHtml.Substring(pFrom, pTo - pFrom).Trim();
                    songName = songName.Replace(System.Environment.NewLine, "");
                    songName = Regex.Replace(songName, @"\t|\n|\r", "");
                    songName = songName.Replace("  ", "");
                    songName = songName.Replace("&#39;", "'");
                    songName = songName.Replace("&amp;", "&");
                    
                    songFound = true;
                }

                if (nNodeSpan.OuterHtml.Contains(@"itemprop=""byArtist""") && nNodeSpan.OuterHtml.Contains(@"class=""artist"""))
                {
                    pFrom = nNodeSpan.OuterHtml.IndexOf(">") + ">".Length;
                    pTo = nNodeSpan.OuterHtml.LastIndexOf("<");

                    artistName = nNodeSpan.OuterHtml.Substring(pFrom, pTo - pFrom).Trim();
                    artistName = artistName.Replace(System.Environment.NewLine, "");
                    artistName = Regex.Replace(artistName, @"\t|\n|\r", "");
                    artistName = artistName.Replace("  ", "");
                    artistName = artistName.Replace("&#39;", "'");
                    artistName = artistName.Replace("&amp;", "&");
                    
                    artistFound = true;
                }

                if (songFound == true && artistFound == true)
                {
                    DataStore x = new DataStore();
                    x.Date = DateTime.Now.ToString("dd/MM/yyyy");
                    x.Time = timePlayed;
                    x.Artist = artistName;
                    x.Song = songName;

                    dataList.Add(x);

                    songFound = false;
                    artistFound = false;

                    artistName = "";
                    songName = "";
                }
            }

            foreach (var nNode in node.Descendants("div"))
            {
                Console.WriteLine(dataList[dataList.Count - 1].Time);
                if (nNode.OuterHtml.Contains(@"class=""apple_music"""))
                {
                    int i = 0;
                    int count = 0;
                    while ((i = nNode.OuterHtml.IndexOf(@"<p class=""publish_date"">", i)) != -1)
                    {
                        int pFrom = nNode.OuterHtml.IndexOf(@"<p class=""publish_date"">", i);
                        int pTo = nNode.OuterHtml.IndexOf("</p>", i);

                        timePlayed = nNode.OuterHtml.Substring(pFrom + 24, (pTo - pFrom) - 24).Trim();
                        dataList[count].Time = timePlayed;

                        count++;
                        i++;
                    }
                }
                if (dataList[dataList.Count - 1].Time != "")
                {
                    break;
                }
            }

            foreach (DataStore x in dataList)
            {
                Console.WriteLine(x.Date);
                Console.WriteLine(x.Time);
                Console.WriteLine(x.Artist);
                Console.WriteLine(x.Song);
                Console.WriteLine("");
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheetData;

            xlApp = new Excel.Application();

            string filePath = @"C:\Users\joshua.whitfield\C# Projects\HeartSongs\HeartPlaylist.xlsx";
            xlWorkBook = xlApp.Workbooks.Open(filePath);

            xlWorkSheetData = (Excel.Worksheet)(xlWorkBook.Sheets[1]);

            Excel.Range last = xlWorkSheetData.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = xlWorkSheetData.get_Range("A1", last);

            int lastUsedRow = last.Row;
            int row = lastUsedRow + 1;
            foreach (DataStore x in dataList)
            {
                //TODO - Fix American date format error when writing to Excel
                xlWorkSheetData.Cells[row, 1] = x.Date;
                xlWorkSheetData.Cells[row, 2] = x.Time;
                xlWorkSheetData.Cells[row, 3] = x.Artist;
                xlWorkSheetData.Cells[row, 4] = x.Song;
                row++;
            }

            last = xlWorkSheetData.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row;
            Excel.Range rng = xlWorkSheetData.Range["A1:D" + lastUsedRow.ToString(), Type.Missing];
            object cols = new object[] { 1, 4 };
            rng.RemoveDuplicates(cols, Excel.XlYesNoGuess.xlYes);

            //TODO - Sort by date then time desc
            dynamic allDataRange = xlWorkSheetData.UsedRange;
            allDataRange.Sort(allDataRange.Columns[2], Excel.XlSortOrder.xlDescending);

            //TODO - Refresh all Pivot Tables

            xlApp.DisplayAlerts = false;
            xlWorkBook.Close(true, filePath);
        }
    }
}