using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Xml.Serialization;
using System.IO;
using AppReviewCollector.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static System.Console;
using System.Threading;
using System.Net;

namespace AppReviewCollector
{
    public static class StringExtensions
    {
        public static string ToShiftJis(this string unicodeStrings)
        {
            var unicode = Encoding.Unicode;
            var unicodeByte = unicode.GetBytes(unicodeStrings);
            var s_jis = Encoding.GetEncoding("shift_jis");
            var s_jisByte = Encoding.Convert(unicode, s_jis, unicodeByte);
            var s_jisChars = new char[s_jis.GetCharCount(s_jisByte, 0, s_jisByte.Length)];
            s_jis.GetChars(s_jisByte, 0, s_jisByte.Length, s_jisChars, 0);
            return new string(s_jisChars);
        }
    }
    public partial class AppReviewController : Form
    {
        public delegate void UpdateStatusHandler(string msg);
        public delegate void ResetProgressHandler(int min, int max);
        public delegate void UpdateProgressHandler(int count);
        private List<Entry> mEntries = new List<Entry>();
        private static Mutex mutex = new Mutex();

        public AppReviewController()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "");
            string id = textAppID.Text;
            progressBar.Value = 0;
            progressBar.Minimum = 0;
            progressBar.Maximum = 10;
            mEntries.Clear();
            //id = "1268959718";
            //id = "com.netmarble.nanatsunotaizai";
            if (string.IsNullOrEmpty(id))
            {
                MessageBox.Show("idを指定してください。idはAppStoreのURLに記載されています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string writePath = namedFile(id);
            if (string.IsNullOrEmpty(writePath))
            {
                MessageBox.Show("保存先ファイル名を指定してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            bool isAppStore = false;
            int iResult = 0;
            if (int.TryParse(id, out iResult))
            {
                if (iResult > 0)
                {
                    isAppStore = true;
                }
            }

            Task.Run(() => {
                /*
                                Uri url = new Uri(@"https://play.google.com/_/PlayStoreUi/data/batchexecute?hl=ja");
                                int sort = 1;
                                int count = 100;
                                string magic = "[[[\"UsvDTd\",\"[null,null,[2," + sort.ToString() + ",[" + count.ToString() + ",null," + "null" + "]],[\\\"" + id + "\\\",7]]\",null,\"generic\"]]]";
                                string formData = "f.req: " + magic;

                                Task<string> ret = post(url, formData);
                                textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), ret.Result);
                */
                if (isAppStore == false)
                {
                    //"https://play.google.com/store/apps/details?id=com.netmarble.nanatsunotaizai&hl=ja"
                    Uri uri = new Uri(@"https://play.google.com/store/apps/details?id=" + id + "&hl=ja");
                    Task<string> ret = get(uri);
                    ret.Wait();
                    if (ret.Result == null)
                    {
                        textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "error: cannot connect");
                        return;
                    }
                    string findStart = "data:[[[\"gp:";
                    string findEnd = ");";
                    string findValue = ret.Result;
                    int start = findValue.IndexOf(findStart) + 8;
                    int end = findValue.IndexOf(findEnd, start + 1);
                    string value = findValue.Substring(start, end - start);
                    value = value.Replace("\n", "");
                    value = value.Replace("[\"gp:", "\n[\"gp:");

                    StreamWriter writer = new StreamWriter(id + ".txt", true, Encoding.UTF8);
                    writer.Write(Utf16ToString(value));
                    writer.Close();

                    string result = Utf16ToString(value);
                    string[] results = result.Split('\n');
                    foreach (var r in results)
                    {
                        string[] columns = r.Split(',');
                        List<string> contents = new List<string>();
                        contents.Add(columns[10].Replace("[", "").Replace("]", "").Replace("\"", ""));
                        contents.Add("");
                        Model.Author a = new Model.Author();
                        a.name = columns[1].Replace("[", "").Replace("]", "").Replace("\"", "");
                        a.uri = columns[7].Replace("[", "").Replace("]", "").Replace("\"", "");
                        string rate = columns[8].Replace("[", "").Replace("]", "").Replace("\"", "");

                        mEntries.Add(new Entry
                        {
                            content = contents,
                            author = a,
                            rating = rate
                        });
                    }
                } else {

                    Uri url = new Uri(@"https://itunes.apple.com/jp/rss/customerreviews/id=" + id + @"/page=1/xml");
                    Task<string> ret = get(url);
                    ret.Wait();
                    if (ret.Result == null)
                    {
                        MessageBox.Show("取得できませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "ダウンロード中...");

                    for (int page = 1; page <= 10; page++)
                    {
                        url = new Uri(@"https://itunes.apple.com/jp/rss/customerreviews/id=" + id + @"/page=" + page.ToString() + "/xml");
                        Task<string> response = get(url);
                        response.Wait();
                        StreamWriter writer = new StreamWriter(id + "_" + page + ".xml", true, Encoding.UTF8);
                        writer.Write(response.Result);
                        var xml = writer.ToString();

                        var xmlSerializer = new XmlSerializer(typeof(AppStore));
                        AppStore appStore = (AppStore)xmlSerializer.Deserialize(new StringReader(response.Result));
                        mutex.WaitOne();
                        foreach (var entry in appStore.entry)
                        {
                            mEntries.Add(entry);
                            progressBar.Invoke(new UpdateProgressHandler(UpdateProgress), 1);
                        }
                        mutex.ReleaseMutex();
                        writer.Close();
                    }
                }
                saveXlsFile(writePath);
                textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "Complete!");
            }
            );

        }

        public static string ToShiftJis(string unicodeStrings)
        {
            var unicode = Encoding.Unicode;
            var unicodeByte = unicode.GetBytes(unicodeStrings);
            var s_jis = Encoding.GetEncoding("shift_jis");
            var s_jisByte = Encoding.Convert(unicode, s_jis, unicodeByte);
            var s_jisChars = new char[s_jis.GetCharCount(s_jisByte, 0, s_jisByte.Length)];
            s_jis.GetChars(s_jisByte, 0, s_jisByte.Length, s_jisChars, 0);
            return new string(s_jisChars);
        }
        public static string Utf16ToString(string utf16String)
        {
            // Storage for the UTF8 string
            string ret = String.Empty;

            // Get UTF16 bytes and convert UTF16 bytes to UTF8 bytes
            byte[] utf16Bytes = Encoding.Unicode.GetBytes(utf16String);
            for (int i = 0;i < utf16Bytes.Length;i += 2)
            {
                byte[] converter = new byte[2];
                converter[0] = utf16Bytes[i];
                converter[1] = utf16Bytes[i + 1];
                if (utf16Bytes[i] == 0x5c && utf16Bytes[i + 1] == 0x00 )
                {
                    i += 2;
                    if (utf16Bytes[i] == 0x75 && utf16Bytes[i + 1] == 0x00)
                    {
                        i += 2;
                        // 0x30 0x31 0x32 0x33 0x34 0x35 0x36 0x37 0x38 0x39
                        //    0    1    2    3    4    5    6    7    8    9
                        // 0x61 0x62 0x63 0x64 0x65 0x66
                        //    a    b    c    d    e    f
                        if (utf16Bytes[i] >= 0x30 && utf16Bytes[i] <= 0x39) { converter[1] = (byte)((utf16Bytes[i] - 0x30) << 4); }
                        if (utf16Bytes[i] >= 0x61 && utf16Bytes[i] <= 0x66) { converter[1] = (byte)((utf16Bytes[i] - 0x57) << 4); }
                        i += 2;
                        if (utf16Bytes[i] >= 0x30 && utf16Bytes[i] <= 0x39) { converter[1] |= (byte)((utf16Bytes[i] - 0x30)); }
                        if (utf16Bytes[i] >= 0x61 && utf16Bytes[i] <= 0x66) { converter[1] |= (byte)((utf16Bytes[i] - 0x57)); }

                        i += 2;
                        if (utf16Bytes[i] >= 0x30 && utf16Bytes[i] <= 0x39) { converter[0] = (byte)((utf16Bytes[i] - 0x30) << 4); }
                        if (utf16Bytes[i] >= 0x61 && utf16Bytes[i] <= 0x66) { converter[0] = (byte)((utf16Bytes[i] - 0x57) << 4); }
                        i += 2;
                        if (utf16Bytes[i] >= 0x30 && utf16Bytes[i] <= 0x39) { converter[0] |= (byte)((utf16Bytes[i] - 0x30)); }
                        if (utf16Bytes[i] >= 0x61 && utf16Bytes[i] <= 0x66) { converter[0] |= (byte)((utf16Bytes[i] - 0x57)); }
                    }

                }
                ret += System.Text.Encoding.Unicode.GetString(converter);
            }

            // Return UTF8
            return ret;
        }

        private void saveXlsFile( string writePath )
        {
            uint recordIndex = 1;
            textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "ファイル保存中...");
            progressBar.Invoke(new ResetProgressHandler(ResetProgress), 0, mEntries.Count());

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(writePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "結果"
                };
                sheets.Append(sheet);
                IEnumerable<Row> row = GetRows(worksheetPart, recordIndex, recordIndex);
                foreach (var r in row)
                {

                }
                InsertCellInWorksheet("A", recordIndex, worksheetPart).CellValue = new CellValue("\"updated\"");
                InsertCellInWorksheet("B", recordIndex, worksheetPart).CellValue = new CellValue("\"id\"");
                InsertCellInWorksheet("C", recordIndex, worksheetPart).CellValue = new CellValue("\"title\"");
                InsertCellInWorksheet("D", recordIndex, worksheetPart).CellValue = new CellValue("\"content\"");
                InsertCellInWorksheet("E", recordIndex, worksheetPart).CellValue = new CellValue("\"content type\"");
                InsertCellInWorksheet("F", recordIndex, worksheetPart).CellValue = new CellValue("\"vote sum\"");
                InsertCellInWorksheet("G", recordIndex, worksheetPart).CellValue = new CellValue("\"vote count\"");
                InsertCellInWorksheet("H", recordIndex, worksheetPart).CellValue = new CellValue("\"rating\"");
                InsertCellInWorksheet("I", recordIndex, worksheetPart).CellValue = new CellValue("\"version\"");
                InsertCellInWorksheet("J", recordIndex, worksheetPart).CellValue = new CellValue("\"author name\"");
                InsertCellInWorksheet("K", recordIndex, worksheetPart).CellValue = new CellValue("\"author url\"");
                InsertCellInWorksheet("L", recordIndex, worksheetPart).CellValue = new CellValue("\"link\"");
                InsertCellInWorksheet("M", recordIndex, worksheetPart).CellValue = new CellValue("\"content\"");
                recordIndex++;
                foreach (var entry in mEntries)
                {
                    List<CellValue> cells = new List<CellValue>();
                    cells.Add(new CellValue("\"" + entry.updated + "\""));
                    cells.Add(new CellValue("\"" + entry.id + "\""));
                    cells.Add(new CellValue("\"" + entry.title + "\""));
                    cells.Add(new CellValue("\"" + entry.content[0] + "\""));
                    cells.Add(new CellValue("\"" + entry.contentType + "\""));
                    cells.Add(new CellValue("\"" + entry.voteSum + "\""));
                    cells.Add(new CellValue("\"" + entry.voteCount + "\""));
                    cells.Add(new CellValue("\"" + entry.rating + "\""));
                    cells.Add(new CellValue("\"" + entry.version + "\""));
                    cells.Add(new CellValue("\"" + entry.author.name + "\""));
                    cells.Add(new CellValue("\"" + entry.author.uri + "\""));
                    cells.Add(new CellValue("\"" + entry.link + "\""));
                    cells.Add(new CellValue("\"" + entry.content[1] + "\""));
                    InsertCellInWorksheet(recordIndex, worksheetPart, cells);
/*
                    InsertCellInWorksheet("A", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.updated + "\"");
                    InsertCellInWorksheet("B", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.id + "\"");
                    InsertCellInWorksheet("C", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.title + "\"");
                    InsertCellInWorksheet("D", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.content[0] + "\"");
                    InsertCellInWorksheet("E", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.contentType + "\"");
                    InsertCellInWorksheet("F", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.voteSum + "\"");
                    InsertCellInWorksheet("G", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.voteCount + "\"");
                    InsertCellInWorksheet("H", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.rating + "\"");
                    InsertCellInWorksheet("I", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.version + "\"");
                    InsertCellInWorksheet("J", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.author.name + "\"");
                    InsertCellInWorksheet("K", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.author.uri + "\"");
                    InsertCellInWorksheet("L", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.link + "\"");
                    InsertCellInWorksheet("M", recordIndex, worksheetPart).CellValue = new CellValue("\"" + entry.content[1] + "\"");
*/
                    progressBar.Invoke(new UpdateProgressHandler(UpdateProgress), 1);
                    recordIndex++;
                }
                workbookpart.Workbook.Save();

            }
        }
        public void UpdateStatus(string msg)
        {
            textResponse.Text = msg;
        }
        public void ResetProgress(int min, int max)
        {
            progressBar.Minimum = min;
            progressBar.Maximum = max;
            progressBar.Value = min;
        }

        public void UpdateProgress(int count)
        {
            progressBar.Increment(count);
            textResponse.Text = progressBar.Value.ToString() + "/" + progressBar.Maximum.ToString();
        }

        private async Task<string> post(Uri url, string formData = "")
        {
            string result;
            try
            {
                using (var client = new HttpClient())
                {
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
                    request.Headers.Add("Accept-Encoding", "gzip, deflate, br");


                    request.Content = new StringContent(formData, Encoding.UTF8, "application/json");
                    var response = await client.SendAsync(request);
                    result = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception e)
            {
                result = null;
            }
            return result;
        }

        private async Task<string> get( Uri url, string formData = "" )
        {
            string result = null;
            int retryCount = 5;
            textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "connecting...");
            while (retryCount-- > 0 && result == null)
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko");
                    client.DefaultRequestHeaders.Add("Accept-Language", "ja-JP");
                    client.Timeout = TimeSpan.FromSeconds(10.0);
                    try
                    {
                        var ret = await client.GetStringAsync(url);
                        result = ret;
                    }
                    catch (Exception e)
                    {
                        result = null;
                        textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "retry:" + (5 - retryCount).ToString() + "/5");
                    }
                }
            }
            return result;
        }
        private string namedFile( string id )
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.FileName = id + ".xlsx";
            sfd.Filter = "Excelファイル(*.xlsx)|*.xlsx;|すべてのファイル(*.*)|*.*";
            sfd.Title = "保存先のファイルを選択してください";
            sfd.RestoreDirectory = true;
            sfd.CheckPathExists = true;


            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            return "";
        }

        private static IEnumerable<Row> GetRows(WorksheetPart wsPart, uint r_base1, uint height)
        {
            uint rmax = r_base1 + height - 1;
            uint nextIndex = r_base1;
            foreach (Row row in wsPart.Worksheet.Descendants<Row>())
            {
                uint rowIndex = row.RowIndex;
                if (rowIndex < r_base1)
                {
                }
                else if (rowIndex > rmax)
                {
                    for (uint ui = nextIndex; ui < rmax; ui++)
                    {
                        yield return null;
                    }
                    break;
                }
                else
                {
                    for (uint ui = nextIndex; ui < rowIndex; ui++)
                    {
                        yield return null;
                    }
                    yield return row;
                    nextIndex++;
                    if (nextIndex > rmax)
                    {
                        break;
                    }
                }

            }
            for (uint ui = nextIndex; ui < rmax; ui++)
            {
                yield return null;
            }
        }
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex.ToString();

            Row row　= sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            Cell refCell = row.Elements<Cell>().FirstOrDefault(c =>
                  c.CellReference.Value == cellReference);
            if (refCell != null)
                return refCell;

            Cell nextCell = row.Elements<Cell>().FirstOrDefault(c =>
                  string.Compare(c.CellReference.Value, cellReference, true) > 0);
            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, nextCell);

            worksheet.Save();
            return newCell;
        }
        private static bool InsertCellInWorksheet(uint rowIndex, WorksheetPart worksheetPart, List<CellValue> cells)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            string[] columnNames = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M" };
            int i = 0;
            foreach (var cell in cells)
            {
                string cellReference = columnNames[i++] + rowIndex.ToString();
                Cell nextCell = row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, cellReference, true) > 0);
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, nextCell);
                newCell.CellValue = cell;
            }

            worksheet.Save();
            return true;
        }
    }
}
