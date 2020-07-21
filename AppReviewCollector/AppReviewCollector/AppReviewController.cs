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
using System.Collections.Generic;
using System.Linq;
using static System.Console;
using System.Threading;

namespace AppReviewCollector
{
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
            string id = textAppID.Text;
            progressBar.Value = 0;
            progressBar.Minimum = 0;
            progressBar.Maximum = 10;
            //id = "1268959718";
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


            Task.Run(() => {
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
                    foreach ( var entry in  appStore.entry)
                    {
                        mEntries.Add(entry);
                        /*
                                                InsertCellInWorksheet("A", recordIndex, worksheetPart).CellValue = new CellValue(entry.updated);
                                                InsertCellInWorksheet("B", recordIndex, worksheetPart).CellValue = new CellValue(entry.id);
                                                InsertCellInWorksheet("C", recordIndex, worksheetPart).CellValue = new CellValue(entry.title);
                                                InsertCellInWorksheet("D", recordIndex, worksheetPart).CellValue = new CellValue(entry.content[0]);
                                                InsertCellInWorksheet("E", recordIndex, worksheetPart).CellValue = new CellValue(entry.contentType);
                                                InsertCellInWorksheet("F", recordIndex, worksheetPart).CellValue = new CellValue(entry.voteSum);
                                                InsertCellInWorksheet("G", recordIndex, worksheetPart).CellValue = new CellValue(entry.voteCount);
                                                InsertCellInWorksheet("H", recordIndex, worksheetPart).CellValue = new CellValue(entry.rating);
                                                InsertCellInWorksheet("I", recordIndex, worksheetPart).CellValue = new CellValue(entry.version);
                                                InsertCellInWorksheet("J", recordIndex, worksheetPart).CellValue = new CellValue(entry.author.name);
                                                InsertCellInWorksheet("K", recordIndex, worksheetPart).CellValue = new CellValue(entry.author.uri);
                                                InsertCellInWorksheet("L", recordIndex, worksheetPart).CellValue = new CellValue(entry.link);
                                                InsertCellInWorksheet("M", recordIndex, worksheetPart).CellValue = new CellValue(entry.content[1]);
                        */
                        /*                        streamWriter.WriteLine(
                                                    "\"" + entry.updated + "\", " +
                                                    "\"" + entry.id + "\", " +
                                                    "\"" + entry.title + "\", " +
                                                    "\"" + entry.content[0] + "\", " +
                                                    "\"" + entry.contentType + "\", " +
                                                    "\"" + entry.voteSum + "\", " +
                                                    "\"" + entry.voteCount + "\", " +
                                                    "\"" + entry.rating + "\", " +
                                                    "\"" + entry.version + "\", " +
                                                    "\"" + entry.author.name + "\", " +
                                                    "\"" + entry.author.uri + "\", " +
                                                    "\"" + entry.link + "\", " +
                                                    "\"" + entry.content[1] + "\""
                                                    );
                        */
                        progressBar.Invoke(new UpdateProgressHandler(UpdateProgress), 1);
                    }
                    mutex.ReleaseMutex();
                    writer.Close();
                }
                saveXlsFile(writePath);
                textResponse.Invoke(new UpdateStatusHandler(UpdateStatus), "Complete!");
            }
            );

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
        }

        private async Task<string> get( Uri url )
        {
            string result;
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko");
                    client.DefaultRequestHeaders.Add("Accept-Language", "ja-JP");
                    client.Timeout = TimeSpan.FromSeconds(10.0);

                    var ret = await client.GetStringAsync(url);
                    result = ret;
                }
            }
            catch(Exception e)
            {
                result = null;
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
    }
}
