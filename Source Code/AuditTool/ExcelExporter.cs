using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Drawing;
namespace XBRL
{
    public class ExcelExporter : IDisposable
    {
        List<string> sheetName;     // 試算表名稱清單
        List<DataTable> table;      // 試算表中的表單清單
        List<string> sheetFooter;     // 試算表名稱清單
        
        MemoryStream ms;            // 將Excel的檔案匯出成Stream
        /// <summary>
        /// 匯出Excel類別建構子
        /// </summary>
        public ExcelExporter()
        {
            sheetName = new List<string>();
            table = new List<DataTable>();
            sheetFooter = new List<string>();
            ms = new MemoryStream();
        }
        /// <summary>
        /// 一開始先新增所有試算表名稱和表單至清單中
        /// </summary>
        /// <param name="sheetName">試算表名稱</param>
        /// <param name="table">表單名稱</param>
        public void Add(string sheetName, DataTable table)
        {
            this.sheetName.Add(sheetName);
            this.table.Add(table);
        }
        /// <summary>
        /// 一開始先新增所有試算表名稱和表單至清單中
        /// </summary>
        /// <param name="sheetName">試算表名稱</param>
        /// <param name="table">表單名稱</param>
        public void Add(string sheetName, DataTable table, string sheetFooter)
        {
            this.sheetName.Add(sheetName);
            this.table.Add(table);
            this.sheetFooter.Add(sheetFooter);
        }
        /// <summary>
        /// 取得Excel的檔案的串流資訊
        /// </summary>
        /// <returns></returns>
        public System.IO.MemoryStream GetExcel()
        {
            string fileGUID = Guid.NewGuid().ToString();
            string fileName = "c:\\" + fileGUID + ".xslx";
            ExportDataTable(fileName);
            FileStream fstream = new FileStream(fileName, FileMode.Open);
            byte[] Data = new Byte[fstream.Length];
            ////Obtain the file into the array of bytes from streams.
            fstream.Read(Data, 0, Data.Length);
            fstream.Close();
            ms = new MemoryStream(Data);
            File.Delete(fileName);
            return ms;
        }
        /// <summary>
        /// 匯出資料儲存成Excel的檔案
        /// </summary>
        /// <param name="exportFile">匯出Excel檔案路徑及名稱</param>
        public void ExportDataTable(string exportFile)
        {
            // 建立試算表
            using (SpreadsheetDocument spreadsheet =
                SpreadsheetDocument.Create(exportFile,SpreadsheetDocumentType.Workbook))
            {
                DataRow contentRow;
                // 建立工作簿
                spreadsheet.AddWorkbookPart();
                spreadsheet.WorkbookPart.Workbook = new Workbook();     // create the worksheet
                // 建立表至工作簿
                spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                // 建立多個工作表
                for (int n = 0; n < table.Count; n++)
                {
                    try
                    {
                        spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                        spreadsheet.WorkbookPart.WorksheetParts.ElementAt(n).Worksheet = new Worksheet();


                        // 最適欄寬(若不先建立讀取會出錯)
                        Columns sheetColumns = new Columns();
                        spreadsheet.WorkbookPart.WorksheetParts.ElementAt(n).Worksheet.AppendChild(sheetColumns);
                        double[] columnFitWidth = new double[table[n].Columns.Count];
                        // 建立工作表資料
                        SheetData sheetData = new SheetData();
                        spreadsheet.WorkbookPart.WorksheetParts.ElementAt(n).Worksheet.AppendChild(sheetData);

                        // 建立標題
                        Cell titleCell = new Cell();
                        titleCell.DataType = new EnumValue<CellValues>(CellValues.String);
                        titleCell.CellValue = new CellValue();
                        titleCell.CellValue.Text = sheetName[n];
                        Row titleRow = new Row();
                        titleRow.AppendChild(titleCell);
                        sheetData.AppendChild(titleRow);
                        MergeCell mergeCell = new MergeCell();
                        MergeCells mergeCells = new MergeCells();
                        mergeCell.Reference = "A1:" + (char)(table[n].Columns.Count + 65 - 1) + "1";
                        mergeCells.AppendChild(mergeCell);
                        // 建立欄位名稱
                        Row header = new Row();
                        header.RowIndex = (UInt32)2;
                        int num = 0;
                        foreach (DataColumn column in table[n].Columns)
                        {
                            Cell headerCell = createTextCell(
                                table[n].Columns.IndexOf(column) + 1,
                                2,
                                column.ColumnName);
                            columnFitWidth[num] = calculateColumnWidth(column.ColumnName);
                            header.AppendChild(headerCell);
                            num = num + 1;
                        }

                        sheetData.AppendChild(header);

                        // 將所有資料寫入                    
                        for (int i = 0; i < table[n].Rows.Count; i++)
                        {
                            contentRow = table[n].Rows[i];

                            sheetData.AppendChild(createContentRow(contentRow, i + 3));
                        }
                        spreadsheet.WorkbookPart.WorksheetParts.ElementAt(n).Worksheet.Save();
                        spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
                        {
                            Id = spreadsheet.WorkbookPart.GetIdOfPart(spreadsheet.WorkbookPart.WorksheetParts.ElementAt(n)),
                            SheetId = (UInt32)(n + 1),
                            Name = sheetName[n].Split(' ')[0]
                        });
                        // 最適欄寬

                        for (int c = 0; c < table[n].Columns.Count; c++)
                        {
                            for (int r = 0; r < table[n].Rows.Count; r++)
                            {
                                if (columnFitWidth[c] < calculateColumnWidth(table[n].Rows[r][c].ToString()))
                                {
                                    columnFitWidth[c] = calculateColumnWidth(table[n].Rows[r][c].ToString());
                                }
                            }
                        }
                        for (int c = 0; c < table[n].Columns.Count; c++)
                        {
                            Column sheetColumn = new Column();
                            sheetColumn.BestFit = true;
                            sheetColumn.Min = (UInt32)c + 1;
                            sheetColumn.Max = (UInt32)c + 1;
                            sheetColumn.Width = columnFitWidth[c];
                            sheetColumn.CustomWidth = true;
                            sheetColumns.AppendChild(sheetColumn);
                        }
                        // 建立頁腳
                        try
                        {
                            Cell footerCell = new Cell();
                            footerCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            footerCell.CellValue = new CellValue();
                            footerCell.CellValue.Text = sheetFooter[n];
                            Row footerRow = new Row();
                            footerRow.AppendChild(footerCell);
                            sheetData.AppendChild(footerRow);
                            mergeCell = new MergeCell();
                            mergeCell.Reference = "A" + (table[n].Rows.Count + 3).ToString() + ":" + (char)(table[n].Columns.Count + 65 - 1) + (table[n].Rows.Count + 3).ToString();
                            mergeCells.AppendChild(mergeCell);
                        }
                        catch
                        {
                        }
                        spreadsheet.WorkbookPart.WorksheetParts.ElementAt(n).Worksheet.AppendChild(mergeCells);
                    }
                    catch
                    {
                    }

                }
                // 儲存
                spreadsheet.WorkbookPart.Workbook.Save();

            }
        }
        /// <summary>
        /// 計算欄位寬度
        /// </summary>
        /// <param name="sILT">判斷的文字</param>
        /// <returns>計算後最適合的寬度</returns>
        private double calculateColumnWidth(string sILT)
        {
            double fDigitWidth = 0.0f;
            double fMaxDigitWidth = 0.0f;
            double fTruncWidth = 0.0f;

            System.Drawing.Font drawfont = new System.Drawing.Font("微軟正黑體", 14);
            // I just need a Graphics object. Any reasonable bitmap size will do.
            Graphics g = Graphics.FromImage(new Bitmap(200, 200));
            for (int i = 0; i < sILT.ToCharArray().Length; ++i)
            {
                fDigitWidth = (double)g.MeasureString(sILT.ToCharArray().GetValue(i).ToString(), drawfont).Width;
                if (fDigitWidth > fMaxDigitWidth)
                {
                    fMaxDigitWidth = fDigitWidth;
                }
            }
            g.Dispose();

            // Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}] / {Maximum Digit Width} * 256) / 256
            fTruncWidth = Math.Truncate((sILT.ToCharArray().Count() * fMaxDigitWidth + 5.0) / fMaxDigitWidth * 256.0) / 256.0 * 2;
            if (Single.IsInfinity((float)fTruncWidth))
            {
                return 0;
            }
            return fTruncWidth;
        }
        /// <summary>
        /// 建立OpenXML對應的Cell
        /// </summary>
        private Cell createTextCell(int columnIndex,int rowIndex ,object cellValue)
        {
            Cell cell = new Cell();

            cell.DataType = CellValues.InlineString;
            cell.CellReference = getColumnName(columnIndex) + rowIndex;
            
            InlineString inlineString = new InlineString();
            Text t = new Text();

            t.Text = cellValue.ToString();

            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);
            
            return cell;
        }
        /// <summary>
        /// 建立OpenXML對應的Row
        /// </summary>
        private Row createContentRow(DataRow dataRow,int rowIndex)
        {
            Row row = new Row

            {
                RowIndex = (UInt32)rowIndex
            };

            for (int i = 0; i < dataRow.Table.Columns.Count; i++)
            {
                Cell dataCell = createTextCell(i + 1, rowIndex, dataRow[i]);
                row.AppendChild(dataCell);
            }
            return row;
        }
        /// <summary>
        /// 取得欄位名稱
        /// </summary>
        private string getColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }
        /// <summary>
        /// 當不用此類別時所做的處置
        /// </summary>
        public void Dispose()
        {
            ms.Close();
            ms.Dispose();
        }
    }
}