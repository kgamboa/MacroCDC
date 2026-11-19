using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Data;
using System.Data.OleDb;

namespace ExcelUtil
{
    public class Utility : IDisposable
    {
        private Application xlExcel;
        private Workbook xlWorkBook;
        private Worksheet xlWorkSheet;

        public string FileName { get; set; }
        public Application StartExcel()
        {
            QuitExcel();
            xlExcel = new Application { Visible = false };
            xlWorkBook = this.xlExcel.Workbooks.Add(Missing.Value);
            return xlExcel;
        }
        public void AddData(string rangeName, DataGridView dgv)
        {

            if (xlExcel != null)
            {
                xlWorkSheet = (Worksheet)this.xlWorkBook.Worksheets[rangeName];
                //Range range = this.xlExcel.get_Range(rangeName, Missing.Value);
                Range last = xlWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Range range = xlWorkSheet.get_Range("A1", last);

                try
                {
                    System.Data.DataTable dt = (System.Data.DataTable)dgv.DataSource;
                    int topRow = range.Rows.Count + 1;
                    //int rowIndex = iRow + dt.Rows.Count;
                    //range = range.get_Resize(rowIndex, dt.ColumnCount );


                    object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        DataRow dr = dt.Rows[r];
                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            arr[r, c] = dr[c];
                        }
                    }

                    Range c1 = (Range)xlWorkSheet.Cells[topRow, 1];
                    Range c2 = (Range)xlWorkSheet.Cells[topRow + dt.Rows.Count - 1, dt.Columns.Count];
                    Range rangeWrite = xlWorkSheet.get_Range(c1, c2);

                    rangeWrite.Value = arr;

                    //CopyGrid(dt);

                    //Range CR = (Range)xlWorkSheet.Cells[topRow, 1];

                    ////CR.Select();
                    ////xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    //xlWorkSheet.PasteSpecial(CR, false, false, Type.Missing, Type.Missing, Type.Missing, true);


                    ////Fill the array.
                    //foreach(DataRow row in dt.Rows)
                    //{
                    //    for (int iCol = 1; iCol <= dt.Columns.Count; iCol++)
                    //    {
                    //        //Put the row and column address in the cell.
                    //        range.Cells[iRow, iCol] = row[iCol - 1].ToString();
                    //    }
                    //    iRow++;
                    //}
                    xlWorkBook.Save();
                    //dt.CurrentCell = dt.Rows[0].Cells[0];

                    //CR.Select();


                }
                catch (SystemException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public void AddDataTable(string tableName, List<string[]> facturasAdeudo)
        {
            xlWorkSheet = xlWorkBook.Worksheets[1];
            Range tableRange = (Range)xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[5, 5]];
            ListObject list = xlWorkSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, tableRange,
                                                     Type.Missing, XlYesNoGuess.xlYes, Type.Missing);
            list.Name = tableName;
            //list.TableStyle = "TableStyleMedium2";
            ListRow row = list.ListRows.AddEx(Type.Missing, Type.Missing);
            foreach (Range r in row.Range.Cells)
            {
                r.Value = "x";
            }
        }

        public void SaveFile(string fileName)
        {
            object misValue = System.Reflection.Missing.Value;

            //Save the workbook
            xlWorkBook.SaveAs(fileName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, fileName + ".pdf");
        }

        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {

            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }


        internal void CreateSheet(string v, System.Data.DataTable dtResultadoTotales)
        {
            // Creation a new Workbook
            var excelworkBook = this.xlExcel.Workbooks.Add(Type.Missing);

            // Workk sheet
            var excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = "Resultado";

            try
            {

                excelSheet.Cells[1, 1] = "Resultado";
                excelSheet.Cells[1, 2] = "Date : " + DateTime.Now.ToShortDateString();
                var rowCount = dtResultadoTotales.Rows.Count;

                // now we resize the columns
                var excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rowCount, dtResultadoTotales.Columns.Count]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                for (int i = 0; i < dtResultadoTotales.Columns.Count; i++)
                    excelSheet.Cells[1, i + 1] = dtResultadoTotales.Columns[i].ColumnName;

                // Content.  
                for (int i = 0; i < dtResultadoTotales.Rows.Count; i++)
                {
                    for (int j = 0; j < dtResultadoTotales.Columns.Count; j++)
                    {
                        excelSheet.Cells[i + 2, j + 1] = dtResultadoTotales.Rows[i][j].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public List<string> GetSheetNames()
        {
            String[] excelSheets = new String[xlExcel.Worksheets.Count];
            int i = 0;
            foreach (Worksheet wSheet in xlExcel.Worksheets)
            {
                excelSheets[i] = wSheet.Name;
                i++;
            }
            return excelSheets.ToList();
        }

        private void CopyGrid(DataGridView dataGridView1)
        {
            // I'm making this up...
            dataGridView1.SelectAll();

            dataGridView1.AutoResizeColumns();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            var data = dataGridView1.GetClipboardContent();

            if (data != null)
            {
                //string clipboardText = dataGridView1.GetClipboardContent().GetText();
                //Clipboard.SetText(clipboardText.Replace("/r/n", ""));

                Clipboard.SetDataObject(data, true);
            }
        }

        public void QuitExcel()
        {
            try
            {
                if (this.xlWorkBook != null)
                {
                    try
                    {
                        this.xlWorkBook.Close();
                    }
                    catch (COMException)
                    {
                    }

                }

                if (this.xlExcel != null)
                {
                    try
                    {
                        this.xlExcel.Quit();
                    }
                    catch (COMException)
                    {
                    }

                }
                this.xlWorkBook = null;
                this.xlWorkSheet = null;
                this.xlExcel = null;
            }
            finally
            {
                Marshal.ReleaseComObject(this.xlWorkBook);
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(this.xlExcel);
                GC.Collect();
            }
        }

        public void OpenFile(string fileName)
        {
            FileName = fileName;
            OpenFile();
        }

        public void OpenFile()
        {
            StartExcel();
            xlWorkBook = xlExcel.Workbooks.Open(FileName);


            // for making Excel visible
            xlExcel.Visible = false;
            xlExcel.DisplayAlerts = false;
        }

        private string ConnectionString(string FileName, string Header)
        {
            OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder();
            if (Path.GetExtension(FileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", string.Format("Excel 8.0;IMEX=1;HDR={0};", Header));
            }
            else
            {
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
                Builder.Add("Extended Properties", string.Format("Excel 12.0;IMEX=1;HDR={0};", Header));
            }

            Builder.DataSource = FileName;
            return Builder.ConnectionString;
        }

        private string ConvertExcelLineBreaks(string text)
        {
            return text.Replace("\n", Environment.NewLine);
        }
        
        public System.Data.DataTable ToDataTable(object index, string query)
        {


            System.Data.DataTable dt = new System.Data.DataTable();
            using (OleDbConnection cn = new OleDbConnection { ConnectionString = ConnectionString(FileName, "No") })
            {
                using (OleDbCommand cmd = new OleDbCommand { CommandText = query, Connection = cn })
                {
                    cn.Open();
                    OleDbDataReader dr = cmd.ExecuteReader();
                    dt.Load(dr);
                }
            }

            if (index is string)
            {
                var x = 0;
                while (dt.Rows.Count > x && !dt.Rows[x][0].ToString().ToLower().Contains(index.ToString().ToLower()))
                {
                    dt.Rows[x].Delete();
                    x++;
                }
                dt.AcceptChanges();

                // remove header
                if (dt.Rows.Count > 0)
                {
                    x =0;
                    foreach (var item in dt.Rows[0].ItemArray)
                    {
                        if (string.IsNullOrEmpty(item.ToString()))
                            dt.Columns.RemoveAt(x);
                        else
                        {
                            dt.Columns[x].ColumnName = item.ToString().ToUpper();
                            x++;
                        }
                    }
                    dt.Rows[0].Delete();
                }
                dt.AcceptChanges();

            }
            else
            {
                var x = 0;
                while (dt.Rows.Count < x && x != (int)index)
                {
                    // remove header
                    dt.Rows[x].Delete();
                    x++;
                }
                dt.AcceptChanges();
            }

            return dt;
        }

        public void Dispose()
        {
            QuitExcel();
        }

        protected string ExportDatatableToHtml(System.Data.DataTable dt)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");
            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-   family:Garamond; font-size:smaller'>");
            strHTMLBuilder.Append("<tr >");
            foreach (DataColumn myColumn in dt.Columns)
            {
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(myColumn.ColumnName);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            foreach (DataRow myRow in dt.Rows)
            {
                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
            }
            //Close tags.  
            strHTMLBuilder.Append("</table>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");
            string Htmltext = strHTMLBuilder.ToString();
            return Htmltext;
        }
    }
}
