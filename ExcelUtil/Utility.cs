using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Common;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Collections;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Globalization;

namespace ExcelUtil
{
    public class Utility : IDisposable
    {
        private Application xlExcel;
        private Workbook xlWorkBook;
        private Worksheet _xlWorkSheet;

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        Hashtable myHashtable = new Hashtable();


        public string FileName { get; set; }


        public Range UsedRange => _xlWorkSheet.UsedRange; 
        //public Range UsedRange
        //{
        //    get
        //    {

        //        Worksheet xlWorkSheet = (Worksheet)xlWorkBook.ActiveSheet;
        //        Range br = xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, System.Type.Missing);

        //        return xlWorkSheet.get_Range("A1", br);// Select using cells
        //    }
        //}

        public Application StartExcel()
        {
            xlExcel = new Application { Visible = true };
            return xlExcel;
        }
        public void AddData(string rangeName, string data)
        {
            if (xlExcel != null)
            {
                _xlWorkSheet = (Worksheet)this.xlWorkBook.Worksheets[1];
                Range rangeWrite = _xlWorkSheet.get_Range(rangeName, rangeName);
                rangeWrite.Value = data;

                xlWorkBook.Save();
            }

        }
        public void AddData(string sheetName, System.Windows.Forms.DataGridView dgv)
        {

            if (xlExcel != null)
            {
                _xlWorkSheet = (Worksheet)this.xlWorkBook.Worksheets[sheetName];
                //Range range = this.xlExcel.get_Range(rangeName, Missing.Value);
                Range last = _xlWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Range range = _xlWorkSheet.get_Range("A1", last);

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
                        if (c == 0)
                        {
                            if (DateTime.TryParse(dr[c].ToString(), out DateTime fecha))
                            {
                                string fechaEspañol = fecha.ToString("d", new CultureInfo("es-MX"));
                                arr[r, c] = fechaEspañol;
                            }
                        }
                        else
                        {
                            arr[r, c] = dr[c];
                        }
                    }
                }

                Range c1 = (Range)_xlWorkSheet.Cells[topRow, 1];
                Range c2 = (Range)_xlWorkSheet.Cells[topRow + dt.Rows.Count - 1, dt.Columns.Count];
                Range rangeWrite = _xlWorkSheet.get_Range(c1, c2);

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
        }

        public string Busca(int columna, string filtro, string sheetName)
        {
            using (OleDbConnection dbConnection = new OleDbConnection(ConnectionString(FileName, "No")))
            {
                dbConnection.Open();
                try
                {
                    // Get the name of the first worksheet:
                    System.Data.DataTable dbSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dbSchema == null || dbSchema.Rows.Count < 1)
                    {
                        throw new Exception("Error: Could not determine the name of the first worksheet.");
                    }
                    string firstSheetName = sheetName == null ? dbSchema.Rows[0]["TABLE_NAME"].ToString() : sheetName;

                    // Now we have the table name; proceed as before:
                    using (OleDbCommand dbCommand = new OleDbCommand("SELECT * FROM[" + firstSheetName + "] WHERE F2=@Folio", dbConnection))
                    {
                        if (filtro == "A PUBLICO GRAL")
                            filtro = "VENTAS";
                        dbCommand.Parameters.AddWithValue("@Folio", filtro);
                        var dbReader = dbCommand.ExecuteReader();

                        while (dbReader.Read())
                        {
                            return dbReader[columna].ToString();
                        }
                        dbReader.Close();

                        dbCommand.Parameters.Clear();
                        if (filtro == "VENTAS")
                        {
                            dbCommand.Parameters.AddWithValue("@Folio", "");
                            dbReader = dbCommand.ExecuteReader();

                            while (dbReader.Read())
                            {
                                return dbReader[columna].ToString();
                            }
                            dbReader.Close();
                        }
                        else
                        {

                            dbCommand.Parameters.AddWithValue("@Folio", int.Parse(filtro).ToString("##,###"));
                            dbReader = dbCommand.ExecuteReader();

                            while (dbReader.Read())
                            {
                                return dbReader[columna].ToString();
                            }
                            dbReader.Close();
                        }
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    dbConnection.Close();
                }
                return string.Empty;
            }
        }

        public void SetActiveSheet(string item)
        {
            foreach (Worksheet wSheet in xlExcel.Worksheets)
            {
                if (wSheet.Name == item)
                {
                    wSheet.Activate();
                    _xlWorkSheet = wSheet;
                    break;
                }
            }
        }
        public void MoveActiveSheet(string item, int index =-1)
        {
            foreach (Worksheet wSheet in xlExcel.Worksheets)
            {
                if (wSheet.Name == item)
                {
                    int totalSheets = index;
                    if (index==-1)
                        totalSheets = xlExcel.Worksheets.Count;
                    wSheet.Move(xlExcel.Worksheets[totalSheets]);
                    _xlWorkSheet = wSheet;
                    break;
                }
            }
        }
        public void SetSheetName(string item, string name)
        {

            foreach (Worksheet wSheet in xlExcel.Worksheets)
            {
                if (wSheet.Name == item)
                {
                    wSheet.Name = name;
                    break;
                }
            }
        }

        public void DeleteTable(string tableName)
        {
            _xlWorkSheet = xlWorkBook.Worksheets[1];
            ListObject table = _xlWorkSheet.ListObjects[tableName];
            table.Range.EntireRow.Delete();


            var range = _xlWorkSheet.get_Range("A10", "H10");
            range.EntireRow.Delete();

            range = _xlWorkSheet.get_Range("A8", "H8");
            range.EntireRow.Delete();


        }
        public void AddDataTable(string tableName, List<string[]> Facturas)
        {
            if (Facturas != null)
            {
                _xlWorkSheet = xlWorkBook.Worksheets[1];
                ListObject table = _xlWorkSheet.ListObjects[tableName];
                foreach (string[] item in Facturas)
                {
                    ListRow row = table.ListRows.AddEx(Type.Missing, Type.Missing);
                    int count = 0;
                    foreach (Range r in row.Range.Cells)
                    {
                        if (item.Count() >= count + 1)
                        {
                            var dato = item[count++];
                            if (dato.Contains("/"))
                                r.Value = UtilityDates.f_RecortaFechaMX(dato);
                            else
                                r.Value = dato;
                        }
                    }
                }
            }
        }

       
        private void SaveFile()
        {
            object misValue = System.Reflection.Missing.Value;
            //Save the workbook
            //xlWorkBook.SaveAs(FileName, misValue, misValue, misValue, misValue, false, XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Save();
        }

        public void SaveFileHTML()
        {
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.SaveAs(FileName.Replace(".xlsx", ".html").Replace(".xls", ".html"), XlFileFormat.xlHtml, misValue, misValue, misValue, false, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, XlTextVisualLayoutType.xlTextVisualLTR, misValue);

        }
        public void SaveFilePDF()
        {
            xlWorkBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, FileName.Replace(".xlsx", ".pdf").Replace(".xls", ".pdf"));
        }

        public void SaveFileAs(string fileName)
        {
            object misValue = System.Reflection.Missing.Value;

            //Save the workbook
            xlWorkBook.SaveAs(fileName, misValue, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);

        }
        public static void SaveFilePDF(string fullName)
        {
            using (Utility excel = new Utility())
            {
                excel.OpenFile(fullName);
                excel.SaveFilePDF();
            };
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

        public bool GetdbReader(Action<OleDbDataReader> bulckData, string sheetName = null)
        {
            using (OleDbConnection dbConnection = new OleDbConnection(ConnectionString(FileName, "No")))
            {
                var procesado = false;
                dbConnection.Open();
                try
                {
                    // Get the name of the first worksheet:
                    System.Data.DataTable dbSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dbSchema == null || dbSchema.Rows.Count < 1)
                    {
                        throw new Exception("Error: Could not determine the name of the first worksheet.");
                    }
                    string firstSheetName = sheetName == null ? dbSchema.Rows[0]["TABLE_NAME"].ToString() : sheetName;

                    // Now we have the table name; proceed as before:
                    using (OleDbCommand dbCommand = new OleDbCommand("SELECT * FROM[" + firstSheetName + "]", dbConnection))
                    {
                        var dbReader = dbCommand.ExecuteReader();
                        bulckData(dbReader);
                        dbReader.Close();
                    }
                    procesado = true;
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    dbConnection.Close();
                }
                return procesado;
            }
        }

        internal void CreateSheet(string v, System.Data.DataTable dtResultadoTotales)
        {
            // Creation a new Workbook
            var excelworkBook = xlExcel.Workbooks.Add(Type.Missing);

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
        public static void ObtieneInfoActiveSheet(string fullFilePath)
        {
            try
            {
                using (Utility excel = new Utility())
                {
                    excel.OpenFile(fullFilePath);

                    excel._xlWorkSheet = excel.xlWorkBook.ActiveSheet;
                    Range br = excel._xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, System.Type.Missing);
                    Range oRange;

                    ListObject table = excel._xlWorkSheet.ListObjects["SaldoCorriente"];
                    ListRow row = table.ListRows.AddEx(Type.Missing, Type.Missing);

                    oRange = excel._xlWorkSheet.get_Range("A1", "I" + table.TotalsRowRange.Row);// Select using cells
                    oRange.Copy(Type.Missing);
                    oRange.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlPicture);
                    excel.SaveFileHTML();
                };
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
        }
        public void InsertRows(int rowIndex, int numRows)
        {
            if (_xlWorkSheet == null)
            {
                throw new InvalidOperationException("Workbook or worksheet not initialized. Call OpenFile first.");
            }
            Console.WriteLine($"rowIndex: {rowIndex}, numRows: {numRows}, Last Row: {_xlWorkSheet.Rows.Count}");
            Console.WriteLine($"Range: {rowIndex}, {rowIndex + numRows - 1}");


            for (int i = 1; i < numRows+1; i++)
            {
                _xlWorkSheet.Rows[rowIndex+1 + i].Insert(XlInsertShiftDirection.xlShiftDown);
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

        private void CopyGrid(System.Windows.Forms.DataGridView dataGridView1)
        {
            // I'm making this up...
            dataGridView1.SelectAll();

            dataGridView1.AutoResizeColumns();
            dataGridView1.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            var data = dataGridView1.GetClipboardContent();

            if (data != null)
            {
                //string clipboardText = dataGridView1.GetClipboardContent().GetText();
                //Clipboard.SetText(clipboardText.Replace("/r/n", ""));

                System.Windows.Forms.Clipboard.SetDataObject(data, true);
            }
        }

        public void QuitExcel()
        {
            try
            {
                if (_xlWorkSheet != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(_xlWorkSheet);
                    }
                    finally { }
                }
                if (xlWorkBook != null)
                {
                    try
                    {
                        xlWorkBook.Close(true);
                        Marshal.ReleaseComObject(xlWorkBook);
                    }
                    catch (COMException)
                    {
                    }

                }

                if (xlExcel != null)
                {
                    try
                    {

                        xlExcel.Quit();
                        KillApplicationProcess(xlExcel);

                        Marshal.ReleaseComObject(xlExcel);
                    }
                    catch (COMException)
                    {
                    }

                }
            }
            finally
            {
                xlWorkBook = null;
                _xlWorkSheet = null;
                xlExcel = null;
                GC.Collect();
            }
        }

        public void OpenFile(string fileName,bool visible =true)
        {
            //visible = true;
            if (xlExcel == null)
                xlExcel = new Application {  Interactive = visible, UserControl = visible, DisplayAlerts = visible, ScreenUpdating = visible };
            try
            {
                FileName = fileName;
                xlWorkBook = xlExcel.Workbooks.Open(fileName, 0,false, 5, "", "", false, XlPlatform.xlWindows, "", true,
                    false, 0, true, false, false);
            }
            catch
            {
                //Create a new workbook if the existing workbook failed to open.
                xlWorkBook = xlExcel.Workbooks.Add();
            }
        }

        public void OpenFile(bool visible=false)
        {
            OpenFile(FileName, visible);
        }
        public void CreateSheet(string item, bool finale = false)
        {
            if (xlWorkBook == null)
            {
                // Creation a new Workbook
                xlWorkBook = xlExcel.Workbooks.Add(Type.Missing);
            }

            _xlWorkSheet = ((Workbook)xlWorkBook).Sheets.Add();
            _xlWorkSheet.Name = item;
            int totalSheets = xlExcel.Worksheets.Count;
            if (finale && totalSheets > 0)
            {
                _xlWorkSheet.Move(xlExcel.Worksheets[totalSheets]);
            }
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
                    x = 0;
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

        public static void KillApplicationProcess(Microsoft.Office.Interop.Excel.Application oXL)
        {
            try
            {
                int hWnd = oXL.Application.Hwnd;
                GetWindowThreadProcessId((IntPtr)hWnd, out uint processID);
                Process.GetProcessById((int)processID).Kill();
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void KillExcel()
        {
            //Process[] AllProcesses = Process.GetProcessesByName("excel");

            //// check to kill the right process
            //foreach (Process ExcelProcess in AllProcesses)
            //{
            //    if (ExcelProcess.MainWindowTitle == "")
            //        //ExcelProcess.Kill();
            //}

            //AllProcesses = null;
        }

        internal void ColorTab( int v2)
        {
            if (_xlWorkSheet != null)
            {
                _xlWorkSheet.Tab.ColorIndex = (XlColorIndex)v2;
            }
        }

        /// <summary>
        /// Obtiene archivos CSV y los junta en uno solo.
        /// </summary>
        internal void GeneraDetalleCompleto()
        {
            var files = UtilityIO.GetFilesByExtensions(@"F:\Facturacion\BCC\2022\09-SEPTIEMBRE\DETALLE_COMPLETO", SearchOption.TopDirectoryOnly, "~", "", ".csv");

            var data = new List<string>();
            foreach (var file in files)
            {
                string[] dataInfo = System.IO.File.ReadAllLines(file.FullName);
                data.AddRange(dataInfo);
            }

            System.IO.File.WriteAllLines(@"F:\Facturacion\BCC\2022\09-SEPTIEMBRE\DETALLE_COMPLETO.csv", data.ToArray());
        }
    } 
}