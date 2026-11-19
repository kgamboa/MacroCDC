using Common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Threading;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Linq.Expressions;

namespace TesoreriaCobranza
{
    public class IdentificaBancos : BaseExcel
    {
        private PleaseWait wait;
        private DirectoryInfo DirInfo;

        public IdentificaBancos(DirectoryInfo  dirInfo,string fileName, string empresa,PleaseWait wait)
        {
            FileName = fileName;
            this.Empresa = empresa;
            this.wait = wait;
            this.DirInfo = dirInfo;
            ErrorServices.singleton.Proceso = "IdentificaBancos " + Empresa;
            this.wait.Title = "Procesando " + ErrorServices.singleton.Proceso;
            Utils.EventLog.LogMessage("IdentificaBancos: " + fileName);
        }

        public void EjecutaProceso(CancellationToken cancellationToken)
        {
            try
            {
                if (listRefBBVA == null || listRefBBVA.Count == 0)
                {
                    ObtieneRefBNTC();
                    // Verificamos si se solicitó la cancelación
                    cancellationToken.ThrowIfCancellationRequested();
                }
                if (listRefCIE == null || listRefCIE.Count == 0)
                {
                    ObtieneRefCIE();
                    // Verificamos si se solicitó la cancelación
                    cancellationToken.ThrowIfCancellationRequested();
                }
                if (listBBVA == null || listBBVA.Count == 0)
                {
                    ObtieneBVBA();
                    // Verificamos si se solicitó la cancelación
                    cancellationToken.ThrowIfCancellationRequested();
                }
                if (listCFDI == null || listCFDI.Count == 0)
                {
                    ObtieneCFDI();
                    // Verificamos si se solicitó la cancelación
                    cancellationToken.ThrowIfCancellationRequested();
                }
                if (listSPEI == null || listSPEI.Count == 0)
                {
                    ObtieneSPEI();
                    // Verificamos si se solicitó la cancelación
                    cancellationToken.ThrowIfCancellationRequested();
                }

                var bancoKey = string.Empty;
                var cuenta = string.Empty;
                using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
                {
                    excel.OpenFile(FileName, false);

                    var sheets = excel.GetSheetNames();
                    var empresa = "CC";
                    foreach (var item in sheets)
                    {
                        // Verificamos si se solicitó la cancelación
                        cancellationToken.ThrowIfCancellationRequested();

                        char separator = ' ';
                        if (item.Contains("_"))
                            separator = '_';

                        var split = item.Split(separator);
                        if (split.Length > 1)
                        {
                            if ((split[1] == "CC" || split[1] == "CI" || split[1] == "CL"))
                                empresa = split[1];
                            else
                                cuenta = split[1];
                        }
                        var banco = split[0];
                        if (banco == "SPEI" || banco == "Flujo" || banco == "Cuadre")
                        {
                            continue;
                        }
                        else
                        {
                            bancoKey = FechaColumn.Keys.Where(t => banco.ToUpper().Contains(t)).FirstOrDefault();
                            if (bancoKey == null)
                                continue;
                            banco = item.Replace("_", "");
                        }

                        if (empresa == null)
                            empresa = "CC";
                        if (empresa != Empresa)
                            continue;

                        excel.SetActiveSheet(item);
                        Range excelRange = excel.UsedRange;
                        SetHeader(excelRange, bancoKey, banco);


                        if (TipoMovimientoColumn.ContainsKey(bancoKey))
                        {
                            Utils.EventLog.LogMessage("Identifica Banco: " + banco);
                            excel.SetActiveSheet(item);
                            excelRange = excel.UsedRange;
                            decimal RowCount = excelRange.Rows.Count;
                            int ColumnCount = excelRange.Columns.Count;
                            int emptyCount = 0;
                            for (int r = 2; r <= RowCount; r++)
                            {
                                // Verificamos si se solicitó la cancelación
                                cancellationToken.ThrowIfCancellationRequested();

                                this.wait.ProgressValue = (r / RowCount) * 100;

                                MovimientoBancario nMov = new MovimientoBancario() { Empresa = this.Empresa, Banco = banco, BancoKey = bancoKey, rowIndex = r, Sheet = item, Signo = null };
                                bool? signo = null;
                                Range excelRange1 = excel.UsedRange;
                                dynamic cell = excelRange1.Cells[nMov.rowIndex, TipoMovimientoColumn[bancoKey]];
                                dynamic content = cell.Value;

                                if (content != null && !string.IsNullOrWhiteSpace(content.ToString()))
                                {
                                    if (content == TipoMovimiento.Cobranza.ToString())
                                    {
                                        #region Obtiene informacion Movimiento Banco
                                        cell = (Range)excelRange.Cells[r, RFCColumn];
                                        nMov.RFC = cell.Value2?.ToString();
                                        if (!string.IsNullOrEmpty(nMov.RFC))
                                            continue;
                                        cell = (Range)excelRange.Cells[r, RazonColumn];
                                        nMov.RazonSocial = cell.Value2?.ToString();
                                        if (!string.IsNullOrEmpty(nMov.RazonSocial))
                                            continue;

                                        if (SecuenciaColumn.ContainsKey(bancoKey))
                                        {
                                            cell = (Range)excelRange.Cells[r, SecuenciaColumn[bancoKey]];
                                            nMov.Secuencia = cell.Value2?.ToString();
                                        }
                                        if (ReferenciaColumn.ContainsKey(bancoKey))
                                        {
                                            cell = (Range)excelRange.Cells[r, ReferenciaColumn[bancoKey]];
                                            nMov.Referencia = cell.Value2 == null ? "" : cell.Value2.ToString();
                                        }
                                        if (ReferenciaAmpliadaColumn.ContainsKey(bancoKey))
                                        {
                                            cell = (Range)excelRange.Cells[r, ReferenciaAmpliadaColumn[bancoKey]];
                                            nMov.ReferenciaAmpliada = cell.Value2?.ToString();
                                            if (nMov.ReferenciaAmpliada != null && nMov.ReferenciaAmpliada.Contains("+"))
                                            {
                                                try
                                                {
                                                    nMov.ReferenciaAmpliada = cell.Value2.ToString("0." + new string('#', 339));
                                                }
                                                catch { }
                                            }
                                        }
                                        if (CodigoColumn.ContainsKey(bancoKey))
                                        {
                                            cell = (Range)excelRange.Cells[r, CodigoColumn[bancoKey]];
                                            nMov.CodigoLeyenda = cell.Value2?.ToString();
                                        }

                                        for (int c = 1; c < ColumnCount; c++)
                                        {
                                            decimal importe = 0;
                                            cell = (Range)excelRange.Cells[r, c];
                                            if (FechaColumn.ContainsKey(bancoKey) && c == FechaColumn[bancoKey])
                                            {
                                                DateTime fecha;
                                                if (cell.value is double)
                                                {
                                                    try
                                                    {
                                                        nMov.FechaAplicacion = UtilityDates.FromExcelSerialDate(cell.value);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ex.ToString();
                                                    }
                                                }
                                                else if (DateTime.TryParse(cell.value?.ToString(), out fecha))
                                                {
                                                    nMov.FechaAplicacion = fecha;
                                                }
                                            }
                                            else if (FechaOperacionColumn.ContainsKey(bancoKey) && c == FechaOperacionColumn[bancoKey])
                                            {
                                                DateTime fecha;
                                                if (cell.value is double)
                                                {
                                                    nMov.FechaAplicacion = UtilityDates.FromExcelSerialDate(cell.value);
                                                }
                                                else if (DateTime.TryParse(cell.value?.ToString(), out fecha))
                                                {
                                                    nMov.FechaAplicacion = fecha;
                                                }
                                            }
                                            else if (ConceptoColumn.ContainsKey(bancoKey) && c == ConceptoColumn[bancoKey])
                                            {
                                                nMov.Concepto = cell.Value2?.ToString();
                                            }
                                            else if (ReferenciaAmpliadaColumn.ContainsKey(bancoKey) && c == ReferenciaAmpliadaColumn[bancoKey])
                                            {
                                                nMov.ReferenciaAmpliada = cell.Value2?.ToString();
                                                if (nMov.ReferenciaAmpliada != null && nMov.ReferenciaAmpliada.Contains("+"))
                                                {
                                                    try
                                                    {
                                                        nMov.ReferenciaAmpliada = cell.Value2.ToString("0." + new string('#', 339));
                                                    }
                                                    catch { }
                                                }
                                            }
                                            else if (SignoColumn.ContainsKey(bancoKey) && c == SignoColumn[bancoKey])
                                                signo = cell.Text == "+";

                                            else if (ImporteColumn.ContainsKey(bancoKey) && c == ImporteColumn[bancoKey])
                                            {
                                                if (decimal.TryParse(cell.Value2?.ToString().Replace("$", ""), out importe))
                                                {
                                                    if (signo.HasValue && !signo.Value)
                                                        importe *= -1;

                                                    if (importe > 0)
                                                        nMov.Abono = importe;
                                                    else
                                                        nMov.Cargo = importe;

                                                    nMov.Signo = nMov.ImporteNeto > 0;
                                                }
                                            }
                                            else if (CargoColumn.ContainsKey(bancoKey) && c == CargoColumn[bancoKey])
                                            {
                                                if (decimal.TryParse(cell.Value2?.ToString().Replace("$", ""), out importe))
                                                {
                                                    if (importe > 0) importe *= -1;
                                                    nMov.Cargo = importe;
                                                    nMov.Signo = nMov.ImporteNeto > 0;
                                                }
                                            }
                                            else if (AbonoColumn.ContainsKey(bancoKey) && c == AbonoColumn[bancoKey])
                                            {
                                                if (decimal.TryParse(cell.Value2?.ToString().Replace("$", ""), out importe) && importe > 0)
                                                    nMov.Abono = importe;
                                                nMov.Signo = nMov.ImporteNeto > 0;
                                            }
                                        }
                                        #endregion

                                        if (nMov.FechaAplicacion.Year == 1 && string.IsNullOrEmpty(nMov.Concepto) && nMov.ImporteNeto == 0)
                                            break;

                                        if (listBBVA != null && listBBVA.ContainsKey(cuenta) && listBBVA[cuenta].Count > 0)
                                        {
                                            if (nMov.CodigoLeyenda == "N06" && nMov.ReferenciaAmpliada.Length >= 15)
                                            {
                                                var llave = nMov.ReferenciaAmpliada.Substring(5, 10);
                                                var itemBBVA = listBBVA[cuenta].Where(c => c.Key == llave).Select(c => c.Value).FirstOrDefault();
                                                if (itemBBVA != null)
                                                {
                                                    nMov.RFC = itemBBVA.RFC;
                                                    nMov.RazonSocial = itemBBVA.Otorgante;
                                                    excelRange.Cells[r, RFCColumn].value = nMov.RFC;
                                                    excelRange.Cells[r, NombreOrdenanteColumn].value = itemBBVA.Otorgante;
                                                    excelRange.Cells[r, CuentaOrdenanteColumn].value = "'" + itemBBVA.CuentaOrdenante;
                                                    excelRange.Cells[r, CuentaClabeOrdenanteColumn].value = "'" + itemBBVA.CuentaClabeOrdenante;
                                                    //excelRange.Cells[r, ComentarioColumn].value = "'" + llave;
                                                }
                                            }
                                            else if (nMov.CodigoLeyenda == "W02" && nMov.Referencia.Length > 0)
                                            {
                                                var llave = nMov.Referencia.Replace(" ", "");
                                                var itemBBVA = listBBVA[cuenta].Where(c => c.Key == llave).Select(c => c.Value).FirstOrDefault();
                                                if (itemBBVA != null)
                                                {
                                                    nMov.RFC = itemBBVA.RFC;
                                                    nMov.RazonSocial = itemBBVA.Otorgante;
                                                    excelRange.Cells[r, RFCColumn].value = nMov.RFC;
                                                    excelRange.Cells[r, NombreOrdenanteColumn].value = itemBBVA.Otorgante;
                                                    //excelRange.Cells[r, CuentaOrdenanteColumn].value = "'" + itemBBVA.CuentaOrdenante;
                                                    //excelRange.Cells[r, CuentaClabeOrdenanteColumn].value = "'" + itemBBVA.CuentaClabeOrdenante;
                                                    //excelRange.Cells[r, ComentarioColumn].value = "'" + llave;
                                                }
                                            }
                                        }
                                        if (listSPEI != null && listSPEI.ContainsKey(cuenta) && listSPEI[cuenta].Count > 0)
                                        {
                                            if (nMov.CodigoLeyenda == "Y16")
                                            {
                                                var llave = nMov.Referencia.Replace(" ", "");
                                                var itemSPEI = listSPEI[cuenta].Where(c => c.Key == llave).Select(c => c.Value).FirstOrDefault();
                                                if (itemSPEI != null)
                                                {
                                                    nMov.RFC = itemSPEI.RFC;
                                                    nMov.RazonSocial = itemSPEI.Otorgante;
                                                    excelRange.Cells[r, RFCColumn].value = nMov.RFC;
                                                    excelRange.Cells[r, NombreOrdenanteColumn].value = itemSPEI.Otorgante;
                                                    excelRange.Cells[r, CuentaOrdenanteColumn].value = "'" + itemSPEI.CuentaOrdenante;
                                                    //excelRange.Cells[r, ComentarioColumn].value = "'" + llave;
                                                }
                                            }
                                            else if (nMov.CodigoLeyenda == "T20")
                                            {
                                                var llave = nMov.Referencia.Replace(" ", "");
                                                //var llave = nMov.Referencia.Substring(0, 10);
                                                var itemSPEI = listSPEI[cuenta].Where(c => c.Key == llave).Select(c => c.Value).FirstOrDefault();
                                                if (itemSPEI != null)
                                                {
                                                    if (!string.IsNullOrEmpty(itemSPEI.RFC))
                                                    {
                                                        nMov.RFC = itemSPEI.RFC;
                                                        excelRange.Cells[r, RFCColumn].value = nMov.RFC;
                                                    }
                                                    if (!string.IsNullOrEmpty(itemSPEI.Otorgante))
                                                    {
                                                        nMov.RazonSocial = itemSPEI.Otorgante;
                                                        excelRange.Cells[r, NombreOrdenanteColumn].value = itemSPEI.Otorgante?.Trim();
                                                    }
                                                    excelRange.Cells[r, CuentaOrdenanteColumn].value = "'" + itemSPEI.CuentaOrdenante;
                                                    //excelRange.Cells[r, ComentarioColumn].value = "'" + llave;
                                                }
                                                else
                                                {
                                                    llave += nMov.ImporteNeto.ToString();
                                                    itemSPEI = listSPEI[cuenta].Where(c => c.Value.Referencia.Substring(0, 10) + c.Value.ImporteNeto == llave).Select(c => c.Value).FirstOrDefault();
                                                    if (itemSPEI != null)
                                                    {
                                                        nMov.RFC = itemSPEI.RFC;
                                                        nMov.RazonSocial = itemSPEI.Otorgante;
                                                        excelRange.Cells[r, RFCColumn].value = nMov.RFC;
                                                        excelRange.Cells[r, NombreOrdenanteColumn].value = itemSPEI.Otorgante?.Trim();
                                                        excelRange.Cells[r, CuentaOrdenanteColumn].value = "'" + itemSPEI.CuentaOrdenante;
                                                        //excelRange.Cells[r, ComentarioColumn].value = "'" + llave;
                                                    }
                                                }
                                            }
                                        }

                                        if (string.IsNullOrEmpty(nMov.RFC) && listRefCIE != null && listRefCIE.Count > 0)
                                        {
                                            string rfc = nMov.SearchReferenciaAmpliada(listRefCIE.Keys.ToArray());
                                            if (rfc == null)
                                            {
                                                rfc = nMov.SearchConcepto(listRefCIE.Keys.ToArray());
                                            }
                                            if (rfc != null && listRefCIE.ContainsKey(rfc))
                                            {
                                                nMov.RFC = rfc;
                                                nMov.RazonSocial = listRefCIE[rfc];
                                                excelRange.Cells[r, RFCColumn].value = nMov.RFC;
                                                excelRange.Cells[r, NombreOrdenanteColumn].value = nMov.RazonSocial;
                                            }
                                        }
                                    }
                                }
                                else { emptyCount++;  if (emptyCount == 10) break; }
                            }

                            // SetAutoFit(excelRange, bancoKey);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR_Bancos, ex, ex.StackTrace);
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
        public void ObtieneCFDI(bool repair = false)
        {
            try
            {
                var fileCFDI = DirInfo.GetFiles("*Bnc_comCFDI*.xlsx", SearchOption.AllDirectories);
                listCFDI = new Dictionary<string, Dictionary<string, CFDI>>();
                if (fileCFDI.Count() == 0) throw (new FileNotFoundException("ERROR CON ARCHVO CFDI_ No se encuentra *Bnc_comCFDI*.xls*"));

                foreach (var f in fileCFDI))
                {
                    string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                    FileInfo file = new FileInfo(tempFileName);

                    //create our connection strings
                    string sexcelconnectionstring = ConnectionString(file.FullName, "NO");

                    string myexceldataquery = "select * from [SPEI_SPID$]";

                    //series of commands to bulk copy data from the excel file into our sql table
                    using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                    {
                        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                        oledbconn.Open();
                        OleDbDataReader dataReader = oledbcmd.ExecuteReader();
                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Load(dataReader);
                        decimal count = dt.Rows.Count;
                        var index = 0;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            var dr = dt.Rows[i];

                            CFDI item = new CFDI()
                            {
                                RFCOrdenante = dr[10].ToString(),
                                CuentaOrdenante = dr[12].ToString(),
                                Cuenta = dr[20].ToString()
                            };
                            item.Cuenta = item.Cuenta.Substring(item.Cuenta.Length - 4);
                            if (!listCFDI.ContainsKey(item.Cuenta))
                            {
                                listCFDI.Add(item.Cuenta, new Dictionary<string, CFDI>());
                            }
                            var llave = item.CuentaOrdenante;
                            if (!listCFDI[item.Cuenta].ContainsKey(llave))
                                listCFDI[item.Cuenta].Add(llave, item);
                            this.wait.ProgressValue = (index++ / count) * 100;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            }
        }

        public void ObtieneBVBA(bool repair = false)
        {
            try
            {
                var fileBVBA = DirInfo.GetFiles("*Bnc_comNEM*.xlsx", SearchOption.AllDirectories);
                listBBVA = new Dictionary<string, Dictionary<string, BBVA>>();
                if (fileBVBA.Count() == 0) throw(new FileNotFoundException("ERROR CON ARCHVO BVBA: No se encuentra *Bnc_comNEM*.xlsx"));

                foreach (var f in fileBVBA)
                {
                    Utils.EventLog.LogMessage("ObtieneBVBA : " + f.Name);

                    string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                    FileInfo file = new FileInfo(tempFileName);

                    //create our connection strings
                    string sexcelconnectionstring = ConnectionString(file.FullName, "NO");

                    string myexceldataquery = "select * from [Propias$]";

                    //series of commands to bulk copy data from the excel file into our sql table
                    using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                    {
                        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                        oledbconn.Open();
                        OleDbDataReader dr = oledbcmd.ExecuteReader();

                        decimal count = dr.FieldCount;
                        var index = 0;
                        while (dr.Read())
                        {
                            if (decimal.TryParse(dr[13].ToString(), out decimal importe))
                            {
                                BBVA item = new BBVA()
                                {
                                    Cuenta = dr[0].ToString(),
                                    CuentaOrdenante = dr[3].ToString(),
                                    CuentaClabeOrdenante = dr[4].ToString(),
                                    Otorgante = dr[6].ToString(),
                                    RFC = dr[7].ToString(),
                                    CodigoLeyenda = dr[10].ToString(),
                                    ImporteNeto = importe,
                                    Referencia = dr[15].ToString(),
                                    ReferenciaAmpliada = dr[16].ToString(),
                                };
                                item.Cuenta = item.Cuenta.Substring(item.Cuenta.Length - 4);
                                if (!listBBVA.ContainsKey(item.Cuenta))
                                {
                                    listBBVA.Add(item.Cuenta, new Dictionary<string, BBVA>());
                                }

                                if (item.ReferenciaAmpliada.StartsWith("BNET"))
                                {
                                    var llave = item.ReferenciaAmpliada.Substring(8, 10);
                                    if (!listBBVA[item.Cuenta].ContainsKey(llave))
                                        listBBVA[item.Cuenta].Add(llave, item);
                                }
                                if (item.CodigoLeyenda.Contains("W02") && listRefBBVA != null && listRefBBVA.ContainsKey(item.Referencia))
                                {
                                    item.Otorgante = listRefBBVA[item.Referencia].Ordenante;
                                    var llave = item.Referencia.Replace(" ", "");
                                    if (!listBBVA[item.Cuenta].ContainsKey(llave))
                                        listBBVA[item.Cuenta].Add(llave, item);
                                }
                            }
                            this.wait.ProgressValue = (index++ / count)*100;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);

            }

        }
        public void ObtieneSPEI(bool repair = false)
        {
            try
            { 
                var fileSPEI = DirInfo.GetFiles("*SPEI*.xlsx", SearchOption.AllDirectories);
                listSPEI = new Dictionary<string, Dictionary<string, SPEI>>();
                if (fileSPEI.Count() == 0) throw(new FileNotFoundException("ERROR CON ARCHVO SPEI: No se encuentra *SPEI*.xls*"));

                foreach (var f in fileSPEI)
                {
                    Utils.EventLog.LogMessage("ObtieneSPEI : " + f.Name);

                    string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                    FileInfo file = new FileInfo(tempFileName);

                    //create our connection strings
                    string sexcelconnectionstring = ConnectionString(file.FullName, "NO");

                    string myexceldataquery = "select * from [Hoja1$]";

                    //series of commands to bulk copy data from the excel file into our sql table
                    using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                    {
                        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                        oledbconn.Open();
                        OleDbDataReader dr = oledbcmd.ExecuteReader();
                        var cuenta = string.Empty;

                        decimal count = dr.FieldCount;
                        var index = 0;
                        while (dr.Read())
                        {
                            if (cuenta != string.Empty)
                            {
                                if (decimal.TryParse(dr[5].ToString(), out decimal importe))
                                {
                                    SPEI item = new SPEI()
                                    {
                                        Referencia = dr[3].ToString(),
                                        Otorgante = dr[8].ToString(),
                                        CuentaOrdenante = dr[9].ToString(),
                                        Cuenta = cuenta,
                                        ImporteNeto = importe,
                                    };
                                    if (item.Cuenta.Length > 1)
                                    {
                                        item.Cuenta = item.Cuenta.Substring(item.Cuenta.Length - 4, 4);
                                        if (!listSPEI.ContainsKey(item.Cuenta))
                                        {
                                            listSPEI.Add(item.Cuenta, new Dictionary<string, SPEI>());
                                        }

                                        if (listCFDI !=null && listCFDI.ContainsKey(item.Cuenta))
                                        {

                                            //if(item.Referencia.Length >= 15)
                                            //{
                                            var itemCFDI = listCFDI[item.Cuenta].Where(c => c.Key == item.CuentaOrdenante).Select(c => c.Value).FirstOrDefault();
                                            item.RFC = itemCFDI?.RFCOrdenante;

                                            var llave = item.Referencia.Replace(" ", "");
                                            if (!listSPEI[item.Cuenta].ContainsKey(llave))
                                                listSPEI[item.Cuenta].Add(llave, item);
                                            //}
                                            //if (item.Referencia.Length >= 10 && listCFDI.ContainsKey(item.Cuenta))
                                            //{
                                            //    var itemCFDI = listCFDI[item.Cuenta].Where(c => c.Key == item.CuentaOrdenante).Select(c => c.Value).FirstOrDefault();
                                            //    item.RFC = itemCFDI?.RFCOrdenante;

                                            //    var llave = item.Referencia.Substring(0, 10);
                                            //    if (!listSPEI[item.Cuenta].ContainsKey(llave))
                                            //        listSPEI[item.Cuenta].Add(llave, item);
                                            //}
                                        }
                                    }
                                }
                            }
                            else
                            {
                                cuenta = dr[1].ToString();
                            }
                            this.wait.ProgressValue = (index++ / count) * 100;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            }
        }

        public void ObtieneRefBNTC()
        {
            try
            {
                var f = DirInfo.GetFiles("*BNTC*.xlsx", SearchOption.AllDirectories).FirstOrDefault();
                listRefBBVA = new Dictionary<string, RefBBVA>();
                if (f != null)
                {
                    Utils.EventLog.LogMessage("ObtieneRefBNTC : " + f.Name);

                    string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                    FileInfo fileRefBBVA = new FileInfo(tempFileName);

                    //create our connection strings
                    string sexcelconnectionstring = ConnectionString(fileRefBBVA.FullName, "NO");

                    string myexceldataquery = "select * from [Hoja1$]";

                    //series of commands to bulk copy data from the excel file into our sql table
                    using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                    {
                        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                        oledbconn.Open();
                        OleDbDataReader dr = oledbcmd.ExecuteReader();

                        decimal count = dr.FieldCount;
                        var index = 0;
                        while (dr.Read())
                        {
                            RefBBVA item = new RefBBVA()
                            {
                                //CuentaOrdenante = dr[0].ToString(),
                                ReferenciaBBVA = dr[0].ToString(),
                                Ordenante = dr[1].ToString()
                            };
                            if (!listRefBBVA.ContainsKey(item.ReferenciaBBVA))
                                listRefBBVA.Add(item.ReferenciaBBVA, item);
                            this.wait.ProgressValue = (index++ / count) * 100;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            }
        }
        
        public void ObtieneRefCIE()
        {
            try
            {
                var f = DirInfo.GetFiles("*CIE*.xlsx", SearchOption.AllDirectories).FirstOrDefault();
                listRefCIE = new Dictionary<string, string>();
                if (f != null)
                {
                    Utils.EventLog.LogMessage("ObtieneRefCIE: " + f.Name);

                    string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                    FileInfo fileRefCIE = new FileInfo(tempFileName);

                    //create our connection strings
                    string sexcelconnectionstring = ConnectionString(fileRefCIE.FullName, "NO");

                    string myexceldataquery = "select * from [Hoja1$]";

                    //series of commands to bulk copy data from the excel file into our sql table
                    using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                    {
                        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                        oledbconn.Open();
                        OleDbDataReader dr = oledbcmd.ExecuteReader();

                        decimal count = dr.FieldCount;
                        var index = 0;
                        while (dr.Read())
                        {
                            try
                            {
                                string RFC = dr[0].ToString();
                                string Ordenante = dr[1].ToString();
                                if (!listRefCIE.ContainsKey(RFC))
                                    listRefCIE.Add(RFC.ToUpper(), Ordenante.ToUpper());
                                this.wait.ProgressValue = (index++ / count) * 100;
                            }catch(Exception ex) 
                            {
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            }
        }
    }
}