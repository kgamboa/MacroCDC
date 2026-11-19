using Common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Globalization;
using CsvHelper;
using System.Security.Cryptography.Xml;
using System.Xml.Linq;
using System.Data.Common;

namespace TesoreriaCobranza
{
    public class IdentificaFacturas : BaseExcel
    {


        private static Dictionary<string, Dictionary<string, Factura>> listFacturaCFDI;
        private static Dictionary<string, List<string>> listUsadaCFDI;
        public List<decimal> listImporteProveedores;
        public List<OpenPayLiquidacion> listOpenPay;
        private static Dictionary<string, List<Otorgante>> otorganteRazonSocialList;
        private static Dictionary<string, List<string>> razonSocialList;
        private decimal difAceptada;
        private PleaseWait wait;
        private CancellationToken CancellationToken;
        public DirectoryInfo DirInfo;

        public IdentificaFacturas(DirectoryInfo dirInfo, string fileName, string empresa, PleaseWait wait)
        {
            this.FileName = fileName;
            this.Empresa = empresa;
            this.DirInfo = dirInfo;

            ErrorServices.singleton.Proceso = "IdentificaFacturas " + Empresa;
            Utils.EventLog.LogMessage("IdentificaFacturas : " + Path.GetFileNameWithoutExtension(FileName));
            difAceptada = 0.01M;

            razonSocialList = new Dictionary<string, List<string>> { { Empresa, new List<string>() } };
            listRefBBVA = new Dictionary<string, List<RefBBVA>> { { Empresa, new List<RefBBVA>() } };
            listUsadaCFDI = new Dictionary<string, List<string>>();
            listOpenPay = new List<OpenPayLiquidacion>();
            listImporteProveedores = new List<decimal>();
            otorganteRazonSocialList = new Dictionary<string, List<Otorgante>> { { Empresa, new List<Otorgante>() } };
            listFacturaCFDI = new Dictionary<string, Dictionary<string, Factura>> { { Empresa, new Dictionary<string, Factura>() } };
            this.wait = wait;
            this.CancellationToken = wait.CancellationToken.Token;
        }


        public new void EjecutaProceso()
        {
            #region IDENTIFICA FILES
            ObtieneReferencias();
            CancellationToken.ThrowIfCancellationRequested();
            ObtieneOtorgantes();
            CancellationToken.ThrowIfCancellationRequested();
            ObtieneSAP();
            CancellationToken.ThrowIfCancellationRequested();

            if (Empresa == "CL")
            {
                ObtieneImportes();
                CancellationToken.ThrowIfCancellationRequested();
            }

            ObtieneOpenPay();
            CancellationToken.ThrowIfCancellationRequested();

            ObtieneFacturas();
            CancellationToken.ThrowIfCancellationRequested();
            #endregion

            #region BANCOS
            ObtieneRefBNTC();
            // Verificamos si se solicitó la cancelación
            CancellationToken.ThrowIfCancellationRequested();

            if (listRefCIE == null || listRefCIE.Count == 0)
            {
                ObtieneRefCIE();
                // Verificamos si se solicitó la cancelación
                CancellationToken.ThrowIfCancellationRequested();
            }
            if (listBBVA == null || listBBVA.Count == 0)
            {
                ObtieneBVBA();
                // Verificamos si se solicitó la cancelación
                CancellationToken.ThrowIfCancellationRequested();
            }
            if (listCFDI == null || listCFDI.Count == 0)
            {
                ObtieneCFDI();
                // Verificamos si se solicitó la cancelación
                CancellationToken.ThrowIfCancellationRequested();
            }
            if (listSPEI == null || listSPEI.Count == 0)
            {
                ObtieneSPEI();
                // Verificamos si se solicitó la cancelación
                CancellationToken.ThrowIfCancellationRequested();
            }
            #endregion

            var bancoKey = string.Empty;
            var cuenta = string.Empty;
            using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
            {
                excel.OpenFile(FileName, false);

                var sheets = excel.GetSheetNames();
                var empresa = "CC";
                var indexSheet = 0;
                foreach (var item in sheets)
                {
                    // Verificamos si se solicitó la cancelación
                    CancellationToken.ThrowIfCancellationRequested();
                    this.wait.ProgressValue = (indexSheet++ / (decimal)sheets.Count) * 100;

                    List<char> separator = new List<char>() { ' ', '_' };

                    var split = item.Split(separator.ToArray());
                    if (split.Length > 1)
                    {
                        if ((split.Last() == "CC" || split.Last() == "CI" || split.Last() == "CL"))
                            empresa = split.Last();
                        else
                            cuenta = split[1];
                    }
                    var banco = split.First();
                    if (banco == "SPEI" || banco == "Flujo" || banco == "Cuadre")
                    {
                        continue;
                    }
                    else
                    {
                        bancoKey = FechaColumn.Keys.Where(t => banco.ToUpper().Contains(t)).FirstOrDefault();
                        if (bancoKey == null)
                            continue;
                        banco = item.Replace("_", "").Replace(" ", "");
                    }

                    if (empresa == null)
                        empresa = "CC";
                    if (empresa != Empresa)
                        continue;


                    excel.SetActiveSheet(item);
                    var excelRange = excel.UsedRange;

                    if (TipoMovimientoColumn.ContainsKey(bancoKey))
                    {

                        Utils.EventLog.LogMessage("Identifica Facturas: " + item);

                        SetHeader(excelRange, bancoKey);
                        excelRange = excel.UsedRange;
                        decimal RowCount = excelRange.Rows.Count;
                        int ColumnCount = excelRange.Columns.Count;
                        int emptyCount = 0;
                        for (int r = 1; r <= RowCount; r++)
                        {
                            // Verificamos si se solicitó la cancelación
                            CancellationToken.ThrowIfCancellationRequested();

                            var progress = (r / (decimal)RowCount) * 100;
                            if (bancoKey.StartsWith("BBVA"))
                                this.wait.ProgressValue = 100 - progress;
                            else
                                this.wait.ProgressValue = progress;

                            MovimientoBancario nMov = new MovimientoBancario() { Empresa = this.Empresa, Banco = banco, BancoKey = bancoKey, rowIndex = r, Sheet = item, Signo = null };
                            bool? signo = null;
                            Range excelRange1 = excel.UsedRange;
                            dynamic cell = excelRange1.Cells[nMov.rowIndex, TipoMovimientoColumn[bancoKey]];
                            dynamic content = cell.Value;

                            if (content != null && !string.IsNullOrWhiteSpace(content.ToString()))
                            {
                                emptyCount = 0;
                                if (content == TipoMovimiento.Cobranza.ToString())
                                {

                                    #region Obtiene informacion Movimiento Banco
                                    cell = (Range)excelRange.Cells[r, RFCColumn];
                                    nMov.RFC = cell.Value2?.ToString();

                                    cell = (Range)excelRange.Cells[r, RazonColumn];
                                    nMov.RazonSocial = cell.Value2?.ToString();

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

                                    Console.WriteLine(nMov.ToString());

                                    if (!string.IsNullOrEmpty(nMov.Concepto))
                                    {
                                        if (nMov.Concepto.StartsWith("CE"))
                                        {
                                            if (nMov.Concepto.Length == 15)
                                            {
                                                nMov.RFC = nMov.Concepto.Substring(2, 12);
                                                var otorgante = otorganteRazonSocialList[nMov.Empresa].FirstOrDefault(c => nMov.Concepto.StartsWith("CE" + c.RFC));
                                                if (otorgante != null)
                                                {
                                                    nMov.RFC = otorgante.RFC.ToString();
                                                    nMov.RazonSocial = string.Join(",", otorganteRazonSocialList[nMov.Empresa].Where(c => c.RFC == otorgante.RFC.ToString()).Select(c => c.RazonSocial).ToArray());
                                                }
                                            }
                                            else if (nMov.Concepto.Length == 16)
                                            {
                                                var match = Regex.Match(nMov.Concepto, @"^(CI|CE)(\d+)(\w{6})"); if (match.Success)
                                                {
                                                    nMov.RFC = match.Groups[3].Value;
                                                    nMov.NumSAP = match.Groups[2].Value;
                                                }
                                                var otorgante = otorganteRazonSocialList[nMov.Empresa].FirstOrDefault(c => c.RFC.Length > 5 && nMov.Concepto.StartsWith("CE" + c.NumSAP + c.RFC.Substring(0, 6)));
                                                if (otorgante != null)
                                                {
                                                    nMov.RFC = otorgante.RFC.ToString();
                                                    nMov.RazonSocial = string.Join(",", otorganteRazonSocialList[nMov.Empresa].Where(c => c.RFC == otorgante.RFC.ToString()).Select(c => c.RazonSocial).ToArray());
                                                }
                                            }
                                        }
                                        if (nMov.Concepto.StartsWith("CI"))
                                        {
                                            if (nMov.Concepto.Length == 19)
                                            {
                                                nMov.RFC = nMov.Concepto.Substring(2, 12);
                                                var otorgante = otorganteRazonSocialList[nMov.Empresa].FirstOrDefault(c => nMov.Concepto.StartsWith("CI" + c.RFC));
                                                if (otorgante != null)
                                                {
                                                    nMov.RFC = otorgante.RFC.ToString();
                                                    nMov.RazonSocial = string.Join(",", otorganteRazonSocialList[nMov.Empresa].Where(c => c.RFC == otorgante.RFC.ToString()).Select(c => c.RazonSocial).ToArray());
                                                }
                                            }
                                            else if (nMov.Concepto.Length == 20)
                                            {
                                                var match = Regex.Match(nMov.Concepto, @"^(CI|CE)(\d+)(\w{6})"); if (match.Success)
                                                {
                                                    nMov.RFC = match.Groups[3].Value;
                                                    nMov.NumSAP = match.Groups[2].Value;
                                                }
                                                var otorgante = otorganteRazonSocialList[nMov.Empresa].FirstOrDefault(c => c.RFC.Length > 5 && nMov.Concepto.StartsWith("CI" + c.NumSAP + c.RFC.Substring(0, 6)));
                                                if (otorgante != null)
                                                {
                                                    nMov.RFC = otorgante.RFC.ToString();
                                                    nMov.RazonSocial = string.Join(",", otorganteRazonSocialList[nMov.Empresa].Where(c => c.RFC == otorgante.RFC.ToString()).Select(c => c.RazonSocial).ToArray());
                                                }
                                            }
                                        }
                                    }
                                    if (Empresa == "CL")
                                    {
                                        if (nMov.Banco.Contains("EUR"))
                                        {
                                            if (string.IsNullOrEmpty(nMov.RFC))
                                                nMov.RFC = "CCR041221JS4";
                                            if (string.IsNullOrEmpty(nMov.RazonSocial))
                                                nMov.RazonSocial = "CIRCULO DE CREDITO";
                                        }
                                    }
                                    if (nMov.ReferenciaAmpliada != null)
                                    {
                                        if (nMov.ReferenciaAmpliada.ToUpper().Contains("PAGO OPENMX"))
                                        {
                                            var cellOpenPay = (Range)excelRange.Cells[nMov.rowIndex, ImporteColumn[nMov.BancoKey]];
                                            cellOpenPay.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);


                                            DateTime fecha = nMov.FechaAplicacion;
                                            var openPayLiquidacions = listOpenPay.Where(c => fecha.Date == c.PayDate.Value.Date && Math.Abs(nMov.ImporteNeto - c.Amount) < 1).ToList();
                                            if (openPayLiquidacions.Count() == 1)
                                            {
                                                var charges = openPayLiquidacions.First().Charges;
                                                excel.InsertRows(nMov.rowIndex, charges.Count());

                                                excelRange = excel.UsedRange;
                                                EditaExcel(excelRange, nMov, charges);
                                            }
                                            else if (openPayLiquidacions.Count() > 0)
                                            {
                                                var liquidaciones = openPayLiquidacions
                                                    .Select(i => i.FileName)
                                                    .ToList();
                                                nMov.Comentario += "OpenPay Archivos: " + string.Join(",", liquidaciones);
                                                EditaExcel(excelRange, nMov);
                                            }
                                            continue;
                                        }
                                    }

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

                                    if (string.IsNullOrEmpty(nMov.RFC) || nMov.Concepto.StartsWith("SPEI"))
                                    {
                                        var dataList = string.Join("¬", nMov.ReferenciaAmpliada, nMov.Concepto, nMov.RFC, nMov.ConceptoSPEI).Split('¬').Where(c => c.Length > 0).Distinct().ToList();
                                        foreach (var data in dataList)
                                        {
                                            // Verificamos si se solicitó la cancelación
                                            CancellationToken.ThrowIfCancellationRequested();

                                            var rfc = nMov.SearchStringSimilitud(data, otorganteRazonSocialList[nMov.Empresa].Where(c => !string.IsNullOrEmpty(c.RFC)).Select(c => c.RFC).ToArray());
                                            if (rfc != null)
                                            {
                                                nMov.RFC = rfc.ToString();
                                                nMov.RazonSocial = string.Join(",", otorganteRazonSocialList[nMov.Empresa].Where(c => c.RFC == rfc.ToString()).Select(c => c.RazonSocial).ToArray());
                                                break;
                                            }

                                        }
                                    }
                                    if (string.IsNullOrEmpty(nMov.RazonSocial) || nMov.Concepto.StartsWith("SPEI"))
                                    {
                                        var dataList2 = string.Join("¬", nMov.NombreOrdenanteSPEI, nMov.ReferenciaAmpliada, nMov.Concepto, nMov.RazonSocial).Split('¬').Where(c => c.Length > 0).Distinct().ToList();
                                        foreach (var data in dataList2)
                                        {
                                            // Verificamos si se solicitó la cancelación
                                            CancellationToken.ThrowIfCancellationRequested();
                                            var razonSocial1 = nMov.SearchStringSimilitud(data, razonSocialList[nMov.Empresa].ToArray());
                                            if (razonSocial1 != null)
                                            {
                                                nMov.RazonSocial = razonSocial1.ToString();
                                                nMov.RFC = otorganteRazonSocialList[nMov.Empresa].FirstOrDefault(c =>
                                                {
                                                    return c.RazonSocial.Contains(nMov.RazonSocial);
                                                })?.RFC;
                                                break;
                                            }
                                            else if (bancoKey.StartsWith("BBVA"))
                                            {
                                                // Verificamos si se solicitó la cancelación
                                                CancellationToken.ThrowIfCancellationRequested();
                                                var referencia = nMov.SearchStringSimilitud(data, listRefBBVA[nMov.Empresa].Where(c => !string.IsNullOrEmpty(c.ReferenciaBBVA)).Select(c => c.ReferenciaBBVA).ToArray());
                                                if (referencia != null)
                                                {
                                                    var o = listRefBBVA[Empresa].FirstOrDefault(c => c.ReferenciaBBVA.Equals(referencia));
                                                    if (o != null)
                                                    {
                                                        nMov.RazonSocial = o.Ordenante;
                                                        razonSocial1 = otorganteRazonSocialList[
                                                            nMov.Empresa].FirstOrDefault(c => nMov.RazonSocial.Contains(c.RazonSocial)).RazonSocial;
                                                        nMov.RazonSocial = razonSocial1?.ToString();
                                                        nMov.RFC = otorganteRazonSocialList[nMov.Empresa].FirstOrDefault(c => c.RazonSocial.Contains(nMov.RazonSocial))?.RFC;

                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    List<double> listaFolios = nMov.ObtieneFolio();
                                    if (listaFolios.Count > 0)
                                    {
                                        var busquedaFolios = listFacturaCFDI[nMov.Empresa].Values.Where(c => listaFolios.Contains(c.FolioNumerico) || listaFolios.Contains(c.FolioSustNumerico));
                                        if (busquedaFolios.GroupBy(c => c.RFC).Count() == 1)
                                        {
                                            if (!string.IsNullOrEmpty(nMov.RFC) && busquedaFolios.Any(c => c.RFC == nMov.RFC))
                                            {
                                                var f = busquedaFolios.FirstOrDefault();
                                                nMov.RFC = f.RFC;
                                                nMov.RazonSocial = f.RazonSocial;
                                            }
                                            else if (!string.IsNullOrEmpty(nMov.RazonSocial) && busquedaFolios.Any(c => nMov.RazonSocial.Contains(c.RazonSocial)))
                                            {
                                                var f = busquedaFolios.FirstOrDefault();
                                                nMov.RFC = f.RFC;
                                                nMov.RazonSocial = f.RazonSocial;
                                            }
                                        }
                                    }

                                    var facturasActivas = listFacturaCFDI[nMov.Empresa]
                                        .Select(c => c.Value)
                                        // .Where(c => c.EstatusFactura != MetodoIdentificacion.Compensada)
                                        .OrderByDescending(c => c.Fecha)
                                        .ToList();
                                    if (!listUsadaCFDI.ContainsKey(nMov.Sheet))
                                        listUsadaCFDI.Add(nMov.Sheet, new List<string>());
                                    else
                                    {
                                        var facturasUsadas = facturasActivas.Where(c => listUsadaCFDI[nMov.Sheet].Contains(c.Folio)).ToList();
                                        facturasActivas.RemoveAll(c => facturasUsadas.Select(x => x.Folio).Contains(c.Folio));
                                    }


                                    var facturasRFC = new List<Factura>();
                                    if (!string.IsNullOrEmpty(nMov.RFC))
                                    {
                                        var list1 = facturasActivas.Where(c => !string.IsNullOrEmpty(c.RFC) && c.RFC.Contains(nMov.RFC)).ToList();
                                        facturasRFC.AddRange(list1);
                                    }
                                    if (!string.IsNullOrEmpty(nMov.RazonSocial))
                                    {
                                        var list2 = facturasActivas.Where(c => !string.IsNullOrEmpty(c.RazonSocial) && c.RazonSocial.Contains(nMov.RazonSocial)).ToList();
                                        if (list2.Count > 0)
                                        {
                                            var lookup = facturasRFC.ToDictionary(x => x.Folio);
                                            var NoRFC = list2.Where(c => !lookup.ContainsKey(c.Folio)).ToList();
                                            foreach (var x in NoRFC)
                                            {
                                                if (string.IsNullOrEmpty(x.RFC))
                                                    x.RFC = nMov.RFC;
                                            }
                                            facturasRFC.AddRange(NoRFC);
                                        }
                                    }
                                    nMov.EncontradoRazonSocial = string.Concat(nMov.RFC ?? "", nMov.RazonSocial ?? "").Length > 0;
                                    nMov.Comentario = string.Empty;
                                    Factura factura = Revisafacturas(nMov.EncontradoRazonSocial ? facturasRFC : facturasActivas, nMov);
                                    nMov.Comentario = nMov.Comentario + string.Join(",", factura?.EstatusFactura, factura?.Folio ?? "", factura?.FolioSust != null ? "Sust:" + factura?.FolioSust : "");//.Replace(",,,","");
                                    if (factura != null)
                                    {
                                        factura.EstatusFactura = factura.ConteoMetodo > 1 ? MetodoIdentificacion.Posible : factura.EstatusFactura;
                                        Console.WriteLine(factura?.ToString());
                                        Console.WriteLine(nMov.Comentario);
                                        if (!string.IsNullOrEmpty(factura?.Folio))
                                            listUsadaCFDI[nMov.Sheet].AddRange(factura.Folio.Split(';'));
                                    }
                                    else
                                    {
                                        facturasRFC = facturasRFC.Where(c => c.Fecha >= DateTime.Now.AddMonths(-1)).ToList();
                                        facturasActivas = facturasActivas.Where(c => c.Fecha >= DateTime.Now.AddMonths(-1)).ToList();
                                        nMov.Comentario = string.Empty;
                                        factura = Revisafacturas(nMov.EncontradoRazonSocial ? facturasRFC : facturasActivas, nMov);
                                        nMov.Comentario = nMov.Comentario + string.Join(",", factura?.EstatusFactura, factura?.Folio ?? "", factura?.FolioSust != null ? "Sust:" + factura?.FolioSust : "");//.Replace(",,,","");
                                        if (factura != null)
                                        {
                                            factura.EstatusFactura = factura.ConteoMetodo > 1 ? MetodoIdentificacion.Posible : factura.EstatusFactura;
                                            Console.WriteLine(factura?.ToString());
                                            Console.WriteLine(nMov.Comentario);
                                            if (!string.IsNullOrEmpty(factura?.Folio))
                                                listUsadaCFDI[nMov.Sheet].AddRange(factura.Folio.Split(';'));
                                        }
                                    }

                                    EditaExcel(excelRange, nMov, factura);


                                    #region BUsueda de varios pagos a una sola factura
                                    //{
                                    //    var lista2 = listBANCOS[item].Where(c => string.IsNullOrEmpty(c.Value.Estatus) && !string.IsNullOrEmpty(c.Value.RFC)).GroupBy(c => c.Value.RFC).Select(x => new
                                    //    {
                                    //        GroupName = x.Key,
                                    //        GroupSize = x.Count(),
                                    //        GroupItems = x.Select(c => c.Value).ToList(),
                                    //        Sum = x.Sum(c => c.Value.ImporteNeto),
                                    //        ImporteList = x.AsEnumerable().Select((c, index) => new UtilitySum.
                                    //        IElement
                                    //        { RowIndex = index, Total = c.Value.ImporteNeto }).ToList()
                                    //    });

                                    //    var lista3 = listBANCOS[item].Where(c => string.IsNullOrEmpty(c.Value.Estatus) && !string.IsNullOrEmpty(c.Value.RazonSocial)).GroupBy(c => c.Value.RazonSocial).Select(x => new
                                    //    {
                                    //        GroupName = x.Key,
                                    //        GroupSize = x.Count(),
                                    //        GroupItems = x.Select(c => c.Value).ToList(),
                                    //        Sum = x.Sum(c => c.Value.ImporteNeto),
                                    //        ImporteList = x.AsEnumerable().Select((c, index) => new UtilitySum.
                                    //        IElement
                                    //        { RowIndex = index, Total = c.Value.ImporteNeto }).ToList()
                                    //    });


                                    //    var facturasActivas = listCFDI[Empresa].Select(c => c.Value).Where(c => c.EstatusFactura == MetodoIdentificacion.NoIdentificado.ToString()).OrderBy(c => c.Fecha).ToList();
                                    //    var facturasUsadas = facturasActivas.Where(c => c.EstatusFactura != MetodoIdentificacion.NoIdentificado.ToString() || listUsadaCFDI[item].Contains(c.Folio)).ToList();
                                    //    facturasActivas.RemoveAll(c => facturasUsadas.Select(x => x.Folio).Contains(c.Folio));
                                    //    var list4 = new List<Factura>();
                                    //    foreach (var o in lista2.Where(c => c.GroupSize > 1).OrderBy(c => c.GroupSize).ToList())
                                    //    {
                                    //        foreach (var f in facturasActivas)
                                    //        {
                                    //            //Si se pagaron varias.
                                    //            var list = UtilitySum.Search(f.Total, o.GroupSize, o.ImporteList);
                                    //            if (list.Count > 0)
                                    //            {
                                    //                foreach (var y in list)
                                    //                {
                                    //                    var nMovFound = o.GroupItems[y];
                                    //                    f.EstatusFactura = MetodoIdentificacion.SumaFacturas.ToString();
                                    //                    f.Total = nMovFound.ImporteNeto;
                                    //                    EditaExcel(excelRange, nMovFound, f);
                                    //                }
                                    //            }
                                    //        }
                                    //    }
                                    //    facturasUsadas = facturasActivas.Where(c => c.EstatusFactura != MetodoIdentificacion.NoIdentificado.ToString() || listUsadaCFDI[item].Contains(c.Folio)).ToList();
                                    //    facturasActivas.RemoveAll(c => facturasUsadas.Select(x => x.Folio).Contains(c.Folio));
                                    //    foreach (var o in lista3.Where(c => c.GroupSize > 1).OrderBy(c => c.GroupSize).ToList())
                                    //    {
                                    //        foreach (var f in facturasActivas)
                                    //        {
                                    //            //Si se pagaron varias.
                                    //            var list = UtilitySum.Search(f.Total, o.GroupSize, o.ImporteList);
                                    //            if (list.Count > 0)
                                    //            {
                                    //                foreach (var y in list)
                                    //                {
                                    //                    var nMovFound = o.GroupItems[y];
                                    //                    if (!nMovFound.Comentario.Contains("\r\n"))
                                    //                    {
                                    //                        f.EstatusFactura = MetodoIdentificacion.SumaFacturas.ToString();
                                    //                        f.Total = nMovFound.ImporteNeto;
                                    //                        EditaExcel(excelRange, nMovFound, f);
                                    //                    }
                                    //                }
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    #endregion

                                    SetAutoFit(excelRange, bancoKey);
                                }
                            }
                            else
                            {
                                emptyCount++; if (emptyCount == 10) break;
                            }
                        }
                    }
                }
            }
        }

        private Factura Revisafacturas(List<Factura> facturas, MovimientoBancario nMov)
        {
            try
            {
                ////COMPENSADA
                //var facturasImporte = facturas.Where(c => c.Diff == 0 && (c.Total == nMov.ImporteNeto || (c.Total <= (nMov.ImporteNeto + 0.01M) && c.Total >= (nMov.ImporteNeto - 0.01M)))).ToList();
                //if (facturasImporte.Count == 1)
                //{
                //    var factura = facturasImporte.First();
                //    nMov.Comentario += "; Compensada:" + factura.Folio;
                //}

                var folioFiscal = nMov.SearchReferenciaAmpliada(facturas.Where(c=>!string.IsNullOrEmpty(c.FolioFiscal) && c.FolioFiscal.Length>0).Select(c => c.FolioFiscal).ToArray());
                if (folioFiscal != null)
                {
                    facturas = facturas.Where(c => c.FolioFiscal == folioFiscal).ToList();
                    if (facturas.Count >= 1)
                    {
                        var factura = facturas.First();
                        factura.EstatusFactura = MetodoIdentificacion.FolioFactura;
                        return factura;
                    }
                }
                List<double> listaFolios = nMov.ObtieneFolio();


                var facturasImporteDiff = facturas.Where(c => Math.Abs(c.ImporteNeto - nMov.ImporteNeto) <= difAceptada || Math.Abs(c.Diff - nMov.ImporteNeto) <= difAceptada).ToList(); if (facturasImporteDiff.Count == 1)
                {
                    var factura = facturasImporteDiff.First();
                    //factura.Total = nMov.ImporteNeto;
                    factura.EstatusFactura = MetodoIdentificacion.ImporteAprox;
                    if (nMov.EncontradoRazonSocial)
                        factura.EstatusFactura |= MetodoIdentificacion.RazonSocial;
                    if (listaFolios.Contains(factura.FolioNumerico) || listaFolios.Contains(factura.FolioSustNumerico))
                        factura.EstatusFactura |= MetodoIdentificacion.FolioFactura;
                    return factura;
                }
                else if (facturasImporteDiff.Count > 1)
                {
                    facturas = facturasImporteDiff;
                }

                if (listaFolios.Count > 0)
                {
                    var busquedaFolios = facturas.Where(c => listaFolios.Contains(c.FolioNumerico) || listaFolios.Contains(c.FolioSustNumerico)).ToList();
                    if (busquedaFolios.Count > 0)
                    {
                        if (busquedaFolios.Count == 1 && busquedaFolios.GroupBy(c => c.RFC).Count() == 1)
                        {
                            var factura = busquedaFolios.First();
                            factura.EstatusFactura = MetodoIdentificacion.FolioFactura;
                            if (nMov.EncontradoRazonSocial)
                                factura.EstatusFactura |= MetodoIdentificacion.RazonSocial;
                            if (Math.Abs(factura.ImporteNeto - nMov.ImporteNeto) <= difAceptada)
                                factura.EstatusFactura |= MetodoIdentificacion.ImporteAprox;
                            return factura;
                        }
                        if (busquedaFolios.GroupBy(c => c.RFC).Count() == 1 || busquedaFolios.GroupBy(c => c.RazonSocial).Count() == 1)
                        {
                            nMov.Comentario += "; Posible Folio:" + string.Join(";", busquedaFolios.Select(c => c.Folio));
                        }
                    }
                    facturas = busquedaFolios;
                }

                if (facturas.Count > 0)
                {
                    #region Por RFC
                    var lista2 = facturas.Where(c => !string.IsNullOrEmpty(c.RFC) && Math.Abs(c.ImporteNeto - nMov.ImporteNeto) <= difAceptada)
                    .GroupBy(c => c.RFC)
                    .Select(x => new
                    {
                        GroupName = x.Key,
                        GroupSize = x.Count(),
                        GroupItems = x.ToList(),
                        Sum = x.Sum(c => c.ImporteNeto),
                        ImporteList = x.AsEnumerable().Select((c, index) => new UtilitySum.
                        IElement
                        { RowIndex = index, Total = c.ImporteNeto }).ToList()
                    }).Where(c => Math.Abs(c.Sum - nMov.ImporteNeto) <= difAceptada)
                    .ToList();

                    //Si se pagaron varias.
                    var lista3 = lista2
                        .OrderBy(c => c.GroupSize)
                        .Where(c => c.GroupSize > 1)//&& c.Sum <= (nMov.ImporteNeto + difAceptada))
                        .ToList();
                    Console.WriteLine(nMov.ImporteNeto.ToString() + "," + lista3.Count());

                    var listFacturasResult = new List<List<Factura>>();
                    foreach (var o in lista3)
                    {
                        // Verificamos si se solicitó la cancelación
                        CancellationToken.ThrowIfCancellationRequested();
                        var list = UtilitySum.Search(nMov.ImporteNeto, o.GroupSize, o.ImporteList);
                        if (list.Count > 0)
                        {
                            var listFacturas1 = new List<Factura>();
                            foreach (var y in list)
                            {
                                var factItem = o.GroupItems[y];
                                listFacturas1.Add(factItem);
                            }
                            Console.WriteLine(listFacturas1.Select(c => c.ImporteNeto).Sum());
                            Console.WriteLine(string.Join(";", listFacturas1.Select(c => c.Folio).ToList()));
                            listFacturasResult.Add(listFacturas1);
                        }
                    }
                    if (listFacturasResult.Count == 1)
                    {
                        var listfactura = listFacturasResult.FirstOrDefault();
                        var factura = new Factura(listfactura.First());
                        factura.EstatusFactura = listfactura.Count > 1 ? MetodoIdentificacion.SumaFacturas : MetodoIdentificacion.ImporteAprox;
                        if (nMov.EncontradoRazonSocial)
                            factura.EstatusFactura |= MetodoIdentificacion.RazonSocial;
                        if(listaFolios.Contains(factura.FolioNumerico) || listaFolios.Contains(factura.FolioSustNumerico))
                            factura.EstatusFactura |= MetodoIdentificacion.FolioFactura;

                        var folios = listfactura.Select(c => c.Folio).ToList();
                        factura.Folio = string.Join(";", folios);
                        factura.ImporteNeto = listfactura.Sum(c => c.ImporteNeto);
                        Console.WriteLine(factura.EstatusFactura + ":" + factura.Folio);
                        return factura;
                    }
                    else if (listFacturasResult.Count > 0)
                    {
                        foreach (var listFacturas1 in listFacturasResult)
                        {
                            nMov.Comentario += "\r\nPosible Factura:" + string.Join(";", listFacturas1.Select(c => c.Folio));
                        }
                    }
                    #endregion

                    #region Por Razon Social
                    var lista4 = facturas.Where(c => !string.IsNullOrEmpty(c.RazonSocial) && string.IsNullOrEmpty(c.RFC) && Math.Abs(c.ImporteNeto - nMov.ImporteNeto) <= difAceptada)
                    .GroupBy(c => c.RazonSocial)
                    .Select(x => new
                    {
                        GroupName = x.Key,
                        GroupSize = x.Count(),
                        GroupItems = x.ToList(),
                        Sum = x.Sum(c => c.ImporteNeto),
                        ImporteList = x.AsEnumerable().Select((c, i) => new UtilitySum.
                        IElement
                        { RowIndex = i, Total = c.ImporteNeto }).ToList()
                    }).Where(c => c.Sum >= (nMov.ImporteNeto + difAceptada))
                    .ToList();

                    var lista5 = lista4.OrderBy(c => c.GroupSize).Where(c => c.GroupSize > 0 && Math.Abs(c.Sum - nMov.ImporteNeto) <= difAceptada).ToList();
                    Console.WriteLine(nMov.ImporteNeto.ToString() + "," + lista5.Count());
                    var listFacturasResult2 = new List<List<Factura>>();
                    foreach (var o in lista5)
                    {
                        var listFacturas2 = new List<Factura>();
                        var list = UtilitySum.Search(nMov.ImporteNeto, o.GroupSize, o.ImporteList);
                        // Verificamos si se solicitó la cancelación
                        CancellationToken.ThrowIfCancellationRequested();
                        if (list.Count > 0)
                        {
                            foreach (var y in list)
                            {
                                var factItem = o.GroupItems[y];
                                //factItem.EstatusFactura = MetodoIdentificacion.SumaRazonSocialFacturas;
                                listFacturas2.Add(factItem);
                            }
                            listFacturasResult2.Add(listFacturas2);
                        }
                    }
                    if (listFacturasResult2.Count == 1)
                    {
                        var listfactura = listFacturasResult2.FirstOrDefault();
                        var factura = new Factura(listfactura.First());
                        factura.EstatusFactura = listfactura.Count > 1 ? MetodoIdentificacion.SumaRazonSocialFacturas : MetodoIdentificacion.ImporteAprox;
                        if (nMov.EncontradoRazonSocial)
                            factura.EstatusFactura |= MetodoIdentificacion.RazonSocial;
                        if(listaFolios.Contains(factura.FolioNumerico) || listaFolios.Contains(factura.FolioSustNumerico))
                            factura.EstatusFactura |= MetodoIdentificacion.FolioFactura;

                        var folios = listfactura.Select(c => c.Folio).ToList();
                        factura.Folio = string.Join(";", folios);
                        factura.ImporteNeto = listfactura.Sum(c => c.ImporteNeto);
                        Console.WriteLine(factura.EstatusFactura + ":" + factura.Folio);
                        return factura;
                    }
                    //else if (listFacturasResult.Count > 0)
                    //{
                    //foreach (var listFacturas1 in listFacturasResult2)
                    //{
                    //    nMov.Comentario += "\r\nPosible Factura:" + string.Join(";", listFacturas1.Select(c => c.Folio));
                    //}
                    //}
                    #endregion

                    if (facturas.GroupBy(c => c.RFC).Count() == 1)
                    {
                        if (facturas.Count == 1)
                        {
                            var factura = facturas.First();
                            factura.EstatusFactura = MetodoIdentificacion.FacturaMesCorriente;
                            if (nMov.EncontradoRazonSocial)
                                factura.EstatusFactura = MetodoIdentificacion.RazonSocial;
                            if(listaFolios.Contains(factura.FolioNumerico) || listaFolios.Contains(factura.FolioSustNumerico))
                                factura.EstatusFactura |= MetodoIdentificacion.FolioFactura;
                            if (Math.Abs(factura.ImporteNeto - nMov.ImporteNeto) <= difAceptada)
                                factura.EstatusFactura |= MetodoIdentificacion.ImporteAprox;

                            return factura;
                        }

                        if (facturas.Count < 3)
                            nMov.Comentario += "; Posible Folio:" + string.Join(",", facturas.Select(c => c.Folio).ToArray());
                    }
                    //else
                    //{
                    //nMov.Comentario += "; Posible Paquetes:" + string.Join(",", facturas.Select(c => string.Join(":", c.Folio, c.RFC)).ToArray());
                    //}
                }


                return null;
            }

            finally { }
        }
        private string ConnectionString(string fileName, string header)
        {
            OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder();
            if (Path.GetExtension(fileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", string.Format("Excel 8.0;IMEX=1;HDR={0};", header));
            }
            else
            {
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0";
                Builder.Add("Extended Properties", string.Format("Excel 12.0;IMEX=1;HDR={0};", header));
            }

            Builder.DataSource = fileName;
            return Builder.ConnectionString;
        }
        //private void ObtieneBANCOS(string empresa, string sheet, string banco, string bancoKey)
        //{
        //    if (!listBANCOS.ContainsKey(sheet))
        //        listBANCOS.Add(sheet, new Dictionary<int, MovimientoBancario>());
        //    if (!listUsadaCFDI.ContainsKey(sheet))
        //        listUsadaCFDI.Add(sheet, new List<string>());


        //    Utils.EventLog.LogMessage("ObtieneBANCOS : " + sheet);

        //    string myexceldataquery = "select * from [" + sheet + "$]";

        //    string sexcelconnectionstring = ConnectionString(this.FileName, "NO");
        //    //series of commands to bulk copy data from the excel file into our sql table
        //    using (OleDbConnection oledbconn =
        //        new OleDbConnection(sexcelconnectionstring))
        //    {
        //        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
        //        oledbconn.Open();
        //        OleDbDataReader dataReader = oledbcmd.ExecuteReader();

        //        System.Data.DataTable dt = new System.Data.DataTable();

        //        dt.Load(dataReader);
        //        decimal count = dt.Rows.Count;
        //        for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
        //        {
        //            var row = dt.Rows[rowIndex];

        //            if (rowIndex != 0)
        //            {
        //                MovimientoBancario nMov = new MovimientoBancario() { Empresa = empresa, Banco = banco, BancoKey = bancoKey, rowIndex = rowIndex + 1, Sheet = sheet, RFC = "" };

        //                // listUsadaCFDI
        //                if (row.ItemArray.Length > FacturaPagadaColumn - 1)
        //                {

        //                    nMov.Folio = row[FacturaPagadaColumn - 1]?.ToString();
        //                    if ((bool)(nMov.Folio?.Contains("INV")))
        //                    {
        //                        nMov.Estatus = row[EstatusColumn - 1]?.ToString();
        //                        if (nMov.Estatus != "POSIBLE")
        //                            listUsadaCFDI[nMov.Sheet].AddRange(nMov.Folio.Split(';'));
        //                    }
        //                }

        //                if (row.ItemArray.Length > CompensacionColumn - 1)
        //                    nMov.Compensacion = row[CompensacionColumn - 1]?.ToString();
        //                if (nMov.Compensacion?.Length > 0)
        //                {
        //                    nMov.Estatus = "COMPENSADO";
        //                    continue;
        //                }
        //                if (nMov.Estatus?.Length > 0 && nMov.Estatus != "POSIBLE")
        //                {
        //                    goto addList;
        //                }

        //                if (TipoMovimientoColumn.ContainsKey(bancoKey))
        //                {
        //                    string cell = row[TipoMovimientoColumn[bancoKey] - 1]?.ToString();
        //                    nMov.TipoMovimiento = TipoMov(cell);
        //                    if (nMov.TipoMovimiento != TipoMovimiento.Cobranza)
        //                        continue;
        //                }
        //                if (SecuenciaColumn.ContainsKey(bancoKey))
        //                    nMov.Secuencia = row[SecuenciaColumn[bancoKey] - 1]?.ToString();

        //                if (ReferenciaColumn.ContainsKey(bancoKey))
        //                    nMov.Referencia = row[ReferenciaColumn[bancoKey] - 1]?.ToString();

        //                if (ReferenciaAmpliadaColumn.ContainsKey(bancoKey))
        //                {
        //                    nMov.ReferenciaAmpliada = row[ReferenciaAmpliadaColumn[bancoKey] - 1]?.ToString();
        //                    //if (nMov.ReferenciaAmpliada != null && nMov.ReferenciaAmpliada.Contains("+"))
        //                    //{
        //                    //    nMov.ReferenciaAmpliada = cell.Value2.ToString("0." + new string('#', 339));
        //                    //}
        //                }

        //                bool signo = false;
        //                decimal importe = 0;
        //                if (FechaColumn.ContainsKey(bancoKey))
        //                {
        //                    var cell = row[FechaColumn[bancoKey] - 1]?.ToString();
        //                    if (double.TryParse(cell, out double vaule))
        //                    {
        //                        try
        //                        {
        //                            nMov.FechaAplicacion = UtilityDates.FromExcelSerialDate(vaule);
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            ex.ToString();
        //                        }
        //                    }
        //                    else if (DateTime.TryParse(cell, out DateTime fecha))
        //                    {
        //                        nMov.FechaAplicacion = fecha;
        //                    }
        //                }
        //                if (FechaOperacionColumn.ContainsKey(bancoKey))
        //                {
        //                    var cell = row[FechaColumn[bancoKey] - 1]?.ToString();
        //                    if (double.TryParse(cell, out double vaule))
        //                    {
        //                        nMov.FechaAplicacion = UtilityDates.FromExcelSerialDate(vaule);
        //                    }
        //                    else if (DateTime.TryParse(cell, out DateTime fecha))
        //                    {
        //                        nMov.FechaAplicacion = fecha;
        //                    }
        //                }
        //                if (ConceptoColumn.ContainsKey(bancoKey))
        //                    nMov.Concepto = row[ConceptoColumn[bancoKey] - 1]?.ToString();

        //                if (SignoColumn.ContainsKey(bancoKey))
        //                    signo = row[SignoColumn[bancoKey] - 1]?.ToString() == "+";

        //                if (ImporteColumn.ContainsKey(bancoKey))
        //                {
        //                    if (decimal.TryParse(row[ImporteColumn[bancoKey] - 1]?.ToString().Replace("$", ""), out importe))
        //                        nMov.Abono = importe;
        //                    if ((SignoColumn.ContainsKey(bancoKey) && !signo) || (bancoKey == "AFIRME"))
        //                        nMov.Cargo *= -1;
        //                }
        //                if (AbonoColumn.ContainsKey(bancoKey))
        //                {
        //                    if (decimal.TryParse(row[AbonoColumn[bancoKey] - 1]?.ToString(), out importe) && importe > 0)
        //                        nMov.Abono = importe;
        //                }
        //                nMov.RFC = row[RFCColumn - 1]?.ToString();
        //                nMov.RazonSocial = row[NombreOrdenanteColumn - 1]?.ToString();
        //                //nMov.RazonSocial = dr[RazonColumn-1]?.ToString();


        //                if (nMov.FechaAplicacion.Year == 1 && string.IsNullOrEmpty(nMov.Concepto) && nMov.ImporteNeto == 0)
        //                {
        //                    break;
        //                }

        //            addList:
        //                if (listBANCOS[nMov.Sheet].ContainsKey(rowIndex))
        //                    listBANCOS[nMov.Sheet][rowIndex] = nMov;
        //                else
        //                    listBANCOS[nMov.Sheet].Add(rowIndex, nMov);
        //                Console.WriteLine(nMov.ToString());
        //            }
        //        }
        //    }

        //    if (listBANCOS.Count() == 0) throw (new Exception("No se encuentra archivo"));
        //}


        private void ObtieneOtorgantes()
        {
            var fileClientes = DirInfo.GetFiles("*generaldecliente*.xlsx", SearchOption.TopDirectoryOnly);
            if (fileClientes.Count() == 0) return;//throw (new Exception("No se encuentra archivo: Facturadeventa*.xlsx"));

            foreach (var f in fileClientes.Where(c=>!c.Name.StartsWith("~$")))
            {
                Utils.EventLog.LogMessage("Otorgantes : " + f.Name);

                string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                FileInfo file = new FileInfo(tempFileName);

                //create our connection strings
                string sexcelconnectionstring = ConnectionString(file.FullName, "YES");

                string myexceldataquery = ObtieneDataQuery(file);
                //series of commands to bulk copy data from the excel file into our sql table
                using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                {
                    OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                    oledbconn.Open();
                    OleDbDataReader dataReader = oledbcmd.ExecuteReader();

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Load(dataReader);
                    var folioAnterior = string.Empty;

                    string[] columnsToFind = { "ID interno","RFC", "Cliente SAP" , "Nombre", "Categoría", "Subsidiaria principal" };
                    // Verificar las columnas que estamos buscando
                    foreach (string columnName in columnsToFind)
                    {
                        if (!dt.Columns.Contains(columnName))
                        {
                            throw new Exception("Error en el archivo: " + f.Name + $"Columna '{columnName}' no existe");
                        }
                    }

                    var index = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        Otorgante item = new Otorgante();
                        if (double.TryParse(dr["ID interno"].ToString(), out double id))
                        {
                            item.ClienteID = id;
                            item.Empresa = dr["Subsidiaria principal"].ToString();
                            if (item.Empresa.ToLower().Contains("laboral"))
                                item.Empresa = "CL";
                            else if (item.Empresa.ToLower().Contains("crédito"))
                                item.Empresa = "CC";
                            else
                                item.Empresa = "CI";

                            if (otorganteRazonSocialList.ContainsKey(item.Empresa))
                            {
                                item.NumSAP = dr["Cliente SAP"].ToString();
                                item.RazonSocial = dr["Nombre"].ToString();
                                item.MesVencido = !dr["Categoría"].ToString().Contains("Prepago");
                                item.RFC = dr["RFC"].ToString().ToUpper();
                                otorganteRazonSocialList[item.Empresa].Add(item);
                            }
                        }

                        this.wait.ProgressValue = (index++ / dt.Rows.Count) * 100;
                    }
                }
            }
        }
        private void ObtieneFacturas()
        {
            var fileCFDI = DirInfo.GetFiles("*Facturadeventa*.xlsx", SearchOption.TopDirectoryOnly);
            if (fileCFDI.Count() == 0) return;//throw (new Exception("No se encuentra archivo: Facturadeventa*.xlsx"));
            if (fileCFDI.Count() > 1) 
                fileCFDI = DirInfo.GetFiles("*Facturadeventa*" + Empresa + "*.xlsx", SearchOption.TopDirectoryOnly);

            var listFactura = new Dictionary<string, Dictionary<string, Factura>>();
            foreach (var f in fileCFDI.Where(c=>!c.Name.StartsWith("~$")))
            {
                Utils.EventLog.LogMessage("Facturas : " + f.Name);
                string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                FileInfo file = new FileInfo(tempFileName);

                //create our connection strings
                string sexcelconnectionstring = ConnectionString(file.FullName, "YES");

                string myexceldataquery = ObtieneDataQuery(file);

                //series of commands to bulk copy data from the excel file into our sql table
                using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                {
                    OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                    oledbconn.Open();
                    OleDbDataReader dataReader = oledbcmd.ExecuteReader();
                    var folioAnterior = string.Empty;

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Load(dataReader);
                    string fechaColumnName = Empresa == "CL" ? "Date" : "Fecha";
                    List<string> columnsToFind = new List<string>() {fechaColumnName,"RFC", "Razón Social", "Total", "Método de pago 3#3", "Obervaciones Refacturación",
                        "Observaciones", "Referencia (Nota trans)", "Otorgante","Cliente SAP","Descuento","Impuesto" };

                    // Verificar las columnas que estamos buscando
                    foreach (string columnName in columnsToFind)
                    {
                        if (!dt.Columns.Contains(columnName))
                        {
                            throw new Exception("Error en el archivo: " + f.Name + $"Columna '{columnName}' no existe");
                        }
                    }


                    decimal count = dt.Rows.Count;
                    var index = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (DateTime.TryParse(dr[fechaColumnName].ToString(), out var dateTime))
                        {
                            Factura item = new Factura();
                            item.Folio = dr["Folio"].ToString().Trim();
                            item.Empresa = Empresa;

                            if (folioAnterior != item.Folio)
                            {
                                folioAnterior = item.Folio;
                            }
                            item.RFC = dr["RFC"].ToString().Trim();
                            item.RazonSocial = UtilityIO.NormalizaNombre(dr["Razón Social"].ToString());


                            if (item.Folio.StartsWith("INV"))
                            {
                                if (decimal.TryParse(dr["Total"].ToString(), out decimal total))
                                    item.ImporteNeto = total;
                                //if (decimal.TryParse(dr[11].ToString(), out decimal importe))
                                //    item.ImporteNeto = importe;

                                if (item.ImporteNeto > 0)
                                {
                                    if (!listFactura.ContainsKey(item.Empresa))
                                        listFactura.Add(item.Empresa, new Dictionary<string, Factura>());

                                    var impuesto = dr["Impuesto"].ToString();
                                    var indexImpuesto = impuesto.IndexOf("$");
                                    if (indexImpuesto > 0)
                                    {
                                        item.Moneda = impuesto.Substring(0, indexImpuesto);
                                        if (decimal.TryParse(impuesto.Substring(indexImpuesto + 1), out decimal iva))
                                            item.IVA = iva;
                                    }
                                    if (decimal.TryParse(dr["Descuento"].ToString(), out decimal desc))
                                        item.Descuento = desc;

                                    if (!listFactura[item.Empresa].ContainsKey(item.Folio))
                                    {

                                        item.Fecha = dateTime;
                                        item.NumSAP = dr["Cliente SAP"].ToString();
                                        item.NumOtorgante = dr["Otorgante"].ToString();
                                        item.Referencias = dr["Referencia (Nota trans)"].ToString();


                                        item.EstatusFactura = MetodoIdentificacion.NoIdentificado;
                                        item.MetodoPago = dr["Método de pago 3#3"].ToString();
                                        item.Observaciones = dr["Observaciones"].ToString();
                                        item.Refacturacion = dr["Obervaciones Refacturación"].ToString();

                                        if(dt.Columns.Contains("UUID CFDI v3#3"))
                                            item.FolioFiscal = dr["UUID CFDI v3#3"].ToString();

                                        List<string> status = new List<string>() { "abierta", "open" };
                                        item.Estatus = status.Contains(dr[20].ToString().ToLower()) ? 1 : 0;
                                        if (item.Estatus == 0)
                                        {
                                            item.Pago = item.ImporteNeto * -1;
                                        }

                                        if (item.Observaciones.Length == 20 && !item.Observaciones.Contains(" "))
                                        {
                                            var liquidacion = listOpenPay.FirstOrDefault(o => o.Charges.Any(charge => charge.TransactionId == item.Observaciones));
                                            item.OpenPayLiquidacion = liquidacion;
                                            if (liquidacion != null)
                                            {
                                                OpenPay openPay = liquidacion.Charges.FirstOrDefault(charge => charge.TransactionId == item.Observaciones);
                                                item.OpenPay = openPay;
                                                openPay.Invoice = item;
                                            }
                                        }


                                        listFactura[item.Empresa].Add(item.Folio, new Factura(item));
                                        //Console.WriteLine(item);
                                    }
                                    else
                                    {
                                        listFactura[item.Empresa][item.Folio].ImporteNeto += item.ImporteNeto;
                                        if (item.Estatus == 0)
                                        {
                                            item.Pago += item.ImporteNeto * -1;
                                        }
                                    }
                                }
                            }
                        }
                        this.wait.ProgressValue = (index++ / count) * 100;
                    }
                }
            }

            if (listFactura.Count() == 0) throw new Exception("No se encuntran facturas en el archivo ");

            var index1 = 0;
            foreach (var item in listFactura[Empresa].Values)
            {
                this.wait.ProgressValue = (index1++ / (decimal)listFactura[Empresa].Values.Count()) * 100;
                if (listFacturaCFDI[Empresa].ContainsKey(item.Folio))
                {
                    var factura = listFacturaCFDI[item.Empresa][item.Folio];
                    if (string.IsNullOrEmpty(factura.RFC) && item.RFC != factura.RFC)
                        factura.RFC = item.RFC;
                    if (string.IsNullOrEmpty(factura.NumSAP) && item.NumSAP != factura.NumSAP)
                        factura.NumSAP = item.NumSAP;
                    if (string.IsNullOrEmpty(factura.RazonSocial) && item.RazonSocial != factura.RazonSocial)
                        factura.RazonSocial = item.RazonSocial;
                    if (!string.IsNullOrEmpty(factura.RazonSocial))
                    {
                        var o = otorganteRazonSocialList[Empresa].FirstOrDefault(c => c.RFC == item.RFC);
                        if (o != null)
                        {
                            o.RazonSocial = item.RazonSocial;
                            if (!item.RazonSocial.Contains(factura.RazonSocial) && !otorganteRazonSocialList[Empresa].Any(c => c.RazonSocial.ToUpper() == factura.RazonSocial.ToUpper()))
                            {
                                otorganteRazonSocialList[Empresa].Add(new Otorgante() { RFC = item.RFC, RazonSocial = factura.RazonSocial });
                            }
                        }
                        else
                        {
                            o = otorganteRazonSocialList[Empresa].FirstOrDefault(c => string.IsNullOrEmpty(c.RFC) && c.NRazonSocial.ToUpper() == item.RazonSocial.ToUpper());
                            if (o != null)
                            {
                                o.RFC = item.RFC;
                            }
                            else
                            {
                                otorganteRazonSocialList[Empresa].Add(new Otorgante() { RFC = item.RFC, RazonSocial = item.RazonSocial });
                            }
                        }
                    }
                    factura.Source = 2;
                    listFacturaCFDI[item.Empresa][item.Folio] = factura;
                }
                else
                {
                    var factura = new Factura(item);
                    factura.Source = 1;
                    listFacturaCFDI[item.Empresa].Add(item.Folio, factura);

                    if (!string.IsNullOrEmpty(item.RazonSocial))
                    {
                        var o = otorganteRazonSocialList[Empresa]
                            .Where(c => !string.IsNullOrEmpty(item.RFC)
                                //&& item.RFC != "XEXX010101000" 
                                && c.RFC.ToUpper() == item.RFC.ToUpper())
                            .FirstOrDefault();
                        if (o != null)
                        {
                            var result = UtilityString.findSimilarity(o.RazonSocial, item.RazonSocial);
                            if (result >= .8 && result < 1)
                            {
                                o.RazonSocial = item.RazonSocial;
                            }
                        }
                        else
                        {
                            o = otorganteRazonSocialList[Empresa].FirstOrDefault(c => string.IsNullOrEmpty(c.RFC) && c.NRazonSocial.ToUpper() == item.RazonSocial.ToUpper());
                            if (o != null)
                            {
                                o.RFC = item.RFC;
                            }
                            else
                            {
                                otorganteRazonSocialList[Empresa].Add(new Otorgante() { RFC = item.RFC, RazonSocial = item.RazonSocial });
                            }
                        }
                    }
                }
            }

            foreach (var otorgante in otorganteRazonSocialList[Empresa])
            {
                if (!razonSocialList[Empresa].Contains(otorgante.NRazonSocial))
                    razonSocialList[Empresa].Add(otorgante.NRazonSocial);
            }
        }

        private string ObtieneDataQuery(FileInfo file, int sheet = 0)
        {
            var sheets = ObtieneSheets(file);
            return $"SELECT * FROM [{sheets[sheet]}$]";
        }
        private List<string> ObtieneSheets(FileInfo file)
        {
            List<string> sheets = new List<string>();
            using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
            {
                excel.OpenFile(file.FullName, false);

                return excel.GetSheetNames();
            }
        }

        private static System.Data.DataTable GetSchemaTable(OleDbConnection connection)
        {
            connection.Open();
            try
            {
                return connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            }
            finally
            {
                connection.Close();
            }
        }

        private void ObtieneSAP()
        {
            var fileSAP = DirInfo.GetFiles("*SAP*" + Empresa + "*.xlsx", SearchOption.TopDirectoryOnly);
            if (fileSAP.Count() == 0) throw (new Exception("No se encuentra archivo: SAP*" + Empresa + "*.xlsx"));


            foreach (var f in fileSAP.Where(c => !c.Name.StartsWith("~$")))
            {
                Utils.EventLog.LogMessage("SAP : " + f.Name);

                string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                FileInfo file = new FileInfo(tempFileName);

                //create our connection strings
                string sexcelconnectionstring = ConnectionString(file.FullName, "YES");

                string myexceldataquery = ObtieneDataQuery(file);

                //series of commands to bulk copy data from the excel file into our sql table
                using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                {
                    OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                    oledbconn.Open();
                    OleDbDataReader dataReader = oledbcmd.ExecuteReader();
                    var folioAnterior = string.Empty;

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Load(dataReader);

                    // Columnas que queremos encontrar
                    string[] columnsToFind = { "Referencia", "Nº documento", "Asignación", "Importe en moneda local", "Clase de documento", "Fecha de documento", "Texto", "Doc#compensación" };

                    // Verificar si las columnas existen en el DataTable
                    Dictionary<string, int> columnIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    // Recorrer todas las columnas del DataTable para identificar su posición
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        columnIndexes[dt.Columns[i].ColumnName] = i;
                    }
                    // Verificar las columnas que estamos buscando
                    foreach (string columnName in columnsToFind)
                    {
                        if (!dt.Columns.Contains(columnName))
                        {
                            throw new Exception("Error en el archivo: " + f.Name + $"Columna '{columnName}' no existe en el DataTable");
                        }
                    }

                    decimal count = dt.Rows.Count;
                    var index = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        Factura item = new Factura();
                        if (DateTime.TryParse(dr["Fecha de documento"].ToString(), out var dateTime))
                        {
                            item.Folio = dr["Referencia"].ToString();
                            if (!string.IsNullOrEmpty(item.Folio) && (item.Folio.StartsWith("INV") || int.TryParse(item.Folio, out int folio)))
                            {
                                var Tipo = dr["Clase de documento"].ToString();
                                var Compensado = dr["Doc#compensación"].ToString();
                                item.Comensado = Compensado;
                                item.NumDocumento = dr["Nº documento"].ToString();
                                item.Empresa = this.Empresa;

                                var texto = dr["Texto"].ToString();
                                var textoSplit = texto.Split(' ');
                                var numero = textoSplit.Count() > 0 ? textoSplit[0] : "";
                                if (int.TryParse(numero, out int x))
                                {
                                    item.NumOtorgante = numero;
                                    item.RazonSocial = texto.Replace(numero, "");
                                    if (!razonSocialList[Empresa].Contains(item.NRazonSocial))
                                    {
                                        razonSocialList[Empresa].Add(item.NRazonSocial);
                                    }
                                }

                                if (decimal.TryParse(dr["Importe en moneda local"].ToString(), out decimal total))
                                {
                                    item.TipoDocumento = Tipo;
                                    if (Tipo == "1A")
                                        item.Total += total;
                                    else if (Tipo == "SA" || Tipo == "1B")
                                        item.Pago += total;
                                    else
                                        continue;

                                    if (Math.Abs(total) > 0)
                                    {

                                        if (listFacturaCFDI.ContainsKey(item.Empresa))
                                        {
                                            if (!listFacturaCFDI[item.Empresa].ContainsKey(item.Folio) && item.TipoDocumento == "1A")
                                            {
                                                item.Fecha = dateTime;
                                                item.NumSAP = dr["Asignación"].ToString();
                                                item.EstatusFactura = MetodoIdentificacion.NoIdentificado;

                                                listFacturaCFDI[item.Empresa].Add(item.Folio, new Factura(item) { Source = 0 });
                                                //Console.WriteLine("Factura : " + item.ToString());

                                            }
                                            //else
                                            //{
                                            //    item.Fecha = dateTime;
                                            //    listCFDI[item.Empresa].Add(Guid.NewGuid().ToString(), new Factura(item));
                                            //    Console.WriteLine("Factura :" + item.NumDocumento + item.ToString());
                                            //}


                                            if (item.TipoDocumento == "1A" && (texto.Contains("SUSTITUYE INV") || texto.Contains("SUST INV")) && string.IsNullOrEmpty(item.FolioSust))
                                            {
                                                var numero2 = textoSplit.LastOrDefault();
                                                if (int.TryParse(numero2?.Replace("INVC", "").Replace("INVL", "").Replace("INVI", ""), out int x1))
                                                {
                                                    item.FolioSust = numero2;
                                                    if (!listFacturaCFDI[item.Empresa].ContainsKey(item.FolioSust))
                                                    {
                                                        listFacturaCFDI[item.Empresa].Add(item.FolioSust, new Factura(item) { Folio = item.FolioSust, FolioSust = item.FolioSust, Source = 0 });
                                                    }
                                                    listFacturaCFDI[item.Empresa][item.Folio].FolioSust = item.FolioSust;
                                                    Console.WriteLine("Factura: " + item.Folio + " SUST:" + item.FolioSust);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            this.wait.ProgressValue = (index++ / count) * 100;
                        }
                    }
                }
            }
            var listPagos = new Dictionary<string, Dictionary<string, decimal>>();
            var list = new Dictionary<string, List<string>>();
            list.Add(Empresa, new List<string>());
            listPagos.Add(Empresa, new Dictionary<string, decimal>());

            foreach (var item in listFacturaCFDI[Empresa].Select(c => c.Value).Where(c => !string.IsNullOrEmpty(c.Comensado) && (c.TipoDocumento == "SA" || c.TipoDocumento == "1B")))
            {
                Console.WriteLine("SA : " + item.Folio + " Complensado:" + item.Comensado);
                if (!listPagos[Empresa].ContainsKey(item.Comensado))
                    listPagos[Empresa].Add(item.Comensado, item.Pago);
                else
                    listPagos[Empresa][item.Comensado] += item.Pago;
            }

            foreach (var item in listPagos[Empresa].Select(c => c.Key).ToArray())
            {
                Console.WriteLine("Compensacion:" + item);
                Console.WriteLine("TOTAL PARA COMPENSAR----:" + listPagos[Empresa][item]);

                foreach (var fact in listFacturaCFDI[Empresa].Select(c => c.Value).OrderBy(c => c.Fecha).Where(c => c.Comensado == item && c.TipoDocumento == "1A" && listPagos[Empresa].ContainsKey(item)))
                {
                    try
                    {
                        Console.WriteLine("SA : " + fact.Folio + " Complensado:" + fact.Comensado);
                        var Diff = fact.Total + listPagos[Empresa][item];
                        fact.Pago = Diff <= 0 ? fact.Total * -1 : listPagos[Empresa][item];
                        listPagos[Empresa][item] += fact.Total;

                        Diff = fact.Diff;
                        listFacturaCFDI[Empresa][fact.Folio].EstatusFactura = ((Diff == 0 || Diff == difAceptada || Diff == -difAceptada) ? MetodoIdentificacion.Compensada : MetodoIdentificacion.NoIdentificado);

                        Console.WriteLine("folio:" + fact.Folio);
                        Console.WriteLine("Total:" + fact.Total);
                        Console.WriteLine("Pago:" + fact.Pago);
                        Console.WriteLine("Diff:" + Diff);
                        Console.WriteLine((listPagos[Empresa][item] > 0 ? "Faltante:" : "Sobrante:") + listPagos[Empresa][item]);
                    }
                    catch
                    {
                    }
                }
            }
            foreach (var fact in listFacturaCFDI[Empresa].Select(c => c.Value).Where(c => !string.IsNullOrEmpty(c.FolioSust) && c.Folio != c.FolioSust))
            {
                if (listFacturaCFDI[Empresa].ContainsKey(fact.FolioSust))
                {
                    var factSust = listFacturaCFDI[Empresa][fact.FolioSust];

                    //if (fact.Diff == factSust.Diff || factSust.Diff == 0)
                    //{
                    list[Empresa].Add(fact.FolioSust);
                    //}
                }
            }
            foreach (var fact in listFacturaCFDI[Empresa].Select(c => c.Value).Where(c => c.Diff <= 0 && c.TipoDocumento == "1A"))
                list[Empresa].Add(fact.Folio);

            //foreach (var fact in listCFDI[empresa].Where(c => Guid.TryParse(c.Key,out Guid result)))
            //    list[empresa].Add(fact.Key);


            foreach (string folio in list[Empresa])
                listFacturaCFDI[Empresa].Remove(folio);



            if (listFacturaCFDI.Count() == 0) throw (new Exception("No se encuentran facturas en el archivo "));

        }

        public void ObtieneReferencias()
        {
            try
            {
                var fileRef = DirInfo.GetFiles("*Referencias*" + Empresa + "*.xlsx", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (fileRef != null)
                {
                    Utils.EventLog.LogMessage("Ref " + Empresa + " : " + fileRef.Name);

                    string tempFileName = UtilityIO.GeneraArchivoTemporal(fileRef.FullName);
                    FileInfo file = new FileInfo(tempFileName);

                    //create our connection strings
                    string sexcelconnectionstring = ConnectionString(file.FullName, "NO");

                    string myexceldataquery = ObtieneDataQuery(file);

                    //series of commands to bulk copy data from the excel file into our sql table
                    using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                    {
                        OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                        oledbconn.Open();
                        OleDbDataReader dr = oledbcmd.ExecuteReader();

                        while (dr.Read())
                        {

                            string referencia = dr[1].ToString().Trim();
                            string referenciaAmpliada = dr[2].ToString().Trim();
                            string otorgante = dr[0].ToString();
                            var index = otorgante.IndexOf(' ');

                            RefBBVA item = new RefBBVA()
                            {
                                //CuentaOrdenante = dr[0].ToString(),
                                ReferenciaBBVA = string.IsNullOrEmpty(referencia) ? referenciaAmpliada : referencia,
                                Ordenante = UtilityIO.NormalizaNombre(otorgante.Substring(index >= 0 ? index : 0)),
                                NumOtorgante = otorgante.Substring(0, index > 0 ? index : otorgante.Length)
                            };
                            var o = listRefBBVA[Empresa].FirstOrDefault(c => c.ReferenciaBBVA.Equals(referencia));
                            if (o == null)
                            {
                                listRefBBVA[Empresa].Add(item);
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

        public void ObtieneImportes()
        {
            try
            {
                var fileRef = DirInfo.GetFiles("*Referencias*" + Empresa + "*.xlsx", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (fileRef == null) throw (new FileNotFoundException("No se encuentra el archivo *Referencias*" + Empresa + ".xlsx"));
                Utils.EventLog.LogMessage("ImportesRef " + Empresa + " : " + fileRef.Name);

                string tempFileName = UtilityIO.GeneraArchivoTemporal(fileRef.FullName);
                FileInfo file = new FileInfo(tempFileName);

                //create our connection strings
                string sexcelconnectionstring = ConnectionString(file.FullName, "NO");

                string myexceldataquery = ObtieneDataQuery(file, 1);

                //series of commands to bulk copy data from the excel file into our sql table
                using (OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring))
                {
                    OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                    oledbconn.Open();
                    OleDbDataReader dr = oledbcmd.ExecuteReader();

                    while (dr.Read())
                    {
                        if (decimal.TryParse(dr[0].ToString(), out decimal importe) && !listImporteProveedores.Contains(importe))
                            listImporteProveedores.Add(importe);
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            }
        }

        public void ObtieneOpenPay()
        {
            try
            {
                var files = DirInfo.GetFiles("*.csv", SearchOption.AllDirectories)
                    .Where(c => c.DirectoryName.ToLower().Contains(("openpay " + Empresa).ToLower()))
                    .ToList();

                foreach (FileInfo f in files)
                {
                    Utils.EventLog.LogMessage("OpenPay " + Empresa + " : " + f.Name);

                    string tempFileName = UtilityIO.GeneraArchivoTemporal(f.FullName);
                    FileInfo file = new FileInfo(tempFileName);

                    OpenPayLiquidacion openPayLiquidacion = new OpenPayLiquidacion()
                    {
                        FileName = file.Name,
                        Amount = 0m,
                        Charges = new List<OpenPay>()
                    };

                    using (var reader = new StreamReader(file.FullName))
                    {
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            while (csv.Read())
                            {
                                if (csv.TryGetField<decimal>(12, out decimal amount))
                                {
                                    string dateTime = csv.GetField(4);
                                    string customer_name = csv.GetField(8);
                                    string description = csv.GetField(6);
                                    string category = csv.GetField(3);
                                    string method = csv.GetField(2);
                                    string status = csv.GetField(1);
                                    string transaction_id = csv.GetField(0);
                                    csv.TryGetField<decimal>(13, out decimal fee);
                                    csv.TryGetField<decimal>(14, out decimal feeTax);
                                    string userName = csv.GetField(17);

                                    if (category == "CHARGE")
                                    {
                                        var openPay = new OpenPay()
                                        {
                                            Category = category,
                                            Method = method,
                                            Status = status,
                                            Description = description,
                                            CustomName = customer_name,
                                            Amount = amount - fee - feeTax,
                                            OriginalAmount = amount,
                                            CreateDate = dateTime,
                                            Fee = fee,
                                            FeeTax = feeTax,
                                            UserName = userName,
                                            TransactionId = transaction_id
                                        };

                                        openPayLiquidacion.Amount += openPay.Amount;
                                        openPayLiquidacion.Charges.Add(openPay);
                                    }
                                    else if (category.Contains("PAYMENT"))
                                    {
                                        if (DateTime.TryParse(dateTime, out DateTime date))
                                        {
                                            openPayLiquidacion.PayDate = date;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    listOpenPay.Add(openPayLiquidacion);
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            }
        }
        private void EditaExcel(Range excelRange, MovimientoBancario movBanco, Factura factura)
        {
            if (factura != null)
            {

                movBanco.MetodoPago = factura.MetodoPago;
                movBanco.RazonSocial = factura.RazonSocial;
                movBanco.NumOtorgante = factura.NumOtorgante;
                movBanco.NumSAP = factura.NumSAP;
                movBanco.RFC = factura.RFC;
                var diff = Math.Abs(movBanco.ImporteNeto - factura.ImporteNeto);
                if (diff != 0)
                    movBanco.Comentario += ";(Dif: " + diff + ")";

                movBanco.Estatus = factura.EstatusFactura == MetodoIdentificacion.Posible ? "POSIBLE" : "IDENTIFICADO";
                movBanco.Folio = factura.Folio;

                //if (factura.EstatusFactura == MetodoIdentificacion.NoIdentificado)
                //    factura.EstatusFactura = movBanco.Estatus;
            }

            if (!string.IsNullOrEmpty(movBanco.Comentario))
            {
                dynamic cell2 = (Range)excelRange.Cells[movBanco.rowIndex, ComentarioColumn];
                cell2.value2 = movBanco.Comentario;
            }


            dynamic cell = (Range)excelRange.Cells[movBanco.rowIndex, RazonColumn];
            if (string.IsNullOrEmpty(cell.value2) && !string.IsNullOrEmpty(movBanco.RazonSocial))
            {
                cell.value2 = movBanco.RazonSocial;
                if (string.IsNullOrEmpty(movBanco.RFC))
                {
                    var o = otorganteRazonSocialList[movBanco.Empresa].FirstOrDefault(c => UtilityIO.NormalizaNombre(c.RazonSocial).Contains(movBanco.RazonSocial));
                    movBanco.RFC = o?.RFC;
                }
            }

            cell = (Range)excelRange.Cells[movBanco.rowIndex, RFCColumn];
            if (string.IsNullOrEmpty(cell.value2) && !string.IsNullOrEmpty(movBanco.RFC))
                cell.value2 = movBanco.RFC;



            if (string.IsNullOrEmpty(movBanco.Folio))
                return;

            if (string.IsNullOrEmpty(movBanco.RazonSocial))
                return;



            cell = (Range)excelRange.Cells[movBanco.rowIndex, SAPColumn];
            cell.value2 = movBanco.NumSAP;

            cell = (Range)excelRange.Cells[movBanco.rowIndex, FechaFacturaColumn];
            cell.value2 = movBanco.FechaAplicacion;

            cell = (Range)excelRange.Cells[movBanco.rowIndex, ImporteFacturaColumn];
            cell.value2 = factura?.ImporteNeto;

            cell = (Range)excelRange.Cells[movBanco.rowIndex, NumOtorganteColumn];
            cell.value2 = movBanco.NumOtorgante;

            cell = (Range)excelRange.Cells[movBanco.rowIndex, MetodoPagoColumn];
            cell.value2 = movBanco.MetodoPago;

            cell = (Range)excelRange.Cells[movBanco.rowIndex, FacturaPagadaColumn];
            cell.value2 = movBanco.Folio;

            cell = (Range)excelRange.Cells[movBanco.rowIndex, EstatusColumn];
            cell.value2 = movBanco.Estatus;



            cell = (Range)excelRange.Cells[movBanco.rowIndex, FormaPagoColumn];
            {
                cell.value2 = "'03";
                if (movBanco.BancoKey == "BAZ")
                {
                    cell.value2 = "'03";
                    if (movBanco.SearchRefereniciaConceptoExacto("CHEQUE"))
                        cell.value2 = "'02";
                    else if (movBanco.SearchRefereniciaConceptoExacto("DEPOSITO", "BANCARIO", "EFECTIVO"))
                        cell.value2 = "'01";
                }
                else if (movBanco.BancoKey == "AFIRME")
                {
                    if (movBanco.SearchRefereniciaConceptoExacto("ABONO", "RASTREO", "REFERENCIA"))
                        cell.value2 = "'03";
                    else if (movBanco.SearchRefereniciaConceptoExacto("EFECTIVO"))
                        cell.value2 = "'01";
                    else if (movBanco.SearchRefereniciaConceptoExacto("CHEQUE"))
                        cell.value2 = "'02";
                }
                else if (movBanco.BancoKey == "BBVA")
                {
                    if (movBanco.SearchRefereniciaConceptoExacto("CHEQUE"))
                        cell.value2 = "'02";
                    else if (movBanco.SearchRefereniciaConceptoExacto("SPEI", "TERCERO", "TEF", "TRANSFER") || (movBanco.Referencia != null && movBanco.Referencia.StartsWith("REFBNT")))
                        cell.value2 = "'03";
                    else if (movBanco.SearchRefereniciaConceptoExacto("EFECTIVO"))
                        cell.value2 = "'01";
                }
            }
        }

        private void EditaExcel(Range excelRange, MovimientoBancario movBanco, List<OpenPay> listOpenPay = null)
        {
            if (listOpenPay != null)
            {
                var index = 1;
                foreach (var openPay in listOpenPay)
                {
                    var rowIndex = movBanco.rowIndex + 1 + index++;

                    dynamic cell1 = (Range)excelRange.Cells[rowIndex, opCreateDateColumn];
                    cell1.value2 = openPay.CreateDate;
                    cell1 = (Range)excelRange.Cells[rowIndex, opDescriptionColumn];
                    cell1.value2 = openPay.Description;
                    cell1 = (Range)excelRange.Cells[rowIndex, opCustomNameColumn];
                    cell1.value2 = openPay.CustomName;
                    cell1 = (Range)excelRange.Cells[rowIndex, opUserNameColumn];
                    cell1.value2 = openPay.UserName;
                    cell1 = (Range)excelRange.Cells[rowIndex, opAmountColumn];
                    cell1.value2 = openPay.Amount;
                    cell1 = (Range)excelRange.Cells[rowIndex, opFeeAuraColumn];
                    cell1.value2 = openPay.FeeAura;
                    cell1 = (Range)excelRange.Cells[rowIndex, opFeeIGHColumn];
                    cell1.value2 = openPay.FeeIGH;
                    cell1 = (Range)excelRange.Cells[rowIndex, opFeeTaxColumn];
                    cell1.value2 = openPay.FeeTax;

                    if (openPay.Invoice != null)
                    {
                        EditaExcel(excelRange, movBanco, openPay.Invoice);
                    }
                }
            }
        }

        public void ObtieneCFDI(bool repair = false)
        {
            try
            {
                var fileCFDI = DirInfo.GetFiles("*Bnc_comCFDI*.xlsx", SearchOption.AllDirectories);
                listCFDI = new Dictionary<string, Dictionary<string, CFDI>>();
                if (fileCFDI.Count() == 0) throw (new FileNotFoundException("ERROR CON ARCHVO CFDI_ No se encuentra *Bnc_comCFDI*.xls*"));

                foreach (var f in fileCFDI.Where(c => !c.Name.StartsWith("~$")))
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
                                RFCOrdenante = dr[10].ToString().Trim(),
                                CuentaOrdenante = dr[12].ToString(),
                                Cuenta = dr[20].ToString()
                            };
                            if (string.IsNullOrEmpty(item.RFCOrdenante))
                                continue;
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
                if (fileBVBA.Count() == 0) throw (new FileNotFoundException("ERROR CON ARCHVO BVBA: No se encuentra *Bnc_comNEM*.xlsx"));

                foreach (var f in fileBVBA.Where(c => !c.Name.StartsWith("~$")))
                {
                    Utils.EventLog.LogMessage("BVBA : " + f.Name);

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

                        while (dr.Read())
                        {
                            if (decimal.TryParse(dr[13].ToString(), out decimal importe))
                            {
                                BBVA item = new BBVA()
                                {
                                    Cuenta = dr[0].ToString(),
                                    CuentaOrdenante = dr[3].ToString(),
                                    CuentaClabeOrdenante = dr[4].ToString(),
                                    Otorgante = dr[6].ToString().Trim(),
                                    RFC = dr[7].ToString().Trim(),
                                    CodigoLeyenda = dr[10].ToString(),
                                    ImporteNeto = importe,
                                    Referencia = dr[15].ToString().Trim(),
                                    ReferenciaAmpliada = dr[16].ToString().Trim(),
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
                                if (item.CodigoLeyenda.Contains("W02") && listRefBBVA != null && listRefBBVA.ContainsKey(Empresa))
                                {
                                    var o = listRefBBVA[Empresa].FirstOrDefault(c => c.ReferenciaBBVA.Equals(item.Referencia));
                                    item.Otorgante = o?.Ordenante ?? item.Otorgante;

                                    var llave = item.Referencia.Replace(" ", "");
                                    if (!listBBVA[item.Cuenta].ContainsKey(llave))
                                        listBBVA[item.Cuenta].Add(llave, item);
                                }
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
        public void ObtieneSPEI()
        {
            try
            {
                var fileSPEI = DirInfo.GetFiles("*SPEI*.xlsx", SearchOption.AllDirectories);
                listSPEI = new Dictionary<string, Dictionary<string, SPEI>>();
                if (fileSPEI.Count() == 0) throw (new FileNotFoundException("ERROR CON ARCHVO SPEI: No se encuentra *SPEI*.xls*"));

                foreach (var f in fileSPEI.Where(c => !c.Name.StartsWith("~$")))
                {
                    Utils.EventLog.LogMessage("SPEI : " + f.Name);

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

                        while (dr.Read())
                        {
                            if (cuenta != string.Empty)
                            {
                                if (decimal.TryParse(dr[5].ToString(), out decimal importe))
                                {
                                    SPEI item = new SPEI()
                                    {
                                        Referencia = dr[3].ToString(),
                                        Otorgante = UtilityIO.NormalizaNombre( dr[8].ToString()),
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

                                        if (listCFDI != null && listCFDI.ContainsKey(item.Cuenta))
                                        {

                                            //if(item.Referencia.Length >= 15)
                                            //{
                                            var itemCFDI = listCFDI[item.Cuenta].Where(c => c.Key == item.CuentaOrdenante).Select(c => c.Value).FirstOrDefault();
                                            item.RFC = itemCFDI?.RFCOrdenante ?? item.RFC;

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
                if (f != null)
                {
                    Utils.EventLog.LogMessage("RefBNTC : " + f.Name);

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

                        while (dr.Read())
                        {
                            RefBBVA item = new RefBBVA()
                            {
                                //CuentaOrdenante = dr[0].ToString(),
                                ReferenciaBBVA = dr[0].ToString().Trim(),
                                Ordenante = UtilityIO.NormalizaNombre( dr[1].ToString())
                            };
                            var o = listRefBBVA[Empresa].FirstOrDefault(c => c.ReferenciaBBVA.Equals(item.ReferenciaBBVA));
                            if (o == null)
                            {
                                listRefBBVA[Empresa].Add(item);
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

        public void ObtieneRefCIE()
        {
            try
            {
                var f = DirInfo.GetFiles("*CIE*.xlsx", SearchOption.AllDirectories).FirstOrDefault();
                listRefCIE = new Dictionary<string, string>();
                if (f != null)
                {
                    Utils.EventLog.LogMessage("RefCIE: " + f.Name);

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
                            string RFC = dr[0].ToString();
                            string Ordenante = dr[1].ToString();
                            if (!listRefCIE.ContainsKey(RFC))
                                listRefCIE.Add(RFC.ToUpper(), Ordenante.ToUpper());
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
    }
}