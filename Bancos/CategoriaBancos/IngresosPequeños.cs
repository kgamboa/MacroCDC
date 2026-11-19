using Common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading;
using System.Globalization;
using CsvHelper;

namespace TesoreriaCobranza
{
    public class IngresosPequeños : BaseExcel
    {
        private PleaseWait wait;
        private CancellationToken CancellationToken;
        DirectoryInfo DirInfo;
        public List<OpenPayLiquidacion> listOpenPay { get; private set; }
        private static Dictionary<string, Dictionary<int, MovimientoBancario>> listMovBancos;

        public IngresosPequeños(DirectoryInfo dirInfo, string fileName, PleaseWait wait)
        {
            this.FileName = fileName;
            this.DirInfo = dirInfo;

            ErrorServices.singleton.Proceso = "IngresosPequeños " + Empresa;
            Utils.EventLog.LogMessage("IngresosPequeños : " + Path.GetFileNameWithoutExtension(FileName));

            listOpenPay = new List<OpenPayLiquidacion>();
            listMovBancos = new Dictionary<string, Dictionary<int, MovimientoBancario>>();

            this.wait = wait;
            this.CancellationToken = wait.CancellationToken.Token;
        }

        public new  void InicializaProceso()
        {
            base.InicializaProceso();
        }

        public new void EjecutaProceso()
        {

            ObtieneLiquidacionesOpenPay();
            CancellationToken.ThrowIfCancellationRequested();

            ObtienePagosOpenPay();
            CancellationToken.ThrowIfCancellationRequested();

            GeneraPoliza();
        }

        private void GeneraPoliza()
        {
            if (listOpenPay.Count > 0)
            {
                if (listMovBancos.Count > 0)
                {
                    List<Poliza> polizas = new List<Poliza>();
                    Poliza polizaIngreso = new Poliza();
                    polizaIngreso.Fecha = DateTime.Now;
                    polizas.Add(polizaIngreso);

                    Poliza polizaTraspaso = new Poliza();
                    polizaTraspaso.Fecha = DateTime.Now;
                    polizaTraspaso.Movimientos = new List<PolizaDetalle>()
                    {
                        new PolizaDetalle() { SocioComercial ="OPENPAY",
                         }
                    };

                    polizas.Add(polizaTraspaso);

                }
            }
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

        private void ObtienePagosOpenPay()        
        {

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
                    
                    if (listBancosOpenPay.Contains(banco))
                    {
                        excel.SetActiveSheet(item);
                        var excelRange = excel.UsedRange;

                        if (TipoMovimientoColumn.ContainsKey(bancoKey))
                        {

                            Utils.EventLog.LogMessage("Identifica OpenPay: " + item);


                            if (!listMovBancos.ContainsKey(item))
                                listMovBancos.Add(item, new Dictionary<int, MovimientoBancario>());

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
                                    if (content == TipoMovimiento.OpenPay.ToString())
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

                                        if (listMovBancos[nMov.Sheet].ContainsKey(nMov.rowIndex))
                                            listMovBancos[nMov.Sheet][nMov.rowIndex] = nMov;
                                        else
                                            listMovBancos[nMov.Sheet].Add(nMov.rowIndex, nMov);

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
        }

        public void ObtieneLiquidacionesOpenPay()
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
                        PayDate = null,
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
                                        openPay.Product = GetProductFromAmount(openPay.OriginalAmount);

                                        openPayLiquidacion.Amount += openPay.Amount;
                                        openPayLiquidacion.Charges.Add(openPay);
                                    }
                                    else if (category.Contains("PAYMENT"))
                                    {
                                        DateTime date = new DateTime (1,1,1);
                                        if (openPayLiquidacion.PayDate == null || (DateTime.TryParse(dateTime, out date) && openPayLiquidacion.PayDate < date))
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




        public string GetProductFromAmount(decimal m51Value)
        {
            if (m51Value == 84.5m || m51Value == 50 || m51Value == -84.5m || m51Value == -50)
            {
                return "CALIFICATE";
            }
            else if (m51Value == 34.5m || m51Value == -34.5m)
            {
                return "RCE";
            }
            else if (m51Value == 250 || m51Value == -250)
            {
                return "AVISAME MUCHO MAS";
            }
            else
            {
                return "REVISAR";
            }
        }
    }
}