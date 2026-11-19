using Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using PTC.Clases;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing.Drawing2D;
using System.Globalization;

namespace TesoreriaCobranza
{
    public class ProcesoPoliza : BaseExcel
    {


        private PleaseWait wait;

        public int AgrupadoPor { get; }

        private CancellationToken CancellationToken;
        public DirectoryInfo DirInfo;
        private Dictionary<string, DateTime[]> dates;

        public ProcesoPoliza(DirectoryInfo dirInfo, string fileName, PleaseWait wait, Dictionary<string, DateTime[]> dates, int agrupadoPor)
        {
            this.FileName = fileName;
            this.DirInfo = dirInfo;
            this.dates = dates;
            this.wait = wait;
            this.AgrupadoPor = agrupadoPor;
            ErrorServices.singleton.Proceso = "ProcesoPoliza";
            Utils.EventLog.LogMessage("ProcesoPoliza : " + Path.GetFileNameWithoutExtension(FileName));
            this.CancellationToken = wait.CancellationToken.Token;
        }

        public ProcesoPoliza(DirectoryInfo dirInfo, string fileName, string empresa,PleaseWait wait)
        {
            this.FileName = fileName;
            this.DirInfo = dirInfo;
            this.Empresa = empresa;

            this.dates = AgregarDiasDelMesActual();
            ErrorServices.singleton.Proceso = "ProcesoPoliza";
            Utils.EventLog.LogMessage("ProcesoPoliza : " + Path.GetFileNameWithoutExtension(FileName));

            this.wait = wait;
            this.CancellationToken = wait.CancellationToken.Token;
        }



        public new void EjecutaProceso()
        {
            Dictionary<string, Dictionary<String, Poliza>> polizas = ObtienePolizas();
            var firstDate = dates.FirstOrDefault().Value.FirstOrDefault();
            var path = this.DirInfo + @"\" + firstDate.ToString("MMMM") + firstDate.ToString("yy");
            UtilityIO.ExisteCreaDirectorio(path);

            var r = 0;
            foreach (var item in polizas)
            {

                // Verificamos si se solicitó la cancelación
                CancellationToken.ThrowIfCancellationRequested();

                var progress = (r++ / (decimal)polizas.Count) * 100;
                this.wait.ProgressValue = progress;

                if (item.Value == null || item.Value.Count == 0)
                    continue;


                var dateData = dates.FirstOrDefault();
                Utils.EventLog.LogMessage("EjecutaPoliza : " + item.Key);                
                string workbookPath = UtilityIO.RenombraArchivo(path + @"\" + item.Key +"com"+ dateData.Key + ".xlsx");

                GeneraPoliza(workbookPath, item.Value.Values.ToList());
            }

            Process.Start(path);
        }

        private void GeneraPoliza(string workbookPath, List<Poliza> polizas)
        {
            try
            {

                Excel.Application excelApp = new Excel.Application() { Visible = false };
                Excel.Workbook excelWorkbook = null;
                try
                {
                    excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
                        false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true,
                        false, 0, true, false, false);
                }
                catch
                {
                    //Create a new workbook if the existing workbook failed to open.
                    excelWorkbook = excelApp.Workbooks.Add();
                }

                // The following gets the Worksheets collection
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                string currentSheet = "Hoja1";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                //var lFecha = DateTime.Now;
                //excelWorksheet.Name = "Poliza " + lFecha.ToString("ddMMyyyy");

                int lRenglon = 1;
                if (polizas == null)
                    return;

                foreach (Poliza poliza in polizas)
                {

                    #region Header
                    //28022025	F-	SA	INGRESO FEBRERO 2025	0754	MXN		28022025
                    //'14102020	GTOS	KF	NOMINA	0754	MXN		14102020
                    // 31032025 EGRESO SA COMISIONES MARZO 2025 0754 MXN 31032025

                    //Fecha de Documento
                    excelWorksheet.Cells[lRenglon, 1].Value = "'" + poliza.Fecha.ToString("ddMMyyyy");
                    excelWorksheet.Cells[lRenglon, 1].NumberFormat = "@";

                    //Referencia
                    excelWorksheet.Cells[lRenglon, 2].Value = poliza.Facturacion;
                    excelWorksheet.Cells[lRenglon, 2].NumberFormat = "@";

                    //Tipo de Documento
                    excelWorksheet.Cells[lRenglon, 3].Value = poliza.TipoDocumento;
                    excelWorksheet.Cells[lRenglon, 3].NumberFormat = "@";

                    //Texto Cabecera
                    excelWorksheet.Cells[lRenglon, 4].Value = poliza.Descripcion;
                    excelWorksheet.Cells[lRenglon, 4].NumberFormat = "@";

                    //Sociedad
                    switch (poliza.Sociedad)
                    {
                        case "CI":
                            excelWorksheet.Cells[lRenglon, 5].Value = "'0761";  //CI
                            break;
                        case "CL":
                            excelWorksheet.Cells[lRenglon, 5].Value = "'0764";  //CL
                            break;
                        case "CC":
                            excelWorksheet.Cells[lRenglon, 5].Value = "'0754";  //CC   
                            break;
                    }
                    excelWorksheet.Cells[lRenglon, 5].NumberFormat = "@";

                    //Moneda
                    excelWorksheet.Cells[lRenglon, 6].Value = poliza.Moneda;
                    excelWorksheet.Cells[lRenglon, 6].NumberFormat = "@";

                    //Espacio
                    excelWorksheet.Cells[lRenglon, 8].Value = "'" + poliza.Fecha.ToString("ddMMyyyy");//Fecha Documento
                    excelWorksheet.Cells[lRenglon, 8].NumberFormat = "@";


                    lRenglon++;
                    #endregion


                    foreach (PolizaDetalle detalle in poliza.Movimientos)
                    {
                        #region Movimientos

                        excelWorksheet.Cells[lRenglon, 2].Value = (int)detalle.TipoMovimiento;
                        excelWorksheet.Cells[lRenglon, 2].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 3].Value = "M";  //Ctas de Mayor
                        excelWorksheet.Cells[lRenglon, 3].NumberFormat = "@";
                        excelWorksheet.Cells[lRenglon, 4].Value = detalle.CentroCostos;
                        excelWorksheet.Cells[lRenglon, 4].NumberFormat = "@";
                        excelWorksheet.Cells[lRenglon, 5].Value = detalle.ImporteNeto.ToString("0.00");
                        excelWorksheet.Cells[lRenglon, 5].NumberFormat = "0.00";

                        excelWorksheet.Cells[lRenglon, 6].Value = detalle.CodigoImpuesto;
                        excelWorksheet.Cells[lRenglon, 6].NumberFormat = "@";

                        if (detalle.Impuesto > 0)
                        {
                            excelWorksheet.Cells[lRenglon, 7].Value = detalle.Impuesto.ToString("0.00");
                            excelWorksheet.Cells[lRenglon, 7].NumberFormat = "0.00";
                        }

                        excelWorksheet.Cells[lRenglon, 8].Value = detalle.SocioComercial;
                        excelWorksheet.Cells[lRenglon, 8].NumberFormat = "@";


                        excelWorksheet.Cells[lRenglon, 9].Value = poliza.Descripcion;
                        excelWorksheet.Cells[lRenglon, 9].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 10].Value = detalle.AreaFuncional;
                        excelWorksheet.Cells[lRenglon, 10].NumberFormat = "@";
                        lRenglon++;

                        #endregion
                    }

                    #region bancos
                    excelWorksheet.Cells[lRenglon, 2].Value = poliza.Ingreso ? "2" : "1";
                    excelWorksheet.Cells[lRenglon, 2].NumberFormat = "@";
                    excelWorksheet.Cells[lRenglon, 3].Value = poliza.CtasMayor;  //Ctas de Mayor
                    excelWorksheet.Cells[lRenglon, 3].NumberFormat = "@";
                    excelWorksheet.Cells[lRenglon, 4].Value = poliza.Catalog;
                    excelWorksheet.Cells[lRenglon, 4].NumberFormat = "@";
                    excelWorksheet.Cells[lRenglon, 5].Value = (poliza.Impuesto + poliza.ImporteNeto).ToString("0.00");
                    excelWorksheet.Cells[lRenglon, 5].NumberFormat = "0.00";

                    excelWorksheet.Cells[lRenglon, 8].Value = poliza.Referencia ?? poliza.Facturacion;
                    excelWorksheet.Cells[lRenglon, 8].NumberFormat = "@";
                    excelWorksheet.Cells[lRenglon, 9].Value = poliza.Descripcion; //Referencia
                    excelWorksheet.Cells[lRenglon, 9].NumberFormat = "@";
                    lRenglon++;
                    #endregion

                }

                object misValue = System.Reflection.Missing.Value;
                excelWorkbook.SaveAs(workbookPath, misValue, misValue, misValue, misValue, false, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);


                if (excelWorkbook != null)
                    excelWorkbook.Close(true);
                if (excelApp != null)
                {
                    excelApp.Quit();

                    if (excelWorksheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorksheet);
                    if (excelWorkbook != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                    if (excelApp != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelWorksheet = null;
                    excelWorkbook = null;
                    excelApp = null;
                    GC.Collect();
                }


            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
        }

        private Dictionary<string, Dictionary<string, Poliza>> ObtienePolizas()
        {
            var bancos = new BankAccountList();

            var result = new Dictionary<string, Dictionary<string, Poliza>>();
            var resultKey = string.Empty;
            var dateData = dates.FirstOrDefault();
            var firstDate = dateData.Value.FirstOrDefault();
            var lastDate = dateData.Value.LastOrDefault();

            Utils.EventLog.LogMessage("Obtiene Información : " + firstDate.ToString("ddMMMMyy") + " - " + lastDate.ToString("ddMMMMyy"));

            var bancoKey = string.Empty;
            using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
            {
                excel.OpenFile(FileName);

                var sheets = excel.GetSheetNames();
                Empresa = "CC";
                foreach (var sheet in sheets)
                {
                    excel.SetActiveSheet(sheet);
                    Microsoft.Office.Interop.Excel.Range range = excel.UsedRange;



                    List<char> separator = new List<char>() { ' ', '_' };
                    var split = sheet.Split(separator.ToArray());
                    if (split.Length > 1)
                    {
                        if ((split.Last() == "CC" || split.Last() == "CI" || split.Last() == "CL"))
                            Empresa = split.Last();
                    }
                    var banco = split[0];
                    if (banco == "SPEI" || banco == "Flujo" || banco == "Cuadre" || banco == "Semanal")
                    {
                        continue;
                    }
                    else
                    {
                        bancoKey = FechaColumn.Keys.Where(t => banco.ToUpper().Contains(t)).FirstOrDefault();
                        if (bancoKey == null)
                            continue;
                        banco = sheet.Replace("_", "");
                    }
                    if (banco.Contains("INTERCAM"))
                        continue;


                    if (TipoMovimientoColumn.ContainsKey(bancoKey))
                    {
                        Utils.EventLog.LogMessage("Obteniendo Información: " + sheet);

                        resultKey = Empresa + "_" + (AgrupadoPor == 1 ? "":  sheet);
                        int count = 0;
                        if(!result.ContainsKey(resultKey))
                            result.Add(resultKey, new Dictionary<string, Poliza>());
                        var bankAccounts = bancos.BankAccounts.Where(c => c.Sheet == sheet).ToList();
                        if (bankAccounts.Count == 0)
                        {
                            Utils.EventLog.LogMessage("No se encontraron cuentas bancarias para la hoja: " + sheet);
                            continue;
                        }

                        for (int r = 1; r <= range.Rows.Count; r++)
                        {
                            // Verificamos si se solicitó la cancelación
                            CancellationToken.ThrowIfCancellationRequested();

                            var progress = (r / (decimal)range.Rows.Count) * 100;
                            if (bancoKey.StartsWith("BBVA"))
                                this.wait.ProgressValue = 100 - progress;
                            else
                                this.wait.ProgressValue = progress;

                            dynamic tipoMov = range.Cells[r, TipoMovimientoColumn[bancoKey]];
                            dynamic content = TipoMov(tipoMov.Value);


                            // Corta solo tras 100 filas vacías seguidas (fin de datos).
                            // Las filas sin clasificar en medio de la hoja se saltan, no detienen el recorrido.
                            if (content == null || string.IsNullOrWhiteSpace(content.ToString()))
                            {
                                count++;
                                if (count >= 100)
                                    break;
                                continue;
                            }
                            count = 0;
                            if (content == TipoMovimiento.SinReconocer)
                                continue;

                            if (FechaColumn.ContainsKey(bancoKey))
                            {
                                string dateString = range.Cells[r, FechaColumn[bancoKey]].value?.ToString();
                                if (dateString!=null && dateString.Contains("'"))
                                {
                                    string formatInput = "ddMMyyyy";
                                    if (DateTime.TryParseExact(dateString.Replace("'",""), formatInput, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                                    {
                                        dateString = parsedDate.ToString("dd/MM/yyyy");
                                    }
                                }

                                if (DateTime.TryParse(dateString, out DateTime date))
                                {

                                    if (firstDate == null || lastDate == null)
                                        break;
                                    else if (date < firstDate || date > lastDate)
                                        continue;

                                    PolizaDetalle detalle = new PolizaDetalle();
                                    var item = new Poliza();

                                    var key = string.Concat(banco, Empresa,content, date.Date.Ticks);
                                    if (content == TipoMovimiento.ISRRetenido)
                                    {
                                        item.Facturacion = "ISR RETENIDO";
                                        item.Ingreso = true;
                                        item.Descripcion = item.Facturacion;
                                    }
                                    else if (content == TipoMovimiento.Comision || content == TipoMovimiento.IVAComision)
                                    {
                                        item.Facturacion = "COMISIONES";
                                        item.Ingreso = true;
                                        key = string.Concat(banco, Empresa,TipoMovimiento.Comision, date.Date.Ticks);
                                        item.Descripcion = "COMISIONES "+bancoKey;
                                    }
                                    else if (content == TipoMovimiento.Intereses)
                                    {
                                        item.Facturacion = "INTERESES";
                                        item.Ingreso = false;
                                        item.Descripcion = item.Facturacion;
                                    }
                                    else if( content  == TipoMovimiento.Compensacion)
                                    {
                                        item.Facturacion = "INGRESO";
                                        item.Ingreso = false;
                                        item.Descripcion = "COMPENSACION POR RETRASO";
                                        item.Referencia = "";
                                    }
                                    else
                                    {
                                        continue;
                                    }

                                    var bankAccount = bankAccounts.Where(c => (c.Type == "Egreso" && item.Ingreso) || (c.Type == "Ingreso" && !item.Ingreso))
                                             .FirstOrDefault();
                                    if (bankAccount == null)
                                    {
                                        Utils.EventLog.LogMessage("No hay cuenta " + (item.Ingreso ? "Egreso" : "Ingreso") + " para la hoja: " + bancoKey + " (" + content + ")");
                                        continue;
                                    }
                                    bankAccount.Prefijo = "3340080";
                                    if (!result[resultKey].ContainsKey(key))
                                    {
                                        item.Movimientos = new List<PolizaDetalle>();
                                        item.CtasMayor = "M";
                                        item.TipoDocumento = bankAccount.GetTypeDoc(content);
                                        item.Fecha = date;
                                        item.Sociedad = Empresa;
                                        result[resultKey].Add(key, item);
                                        item.Movimientos.Add(detalle);
                                        item.Catalog = bankAccount?.Catalog;
                                        item.Moneda = bankAccount?.Currency;
                                    }
                                    else
                                    {
                                        item = result[resultKey][key];
                                        detalle = item.Movimientos.First();
                                    }


                                    if (content == TipoMovimiento.IVAComision)
                                    {
                                        if (CargoColumn.ContainsKey(bancoKey))
                                        {
                                            decimal.TryParse(range.Cells[r, CargoColumn[bancoKey]].value2?.ToString(),out decimal cargo);
                                            decimal.TryParse(range.Cells[r, AbonoColumn[bancoKey]].value2?.ToString(), out decimal abono);
                                            detalle.Impuesto += Math.Abs(abono) + Math.Abs(cargo);

                                        }
                                        else if (ImporteColumn.ContainsKey(bancoKey))
                                        {
                                            decimal.TryParse(range.Cells[r, ImporteColumn[bancoKey]].value2?.ToString(), out decimal IVAComision);
                                            detalle.Impuesto += Math.Abs(IVAComision);
                                        }
                                    }
                                    else
                                    {
                                        if (CargoColumn.ContainsKey(bancoKey))
                                        {
                                            decimal.TryParse(range.Cells[r, CargoColumn[bancoKey]].value2?.ToString(),out decimal cargo);
                                            decimal.TryParse(range.Cells[r, AbonoColumn[bancoKey]].value2?.ToString(),out decimal abono);
                                            detalle.ImporteNeto +=  Math.Abs(cargo)+ Math.Abs(abono);
                                        }
                                        if (ImporteColumn.ContainsKey(bancoKey))
                                        {
                                            decimal.TryParse(range.Cells[r, ImporteColumn[bancoKey]].value2?.ToString(), out decimal importe);
                                            detalle.ImporteNeto += Math.Abs(importe);
                                        }
                                    }

                                    detalle.CodigoImpuesto = (content == TipoMovimiento.Comision || content == TipoMovimiento.IVAComision) ? "AI" : "";
                                    detalle.CentroCostos = bankAccount.GetAccount(content, item.Ingreso);
                                    detalle.SocioComercial = content == TipoMovimiento.Compensacion ? "": item.Facturacion;
                                    detalle.AreaFuncional = content != TipoMovimiento.ISRRetenido ? bankAccount.CeCo :"";
                                    detalle.TipoMovimiento = (content == TipoMovimiento.Intereses || content == TipoMovimiento.Compensacion) ? TipoDetalle.Cargo : TipoDetalle.Abono;
                                    
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        public Dictionary<string, DateTime[]> AgregarDiasDelMesActual()
        {
            var mesActual = DateTime.Now.Month;
            var anioActual = DateTime.Now.Year;
            var diasDelMes = DateTime.DaysInMonth(anioActual, mesActual);

            var fechasDelMes = new DateTime[diasDelMes];
            for (int dia = 1; dia <= diasDelMes; dia++)
            {
                fechasDelMes[dia - 1] = new DateTime(anioActual, mesActual, dia);
            }

            if (dates == null)
            {
                dates = new Dictionary<string, DateTime[]>();
            }

            var claveMes = $"{DateTime.Now.ToString("MMMM")} {anioActual}";
            dates[claveMes] = fechasDelMes;
            return dates;
        }
    }
}