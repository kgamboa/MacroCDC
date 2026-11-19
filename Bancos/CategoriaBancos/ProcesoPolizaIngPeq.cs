using Common;
using Microsoft.Office.Interop.Excel;
using PTC.Clases;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Excel = Microsoft.Office.Interop.Excel;

namespace TesoreriaCobranza
{
    public class ProcesoPolizaIngPeq : BaseExcel
    {


        private PleaseWait wait;

        public int AgrupadoPor { get; }

        private CancellationToken CancellationToken;
        public DirectoryInfo DirInfo;
        private Dictionary<string, DateTime[]> dates;

        public ProcesoPolizaIngPeq(DirectoryInfo dirInfo, string fileName, PleaseWait wait, Dictionary<string, DateTime[]> dates, int agrupadoPor)
        {
            this.FileName = fileName;
            this.DirInfo = dirInfo;
            this.dates = dates;
            this.wait = wait;
            this.AgrupadoPor = agrupadoPor;
            ErrorServices.singleton.Proceso = "ProcesoPolizaIngPeq";
            Utils.EventLog.LogMessage("ProcesoPolizaIngPeq : " + Path.GetFileNameWithoutExtension(FileName));
            this.CancellationToken = wait.CancellationToken.Token;
        }

        public new void EjecutaProceso()
        {
            // La expresión regular busca:
            // ([A-Za-z]{3})   -> Tres letras (mayúsculas o minúsculas) para el mes. Es el Grupo 1.
            // _              -> Un guion bajo.
            // (\d{4})        -> Cuatro dígitos para el año. Es el Grupo 2.
            string patron = @"([A-Za-z]{3})_(\d{4})";
            var mesaño = string.Empty;
            Match match = Regex.Match(this.FileName, patron);
            DateTime lastDayOfMonth;
            if (match.Success)
            {
                string mesAbreviado = match.Groups[1].Value; // Captura el primer grupo (el mes)
                string año = match.Groups[2].Value;          // Captura el segundo grupo (el año)mesaño
                mesaño = mesAbreviado + año;

               var Month= ObtenerPrimerNumeroDeFileName(this.FileName);

                // 2. Get the number of days in that month
                int daysInMonth = DateTime.DaysInMonth(int.Parse(año), int.Parse(Month));
                // 3. Construct a new DateTime object for the last day
                 lastDayOfMonth = new DateTime(int.Parse(año), int.Parse(Month), daysInMonth);
            }
            else
            {
                DateTime firstDate = dates.FirstOrDefault().Value.FirstOrDefault();
                mesaño= firstDate.ToString("MMMM") + firstDate.ToString("yy");

                // 2. Get the number of days in that month
                int daysInMonth = DateTime.DaysInMonth(firstDate.Year, firstDate.Month);
                // 3. Construct a new DateTime object for the last day
                lastDayOfMonth = new DateTime(firstDate.Year, firstDate.Month, daysInMonth);
            }
            



            var path = this.DirInfo + @"\"+ mesaño;
            Dictionary<string, Dictionary<String, Poliza>> polizas = ObtienePolizas(lastDayOfMonth);

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
                string workbookPath = UtilityIO.RenombraArchivo(path + @"\" + item.Key +"IngPeq"+mesaño+ ".xlsx");

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

                        excelWorksheet.Cells[lRenglon, 8].Value = poliza.Referencia;
                        excelWorksheet.Cells[lRenglon, 8].NumberFormat = "@";


                        excelWorksheet.Cells[lRenglon, 9].Value = detalle.Descripcion?.ToUpper();
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
                    if (poliza.Facturacion == "COM AFIRME TPV")
                    {

                        excelWorksheet.Cells[lRenglon, 4].Value = "7120000001";
                        excelWorksheet.Cells[lRenglon, 4].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 5].Value = (poliza.ImporteNeto).ToString("0.00");
                        excelWorksheet.Cells[lRenglon, 5].NumberFormat = "0.00";

                        excelWorksheet.Cells[lRenglon, 6].Value = "AI";
                        excelWorksheet.Cells[lRenglon, 6].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 7].Value = (poliza.Impuesto).ToString("0.00");
                        excelWorksheet.Cells[lRenglon, 7].NumberFormat = "0.00";

                        excelWorksheet.Cells[lRenglon, 8].Value = poliza.Referencia;
                        excelWorksheet.Cells[lRenglon, 8].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 9].Value = "AJUSTE COMISIONES TPV";
                        excelWorksheet.Cells[lRenglon, 9].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 10].NumberFormat = "@";
                        //Sociedad
                        switch (poliza.Sociedad)
                        {
                            case "CI":
                                excelWorksheet.Cells[lRenglon, 10].Value = "3359810761";  //CI
                                break;
                            case "CL":
                                excelWorksheet.Cells[lRenglon, 10].Value = "3359810764";  //CL
                                break;
                            case "CC":
                                excelWorksheet.Cells[lRenglon, 10].Value = "3359810754";  //CC   
                                break;
                        }
                    }
                    else
                    {

                        excelWorksheet.Cells[lRenglon, 5].Value = (poliza.Impuesto + poliza.ImporteNeto).ToString("0.00");
                        excelWorksheet.Cells[lRenglon, 5].NumberFormat = "0.00";
                        excelWorksheet.Cells[lRenglon, 8].Value = poliza.Referencia;
                        excelWorksheet.Cells[lRenglon, 8].NumberFormat = "@";

                        excelWorksheet.Cells[lRenglon, 9].Value = poliza.DescripcionReferencia; //Referencia
                        excelWorksheet.Cells[lRenglon, 9].NumberFormat = "@";
                    }

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

        private Dictionary<string, Dictionary<string, Poliza>> ObtienePolizas(DateTime date)
        {
            var bancos = new BankAccountList();

            var result = new Dictionary<string, Dictionary<string, Poliza>>();
            var resultKey = string.Empty;


            var bancoKey = string.Empty;
            using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
            {
                excel.OpenFile(FileName);
                var sheet = "CF";  //Antes Ing Peg
                excel.SetActiveSheet(sheet);
                Microsoft.Office.Interop.Excel.Range range = excel.UsedRange;

                if (range == null || range.Rows.Count < 2 || range.Columns.Count < 4)
                {
                    Utils.EventLog.LogError(ErrorType.ERROR_Bancos, "No se encontraron datos en el archivo de Excel: " + FileName);
                    return result;
                }

                resultKey = Empresa + "_" + sheet;
                if (!result.ContainsKey(resultKey))
                    result.Add(resultKey, new Dictionary<string, Poliza>());
                var lastBanco = string.Empty;
                var lastKey = string.Empty;
                for (int r = 5; r <= range.Rows.Count; r++)
                {
                    string banco = range.Cells[r, 1].value2?.ToString();
                    string producto = range.Cells[r, 2].value2?.ToString();
                    if (banco != null || (banco == null && lastBanco != null && producto != null))
                    {
                        //if (producto == "BBVA 0798 OpenPay adicional")
                        //    producto = "BBVA 6056";
                        if (producto == null)
                            continue;
                        if (banco == "Para Facturación" || producto == "Para Facturación")
                            break;

                        producto = producto.UpperRemoveAccents();
                        if (banco == null)
                            banco = lastBanco;

                        if (banco != null)
                        {
                            lastBanco = banco.ToUpper();
                            banco = banco == "BAZ 142" ? "BAZ_42" :
                            banco.ToUpper() == "SANTANDER 6014" ? "SANTANDER" :
                            banco.ToUpper() == "PAYPAL STD" ? "BBVA_6056" :
                            banco.ToUpper().StartsWith("AFIRME 676") ? "AFIRME_CC" : banco.Replace(" ", "_");
                        }


                        var key = string.Concat(lastBanco,
                                producto.Contains("VENTANILLA") ? "V" :
                                producto.Contains("OPENPAY") ? "O" :
                                producto.Contains("TPV") ? "T" :
                                producto.Contains("ADICIONAL") ? "A" :
                                producto.Contains("PAYPAL") ? "P" : "K", DateTime.Now.Date.Ticks);

                        if (lastKey != key && !key.Contains("K"))
                            lastKey = key;
                        if (key.Contains("K"))
                            key = lastKey;

                        var bankAccounts = bancos.BankAccounts.Where(c => c.Sheet == banco).ToList();
                        Empresa = bankAccounts.FirstOrDefault()?.Company;
                        resultKey = Empresa + "_" + (AgrupadoPor == 1 ? "" : banco);
                        if (!result.ContainsKey(resultKey))
                            result.Add(resultKey, new Dictionary<string, Poliza>());

                        var tipoMovimeinto = producto.Contains("RCE") ? TipoMovimiento.RCE :
                            producto.Contains("CALIFICATE") ? TipoMovimiento.CALIFICATE :
                            producto.Contains("AVISAME MUCHO MAS") ? TipoMovimiento.AvisameMuchoMas :
                            producto.Contains("AVISAME") ? TipoMovimiento.AVISAME :
                            TipoMovimiento.Cobranza;

                        if (decimal.TryParse(range.Cells[r, 5].value2?.ToString(), out decimal Total))
                        {
                            if (lastBanco.Contains(producto) 
                                || producto.Contains(lastBanco) 
                                || (producto.StartsWith("AFIRME") && lastBanco.StartsWith("AFIRME"))
                                //|| (producto.Contains("BBVA 0798") && lastBanco.StartsWith("BBVA 6056"))
                                || (producto.Contains("BBVA 6056") && lastBanco.StartsWith("PAYPAL STD")))
                            {
                                foreach (var resulValue in result.Values)
                                {
                                    if (resulValue.ContainsKey(key))
                                    {
                                        var item = resulValue[key];
                                        switch (producto)
                                        {
                                            case "AFIRME TPV":
                                                item.Facturacion = "AFIRME TPV";
                                                foreach (var movItem in item.Movimientos)
                                                    movItem.Descripcion += " TPV";
                                                item.Referencia = item.Facturacion;
                                                break;
                                            case "BAZ 450 OPENPAY":
                                                item.Facturacion = "OPENPAY CF APP";
                                                item.Descripcion = "COBRANZA OPENPAY CF APP";
                                                item.TipoDocumento = "LP";
                                                item.CtasMayor = "D";
                                                item.DescripcionReferencia = "VENTAS OPENPAY CF APP";
                                                item.Referencia = item.Facturacion;
                                                item.Catalog = "2864811";
                                                foreach (var movItem in item.Movimientos)
                                                    movItem.Descripcion += " " + item.Facturacion;
                                                break;
                                            case "BBVA 0798 OPENPAY":
                                                item.Descripcion = "COBRANZA OPENPAY CF WEB";
                                                item.Facturacion = "OP CF WEB";
                                                item.TipoDocumento = "LP";
                                                item.CtasMayor = "D";
                                                item.DescripcionReferencia = "VENTAS OPENPAY CF WEB";
                                                item.Referencia = item.Facturacion;
                                                item.Catalog = "2864811";
                                                foreach (var movItem in item.Movimientos)
                                                    movItem.Descripcion += " " + item.Facturacion;
                                                break;
                                            case "BBVA 6056 PAYPAL":
                                            case "PAYPAL STD":
                                                item.Descripcion = "COBRANZA PAYPAL CF";
                                                item.Facturacion = "CF PAYPAL";
                                                item.TipoDocumento = "LP";
                                                item.CtasMayor = "D";
                                                item.DescripcionReferencia = "VENTAS PAYPAL CF";
                                                item.Referencia = item.Facturacion;
                                                item.Catalog = "2006527";

                                                foreach (var movItem in item.Movimientos)
                                                    movItem.Descripcion += " PAYPAL";
                                                break;
                                            case "BBVA 0798 OpenPay adicional":
                                                item.Descripcion = "COBRANZA";
                                                item.Facturacion = "CF PAYPAL";
                                                item.TipoDocumento = "LP";
                                                item.CtasMayor = "D";
                                                item.DescripcionReferencia = "VENTAS PAYPAL";
                                                item.Referencia = "PAYPAL";
                                                item.Catalog = "2006527";

                                                foreach (var movItem in item.Movimientos)
                                                    movItem.Descripcion = "CF PAYPAL";
                                                break;
                                            default:
                                                break;
                                        }
                                    }

                                    if (resulValue.ContainsKey(key + "com"))
                                    {
                                        var itemComision = resulValue[key + "com"];

                                        switch (producto)
                                        {
                                            case "AFIRME TPV":
                                                itemComision.Facturacion = "COM AFIRME TPV";
                                                itemComision.Descripcion = "COMISIONES";

                                                foreach (var movItem in itemComision.Movimientos)
                                                    movItem.Descripcion += " TPV";
                                                itemComision.Referencia = itemComision.Facturacion;
                                                itemComision.Catalog = "2864811";
                                                break;
                                            case "BAZ 450 OPENPAY":
                                                itemComision.Facturacion = "COM OPENPAY CF APP";
                                                itemComision.Descripcion = "COM OP CF APP";
                                                itemComision.TipoDocumento = "LR";
                                                itemComision.DescripcionReferencia = "COM VENTAS OP CF APP";
                                                itemComision.CtasMayor = "D";
                                                itemComision.Referencia = itemComision.Facturacion;

                                                itemComision.Catalog = "2864811";
                                                foreach (var movItem in itemComision.Movimientos)
                                                    movItem.Descripcion += " OP CF APP";

                                                break;
                                            case "BBVA 0798 OPENPAY":
                                                itemComision.Descripcion = "COM OPENPAY CF WEB";
                                                itemComision.Facturacion = "COM OP CF WEB";
                                                itemComision.TipoDocumento = "LR";
                                                itemComision.DescripcionReferencia = "COM OPENPAY CF WEB";
                                                itemComision.CtasMayor = "D";
                                                itemComision.Referencia = itemComision.Facturacion;

                                                itemComision.Catalog = "2864811";
                                                foreach (var movItem in itemComision.Movimientos)
                                                    movItem.Descripcion += " OP CF WEB";

                                                break;
                                            case "PAYPAL STD":
                                            case "BBVA 6056 PAYPAL":
                                                itemComision.Descripcion = "COM VENTA PAYPAL CF";
                                                itemComision.Facturacion = "COM CF PAYPAL";
                                                itemComision.TipoDocumento = "LR";
                                                itemComision.CtasMayor = "D";
                                                itemComision.Referencia = itemComision.Facturacion;
                                                itemComision.DescripcionReferencia = "COM VENTA PAYPAL CF";
                                                //itemComision.DescripcionReferencia = "COM VENTAS AL PUBLICO PAYPAL";
                                                //itemComision.Referencia = itemComision.Facturacion;
                                                itemComision.Catalog = "2006527";
                                                foreach (var movItem in itemComision.Movimientos)
                                                    movItem.Descripcion += " PAYPAL";
                                                break;
                                            case "BBVA 0798 OpenPay adicional":
                                                itemComision.Descripcion = "COSTO PAYPAL";
                                                itemComision.Facturacion = "COM PAYPAL";
                                                itemComision.TipoDocumento = "LR";
                                                itemComision.CtasMayor = "D";
                                                itemComision.DescripcionReferencia = "COM VENTAS PAYPAL";
                                                itemComision.Referencia = itemComision.Facturacion;
                                                itemComision.Catalog = "2006527";
                                                foreach (var movItem in itemComision.Movimientos)
                                                    movItem.Descripcion += " PAYPAL";
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }

                                continue;
                            }
                        }

                        if (decimal.TryParse(range.Cells[r, 8].value2?.ToString(), out decimal importeNeto))
                        {
                            var item = new Poliza();
                            item.Ingreso = false;
                            PolizaDetalle detalle = new PolizaDetalle();
                            detalle.ImporteNeto += Math.Abs(importeNeto);
                            if (decimal.TryParse(range.Cells[r, 9].value2?.ToString(), out decimal impuesto))
                            {
                                detalle.Impuesto += Math.Abs(impuesto);
                            }
                            detalle.TipoMovimiento = TipoDetalle.Cargo;
                            detalle.CodigoImpuesto = "AA";
                            var bankAccount = bankAccounts.Where(c => (c.Type == "Egreso" && item.Ingreso) || (c.Type == "Ingreso" && !item.Ingreso)).FirstOrDefault();
                            if (bankAccount == null)
                            {
                                Utils.EventLog.LogMessage("No se registró movimiento en la poliza: " + key + banco + producto + tipoMovimeinto);
                                continue;
                            }
                            bankAccount.Prefijo = "3359810";

                            detalle.CentroCostos = bankAccount?.GetAccount(tipoMovimeinto, item.Ingreso);
                            detalle.SocioComercial = item.Facturacion;
                            detalle.AreaFuncional = bankAccount?.CeCo;
                            detalle.Descripcion = "VENTA " + tipoMovimeinto;

                            if (importeNeto > 0 || impuesto > 0)
                            {

                                if (!result[resultKey].ContainsKey(key))
                                {
                                    item.Movimientos = new List<PolizaDetalle>() { detalle };
                                    item.TipoDocumento = "59";
                                    item.Facturacion = "CF";
                                    item.CtasMayor = "M";
                                    item.Sociedad = bankAccount.Company;
                                    item.Fecha = date;

                                    result[resultKey].Add(key, item);
                                    item.Catalog = bankAccount?.Catalog;
                                    item.Moneda = bankAccount?.Currency;
                                    item.Descripcion = "COBRANZA CF";
                                    item.DescripcionReferencia = "VENTAS CF";
                                }
                                else
                                {
                                    item = result[resultKey][key];
                                    if (banco != null)
                                        item.Movimientos.Add(detalle);
                                }
                            }
                        }


                        if (decimal.TryParse(range.Cells[r, 13].value2?.ToString(), out decimal importeNetoCOM))
                        {
                            var keyCom = key + "com";

                            var itemComision = new Poliza();
                            itemComision.Ingreso = true;
                            PolizaDetalle detalleComision = new PolizaDetalle();
                            detalleComision.ImporteNeto += Math.Abs(importeNetoCOM);
                            if (decimal.TryParse(range.Cells[r, 14].value2?.ToString(), out decimal impuestoCOM))
                                detalleComision.Impuesto += Math.Abs(impuestoCOM);
                            detalleComision.TipoMovimiento = TipoDetalle.Abono;
                            detalleComision.CodigoImpuesto = "AI";
                            var bankAccount = bankAccounts.Where(c => (c.Type == "Egreso" && itemComision.Ingreso) || (c.Type == "Ingreso" && !itemComision.Ingreso)).FirstOrDefault();
                            if (bankAccount != null)
                            {

                                bankAccount.Prefijo = "3359810";
                                detalleComision.CentroCostos = bankAccount?.GetAccount(tipoMovimeinto, itemComision.Ingreso);
                                detalleComision.SocioComercial = itemComision.Facturacion;
                                detalleComision.AreaFuncional = bankAccount?.CeCo;
                                detalleComision.Descripcion = "COM VENTA " + tipoMovimeinto;
                            }


                            if (!result[resultKey].ContainsKey(keyCom))
                            {
                                itemComision.Movimientos = new List<PolizaDetalle>() { detalleComision };
                                itemComision.TipoDocumento = "D2";
                                itemComision.Facturacion = "CF";
                                itemComision.CtasMayor = "M";
                                itemComision.Sociedad = bankAccount?.Company;
                                itemComision.Fecha = date;

                                result[resultKey].Add(keyCom, itemComision);
                                itemComision.Catalog = bankAccount?.Catalog;
                                itemComision.Moneda = bankAccount?.Currency;
                                itemComision.Descripcion = "COBRANZA CF";
                            }
                            else
                            {
                                itemComision = result[resultKey][keyCom];
                                itemComision.Movimientos.Add(detalleComision);
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
        // Método para obtener el primer número del nombre del archivo (this.FileName)
        private string ObtenerPrimerNumeroDeFileName(string fileName)
        {
            // Extrae solo el nombre del archivo sin la ruta ni extensión
            var nombreArchivo = Path.GetFileNameWithoutExtension(fileName);
            // Busca el primer grupo de dígitos en el nombre
            var match = Regex.Match(nombreArchivo, @"\d+");
            return match.Success ? match.Value : string.Empty;
        }
    }
}