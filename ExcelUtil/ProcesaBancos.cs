//using Common;
//using CoreType.Clases;
//using CoreType.ClasesCobranza;
//using Microsoft.Office.Interop.Excel;
//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.Globalization;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace ExcelUtil
//{
//    public static class ProcesaBancos
//    {
//        private static readonly object FechaColumn = 1;
//        private static readonly object FechaOperacionColumn = 2;
//        private static readonly object ConceptoColumn = 3;
//        private static readonly object ReferenciaColumn = 4;
//        private static readonly object CargoColumn = 5;
//        private static readonly object AbonoColumn = 6;
//        private static readonly object SaldoColumn = 7;
//        private static readonly object SecuenciaColumn = 8;
//        private static readonly object CuentaColumn = 9;
//        private static readonly object ReferenciaAmpliadaColumn = 10;
//        private static readonly object SemanaColumn = 11;

//        private static readonly int TipoMovimientoColumn = 12;
//        private static readonly int RazonSocialColumn = 13;
//        private static readonly int NumeroSAPColumn = 14;
//        private static readonly int FacturaColumn = 15;
//        private static readonly int OtorganteColumn = 16;
//        private static readonly int FormaPagoColumn = 17;
//        private static readonly int CompensacionColumn = 18;
//        private static readonly int MetodoColumn = 19;
//        private static readonly int EstatusPolizaColumn = 20;
//        private static readonly int ObservacionesColumn= 21;
//        public static void SaveFile(string outfile, Dictionary<string, List<MovimientoBancario>> lista)
//        {

//            if (lista.Count > 0)
//            {

//                //Guarda EXCEL
//                FileInfo file1 = new FileInfo(outfile);
//                using (Utility excel = new Utility())
//                {
//                    excel.OpenFile(file1.FullName);

//                    var sheets = excel.GetSheetNames();

//                    var ordenLista = new List<string>()
//                    {
//                        "BAZ9100CC","BAZ7142CC","BAZ6957CC","BBVA4065CC","BBVA6056CC","AfirmeCC","SantanderCC","BanCoppelCC","BanorteCC",
//                        "IntercamCC_USD","BAZCC_USD",
//                        "BAZ1105CL","BBVA3454CL","AfirmeCL",
//                        "BAZ7763CI", "AfirmeCI"
//                    };

//                    foreach (var item in ordenLista)
//                    {
//                        if (sheets.Where(c => c.Contains(item)).FirstOrDefault() == null)
//                             excel.CreateSheet(item,true);
//                        else
//                            excel.MoveActiveSheet(item);

//                        if(item.EndsWith("CC"))
//                            excel.ColorTab( 10);
//                        if (item.EndsWith("CC_USD"))
//                            excel.ColorTab(1);
//                        if (item.EndsWith("CL"))
//                            excel.ColorTab( 21);
//                        if (item.EndsWith("CI"))
//                            excel.ColorTab(46);

//                        Range excelRange = excel.UsedRange;
//                        int RowCount = excelRange.Rows.Count;
//                        int ColumnCount = excelRange.Columns.Count;
//                        var r = RowCount;

//                        if (RowCount == 1)
//                        {

//                            //agrega nombres de columnnas
//                            ((Range)excelRange.Cells[r, FechaOperacionColumn]).Value = "FechaOperacion";
//                            ((Range)excelRange.Cells[r, FechaColumn]).Value = "FechaAplicacion";
//                            ((Range)excelRange.Cells[r, ConceptoColumn]).Value = "Concepto";
//                            ((Range)excelRange.Cells[r, ReferenciaColumn]).Value = "Referencia";
//                            ((Range)excelRange.Cells[r, CargoColumn]).Value = "Cargo";
//                            ((Range)excelRange.Cells[r, AbonoColumn]).Value = "Abono";
//                            ((Range)excelRange.Cells[r, SaldoColumn]).Value = "Saldo";
//                            ((Range)excelRange.Cells[r, SecuenciaColumn]).Value = "Secuencia";
//                            ((Range)excelRange.Cells[r, CuentaColumn]).Value = "Cuenta";
//                            ((Range)excelRange.Cells[r, ReferenciaAmpliadaColumn]).Value = "ReferenciaAmpliada";

//                            ((Range)excelRange.Cells[r, SemanaColumn]).Value = "Semana";
//                            ((Range)excelRange.Cells[r, TipoMovimientoColumn]).Value = "Tipo";
//                            ((Range)excelRange.Cells[r, OtorganteColumn]).Value = "Otorgante";
//                            ((Range)excelRange.Cells[r, RazonSocialColumn]).Value = "Razon Social";
//                            ((Range)excelRange.Cells[r, NumeroSAPColumn]).Value = "SAP";
//                            ((Range)excelRange.Cells[r, FacturaColumn]).Value = "Facturación";
//                            ((Range)excelRange.Cells[r, FormaPagoColumn]).Value = "Forma Pago";
//                            ((Range)excelRange.Cells[r, CompensacionColumn]).Value = "Documento de Compensación";
//                            ((Range)excelRange.Cells[r, MetodoColumn]).Value = "Metodo";
//                            ((Range)excelRange.Cells[r, EstatusPolizaColumn]).Value = "Estatus Poliza";
//                            ((Range)excelRange.Cells[r, ObservacionesColumn]).Value = "Observaciones";

//                        }
//                        if(!lista.ContainsKey(item)) continue;
//                        if (item.ToUpper().Contains("AFIRME"))
//                            r = 1;
//                        r++;
//                        foreach (var movBanc in lista[item].OrderByDescending(c => c.rowIndex).ThenBy(c => c.FechaAplicacion))
//                        {
//                            ((Range)excelRange.Cells[r, ReferenciaAmpliadaColumn]).EntireColumn.NumberFormat = "@";
//                            ((Range)excelRange.Cells[r, CuentaColumn]).EntireColumn.NumberFormat = "@";
//                            ((Range)excelRange.Cells[r, AbonoColumn]).EntireColumn.NumberFormat = "$#,##0.00";
//                            ((Range)excelRange.Cells[r, AbonoColumn]).EntireColumn.AutoFit();
//                            ((Range)excelRange.Cells[r, SaldoColumn]).EntireColumn.NumberFormat = "$#,##0.00";
//                            ((Range)excelRange.Cells[r, SaldoColumn]).EntireColumn.AutoFit();
//                            ((Range)excelRange.Cells[r, CargoColumn]).EntireColumn.NumberFormat = "$#,##0.00";
//                            ((Range)excelRange.Cells[r, CargoColumn]).EntireColumn.AutoFit();
//                            ((Range)excelRange.Cells[r, FechaColumn]).EntireColumn.NumberFormat = "dd/MM/yyyy";
//                            ((Range)excelRange.Cells[r, FechaOperacionColumn]).EntireColumn.NumberFormat = "dd/MM/yyyy";

//                            ((Range)excelRange.Cells[r, FechaColumn]).Value2 = movBanc.FechaAplicacion;
//                            ((Range)excelRange.Cells[r, FechaOperacionColumn]).Value2 = movBanc.FechaOperacion;
//                            ((Range)excelRange.Cells[r, ConceptoColumn]).Value2 = movBanc.Concepto;
//                            ((Range)excelRange.Cells[r, ReferenciaColumn]).Value2 = movBanc.Referencia;
//                            ((Range)excelRange.Cells[r, CargoColumn]).Value2 = Math.Abs(movBanc.Cargo)*-1;
//                            ((Range)excelRange.Cells[r, AbonoColumn]).Value2 = movBanc.Abono;
//                            ((Range)excelRange.Cells[r, SaldoColumn]).Value2 = movBanc.Saldo;
//                            ((Range)excelRange.Cells[r, SecuenciaColumn]).Value2 = movBanc.rowIndex;
//                            ((Range)excelRange.Cells[r, CuentaColumn]).Value2 = movBanc.Cuenta;
//                            ((Range)excelRange.Cells[r, ReferenciaAmpliadaColumn]).Value2 = movBanc.ReferenciaAmpliada;
//                            ((Range)excelRange.Cells[r, ReferenciaAmpliadaColumn]).EntireColumn.AutoFit();

//                            ((Range)excelRange.Cells[r, TipoMovimientoColumn]).Value2 = TipoMovString(movBanc.TipoMovimiento);


//                            int NumSemana = System.Globalization.CultureInfo.CurrentUICulture.Calendar.GetWeekOfYear(movBanc.FechaAplicacion, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
//                            DateTime firtDateMonth = new DateTime(movBanc.FechaAplicacion.Year, movBanc.FechaAplicacion.Month, 1);
//                            int x = System.Globalization.CultureInfo.CurrentUICulture.Calendar.GetWeekOfYear(firtDateMonth, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday  );
//                            ((Range)excelRange.Cells[r, SemanaColumn]).Value = "Sem" + (NumSemana - x + 1);

//                            r++;
//                        }
//                    }
//                }
//            }

//        }

//        public static void SaveSemanal(string outfile)
//        {
//            FileInfo file1 = new FileInfo(outfile);
//            using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
//            {
//                excel.OpenFile(file1.FullName);

//                var sheets = excel.GetSheetNames();
//                var sheetNameSem = "Semanal";
//                if (sheets.Where(c => c.Contains(sheetNameSem)).FirstOrDefault() == null)
//                    excel.CreateSheet(sheetNameSem);
//                else
//                {
//                    excel.SetActiveSheet(sheetNameSem);
//                }
//                excel.MoveActiveSheet(sheetNameSem, 1);

//                Range excelRange = excel.UsedRange;
//                int RowCount = excelRange.Rows.Count;
//                int ColumnCount = excelRange.Columns.Count;
//                if (RowCount > 1 && ColumnCount > 1)
//                    return;


//                List<TipoMovimiento> tipoMov = new List<TipoMovimiento>();
//                tipoMov.AddRange((TipoMovimiento[])Enum.GetValues(typeof(TipoMovimiento)));
//                tipoMov.Remove(TipoMovimiento.ErrorSPEI);

//                var bancosCount = 0;
//                var x = 0;
//                var empresa = "";

//                var sheetsList = excel.GetSheetNames();
//                foreach (var sheetName in sheetsList)
//                {
//                    var empresaAct = sheetName.Contains("CC_USD") ? sheetName.Substring(sheetName.Length - 6) : sheetName.Substring(sheetName.Length - 2);

//                    var r = 4;

//                    if (empresaAct == "CC")
//                        r = 4;
//                    else if (empresaAct == "CC_USD")
//                        r += 8 + tipoMov.Count;
//                    else if (empresaAct == "CL")
//                        r += 12+ (tipoMov.Count *2);
//                    else if (empresaAct == "CI")
//                        r += 16 + (tipoMov.Count *3);
//                    else continue;

//                    if (empresa != empresaAct)
//                    {
//                        bancosCount = 1;
//                        empresa = empresaAct;

//                        ((Range)excelRange.Cells[r, bancosCount]).Value = "Concepto";
//                        x = r + 1;
//                        foreach (TipoMovimiento tipo in tipoMov)
//                        {
//                            ((Range)excelRange.Cells[x++, bancosCount]).Value = TipoMovString(tipo);
//                            ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();
//                        }
//                        bancosCount++;
//                        ((Range)excelRange.Cells[r - 1, bancosCount++]).Value = "TOTAL";
//                        for (int y = 1; y <= 5; y++)
//                        {
//                            ((Range)excelRange.Cells[r, bancosCount]).Value = "Sem" + y;
//                            x = r;
//                            foreach (TipoMovimiento tipo in tipoMov)
//                            {
//                                ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                                ((Range)excelRange.Cells[x, y + 2]).FormulaR1C1 = "=0";
//                                ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();
//                            }

//                            ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                            ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C{1}:R{2}C{1})", (r + 1), bancosCount, x - 1);
//                            ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();

//                            bancosCount++;
//                        }

//                        ((Range)excelRange.Cells[r, bancosCount]).Value = "Total";
//                        x = r;
//                        foreach (TipoMovimiento tipo in tipoMov)
//                        {
//                            ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                            ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C2:R{0}C{1})", x, (bancosCount - 1));
//                            ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();
//                        }

//                        ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                        ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C{1}:R{2}C{1})", (r + 1), bancosCount, x - 1);
//                        ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();

//                        //Suma Validaciones
//                        ((Range)excelRange.Cells[x++, 8]).EntireColumn.NumberFormat = "$#,##0.00";
//                        ((Range)excelRange.Cells[x, 8]).FormulaR1C1 = "=0";
//                        ((Range)excelRange.Cells[x, 8]).EntireColumn.AutoFit();

//                        //Validacion Final
//                        ((Range)excelRange.Cells[x, 9]).EntireColumn.NumberFormat = "$#,##0.00";
//                        ((Range)excelRange.Cells[x, 9]).FormulaR1C1 = string.Format("=R{0}C8-R{1}C8", x, x - 1);
//                        ((Range)excelRange.Cells[x, 9]).EntireColumn.AutoFit();
//                        bancosCount++;

//                    }


//                    bancosCount++;
//                    ((Range)excelRange.Cells[r - 1, bancosCount++]).Value = sheetName;
//                    var bancosIndex = bancosCount;
//                    for (int y = 1; y <= 5; y++)
//                    {
//                        ((Range)excelRange.Cells[r, bancosCount]).Value = "Sem" + y;
//                        x = r;
//                        foreach (TipoMovimiento tipo in tipoMov)
//                        {
//                            ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                            ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUMIFS({0}!C6:C6,{0}!C11:C11,R{1}C{2},{0}!C12:C12,R{3}C1)+SUMIFS({0}!C5:C5,{0}!C11:C11,R{1}C{2},{0}!C12:C12,R{3}C1)", sheetName, r, bancosCount, x);
//                            ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();

//                            ((Range)excelRange.Cells[x, y + 2]).FormulaR1C1 += string.Format("+R{0}C{1}", x, bancosCount);
//                            ((Range)excelRange.Cells[x, y + 2]).EntireColumn.AutoFit();

//                        }

//                        ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                        ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C{1}:R{2}C{1})", (r + 1), bancosCount, x - 1);
//                        ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();
//                        bancosCount++;
//                    }

//                    ((Range)excelRange.Cells[r, bancosCount]).Value = "Importe2";
//                    x = r;
//                    foreach (TipoMovimiento tipo in tipoMov)
//                    {
//                        ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                        ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUMIFS({0}!C6:C6,{0}!C12:C12,R{1}C1)+SUMIFS({0}!C5:C5,{0}!C12:C12,R{1}C1)", sheetName, x);
//                        ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();
//                    }

//                    ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                    ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C{1}:R{2}C{1})", (r + 1), bancosCount, x - 1);
//                    ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();

//                    //Validacion 
//                    ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                    ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM({0}!C6:C6)+SUM({0}!C5:C5)", sheetName);
//                    ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();

//                    //Suma Validacion
//                    ((Range)excelRange.Cells[x, 8]).FormulaR1C1 += string.Format("+R{0}C{1}", x, bancosCount);
//                    ((Range)excelRange.Cells[x, 8]).EntireColumn.AutoFit();

//                    bancosCount++;
//                    ((Range)excelRange.Cells[r, bancosCount]).Value = "Var";
//                    x = r;
//                    foreach (TipoMovimiento tipo in tipoMov)
//                    {
//                        ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                        ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C{1}:R{0}C{2})-R{0}C{3}", x, bancosIndex, (bancosCount - 2), (bancosCount - 1));
//                        ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();
//                    }

//                    ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                    ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=SUM(R{0}C{1}:R{2}C{1})", (r + 1), bancosCount, x - 1);
//                    ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();


//                    //Validacion 
//                    ((Range)excelRange.Cells[x++, bancosCount]).EntireColumn.NumberFormat = "$#,##0.00";
//                    ((Range)excelRange.Cells[x, bancosCount]).FormulaR1C1 = string.Format("=R{0}C{1}-R{2}C{1}", x, bancosCount - 1, x - 1);
//                    ((Range)excelRange.Cells[x, bancosCount]).EntireColumn.AutoFit();


//                    bancosCount++;

//                }

//            }
//        }

//        public static string BuscaCorreoElectronico(string nombre)
//        {
//            using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
//            {
//                var fileName = @"C:\Pago Proveedores\BaseDatosCorreosElectronicos.xlsx"; 
//                excel.OpenFile(fileName);


//                var sheets = excel.GetSheetNames();
//                var item = sheets.FirstOrDefault();
//                excel.SetActiveSheet(item);
//                Microsoft.Office.Interop.Excel.Range range = excel.UsedRange;

//                for (int r = 1; r <= range.Rows.Count; r++)
//                {
//                    string nombreBD = range.Cells[r, 2].value2?.ToString();
//                    if (nombreBD != null)
//                    {
//                        nombreBD = nombreBD.Replace(" ", "");
//                        var nombreN = nombre.Replace("¿", "").Replace("|", "").Replace("'", "").Replace("?", "").Replace("!", "").Replace(",", "").Replace(".", "").Replace(" ", "").Replace(";", "").Replace("…", "").Replace("-", "").Replace("_", "").Replace("!", "").Trim();
//                        if (nombreN.ToUpper() == nombreBD.ToUpper())
//                            return range.Cells[r, 3].value2;
//                    }
//                }
//            }
//            return null;
//        }
//        public static string TipoMovString(TipoMovimiento mov)
//        {
//            switch (mov)
//            {
//                case TipoMovimiento.Compensacion: return "Compensacion por retraso\r\n";
//                case TipoMovimiento.Cobranza: return "Cobranza";
//                case TipoMovimiento.FondeoMas: return "Fondeo+";
//                case TipoMovimiento.FondeoMenos: return "Fondeo-";
//                case TipoMovimiento.RetiroInversion: return "Retiro Inversion";
//                case TipoMovimiento.Comision: return "Comision";
//                case TipoMovimiento.Intereses: return "Intereses";
//                case TipoMovimiento.ISRRetenido: return "ISR Retenido";
//                case TipoMovimiento.Devolucion: return "Devolución";
//                case TipoMovimiento.RCE: return "RCE";
//                case TipoMovimiento.IVAComision: return "IVA Comision";
//                case TipoMovimiento.PagoProveedor: return "Pago Proveedor";
//                case TipoMovimiento.ErrorSPEI: return "Error SPEI";
//                case TipoMovimiento.Inversion: return "Inversion";
//                case TipoMovimiento.CALIFICATE: return "CALIFICATE";
//                case TipoMovimiento.AVISAME: return "AVISAME";
//                case TipoMovimiento.AvisameMuchoMas: return "Avisame Mucho Mas";
//                case TipoMovimiento.ChequeDevuelto: return "Cheque Devuelto";
//                case TipoMovimiento.Tesoreria: return "Tesoreria";
//                case TipoMovimiento.TraspasoPAYPAL: return "Traspaso PAYPAL";
//                case TipoMovimiento.PagoImpuesto: return "Pago Impuesto";
//                case TipoMovimiento.Nomina: return "Nomina";
//                case TipoMovimiento.Tarjeta: return "Tarjeta";
//                case TipoMovimiento.Empleados: return "Empleados";
//            }

//            return mov.ToString();
//        }
//    }
//}
