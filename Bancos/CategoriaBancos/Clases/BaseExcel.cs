using TesoreriaCobranza;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{
    public class BaseExcel: IProceso
    {
        public string FileName;
        public string Empresa;
        public Dictionary<string, int> ConceptoColumn;
        public Dictionary<string, int> FechaColumn;
        //public Dictionary<string, int> FechaOperacionColumn;
        public Dictionary<string, int> ReferenciaColumn;
        public Dictionary<string, int> ReferenciaAmpliadaColumn;
        public Dictionary<string, int> ImporteColumn;
        public Dictionary<string, int> CargoColumn;
        public Dictionary<string, int> AbonoColumn;
        public Dictionary<string, int> SignoColumn;
        public Dictionary<string, int> SecuenciaColumn;
        public Dictionary<string, int> CodigoColumn;
        public Dictionary<string, int> TipoMovimientoColumn;
        public Dictionary<string, string> listRefCIE; 
        public Dictionary<string, List<RefBBVA>> listRefBBVA;
        public Dictionary<string, Dictionary<string, BBVA>> listBBVA;
        public Dictionary<string, Dictionary<string, SPEI>> listSPEI;
        public Dictionary<string, Dictionary<string, CFDI>> listCFDI;
        public Dictionary<string, int> SemanaColumn;
        public Dictionary<string, int> SemanaKeyColumn;
        public List<string> listBancosOpenPay;

        public readonly int CuentaOrdenanteColumn = 15;
        public readonly int CuentaClabeOrdenanteColumn = 16;
        public readonly int NombreOrdenanteColumn = 17;
        public readonly int RFCColumn = 18;
        public readonly int EspacioColumn = 19;

        //Razon social	Ciente SAP	Documento de Compensación	Fecha Póliza	Fecha	Importe	Forma de pago (01,02,03)	
        //NÚMERO;OTORGANTE	Factura Pagada	Método Pago (PPD / PUE)	REP relacionado
        public readonly int RazonColumn = 20;
        public readonly int SAPColumn = 21;
        public readonly int CompensacionColumn = 22;
        public readonly int FechaPolizaColumn = 23;
        public readonly int FechaFacturaColumn = 24;
        public readonly int ImporteFacturaColumn = 25;
        public readonly int FormaPagoColumn = 26;
        public readonly int NumOtorganteColumn = 27;
        public readonly int FacturaPagadaColumn = 28;
        public readonly int MetodoPagoColumn = 29;
        public readonly int REPColumn = 30;
        public readonly int EstatusColumn = 31;
        public readonly int ComentarioColumn = 32;

        public readonly int opCreateDateColumn = 34;
        public readonly int opDescriptionColumn = 35;
        public readonly int opCustomNameColumn = 36;
        public readonly int opUserNameColumn = 37;
        public readonly int opAmountColumn = 38;
        public readonly int opFeeAuraColumn = 39;
        public readonly int opFeeIGHColumn = 40;
        public readonly int opFeeTaxColumn = 41;
        public readonly int opSAPOpenPayColumn = 42;
        public readonly int opCompensa1Column = 43;
        public readonly int opCompensa2Column = 44;
        public readonly int EstatusPolizaColumn = 45;


        public void SetHeader(Range excelRange, string bancoKey, string banco = null)
        {

            #region HEADER
            var r = 1;

            for (int i = 1; i <= 44; i++)
            {
                // Center the text horizontally and vertically
                excelRange.Cells[r, i].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                excelRange.Cells[r, i].VerticalAlignment = XlVAlign.xlVAlignCenter;
            }


            //agrega nombres de columnnas

            excelRange.Cells[r, RFCColumn].value = "RFC";
            excelRange.Cells[r, RFCColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            excelRange.Cells[r, NombreOrdenanteColumn].value = "Nombre Ordenante".ToUpper();
            excelRange.Cells[r, NombreOrdenanteColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);            
            
            excelRange.Cells[r, CuentaOrdenanteColumn].value = "Cuenta Ordenante".ToUpper();
            excelRange.Cells[r, CuentaOrdenanteColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            excelRange.Cells[r, CuentaClabeOrdenanteColumn].value = "Cuenta Cable Ordenante".ToUpper();
            excelRange.Cells[r, CuentaClabeOrdenanteColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            excelRange.Cells[r, EspacioColumn].value = "".ToUpper();
            excelRange.Cells[r, EspacioColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            //agrega nombres de columnnas
            excelRange.Cells[r, TipoMovimientoColumn[bancoKey]].value = "Tipo".ToUpper();
            excelRange.Cells[r, TipoMovimientoColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            excelRange.Cells[r, SemanaColumn[bancoKey]].value = "Semana".ToUpper();
            excelRange.Cells[r, SemanaColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            excelRange.Cells[r, SemanaKeyColumn[bancoKey]].value = "Semana Llave".ToUpper();
            excelRange.Cells[r, SemanaKeyColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

            if (FechaColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, FechaColumn[bancoKey]].value = "Fecha".ToUpper();
                excelRange.Cells[r, FechaColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (ConceptoColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, ConceptoColumn[bancoKey]].value = "Concepto".ToUpper();
                excelRange.Cells[r, ConceptoColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (CodigoColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, CodigoColumn[bancoKey]].value = "Codigo".ToUpper();
                excelRange.Cells[r, CodigoColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (ReferenciaAmpliadaColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, ReferenciaAmpliadaColumn[bancoKey]].value = "Referencia Ampliada".ToUpper();
                excelRange.Cells[r, ReferenciaAmpliadaColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (AbonoColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, AbonoColumn[bancoKey]].value = "Abono".ToUpper();
                excelRange.Cells[r, AbonoColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (SecuenciaColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, SecuenciaColumn[bancoKey]].value = "Secuencia".ToUpper();
                excelRange.Cells[r, SecuenciaColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (SignoColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, SignoColumn[bancoKey]].value = "Signo".ToUpper();
                excelRange.Cells[r, SignoColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (CargoColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, CargoColumn[bancoKey]].value = "Cargo".ToUpper();
                excelRange.Cells[r, CargoColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (ImporteColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, ImporteColumn[bancoKey]].value = "Importe".ToUpper();
                excelRange.Cells[r, ImporteColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }
            if (ReferenciaColumn.ContainsKey(bancoKey))
            {
                excelRange.Cells[r, ReferenciaColumn[bancoKey]].value = "Referencia".ToUpper();
                excelRange.Cells[r, ReferenciaColumn[bancoKey]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            }

           
                //Razon social	Ciente SAP	Documento de Compensación	Fecha Póliza	Fecha	Importe	Forma de pago (01,02,03)	
                //NÚMERO;OTORGANTE	Factura Pagada	Método Pago (PPD / PUE)	REP relacionado
                excelRange.Cells[r, RazonColumn].value = "Razon social".ToUpper();
                excelRange.Cells[r, SAPColumn].value = "SAP".ToUpper();
                excelRange.Cells[r, CompensacionColumn].value = "Documento \r\nCompensación".ToUpper();
                excelRange.Cells[r, FechaPolizaColumn].value = "Fecha Póliza".ToUpper();
                excelRange.Cells[r, RazonColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
                excelRange.Cells[r, SAPColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
                excelRange.Cells[r, CompensacionColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
                excelRange.Cells[r, FechaPolizaColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);



                excelRange.Cells[r, FechaFacturaColumn].value = "Fecha".ToUpper();
                excelRange.Cells[r, ImporteFacturaColumn].value = "Importe".ToUpper();
                excelRange.Cells[r, FormaPagoColumn].value = "Forma pago \r\n(01, 02, 03)".ToUpper();
                excelRange.Cells[r, NumOtorganteColumn].value = "Número \r\nOtorgante".ToUpper();
                excelRange.Cells[r, FacturaPagadaColumn].value = "Factura \r\nPagada".ToUpper();
                excelRange.Cells[r, FechaFacturaColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                excelRange.Cells[r, ImporteFacturaColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                excelRange.Cells[r, FormaPagoColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                excelRange.Cells[r, NumOtorganteColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                excelRange.Cells[r, FacturaPagadaColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);



                excelRange.Cells[r, MetodoPagoColumn].value = "Método Pago \r\n(PPD / PUE)".ToUpper();
                excelRange.Cells[r, REPColumn].value = "REP relacionado".ToUpper();
                excelRange.Cells[r, MetodoPagoColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, REPColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);

                excelRange.Cells[r, EstatusColumn].value = "Estatus".ToUpper();
                excelRange.Cells[r, EstatusColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);
                excelRange.Cells[r, ComentarioColumn].value = "Comentario".ToUpper();
                excelRange.Cells[r, ComentarioColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);
            
            if (listBancosOpenPay.Contains(banco))
            {
                excelRange.Cells[r, opCreateDateColumn].value = "creation_date".ToUpper();
                excelRange.Cells[r, opDescriptionColumn].value = "descrption".ToUpper();
                excelRange.Cells[r, opCustomNameColumn].value = "customer_name".ToUpper();
                excelRange.Cells[r, opUserNameColumn].value = "correo".ToUpper();
                excelRange.Cells[r, opCreateDateColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, opDescriptionColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, opCustomNameColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, opUserNameColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);

                excelRange.Cells[r, 24].value = "Razón social".ToUpper();
                excelRange.Cells[r, opSAPOpenPayColumn].value = "SAP \r\nOPENPAY \r\n2864811".ToUpper();

                excelRange.Cells[r, opAmountColumn].value = "Amount".ToUpper();
                excelRange.Cells[r, opFeeAuraColumn].value = "Fee \r\n5003500049 \r\nAURA".ToUpper();
                excelRange.Cells[r, opFeeIGHColumn].value = "Fee \r\n5003500043 \r\nIGH".ToUpper();
                excelRange.Cells[r, opFeeTaxColumn].value = "fee_tax".ToUpper();
                excelRange.Cells[r, opAmountColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, opFeeAuraColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, opFeeIGHColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);
                excelRange.Cells[r, opFeeTaxColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);


                excelRange.Cells[r, opCompensa1Column].value = "Compensación 1".ToUpper();
                excelRange.Cells[r, opCompensa2Column].value = "Compensación 2".ToUpper();
                excelRange.Cells[r, EstatusPolizaColumn].value = "Estatus Poliza".ToUpper();
                excelRange.Cells[r, opCompensa1Column].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);
                excelRange.Cells[r, opCompensa2Column].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);
                excelRange.Cells[r, EstatusPolizaColumn].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);

            }
            #endregion

            SetAutoFit(excelRange, bancoKey);
        }

        public void SetAutoFit(Range excelRange, string bancoKey)
        {

            excelRange.Cells[1, CuentaClabeOrdenanteColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, NombreOrdenanteColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, CuentaOrdenanteColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, RFCColumn].EntireColumn.AutoFit();

            excelRange.Cells[1, SAPColumn].EntireColumn.NumberFormat = "@";
            excelRange.Cells[1, ImporteFacturaColumn].EntireColumn.NumberFormat = "$#,##0.00";
            excelRange.Cells[1, NumOtorganteColumn].EntireColumn.NumberFormat = "@";
            excelRange.Cells[1, FechaPolizaColumn].EntireColumn.NumberFormat = "DD/MM/aaaa";
            excelRange.Cells[1, FechaFacturaColumn].EntireColumn.NumberFormat = "DD/MM/aaaa";

            excelRange.Cells[1, TipoMovimientoColumn[bancoKey]].EntireColumn.AutoFit();
            excelRange.Cells[1, SemanaColumn[bancoKey]].EntireColumn.AutoFit();
            excelRange.Cells[1, SemanaKeyColumn[bancoKey]].EntireColumn.AutoFit();
            excelRange.Cells[1, RazonColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, SAPColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, CompensacionColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, FechaPolizaColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, FechaFacturaColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, ImporteFacturaColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, FormaPagoColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, NumOtorganteColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, FacturaPagadaColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, MetodoPagoColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, REPColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, EstatusColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, ComentarioColumn].EntireColumn.AutoFit();


            excelRange.Cells[1, opAmountColumn].EntireColumn.NumberFormat = "$#,##0.00";
            excelRange.Cells[1, opFeeAuraColumn].EntireColumn.NumberFormat = "$#,##0.00";
            excelRange.Cells[1, opFeeIGHColumn].EntireColumn.NumberFormat = "$#,##0.00";
            excelRange.Cells[1, opFeeTaxColumn].EntireColumn.NumberFormat = "$#,##0.00";
            excelRange.Cells[1, opSAPOpenPayColumn].EntireColumn.NumberFormat = "@";
            excelRange.Cells[1, opCreateDateColumn].EntireColumn.NumberFormat = "DD/MM/aaaa";

            excelRange.Cells[1, opAmountColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opCompensa1Column].EntireColumn.AutoFit();
            excelRange.Cells[1, opCompensa2Column].EntireColumn.AutoFit();
            excelRange.Cells[1, opCreateDateColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opCustomNameColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opDescriptionColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opFeeAuraColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opFeeIGHColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opFeeTaxColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opSAPOpenPayColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, opUserNameColumn].EntireColumn.AutoFit();
            excelRange.Cells[1, EstatusPolizaColumn].EntireColumn.AutoFit();
        }
        public void InicializaProceso()
        {
            listBancosOpenPay = new List<string>() { "BBVA4254", "BBVA0895", "BBVA0798" };
            FechaColumn = new Dictionary<string, int>();
            FechaColumn.Add("BANCOPPEL", 1);
            FechaColumn.Add("BAZ", 1);
            FechaColumn.Add("BBVA", 3);  //Fecha de operación
            FechaColumn.Add("AFIRME", 2);
            FechaColumn.Add("DONDÉ", 2);
            FechaColumn.Add("BANORTE", 2);
            FechaColumn.Add("SANTANDER", 2);
            FechaColumn.Add("INTERCAM", 1);

            TipoMovimientoColumn = new Dictionary<string, int>();
            TipoMovimientoColumn.Add("BANCOPPEL", 12);
            TipoMovimientoColumn.Add("BAZ", 12);
            TipoMovimientoColumn.Add("BBVA", 12);
            TipoMovimientoColumn.Add("AFIRME", 12);
            TipoMovimientoColumn.Add("DONDÉ", 12);
            TipoMovimientoColumn.Add("BANORTE", 13);
            TipoMovimientoColumn.Add("SANTANDER", 12);
            TipoMovimientoColumn.Add("INTERCAM", 12);

            SemanaKeyColumn = new Dictionary<string, int>();
            SemanaKeyColumn.Add("BANCOPPEL", 14);
            SemanaKeyColumn.Add("BAZ", 14);
            SemanaKeyColumn.Add("BBVA", 14);
            SemanaKeyColumn.Add("AFIRME", 14);
            SemanaKeyColumn.Add("DONDÉ", 14);
            SemanaKeyColumn.Add("BANORTE", 15);
            SemanaKeyColumn.Add("SANTANDER", 14);
            SemanaKeyColumn.Add("INTERCAM", 14);

            SemanaColumn = new Dictionary<string, int>();
            SemanaColumn.Add("BANCOPPEL", 13);
            SemanaColumn.Add("BAZ", 13);
            SemanaColumn.Add("BBVA", 13);
            SemanaColumn.Add("AFIRME", 13);
            SemanaColumn.Add("DONDÉ", 13);
            SemanaColumn.Add("BANORTE", 14);
            SemanaColumn.Add("SANTANDER", 13);
            SemanaColumn.Add("INTERCAM", 13);

            //FechaOperacionColumn = new Dictionary<string, int>();
            //FechaOperacionColumn.Add("BAZ", 1);
            //FechaOperacionColumn.Add("BBVA", 3);


            ConceptoColumn = new Dictionary<string, int>();
            ConceptoColumn.Add("BAZ", 3);
            ConceptoColumn.Add("BANCOPPEL", 3);
            ConceptoColumn.Add("BBVA", 7);
            ConceptoColumn.Add("AFIRME", 1);
            ConceptoColumn.Add("SANTANDER", 10);
            ConceptoColumn.Add("INTERCAM", 3);
            ConceptoColumn.Add("BANORTE", 5);
            ConceptoColumn.Add("DONDÉ", 3);


            ReferenciaColumn = new Dictionary<string, int>();
            ReferenciaColumn.Add("BBVA", 8);
            ReferenciaColumn.Add("AFIRME", 3);
            ReferenciaColumn.Add("BANCOPPEL", 2);
            ReferenciaColumn.Add("SANTANDER", 5);
            ReferenciaColumn.Add("INTERCAM", 4);
            ReferenciaColumn.Add("BANORTE", 4);


            ImporteColumn = new Dictionary<string, int>();
            ImporteColumn.Add("BAZ", 4);
            ImporteColumn.Add("BANCOPPEL", 4);
            ImporteColumn.Add("BBVA", 6);
            ImporteColumn.Add("SANTANDER", 7);

            CargoColumn = new Dictionary<string, int>();
            CargoColumn.Add("AFIRME", 4);
            CargoColumn.Add("INTERCAM", 7);
            CargoColumn.Add("BANORTE", 9);
            CargoColumn.Add("DONDÉ", 5);

            SignoColumn = new Dictionary<string, int>();
            SignoColumn.Add("SANTANDER", 6);
            //SignoColumn.Add("BANORTE", 8);

            SecuenciaColumn = new Dictionary<string, int>();
            SecuenciaColumn.Add("BBVA", 11);
            SecuenciaColumn.Add("AFIRME", 9);

            AbonoColumn = new Dictionary<string, int>();
            AbonoColumn.Add("AFIRME", 5);
            AbonoColumn.Add("INTERCAM", 8);
            AbonoColumn.Add("BANORTE", 8);
            AbonoColumn.Add("DONDÉ", 4);

            ReferenciaAmpliadaColumn = new Dictionary<string, int>();
            ReferenciaAmpliadaColumn.Add("BBVA", 9);
            ReferenciaAmpliadaColumn.Add("BANORTE", 11);

            CodigoColumn = new Dictionary<string, int>();
            CodigoColumn.Add("BBVA", 5);
            CodigoColumn.Add("AFIRME", 8);

        }

        public void EjecutaProceso()
        {

        }

        public static TipoMovimiento TipoMov(string value)
        {
            var tipoMovimiento = TipoMovimiento.SinReconocer;
            if (Enum.TryParse<TipoMovimiento>(value, true, out TipoMovimiento tipo))
                tipoMovimiento = tipo;
            else
            {
                switch (value)
                {
                    case "Compensacion por retraso\r\n": return TipoMovimiento.Compensacion;
                    case "Fondeo +": 
                    case "Fondeo+":
                        return TipoMovimiento.FondeoMas;
                    case "Fondeo-":
                    case "Fondeo -": 
                        return TipoMovimiento.FondeoMenos;
                    case "Retiro Inversion": return TipoMovimiento.RetiroInversion;
                    case "ISR Retenido": return TipoMovimiento.ISRRetenido;
                    case "Devolución": return TipoMovimiento.Devolucion;
                    case "IVA Comision": return TipoMovimiento.IVAComision;
                    case "Pago Proveedor": return TipoMovimiento.PagoProveedor;
                    case "Error SPEI": return TipoMovimiento.ErrorSPEI;
                    case "Avisame Mucho Mas": return TipoMovimiento.AvisameMuchoMas;
                    case "Cheque Devuelto": return TipoMovimiento.ChequeDevuelto;
                    case "Traspaso PAYPAL": return TipoMovimiento.TraspasoPAYPAL;
                    case "Pago Impuesto": return TipoMovimiento.PagoImpuesto;
                }
            }

            return tipoMovimiento;
        }
        public static string TipoMovString(TipoMovimiento mov)
        {
            switch (mov)
            {
                case TipoMovimiento.Compensacion: return "Compensacion por retraso\r\n";
                case TipoMovimiento.Cobranza: return "Cobranza";
                case TipoMovimiento.FondeoMas: return "Fondeo +";
                case TipoMovimiento.FondeoMenos: return "Fondeo -";
                case TipoMovimiento.RetiroInversion: return "Retiro Inversion";
                case TipoMovimiento.Comision: return "Comision";
                case TipoMovimiento.Intereses: return "Intereses";
                case TipoMovimiento.ISRRetenido: return "ISR Retenido";
                case TipoMovimiento.Devolucion: return "Devolución";
                case TipoMovimiento.RCE: return "RCE";
                case TipoMovimiento.IVAComision: return "IVA Comision";
                case TipoMovimiento.PagoProveedor: return "Pago Proveedor";
                case TipoMovimiento.ErrorSPEI: return "Error SPEI";
                case TipoMovimiento.Inversion: return "Inversion";
                case TipoMovimiento.CALIFICATE: return "CALIFICATE";
                case TipoMovimiento.AVISAME: return "AVISAME";
                case TipoMovimiento.AvisameMuchoMas: return "Avisame Mucho Mas";
                case TipoMovimiento.ChequeDevuelto: return "Cheque Devuelto";
                case TipoMovimiento.Tesoreria: return "Tesoreria";
                case TipoMovimiento.TraspasoPAYPAL: return "Traspaso PAYPAL";
                case TipoMovimiento.PagoImpuesto: return "Pago Impuesto";
                case TipoMovimiento.Nomina: return "Nomina";
                case TipoMovimiento.Tarjeta: return "Tarjeta";
                case TipoMovimiento.Empleados: return "Empleados";
                case TipoMovimiento.TransferenciaSPEI: return "TransferenciaSPEI";
            }

            return mov.ToString();
        }
    }

}
