using Common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace TesoreriaCobranza
{
    public class ClasificacionBancos : BaseExcel
    {
        private static Dictionary<string, List<MovimientoBancario>> listaPago;
        private static Dictionary<string, List<MovimientoBancario>> listaDevolucionProv;

        // Patrones de factura centralizados para clasificación bancaria
        // Incluye todos los patrones utilizados en el sistema para detectar facturas
        // NOTA: Los patrones cortos "F", "F-", "F/", "FC", "FA", "PAGO F" y "PAGO FA" se manejan por separado con regex
        // NOTA: "PAGO" genérico removido para evitar falsos positivos
        private static readonly string[] PatronesFactura = { 
            "INV", "FACTURA", "FACT", "FAC",                            // Patrones originales específicos
            "INVOICE", "INVC", "INVL", "INVI", "INC",                   // Variaciones de Invoice
            "FACTUR", "FACTU",                                          // Variaciones de Factura
            "PAGO FACT", "PAGO FAC",                                    // Patrones específicos de pago de facturas (sin genéricos)
            "FOLIO", "FOL"                                              // Patrones de folio
        };

        public ClasificacionBancos(string fileName)
        {
            FileName = fileName;
            ErrorServices.singleton.Proceso = "CategorizaBancos";
            Utils.EventLog.LogMessage("Categoriza Bancos : " + fileName);
        }

        /// <summary>
        /// Rangos de importe neto (MXN) con su tipo de movimiento y precio de referencia.
        /// Cada rango va de Desde (incluido) hasta el Desde del siguiente (excluido), así no quedan huecos.
        /// El último rango no tiene tope superior.
        /// </summary>
        private static readonly List<(decimal Desde, TipoMovimiento Tipo, int Precio)> RangosClasificacion =
            new List<(decimal Desde, TipoMovimiento Tipo, int Precio)>
            {
                (0M,    TipoMovimiento.RCE,        50),
                (58M,   TipoMovimiento.CALIFICATE, 58),
                (61M,   TipoMovimiento.RCE,        61),
                (110M,  TipoMovimiento.CALIFICATE, 110),
                (119M,  TipoMovimiento.CALIFICATE, 119),
                (150M,  TipoMovimiento.RCE,        150),
                (190M,  TipoMovimiento.RCE,        190),
                (208M,  TipoMovimiento.CALIFICATE, 208),
                (211M,  TipoMovimiento.RCE,        211),
                (248M,  TipoMovimiento.CALIFICATE, 248),
                (251M,  TipoMovimiento.RCE,        251),
                (269M,  TipoMovimiento.CALIFICATE, 269),
                (309M,  TipoMovimiento.CALIFICATE, 309),
            };

        /// <summary>Solo Santander, BBVA 0798 y BAZ 450 usan rangos de clasificación.</summary>
        private static bool AplicaRangosClasificacion(MovimientoBancario movBanco)
        {
            return movBanco.BancoKey == TipoBanco.SANTANDER.ToString()
                || movBanco.Banco.EndsWith("0798")
                || movBanco.Banco.EndsWith("450");
        }

        /// <summary>
        /// Verifica si el movimiento bancario contiene patrones de factura en concepto o referencia
        /// </summary>
        /// <param name="movBanco">Movimiento bancario a verificar</param>
        /// <returns>True si contiene patrones de factura, False en caso contrario</returns>
        private static bool ContienePatronesFactura(MovimientoBancario movBanco)
        {
            // Primero verificar patrones específicos largos (que no necesitan regex)
            var patronesEspecificos = new string[] { 
                "INV", "FACTURA", "FACT", "FAC",                        // Patrones originales específicos
                "INVOICE", "INVC", "INVL", "INVI", "INC",               // Variaciones de Invoice
                "FACTUR", "FACTU",                                      // Variaciones de Factura
                "PAGO FACT", "PAGO FAC",                                // Patrones específicos de pago de facturas
                "FOLIO", "FOL"                                          // Patrones de folio largos
            };
            
            if (movBanco.SearchRefereniciaConceptoExacto(patronesEspecificos) ||
                movBanco.ContainsReferenciaAmpliada(patronesEspecificos))
            {
                return true;
            }

            // Verificación especial para patrones cortos usando regex
            string textoCompleto = $"{movBanco.Concepto} {movBanco.Referencia} {movBanco.ReferenciaAmpliada}".ToUpper();
            
            // Patrones regex para letras cortas seguidas de números
            // F seguida de números: F123, F-123, F/123, F 123
            var regexPatternF = @"\bF[\s\-/]?\d+\b";
            
            // FC seguida de números: FC123, FC-123, FC/123, FC 123  
            var regexPatternFC = @"\bFC[\s\-/]?\d+\b";
            
            // FA seguida de números: FA123, FA-123, FA/123, FA 123
            var regexPatternFA = @"\bFA[\s\-/]?\d+\b";

            // PAGO F seguida de números: "PAGO F123", "PAGO F-123", "PAGO F/123", "PAGO F 123"
            var regexPatternPagoF = @"\bPAGO\s+F[\s\-/]?\d+\b";

            // PAGO FA seguida de números: "PAGO FA123", "PAGO FA-123", "PAGO FA/123", "PAGO FA 123"  
            var regexPatternPagoFA = @"\bPAGO\s+FA[\s\-/]?\d+\b";
            
            return System.Text.RegularExpressions.Regex.IsMatch(textoCompleto, regexPatternF) ||
                   System.Text.RegularExpressions.Regex.IsMatch(textoCompleto, regexPatternFC) ||
                   System.Text.RegularExpressions.Regex.IsMatch(textoCompleto, regexPatternFA) ||
                   System.Text.RegularExpressions.Regex.IsMatch(textoCompleto, regexPatternPagoF) ||
                   System.Text.RegularExpressions.Regex.IsMatch(textoCompleto, regexPatternPagoFA);
        }

        public new void EjecutaProceso()
        {
            try
            {
                var bancoKey = string.Empty;
                var cuenta = string.Empty;
                using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
                {
                    excel.OpenFile(FileName, false);

                    var sheets = excel.GetSheetNames();

                    foreach (var item in sheets)
                    {

                        char separator = ' ';
                        if (item.Contains("_"))
                            separator = '_';

                        var split = item.Split(separator);
                        if (split.Length > 1)
                        {
                            if ((split[1] == "CC" || split[1] == "CI" || split[1] == "CL"))
                                Empresa = split[1];
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

                        if (Empresa == null)
                            Empresa = "CC";

                        var bancos = new PTC.Clases.BankAccountList();
                        var bankAccounts = bancos.BankAccounts.Where(c => c.Sheet == item).ToList();

                        if (TipoMovimientoColumn.ContainsKey(bancoKey))
                        {
                            try
                            {
                                Utils.EventLog.LogMessage("Categoriza Banco: " + banco);
                                excel.SetActiveSheet(item);
                            }
                            catch (Exception e)
                            {

                                Utils.EventLog.LogMessage("Error Banco: " + banco + " " + e.Message);
                                continue;
                            }


                                Range excelRange = excel.UsedRange;
                            int RowCount = excelRange.Rows.Count;
                            int ColumnCount = excelRange.Columns.Count;

                            SetHeader(excelRange, bancoKey);

                            for (int r = 1; r <= RowCount; r++)
                            {
                                MovimientoBancario nMov = new MovimientoBancario() { Empresa = Empresa, Banco = banco, BancoKey = bancoKey, rowIndex = r, Sheet = item, Signo = null };
                                bool? signo = null;
                                Range excelRange1 = excel.UsedRange;
                                dynamic cell = excelRange1.Cells[nMov.rowIndex, TipoMovimientoColumn[bancoKey]];
                                dynamic content = cell.Value;
                                if (content != null && !string.IsNullOrWhiteSpace(content.ToString()))
                                {
                                    continue;
                                }


                                #region Obtiene informacion Movimiento Banco

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


                                //Reclasifica los tipos de MOvimientos
                                cell = excelRange1.Cells[nMov.rowIndex, TipoMovimientoColumn[bancoKey]];
                                nMov.TipoMovimiento = ReclasificaMovimientos(nMov);
                                cell.value = TipoMovString(nMov.TipoMovimiento);
                            }

                            SetAutoFit(excelRange, bancoKey);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR_Bancos, ex, ex.StackTrace);
            }
        }

        // Método auxiliar para verificar si la cuenta es de moneda MXN
        private static bool EsCuentaMXN(MovimientoBancario movBanco)
        {
            try
            {
                var bancos = new PTC.Clases.BankAccountList();
                var bankAccount = bancos.BankAccounts.FirstOrDefault(c => c.Sheet == movBanco.Sheet);
                
                // Si no se encuentra la cuenta bancaria o no tiene moneda especificada, asume MXN
                if (bankAccount == null || string.IsNullOrEmpty(bankAccount.Currency))
                    return true;
                
                // Retorna true si la moneda es MXN (o null/vacío que se considera MXN por defecto)
                return bankAccount.Currency.ToUpper() == "MXN" || bankAccount.Currency.ToUpper() == "PESO" || 
                       bankAccount.Currency.ToUpper() == "PESOS" || string.IsNullOrWhiteSpace(bankAccount.Currency);
            }
            catch
            {
                // En caso de error, asume MXN por defecto
                return true;
            }
        }

        public static TipoMovimiento ReclasificaMovimientos(MovimientoBancario movBanco)
        {
            try
            {
                // Initialize both dictionaries if needed
                if (listaPago == null)
                    listaPago = new Dictionary<string, List<MovimientoBancario>>();
                
                if (listaDevolucionProv == null)
                    listaDevolucionProv = new Dictionary<string, List<MovimientoBancario>>();

                if (movBanco.Sheet != null && !listaPago.ContainsKey(movBanco.Sheet))
                    listaPago.Add(movBanco.Sheet, new List<MovimientoBancario>());
                
                if (movBanco.Sheet != null && !listaDevolucionProv.ContainsKey(movBanco.Sheet))
                    listaDevolucionProv.Add(movBanco.Sheet, new List<MovimientoBancario>());

                movBanco.TipoMovimiento = TipoMovimiento.SinReconocer;

                // GENERAL DEVOLUTION HANDLING - Should be checked early for all banks
                if (movBanco.SearchRefereniciaConceptoExacto("DEVSPEI", "DEVOLUCION", "DEV SPEI"))
                {
                    if (movBanco.ImporteNeto > 0)
                    {
                        // Positive amount with devolution pattern = DevolucionProv
                        movBanco.TipoMovimiento = TipoMovimiento.DevolucionProv;

                        // Check for matching PagoDevuelto with same amount
                        if (listaDevolucionProv.ContainsKey(movBanco.Sheet))
                        {
                            var pagoDevuelto = listaDevolucionProv[movBanco.Sheet].FirstOrDefault(c =>
                                Math.Abs(c.ImporteNeto) == Math.Abs(movBanco.ImporteNeto) &&
                                c.TipoMovimiento == TipoMovimiento.PagoDevuelto);

                            if (pagoDevuelto != null)
                            {
                                // Match found, keep both as their respective types
                                Utils.EventLog.LogMessage($"DevolucionProv matched with PagoDevuelto: Amount {movBanco.ImporteNeto}");
                            }
                            else
                            {
                                // Add to tracking list for future matching
                                listaDevolucionProv[movBanco.Sheet].Add(movBanco);
                            }
                        }

                        return movBanco.TipoMovimiento;
                    }
                    else if (movBanco.ImporteNeto < 0)
                    {
                        // Negative amount with devolution pattern = PagoDevuelto
                        movBanco.TipoMovimiento = TipoMovimiento.PagoDevuelto;

                        // Check for matching DevolucionProv with same amount
                        if (listaDevolucionProv.ContainsKey(movBanco.Sheet))
                        {
                            var devolucionProv = listaDevolucionProv[movBanco.Sheet].FirstOrDefault(c =>
                                Math.Abs(c.ImporteNeto) == Math.Abs(movBanco.ImporteNeto) &&
                                c.TipoMovimiento == TipoMovimiento.DevolucionProv);

                            if (devolucionProv != null)
                            {
                                // Match found, confirm this is PagoDevuelto
                                Utils.EventLog.LogMessage($"PagoDevuelto matched with DevolucionProv: Amount {movBanco.ImporteNeto}");
                            }
                            else
                            {
                                // Add to tracking list for future matching
                                listaDevolucionProv[movBanco.Sheet].Add(movBanco);
                            }
                        }

                        return movBanco.TipoMovimiento;
                    }
                }

                // Check for PRONTOFONDEO first - this should NOT be classified as Fondeo
                if (movBanco.SearchRefereniciaConceptoExacto("PRONTOFONDEO") || movBanco.ContainsReferenciaAmpliada("PRONTOFONDEO"))
                {
                    // PRONTOFONDEO should not be classified as fondeo, continue to other classification rules
                    // This prevents PRONTOFONDEO from being caught by the general FONDEO logic below
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("FONDEO", "FONDOE") || movBanco.ContainsReferenciaAmpliada("FONDEO", "FONDOE"))// || (movBanco.BancoKey =="AFIRME" && movBanco.SearchRefereniciaConceptoExacto("SPEI ENVIADO")))
                {
                    if (movBanco.ImporteNeto > 0)
                        movBanco.TipoMovimiento = TipoMovimiento.FondeoMas;
                    else
                        movBanco.TipoMovimiento = TipoMovimiento.FondeoMenos;
                    goto EditaExcel;
                }
                
                // VERIFICACIONES ESPECÍFICAS DE TRASPASO PARA FONDEO
                if (movBanco.SearchRefereniciaConceptoExacto("TRASPASO 01720117487142") || 
                    movBanco.ContainsReferenciaAmpliada("TRASPASO 01720117487142"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.FondeoMas;
                    goto EditaExcel;
                }
                
                if (movBanco.SearchRefereniciaConceptoExacto("TRASPASO 01720117409100") || 
                    movBanco.ContainsReferenciaAmpliada("TRASPASO 01720117409100"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.FondeoMenos;
                    goto EditaExcel;
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("IVA COMISION", "IVA COM", "IVA DE COM", "IVA POR" , "IVA_TAX", "IVAUSD", "IVA USD", "IVAEUR", "IVA EUR", "IVA DE LA COM", "I V A  POR COMISION"))                
                {
                    movBanco.TipoMovimiento = TipoMovimiento.IVAComision;
                    goto EditaExcel;
                }

                if (movBanco.ContainsReferenciaAmpliada("PAGO OPENMX", "OPENPAY") || movBanco.SearchRefereniciaConceptoExacto("PAGO TBIXUTLGZFGVGUD4", "OPENPAY", "PAGO OPENMX"))
                {
                    if (movBanco.Banco.EndsWith("798") || movBanco.Banco.EndsWith("450") || movBanco.BancoKey == TipoBanco.SANTANDER.ToString())
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.OpenPay;
                        return movBanco.TipoMovimiento;
                    }
                    else if (movBanco.BancoKey == TipoBanco.BBVA.ToString())
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.CobranzaOP;
                        return movBanco.TipoMovimiento;
                    }
                }

                if (movBanco.BancoKey == TipoBanco.INTERCAM.ToString())
                {
                    if (Math.Abs(movBanco.ImporteNeto) == 30)
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Comision;
                    }
                    else if (movBanco.Cargo != 0 && movBanco.Cargo != 30)
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.PagoProveedor;

                        if (listaPago.ContainsKey(movBanco.Sheet))
                        {
                            var devPago = listaPago[movBanco.Sheet].FirstOrDefault(c => Math.Abs(c.ImporteNeto) == Math.Abs(movBanco.ImporteNeto) && c.TipoMovimiento == TipoMovimiento.Cobranza);
                            if (devPago != null)
                            {
                                movBanco.TipoMovimiento = TipoMovimiento.PagoDevuelto;
                            }
                            else
                            {
                                listaPago[movBanco.Sheet].Add(movBanco);
                            }
                        }
                    }
                    else if (movBanco.Abono != 0)
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                        if (listaPago.ContainsKey(movBanco.Sheet))
                        {
                            var devPago = listaPago[movBanco.Sheet].FirstOrDefault(c => Math.Abs(c.ImporteNeto) == Math.Abs(movBanco.ImporteNeto) && c.TipoMovimiento == TipoMovimiento.PagoProveedor);
                            if (devPago != null)
                            {
                                devPago.TipoMovimiento = TipoMovimiento.PagoDevuelto;
                            }
                            else
                            {
                                listaPago[movBanco.Sheet].Add(movBanco);
                            }
                        }
                    }

                    goto EditaExcel;
                }
                else if (movBanco.BancoKey == TipoBanco.AFIRME.ToString())
                {
                    if (movBanco.SearchRefereniciaConceptoExacto("COMISION"))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Comision;
                    }
                    else if (movBanco.Cargo != 0 && (movBanco.SearchRefereniciaConceptoExacto("DISPERSION DE FONDOS") || (movBanco.SearchRefereniciaConceptoExacto("RASTREO"))))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Nomina;
                        goto EditaExcel;
                    }
                    else if (movBanco.Cargo != 0 && movBanco.CodigoLeyenda == "CN" && movBanco.SearchRefereniciaConceptoExacto("FIN","FINIQUITO", "FNQT", "CONTRATACIÓN - DISPERSIÓN DE NÓMINA", "CONTRATACIÓN DISPERSIÓN DE NÓMINA", "CONTRATACIÓN DISPERSIÓN", "DISPERSIÓN DE NÓMINA", "4W CONTRATACIÓN - DISPERSIÓN DE NÓMINA"))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Finiquito;
                        goto EditaExcel;
                    }
                    else if (movBanco.Cargo != 0 && movBanco.CodigoLeyenda == "CK" && movBanco.SearchRefereniciaConceptoExacto("CHEQUE"))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Bonos;
                        goto EditaExcel;
                    }

                    if (movBanco.Empresa == "CC")
                    {
                        if (movBanco.Abono > 0)
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                            if (movBanco.CodigoLeyenda == "QG" && movBanco.SearchRefereniciaConceptoExacto("VENTA TPV"))
                            {
                                movBanco.TipoMovimiento = TipoMovimiento.TPV;
                            }
                            goto EditaExcel;
                        }
                    }
                }
                else if (movBanco.BancoKey == TipoBanco.BAZ.ToString())
                {
                    if (movBanco.Banco.EndsWith("42") && ContienePatronesFactura(movBanco))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                        return movBanco.TipoMovimiento;
                    }

                    // Note: DEVSPEI and DEVOLUCION patterns are now handled in the general devolution logic above
                    // No need to duplicate that logic here

                    if (movBanco.SearchRefereniciaConceptoExacto("Transferencia envio SPEI banca"))
                    {
                        if (movBanco.ImporteNeto == -8)
                            movBanco.TipoMovimiento = TipoMovimiento.Comision;
                        else
                            movBanco.TipoMovimiento = TipoMovimiento.TransferenciaSPEI;
                        goto EditaExcel;
                    }
                    if (movBanco.Banco.EndsWith("57") || movBanco.Banco.EndsWith("307"))
                    {
                        if (movBanco.SearchRefereniciaConceptoExacto("COMISION SPEI"))
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Comision;
                            goto EditaExcel;
                        }
                        else if (movBanco.ImporteNeto > 200)
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Empleados;
                            goto EditaExcel;
                        }
                        else if (movBanco.ImporteNeto == 200)
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Credencial;
                            goto EditaExcel;
                        }
                        else if (movBanco.ImporteNeto < 200)
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Tarjeta;
                            goto EditaExcel;
                        }
                    }
                    if (movBanco.Banco.EndsWith("00"))
                    {
                        if (movBanco.SearchReferenciaConcepto("00000 0500"))
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Devolucion;
                            goto EditaExcel;
                        }
                    }
                    if (movBanco.Banco.EndsWith("00") || movBanco.Empresa.Contains("CI") || movBanco.Empresa.Contains("CL"))
                    {
                        //if (movBanco.Concepto.Trim().Length == 7 && int.TryParse(movBanco.Concepto, out int numero))
                        if (movBanco.ContainsConcepto("Vencimiento MD", "Inversion MD"))
                        {
                            if (movBanco.ImporteNeto > 0)
                                movBanco.TipoMovimiento = TipoMovimiento.RetiroInversion;
                            else
                                movBanco.TipoMovimiento = TipoMovimiento.Inversion;
                            goto EditaExcel;
                        }
                    }
                }
                else if (movBanco.BancoKey == TipoBanco.BANCOPPEL.ToString())
                {
                    if (movBanco.SearchRefereniciaConceptoExacto("INTERESES"))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Intereses;
                        goto EditaExcel;
                    }
                }
                else if (movBanco.BancoKey == TipoBanco.BANORTE.ToString())
                {
                    if (movBanco.SearchRefereniciaConceptoExacto("I.V.A."))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.IVAComision;
                        goto EditaExcel;
                    }
                    if (movBanco.SearchRefereniciaConceptoExacto("COMPRA ORDEN DE PAGO SPEI"))
                    {
                        if (movBanco.ImporteNeto > 0)
                            movBanco.TipoMovimiento = TipoMovimiento.FondeoMas;
                        else
                            movBanco.TipoMovimiento = TipoMovimiento.FondeoMenos;
                        goto EditaExcel;
                    }
                }
                else if (movBanco.BancoKey == TipoBanco.BBVA.ToString())
                {
                    if (movBanco.Banco.EndsWith("4065") || movBanco.Banco.EndsWith("0798"))
                    {
                        if (ContienePatronesFactura(movBanco))
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                            return movBanco.TipoMovimiento;
                        }
                    }
                }
                else if (movBanco.BancoKey == TipoBanco.SANTANDER.ToString())
                {
                    // Nueva regla: Si la referencia es exactamente "106180000124550108", clasificar como TraspasoPAYPAL
                    if (movBanco.ImporteNeto > 0 && movBanco.SearchRefereniciaConceptoExacto("106180000124550108"))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.TraspasoPAYPAL;
                        return movBanco.TipoMovimiento;
                    }
                    if(movBanco.SearchRefereniciaConceptoExacto("ADMINISTRACION RENTA"))
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Comision;
                        return movBanco.TipoMovimiento;
                    }
                }

                if (movBanco.SearchRefereniciaConceptoExacto("PENSION"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.Pensiones;
                }
                else if (movBanco.ImporteNeto > 0 && movBanco.SearchRefereniciaConceptoExacto("1920677592 106", "TEF RECIBIDO BANK OF AMER"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.TraspasoPAYPAL;
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("I.S.R. RETENIDO", "RETENCION ISR", "RETENCION I S R"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.ISRRetenido;
                }
                else if (movBanco.ImporteNeto < 0 && (movBanco.CodigoLeyenda == "P13" || movBanco.CodigoLeyenda == "P14" || movBanco.CodigoLeyenda == "X01")
                    && movBanco.SearchRefereniciaConceptoExacto("CDMX", "IMSS", "SAT", "IMPUESTO", "RECAUDACION DE IMPUE", "2%SN", "PAGO IMSS", "04200GRD600026518417", "IMSS/INF/AFORE"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.PagoImpuesto;
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("INTERESES GANADOS", "AB INTERESES", "ABO POR INTERESES"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.Intereses;
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("COMPENSACION POR RETRASO", "MORA SPEI NORMABANXICO", "MORA SPEI", "MORA"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.Compensacion;
                }
                else if (movBanco.CodigoLeyenda == "Y45" || movBanco.SearchRefereniciaConceptoExacto("COMVTA", "COMISION SPEI", "COMISION", "CGO.", "CGO.VTAS", "COM MEM", "COM SDO", "COM ANUALIDAD", "COM ", "Compensacion", "INTERES RECEPCION SPEI TARDIO"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.Comision;
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("ANN SPEI", "CAN SPEI BCO DEST NO DISPONIBL"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.ErrorSPEI;
                }
                else if (movBanco.SearchRefereniciaConceptoExacto("DEVUELTO"))//, "DEP.CHEQUES DE OTRO BANCO"))
                {
                    movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                }
                //else if (movBanco.CodigoLeyenda == "QG" || movBanco.SearchRefereniciaConceptoExacto("VENTA TPV", "VENTAS TPV", "VENTAS DEBITO", "VENTAS DEBITO", "TERMINALES PUNTO DE VENTA") || movBanco.SearchReferenciaAmpliada("TERMINALES PUNTO DE VENTA"))
                //{
                //    movBanco.TipoMovimiento = TipoMovimiento.Tesoreria;
                //}

                else if (movBanco.ImporteNeto < 0)
                {
                    if (movBanco.BancoKey == TipoBanco.AFIRME.ToString())
                    {
                        movBanco.TipoMovimiento = TipoMovimiento.Nomina;
                    }
                    else
                    {
                        // Check first if this could be a PagoDevuelto based on specific patterns
                        // For BAZ bank, check if this is specifically a devolution-related negative transaction
                        if (movBanco.BancoKey == TipoBanco.BAZ.ToString() && 
                            movBanco.SearchRefereniciaConceptoExacto("DEVSPEI", "DEVOLUCION", "DEV", "DEVOL", "COMPROBACION"))
                        {
                            // This is specifically a devolution-related negative transaction, likely a PagoDevuelto
                            movBanco.TipoMovimiento = TipoMovimiento.PagoDevuelto;

                            // Check for matching DevolucionProv with same amount
                            if (listaDevolucionProv.ContainsKey(movBanco.Sheet))
                            {
                                var devolucionProv = listaDevolucionProv[movBanco.Sheet].FirstOrDefault(c =>
                                    Math.Abs(c.ImporteNeto) == Math.Abs(movBanco.ImporteNeto) &&
                                    c.TipoMovimiento == TipoMovimiento.DevolucionProv);

                                if (devolucionProv != null)
                                {
                                    // Match found, confirm this is PagoDevuelto
                                    Utils.EventLog.LogMessage($"PagoDevuelto matched with DevolucionProv: Amount {movBanco.ImporteNeto}");
                                }
                                else
                                {
                                    // Add to tracking list for future matching
                                    listaDevolucionProv[movBanco.Sheet].Add(movBanco);
                                }
                            }
                        }
                        else
                        {
                            // Default to PagoProveedor for other negative transactions
                            movBanco.TipoMovimiento = TipoMovimiento.PagoProveedor;

                            // Check if this could be a PagoDevuelto by looking for matching DevolucionProv
                            if (listaDevolucionProv.ContainsKey(movBanco.Sheet))
                            {
                                var devolucionProv = listaDevolucionProv[movBanco.Sheet].FirstOrDefault(c =>
                                    Math.Abs(c.ImporteNeto) == Math.Abs(movBanco.ImporteNeto) &&
                                    c.TipoMovimiento == TipoMovimiento.DevolucionProv);

                                if (devolucionProv != null)
                                {
                                    movBanco.TipoMovimiento = TipoMovimiento.PagoDevuelto;
                                    Utils.EventLog.LogMessage($"PagoDevuelto matched with DevolucionProv: Amount {movBanco.ImporteNeto}");
                                }
                                else
                                {
                                    // Add to tracking list for future matching
                                    listaDevolucionProv[movBanco.Sheet].Add(movBanco);
                                }
                            }
                        }
                    }
                }
                else if (movBanco.ImporteNeto > 0)
                {
                    movBanco.TipoMovimiento = TipoMovimiento.Cobranza;

                    // Verificar si contiene patrones de factura - si es así, mantener como Cobranza y no aplicar clasificación por rangos
                    if (!ContienePatronesFactura(movBanco))
                    {
                        // Los rangos de clasificación (RCE, CALIFICATE) solo aplican a Santander, BBVA 0798 y BAZ 450, y solo para moneda MXN
                            // y cuando NO hay indicadores de factura en el concepto
                            // EXCEPCIÓN: BBVA 9854 no aplica rangos, siempre es Cobranza puro
                            if (AplicaRangosClasificacion(movBanco) && EsCuentaMXN(movBanco))
                            {
                                // Rango aplicable: el de mayor Desde que sea <= importe (sin huecos).
                                // Los rangos terminan en 310 (importes de 310.99 o menos); arriba de eso no hay regla.
                                var candidatos = movBanco.ImporteNeto < 311M
                                    ? RangosClasificacion
                                        .Where(r => movBanco.ImporteNeto >= r.Desde)
                                        .ToList()
                                    : new List<(decimal Desde, TipoMovimiento Tipo, int Precio)>();
                                candidatos = candidatos
                                    .OrderByDescending(r => r.Desde)
                                    .ToList();
                                if (candidatos.Count > 0)
                                    movBanco.TipoMovimiento = candidatos[0].Tipo;
                            }
                    }
                    // Si contiene patrones de factura, se mantiene como TipoMovimiento.Cobranza (se pasa a la siguiente regla)
                }

                else if (movBanco.BancoKey == TipoBanco.BBVA.ToString())
                {
                    if (movBanco.Banco.EndsWith("4065"))
                    {
                        if (ContienePatronesFactura(movBanco))
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                            return movBanco.TipoMovimiento;
                        }
                    }

                    if (movBanco.Banco.EndsWith("0798"))
                    {
                        if (ContienePatronesFactura(movBanco))
                        {
                            movBanco.TipoMovimiento = TipoMovimiento.Cobranza;
                            return movBanco.TipoMovimiento;
                        }
                    }
                }

            EditaExcel:
                return movBanco.TipoMovimiento;

            }
            catch (Exception ex)
            {
                // we are using dynamic type for cell variable so
                // the variable might not have all the properties we used in our code
                throw (ex);
            }
        }
    }
}
