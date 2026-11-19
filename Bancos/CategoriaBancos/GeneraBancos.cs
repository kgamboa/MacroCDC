using System;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using TesoreriaCobranza;
using static System.Net.WebRequestMethods;
using Common;
using ExcelUtil;

namespace TesoreriaCobranza
{

    public class GeneraBancos
    {

#if localhost        
        public static Dictionary<string, List<MovimientoBancario>> GeneraArchivoBancos(string pathDirectory = @"C:\Tesoreria\Bancos\", string outfile = @"C:\Tesoreria\Bancos\Test.xlsx")
#else        
        public static Dictionary<string, List<MovimientoBancario>> GeneraArchivoBancos(string pathDirectory = @"F:\Tesoreria\Bancos\", string outfile = @"F:\Tesoreria\Bancos\Test.xlsx")
#endif
        {

            Dictionary<string, List<MovimientoBancario>> lista = new Dictionary<string, List<MovimientoBancario>>();
            var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
            {
                HasHeaderRecord = false
            };

            //Obtiene EXCEL
            #region ObtieneExcel
            string sheet = string.Empty;
            string empresa = "CC";
            var directorio = new DirectoryInfo(pathDirectory);
            if (!directorio.Exists) return lista;

            lista.Add("BanCoppelCC", new List<MovimientoBancario>());
            lista.Add("BanorteCC", new List<MovimientoBancario>());
            lista.Add("BAZCC_USD", new List<MovimientoBancario>());

            var files = UtilityIO.GetFilesByExtensions2(directorio, SearchOption.AllDirectories, "~", "", ".exp", ".csv", ".xlsx", ".xls");
            var count = files.Count();
            foreach (var file in files)
            {
                count--;
                if (file.Name.StartsWith("~")) continue;
                
                if (file.Name.StartsWith("mov"))
                {
                    #region BAZ
                    string sexcelconnectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file.FullName + "; Extended Properties =\"Excel 12.0;HDR=Yes;\"";

                    using (OleDbConnection dbConnection = new OleDbConnection(sexcelconnectionstring))
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
                            string firstSheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();

                            // Now we have the table name; proceed as before:
                            using (var adapter = new OleDbDataAdapter("SELECT * FROM [" + firstSheetName + "]", dbConnection))
                            {
                                OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
                                var ds = new DataSet();
                                adapter.Fill(ds, firstSheetName);

                                var data = ds.Tables[firstSheetName].AsEnumerable();

                                var i = 0;
                                foreach (var item in data)
                                {
                                    string observaciones = string.Empty;

                                    //NUMERO DE CUENTA FECHA DE OPERACION  FECHA DE APLICACION CONCEPTO    IMPORTE SALDO   MOVIMIENTO
                                    //Fecha (DD/MM/AA)
                                    if (DateTime.TryParse(item[1].ToString(), out var fechaPago))
                                    {
                                        string cuenta = item[0].ToString().ToUpper();
                                        if (cuenta.Contains("01720117409100")
                                            || cuenta.Contains("01720117487142")
                                            || cuenta.Contains("01720118646957")
                                            || cuenta.Contains("01720500845190"))
                                            empresa = "CC";
                                        else if (cuenta.Contains("01720173481105"))
                                            empresa = "CL";
                                        else if (cuenta.Contains("01720118077763"))
                                            empresa = "CI";

                                        sheet = TipoBanco.BAZ.ToString() + cuenta.Substring(cuenta.Length - 4) + empresa;
                                        if (!lista.ContainsKey(sheet))
                                        {
                                            lista.Add(sheet, new List<MovimientoBancario>());
                                        }

                                        if (DateTime.TryParse(item[2].ToString(), out var fechaAplicacion))
                                        {
                                            string concepto = item[3].ToString().ToUpper();
                                            if (decimal.TryParse(item[4].ToString(), out var importe))
                                                if (decimal.TryParse(item[5].ToString(), out var saldo))
                                                {
                                                    if (int.TryParse(item[6].ToString(), out var movimiento))
                                                    {

                                                        MovimientoBancario nuevoMoviemiento = new MovimientoBancario()
                                                        {
                                                            Banco = TipoBanco.BAZ.ToString(),
                                                            BancoKey = ((int)TipoBanco.BAZ).ToString(),
                                                            rowIndex = movimiento,
                                                            Concepto = concepto,
                                                            Empresa = empresa,
                                                            FechaAplicacion = fechaAplicacion,
                                                            FechaOperacion = fechaPago,
                                                            Cargo = importe < 0 ? importe : 0,
                                                            Abono = importe > 0 ? importe : 0,
                                                            Saldo = saldo,
                                                            Referencia = null,
                                                            ReferenciaAmpliada = null,
                                                            Cuenta = cuenta,
                                                            TipoMovimiento = TipoMovimiento.Tesoreria,
                                                            Sheet = sheet
                                                        };
                                                        nuevoMoviemiento.TipoMovimiento = ProcesoBancos.ReclasificaMovimientos(nuevoMoviemiento);
                                                        lista[sheet].Add(nuevoMoviemiento);
                                                        //Console.WriteLine(nuevoMoviemiento.ToString());
                                                    }
                                                }
                                        }
                                    }
                                    i++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine();
                            Console.WriteLine("ex:" + ex.Message);
                        }
                        finally
                        {
                            dbConnection.Close();
                            Console.WriteLine();
                            Console.WriteLine("CUENTA:" + sheet);
                            Console.WriteLine("Total:" + lista[sheet].Count);
                        }
                    }
                    #endregion
                }
                else if (file.Name.Contains("Cheque"))
                {
                    #region Santander
                    try
                    {
                        sheet = TipoBanco.SANTANDER.ToString() + empresa;
                        if (!lista.ContainsKey(sheet))
                            lista.Add(sheet, new List<MovimientoBancario>());
                        //Cuenta Fecha   Hora Sucursal    Descripción Cargo/ Abono Importe Saldo   Referencia Concepto
                        using (var streamReader = file.OpenText())
                        using (var csvReader = new CsvReader(streamReader, csvConfig))
                        {

                            int secuencia = 0;
                            while (csvReader.Read())
                            {
                                secuencia++;
                                csvReader.TryGetField<string>(0, out var cuenta);
                                csvReader.TryGetField<string>(1, out var fecha);
                                csvReader.TryGetField<string>(2, out var Hora);
                                if (DateTime.TryParse(fecha.Replace("'", "").Insert(2, "/").Insert(5, "/") + " " + Hora, out var fechaPago))
                                {

                                    csvReader.TryGetField<string>(4, out var concepto);
                                    csvReader.TryGetField<string>(5, out var tipo);
                                    csvReader.TryGetField<decimal>(6, out var importe);
                                    csvReader.TryGetField<decimal>(7, out var saldo);
                                    csvReader.TryGetField<string>(8, out var referencia);
                                    csvReader.TryGetField<string>(9, out var referenciaAmpliada);



                                    MovimientoBancario nuevoMoviemiento = new MovimientoBancario()
                                    {
                                        Banco = TipoBanco.SANTANDER.ToString(),
                                        BancoKey = ((int)TipoBanco.SANTANDER).ToString(),
                                        rowIndex = secuencia,
                                        Concepto = concepto,
                                        Empresa = empresa,
                                        FechaAplicacion = fechaPago,
                                        FechaOperacion = fechaPago,
                                        Cargo = tipo == "-" ? importe : 0,
                                        Abono = tipo == "+" ? importe : 0,
                                        Saldo = saldo,
                                        Referencia = referencia,
                                        ReferenciaAmpliada = referenciaAmpliada,
                                        Cuenta = cuenta,
                                        TipoMovimiento = TipoMovimiento.Tesoreria,
                                        Sheet = sheet
                                    };
                                    nuevoMoviemiento.TipoMovimiento = ProcesoBancos.ReclasificaMovimientos(nuevoMoviemiento);
                                    lista[sheet].Add(nuevoMoviemiento);
                                    //Console.WriteLine(nuevoMoviemiento.ToString());
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine();
                    Console.WriteLine("CUENTA:" + sheet);
                    Console.WriteLine("Total:" + lista[sheet].Count);
                    #endregion
                }
                else if (file.Name.Contains("ECta"))
                {
                    #region Intercam
                    try
                    {
                        //FECHA, HORA, DESCRIPCION, REFERENCIA, ORIGEN, SUCURSAL, CARGOS, ABONOS, SALDO, OPERACION, TRANSACCION
                        using (var streamReader = file.OpenText())
                        using (var csvReader = new CsvReader(streamReader, csvConfig))
                        {
                            int secuencia = 0;
                            decimal saldo = 0;
                            while (csvReader.Read())
                            {
                                csvReader.TryGetField<string>(3, out var RFC);

                                if (RFC != null && RFC.Contains("RFC"))
                                {
                                    if (RFC.Contains("CLA061205BG3"))
                                        empresa = "CL";
                                    else if (RFC.Contains("CCR041221JS4"))
                                        empresa = "CC";

                                    sheet = TipoBanco.INTERCAM.ToString() + empresa+"_USD";
                                    if (!lista.ContainsKey(sheet))
                                        lista.Add(sheet, new List<MovimientoBancario>());
                                }

                                csvReader.TryGetField<string>(1, out var fecha);
                                csvReader.TryGetField<string>(2, out var Hora);
                                if (fecha?.Length >= 9 && DateTime.TryParse(fecha.Replace("_", "").Insert(2, "/").Insert(5, "/") + " " + Hora.Replace("_", ""), out var fechaOperacion))
                                {
                                    secuencia++;
                                    csvReader.TryGetField<string>(3, out var concepto);
                                    csvReader.TryGetField<string>(4, out var referencia);
                                    csvReader.TryGetField<string>(5, out var cuenta);
                                    csvReader.TryGetField<decimal>(6, out var x);
                                    csvReader.TryGetField<decimal>(7, out var cargo);
                                    csvReader.TryGetField<decimal>(8, out var abono);
                                    csvReader.TryGetField<string>(9, out var operacion);
                                    csvReader.TryGetField<string>(10, out var transaccion);
                                    cargo = Math.Abs(cargo) * -1;
                                    saldo += abono + cargo;

                                    MovimientoBancario nuevoMoviemiento = new MovimientoBancario()
                                    {
                                        Banco = TipoBanco.INTERCAM.ToString(),
                                        BancoKey = ((int)TipoBanco.INTERCAM).ToString(),
                                        rowIndex = secuencia,
                                        Concepto = concepto.Replace("_",""),
                                        Empresa = empresa,
                                        FechaAplicacion = fechaOperacion,
                                        FechaOperacion = fechaOperacion,
                                        Cargo = cargo,
                                        Abono = abono,
                                        Saldo = saldo,
                                        Referencia = referencia.Replace("_",""),
                                        ReferenciaAmpliada = operacion,
                                        Cuenta = cuenta.Replace("_",""),
                                        TipoMovimiento = TipoMovimiento.Tesoreria,
                                        Sheet = sheet
                                    };
                                    nuevoMoviemiento.TipoMovimiento = ProcesoBancos.ReclasificaMovimientos(nuevoMoviemiento);
                                    lista[sheet].Add(nuevoMoviemiento);
                                    //Console.WriteLine(nuevoMoviemiento.ToString());
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine();
                    Console.WriteLine("CUENTA:" + sheet);
                    Console.WriteLine("Total:" + lista[sheet].Count);
                    #endregion
                }
                else if (file.Name.ToUpper().Contains("MH"))
                {
                    #region BBVA
                    try
                    {
                        //FECHA, HORA, DESCRIPCION, REFERENCIA, ORIGEN, SUCURSAL, CARGOS, ABONOS, SALDO, OPERACION, TRANSACCION
                        using (var streamReader = file.OpenText())
                        using (var csvReader = new CsvReader(streamReader, csvConfig))
                        {
                            string value =string.Empty, cuenta = string.Empty, concepto = string.Empty;
                            DateTime? fechaOperacion = null;
                            DateTime? fechaAplicacion = null;
                            int secuencia = 0;
                            decimal saldo = 0;
                            decimal cargo = 0, abono = 0;
                            while (csvReader.Read())
                            {
                                csvReader.TryGetField<string>(0, out value);

                                var registro = value.Substring(0, 2);
                                if (registro == "11")
                                {
                                    cuenta = value.Substring(10, 10);
                                    saldo = decimal.Parse(value.Substring(33, 14).Insert(12, "."));

                                    if (cuenta == "0148354065")
                                        empresa = "CC";
                                    else if (cuenta == "0193853454")
                                        empresa = "CL";
                                    else if (cuenta == "0193236056")
                                        empresa = "CC";

                                    sheet = TipoBanco.BBVA.ToString() + cuenta.Substring(cuenta.Length - 4) + empresa;
                                    if (!lista.ContainsKey(sheet))
                                        lista.Add(sheet, new List<MovimientoBancario>());
                                }
                                else if (registro == "22")
                                {
                                    cargo = abono = 0;
                                    concepto = value.Substring(52, 28);
                                    var fecha = value.Substring(10, 6);

                                    var year = int.Parse(fecha.Substring(0, 2))+2000;
                                    var month = int.Parse(fecha.Substring(2, 2));
                                    var day = int.Parse(fecha.Substring(4, 2));

                                        fechaOperacion = new DateTime(year,month,day);
                                    var cargoAbono = value.Substring(27, 1);//1:Cargo , 2:Abono
                                    var importe = decimal.Parse(value.Substring(28, 14).Insert(12, "."));
                                    var fAplicacion = value.Substring(16, 6);


                                     year = int.Parse(fAplicacion.Substring(0, 2)) + 2000;
                                     month = int.Parse(fAplicacion.Substring(2, 2));
                                     day = int.Parse(fAplicacion.Substring(4, 2));

                                    fechaAplicacion = new DateTime(year, month, day);
                                    if (cargoAbono == "1")
                                    {
                                        cargo = Math.Abs(importe)*-1;
                                        saldo -= cargo;
                                    }
                                    else
                                    {
                                        abono = importe;
                                        saldo += abono;
                                    }
                                }
                                else if (registro == "23")
                                {
                                    var referenciaAmpliada = value.Length==80 ?  value.Substring(4, 38):"";
                                    var referencia = value.Length==80?  value.Substring(42, 38): value;

                                    MovimientoBancario nuevoMoviemiento = new MovimientoBancario()
                                    {
                                        Banco = TipoBanco.BBVA.ToString(),
                                        BancoKey = ((int)TipoBanco.BBVA).ToString(),
                                        rowIndex = secuencia,
                                        Concepto = concepto,
                                        Empresa = empresa,
                                        FechaAplicacion = fechaAplicacion.HasValue ? fechaAplicacion.Value : DateTime.Now,
                                        FechaOperacion = fechaOperacion,
                                        Cargo = cargo,
                                        Abono = abono,
                                        Saldo = saldo,
                                        Referencia = referencia,
                                        ReferenciaAmpliada = referenciaAmpliada,
                                        Cuenta = cuenta,
                                        TipoMovimiento = TipoMovimiento.Tesoreria,
                                        Sheet = sheet
                                    };
                                    nuevoMoviemiento.TipoMovimiento = ProcesoBancos.ReclasificaMovimientos(nuevoMoviemiento);
                                    lista[sheet].Add(nuevoMoviemiento);
                                    //Console.WriteLine(nuevoMoviemiento.ToString());
                                }
                                else if (registro == "33")
                                {
                                    var saldoFinal = decimal.Parse(value.Substring(59, 14).Insert(12, "."));
                                    if (saldo != saldoFinal)
                                    {
                                        //algo mal
                                    }
                                }
                                secuencia++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine();
                    Console.WriteLine("CUENTA:" + sheet);
                    Console.WriteLine("Total:" + lista[sheet].Count);
                    #endregion
                }
                else if (file.Name.ToLower().StartsWith("afirme"))
                {
                    #region Afirme
                    try
                    {
                        using (var streamReader = file.OpenText())
                        using (var csvReader = new CsvReader(streamReader, csvConfig))
                        {
                            string value = string.Empty;

                            while (csvReader.Read())
                            {
                                csvReader.TryGetField<string>(0, out value);
                                var items = value.Split('\t');

                                //Fecha (DD/MM/AA)
                                if (DateTime.TryParse(items[1], out var fechaOperacion))
                                {
                                    var concepto = items[0];
                                    var referencia = items[2];
                                    var cargo = Math.Abs(decimal.Parse(items[3]))*-1;
                                    var abono = decimal.Parse(items[4]);
                                    var saldo = decimal.Parse(items[5]);
                                    var cuenta = items[6];
                                    var codigo = items[7];
                                    var secuencia = int.Parse(items[8]);

                                    if (cuenta == "137102676")
                                        empresa = "CC";
                                    if (cuenta == "128110917")
                                        empresa = "CL";
                                    if (cuenta == "128110690")
                                        empresa = "CI";

                                    sheet = TipoBanco.AFIRME.ToString() + empresa;
                                    if (!lista.ContainsKey(sheet))
                                        lista.Add(sheet, new List<MovimientoBancario>());

                                    MovimientoBancario nuevoMoviemiento = new MovimientoBancario()
                                    {
                                        Banco = TipoBanco.AFIRME.ToString(),
                                        BancoKey = ((int)TipoBanco.AFIRME).ToString(),
                                        rowIndex = secuencia,
                                        Concepto = concepto,
                                        Empresa = empresa,
                                        FechaAplicacion = fechaOperacion,
                                        FechaOperacion = fechaOperacion,
                                        Cargo = cargo,
                                        Abono = abono,
                                        Saldo = saldo,
                                        Referencia = referencia,
                                        ReferenciaAmpliada = codigo,
                                        Cuenta = cuenta,
                                        TipoMovimiento = TipoMovimiento.Tesoreria,
                                        Sheet = sheet
                                    };
                                    nuevoMoviemiento.TipoMovimiento = ProcesoBancos.ReclasificaMovimientos(nuevoMoviemiento);
                                    lista[sheet].Add(nuevoMoviemiento);
                                    //Console.WriteLine(nuevoMoviemiento.ToString());
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine();
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine();
                    Console.WriteLine("CUENTA:" + sheet);
                    Console.WriteLine("Total:" + lista[sheet].Count);
                    #endregion
                }

                #region Obtiene OutFileName
                var first = lista.Where(c => c.Key == sheet).Select(c => c.Value).FirstOrDefault();
                if (first != null)
                {
                    var firstMov = first.FirstOrDefault();
                    if (firstMov != null)
                    {
                        var periodo = firstMov.FechaOperacion.Value.ToString("yyyyMMMM");
                        var month = firstMov.FechaOperacion.Value.ToString("MM");

                        if (outfile.Contains("Test"))
                        {
                            outfile = outfile.Replace("Test", month.PadLeft(2, '0') + " Mov Resultado_" + periodo);
                        }

                        if (outfile.Contains(periodo))
                        {
                            UtilityIO.MarcaDirectorioProcesado(file);
                        }
                        else
                        {
                            lista.Remove(sheet);
                        }
                    }
                }
                #endregion
            }
            #endregion



            //ExcelUtil.ProcesaBancos.SaveFile(outfile, lista);


            //ExcelUtil.ProcesaBancos.SaveSemanal(outfile);


            return lista;
        }
    }
}