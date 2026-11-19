using Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TesoreriaCobranza;

namespace PTC.Clases
{
    public static class Account
    {
        public static string CC => "754";
        public static string CL => "764";
        public static string CI => "761";
    }
    public class BankAccount
    {
        public string BankName { get; set; }
        public string Type { get; set; }
        public string Catalog { get; set; }
        public string Sheet { get; set; }
        public string Currency { get; set; }
        public string Description { get; set; }
        public string Company { get; set; }

        public string Prefijo { get; set; }
        public string CeCo
        {
            get {
                if (Company == "CC")
                    return Prefijo + Account.CC;
                else if (Company == "CL")
                    return Prefijo + Account.CL;
                else if (Company == "CI")
                    return Prefijo + Account.CI;

                return "";
            }
        }


        public string GetAccount(TipoMovimiento tipoMovimiento, bool ingreso)
        {

            if (tipoMovimiento == TipoMovimiento.Intereses)
            {
                if (ingreso)
                    return Catalog;
                return "7190000100";
            }
            else if (tipoMovimiento == TipoMovimiento.RCE)
            {
                if (ingreso)
                    return "5003500023";
                return "4003500023";
            }
            else if (tipoMovimiento == TipoMovimiento.CALIFICATE)
            {
                if (ingreso)
                    return "5003500009";
                return "4003500009";
            }
            else if (tipoMovimiento == TipoMovimiento.AvisameMuchoMas)
            {
                if (ingreso)
                    return "5003500075";
                return "4003500010";
            }
            else if (tipoMovimiento == TipoMovimiento.AVISAME)
            {
                if (ingreso)
                    return "PENDIENTE";
                return "PENDIENTE";
            }

            else if (tipoMovimiento == TipoMovimiento.Compensacion)
            {
                if (ingreso)
                    return Catalog;
                return "7010000088";
            }   
            


            if (ingreso)
            {
                if (tipoMovimiento == TipoMovimiento.Comision || tipoMovimiento == TipoMovimiento.IVAComision)
                {
                    return "7120000001";
                }
                else if (tipoMovimiento == TipoMovimiento.ISRRetenido)
                {
                    return "1008060005";
                }
                return "";
            }
            else
            {
                return Catalog;
            }
        }


        internal string GetTypeDoc(TipoMovimiento tipoMovimiento)
        {
            switch (tipoMovimiento)
            {

                case TipoMovimiento.Comision:
                    return "D2";
                case TipoMovimiento.Intereses:
                    return "FH";
                case TipoMovimiento.ISRRetenido:
                    return "K2";
                case TipoMovimiento.OpenPay:
                case TipoMovimiento.TraspasoPAYPAL:
                case TipoMovimiento.Empleados:
                case TipoMovimiento.Tarjeta:
                case TipoMovimiento.Credencial:
                case TipoMovimiento.Compensacion:
                    return "3C";
                default:
                    return "";
            }
        }

        //Concepto                      Cuenta      Tipo Doc    CeCo CdC    CeCo C Lab  CeCo C Inf
        //Comisiones                    7120000001  D2          3340080754  3340080764  3340080761
        //Intereses ganados terceros    7190000100  FH          3340080754  3340080764  3340080761
        //ISR Retenido                  1008060005  K2
    }


    public class BankAccountList
    {
        public List<BankAccount> BankAccounts { get; set; }

        public BankAccountList()
        {
            BankAccounts = new List<BankAccount>();

            try
            {
                // Leer el archivo JSON directamente
                //BankAccounts = JsonConvert.DeserializeObject<List<BankAccount>>(
                //    File.ReadAllText(UtilityIO.MacroPath + @"/CuentasBancos.json"));

                BankAccounts = JsonConvert.DeserializeObject<List<BankAccount>>(
    File.ReadAllText(UtilityIO.MacroPath + @"/CuentasBancos.json"))
    .GroupBy(account => account.Catalog) // Agrupar por 'catalog'
    .Select(group => group.First()) // Seleccionar el primer elemento de cada grupo
    .ToList();

                //File.WriteAllText(UtilityIO.MacroPath + @"/2CuentasBancos.json", JsonConvert.SerializeObject(BankAccounts));

                // Trabajar con los datos
                Console.WriteLine($"Se cargaron {BankAccounts.Count} cuentas bancarias.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
    

}
