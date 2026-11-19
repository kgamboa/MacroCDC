using Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
namespace TesoreriaCobranza
{
    public partial class MovimientoBancario
    {
        public string Banco { get; set; }
        public TipoMovimiento TipoMovimiento { get; set; }
        public string Referencia { get; set; }
        public DateTime FechaAplicacion { get; set; }
        public string Concepto { get; set; }
        public int rowIndex { get; set; }
        public string Secuencia { get; set; }
        public decimal ImporteNeto
        {
            get
            {
                return Math.Abs(Abono) > Math.Abs(Cargo) ? Abono : Cargo;
            }
        }

        public string ReferenciaAmpliada { get; set; }
        public string BancoKey { get; set; }
        public Double? ReferenciaNumerica
        {
            get
            {
                if (double.TryParse(Referencia?.Replace(" ", ""), out double refOut))
                    return refOut;
                else return null;
            }
        }

        public override string ToString()
        {
            return string.Join(",", "RowIndex:" + rowIndex, "Empresa:" + Empresa, "Folio" + Folio, "RazonSocial",RazonSocial, "RFC",RFC, "NumOtorgante",NumOtorgante, "NumSAP" + NumSAP, "ImporteNeto" + ImporteNeto, "Concepto",Concepto, "ReferenciaAmpliada", ReferenciaAmpliada, "Estatus" + Estatus);
        }
        private string _nombreOrdenanteSPEI;
        public string NombreOrdenanteSPEI
        {
            get
            {
                return _nombreOrdenanteSPEI;
            }
            set
            {
                _nombreOrdenanteSPEI = UtilityIO.NormalizaNombre(value);
            }
        }
        public string ConceptoSPEI { get; set; }
        public string Sheet { get; set; }
        public string Empresa { get; set; }
        //public DateTime? FechaOperacion { get; set; }
        public decimal Cargo { get; set; }
        public decimal Abono { get; set; }
        public decimal Saldo { get; set; }
        public string Cuenta { get; set; }
        public string CodigoCliente { get; set; }
        public string FormaPago { get; set; }
        public string Folio { get; set; }
        public string NumSAP { get; set; }
        private string _rfc { get; set; }
        public string RFC
        {
            get { return _rfc; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _rfc = value?.Replace(" ", "");
                    if (!string.IsNullOrEmpty(_rfc))
                    {
                        Comentario += ";" + _rfc;
                        Console.WriteLine("rfc: " + _rfc);
                    }
                }
            }
        }
        public bool? Signo { get; set; }
        public dynamic CodigoLeyenda { get; set; }
        public string MetodoPago { get; set; }
        private string _razonSocial { get; set; }
        public string RazonSocial
        {
            get
            {
                return _razonSocial;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _razonSocial = UtilityIO.NormalizaNombre(value);
                    if (!string.IsNullOrEmpty(_razonSocial))
                    {
                        Comentario += ";" + _razonSocial;
                        Console.WriteLine("razonSocial: " + _razonSocial);
                    }
                }
            }
        }
        public string Estatus { get; set; }
        public string Compensacion { get; set; }
        public string NumOtorgante { get; set; }
        public string Comentario { get; set; }
        public bool EncontradoRazonSocial { get; internal set; }

        public bool SearchRefereniciaConceptoExacto(params object[] busqueda)
        {
            foreach (var param in busqueda)
            {
                List<string> datos = new List<string>();
                if (param is string)
                    datos.Add(param as string);

                if (param is string[])
                    datos.AddRange(param as string[]);

                foreach (var dato in datos)
                {
                    if (string.IsNullOrEmpty(dato.Trim())) continue;
                    if (dato.Trim().Length <= 0) continue;

                    if ((Referencia != null && Referencia.ToUpper().Contains(dato.ToUpper())) ||
                        (Concepto != null && Concepto.ToUpper().Contains(dato.ToUpper())))
                        return true;
                }
            }
            return false;
        }
        public bool SearchReferenciaConcepto(params object[] busqueda)
        {
            return SearchRefereniciaConcepto(6, busqueda);
        }
        private bool SearchRefereniciaConcepto(int length = 6, params object[] busqueda)
        {
            foreach (var param in busqueda)
            {
                List<string> datos = new List<string>();
                if (param is string)
                    datos.Add(param as string);

                if (param is string[])
                    datos.AddRange(param as string[]);

                foreach (var dato in datos)
                {
                    if (string.IsNullOrEmpty(dato.Trim())) continue;
                    if (dato.Trim().Length <= length) continue;

                    if ((NombreOrdenanteSPEI != null && NombreOrdenanteSPEI.Replace(" ", "").ToUpper().Contains(dato.Replace(" ", ""))) ||
                        (Referencia != null && Referencia.Replace(" ", "").ToUpper().Contains(dato.Replace(" ", ""))) ||
                        (Concepto != null && Concepto.Replace(" ", "").ToUpper().Contains(dato.Replace(" ", ""))))
                        return true;
                }
            }
            return false;
        }

        //public bool SearchReferencia(params object[] busqueda)
        //{
        //    return SearchData(Referencia, busqueda);
        //}
        public bool ContainsConcepto(params object[] busqueda)
        {
            string concepto = SearchConcepto(busqueda);
            return !string.IsNullOrEmpty(concepto);
        }
        public string SearchConcepto(params object[] busqueda)
        {
            return SearchData(Concepto, busqueda);
        }
        //public bool SearchRFC(params object[] busqueda)
        //{
        //    return SearchData(RFC, busqueda);
        //}

        public bool ContainsReferenciaAmpliada(params object[] busqueda)
        {
            return !string.IsNullOrEmpty(SearchReferenciaAmpliada(busqueda));
        }
        public string SearchReferenciaAmpliada(params object[] busqueda)
        {
            return SearchData(ReferenciaAmpliada?.Replace("BMRCASH", "").Replace(",", ""), busqueda);
        }
        public string SearchData(string data, params object[] busqueda)
        {
            if (!string.IsNullOrEmpty(data))
            {
                List<string> datos = new List<string>();
                foreach (var param in busqueda)
                {
                    if (param is string)
                        datos.Add(param as string);

                    if (param is string[])
                        datos.AddRange(param as string[]);

                }

                foreach (var dato in datos)
                {
                    if (string.IsNullOrEmpty(dato.Trim())) continue;
                    if (dato == "INV" || dato == "FAC" || dato == "FACT" || dato.Length >= 6)
                    {
                        if (dato.ToUpper().Contains(data.ToUpper()) && data.Length > 10)
                            return dato;
                        if (data.ToUpper().Contains(dato.ToUpper()))
                            return dato;
                        if (data.Replace(" ", "").ToUpper().Contains(dato.ToUpper().Replace(" ", "")) && !data.Contains("TRANSF A CIRC") && !data.Contains("CIRCULO DE CREDITO") && !data.Contains("PAGO PROVEEDOR"))
                            return dato;
                    }
                }
            }
            return null;
        }


        public List<double> ObtieneFolio(bool buscaTexto = false, double minValue = 0, double maxValue = double.MaxValue)
        {
            if (maxValue == double.MaxValue)
                maxValue = 999999;
            List<string> dataItems = new List<string>();
            var itemsControl = listaFact.ToArray();
            int follioRefA = 0;
            if (ReferenciaAmpliada != null)
            {
                var list = ReferenciaAmpliada.Split(itemsControl, StringSplitOptions.RemoveEmptyEntries).ToList();
                foreach (var item in list)
                {
                    var x = item.Split(' ');
                    foreach (var y in x)
                    {
                        if (itemsControl.Count(c => y.ToUpper().Contains(c.Trim()) && int.TryParse(y.Replace(c.Trim().ToString(), ""), out follioRefA)) > 0 && follioRefA > 0)
                            dataItems.Add(follioRefA.ToString());
                        if (int.TryParse(y, out follioRefA))
                            dataItems.Add(follioRefA.ToString());
                    }
                }
            }

            if (Referencia != null)
            {
                var list = Referencia.Split(itemsControl, StringSplitOptions.RemoveEmptyEntries).ToList();
                foreach (var item in list)
                {
                    var x = item.Split(' ');
                    foreach (var y in x)
                    {
                        if (itemsControl.Count(c => y.ToUpper().Contains(c.Trim()) && int.TryParse(y.Replace(c.Trim().ToString(), ""), out follioRefA)) > 0 && follioRefA > 0)
                            dataItems.Add(follioRefA.ToString());

                        if (int.TryParse(y, out follioRefA))
                            dataItems.Add(follioRefA.ToString());
                    }
                }
            }
            if (Concepto != null)
            {
                List<string> list = Concepto.Split(itemsControl, StringSplitOptions.RemoveEmptyEntries).ToList();
                foreach (var item in list)
                {
                    var x = item.Split(' ');
                    foreach (var y in x)
                    {
                        if (itemsControl.Count(c => !string.IsNullOrEmpty(c) && y.ToUpper().Contains(c.Trim()) && int.TryParse(y.Replace(c.Trim().ToString(), ""), out follioRefA)) > 0 && follioRefA > 0)
                            dataItems.Add(follioRefA.ToString());

                        if (int.TryParse(y, out follioRefA))
                            dataItems.Add(follioRefA.ToString());
                    }
                }
            }
            if (ConceptoSPEI != null)
            {
                List<string> list = ConceptoSPEI.Split(itemsControl, StringSplitOptions.RemoveEmptyEntries).ToList();
                foreach (var item in list)
                {
                    var x = item.Split(' ');
                    foreach (var y in x)
                    {
                        if (itemsControl.Count(c => y.ToUpper().Contains(c.Trim()) && int.TryParse(y.Replace(c.Trim().ToString(), ""), out follioRefA)) > 0 && follioRefA > 0)
                            dataItems.Add(follioRefA.ToString());
                        if (int.TryParse(y, out follioRefA))
                            dataItems.Add(follioRefA.ToString());
                    }
                }
            }



            List<double> folios = new List<double>();
            if (!buscaTexto)
            {
                foreach (var item2 in dataItems)
                {

                    if (double.TryParse(item2, out double folio))
                    {
                        if (!folios.Contains(folio))
                        {
                            folios.Add(folio);
                            continue;
                        }
                    }
                }
            }

            foreach (var dataItem in dataItems)
            {

                if (double.TryParse(dataItem, out double x))
                {
                    folios.Add(x);
                }
                else
                {

                    foreach (var item in itemsControl)
                    {
                        List<string> foliosRevision = new List<string>();
                        var info = dataItem.Replace('#', ' ').Replace(',', ' ').Replace('.', ' ');

                        var itemControl = item + " ";
                        if (itemControl.Trim().Length > 0 && info.Contains(itemControl))
                        {
                            if (double.TryParse(item, out double d))
                            {
                                foliosRevision.Add(item);
                            }
                            else
                            {
                                var split = info.Substring(info.IndexOf(itemControl) + item.Length).Split(' ');
                                foliosRevision.AddRange(split);
                            }
                        }
                        else
                        {
                            itemControl = item + "-";
                            if (info.Contains(itemControl))
                            {
                                var split = info.Substring(info.IndexOf(itemControl) + item.Length).Split(' ');
                                foliosRevision.AddRange(split);
                            }
                            else if (info.Contains(item))
                            {
                                if (item == " ")
                                {
                                    var split = info.Split(' ');
                                    foliosRevision.AddRange(split);
                                }
                                else if (double.TryParse(item, out double d))
                                {
                                    foliosRevision.Add(item);
                                }
                                else
                                {
                                    var split = info.Substring(info.IndexOf(item) + item.Length).Split(' ');
                                    foliosRevision.AddRange(split);
                                }
                            }
                        }

                        foreach (var data in foliosRevision)
                        {
                            var a = data;
                            if (data.Length == 5)
                                a = data.Substring(0, 5);
                            if (double.TryParse(a.Replace(item + " ", "").Replace(item + "-", "").Replace("-", "").Replace(";", "").Trim(), out x))
                            {
                                if (!folios.Contains(x) && x.ToString().Length > 4)
                                    folios.Add(x);
                            }
                        }
                    }
                }
            }

            return folios.Where(c => c >= minValue && c <= maxValue && c != DateTime.Now.Year && c.ToString().Length > 3).Distinct().ToList();
        }
        private readonly List<string> listaFact = string.Join("¬", "-", "_", ".", ",", "PAGO FACT", "PAGO FAC", "PAGO FA", "PAGO F", "PAGO", "INVOICE", "INVC", "INVL", "INVI", "INV", "INC", "FACTURA", "FACTUR", "FACTU", "FACT", "FAC", "Y", "FA ", "FC", "FOLIO", "FOL","F/", "F-", "F").Split('¬').ToList();
        private readonly List<string> listaRemove = string.Join("¬", "XEXX010101000", "XAXX010101000", "CONSULTAS CIRCULO DE CREDITO BMRCASH", "CONSULTA DE BURO BMRCASH", "REPORTES DE BURO BMRCASH", "BMRCASH","CHEQUE DEVUELTO", "DEP.CHEQUES DE OTRO BANCO", "PAGO CUENTA DE TERCERO", "DEPOSITO CHEQUE BBVA", "DEPOSITO DE TERCERO", "DEPOSITO EN EFECTIVO", "DEPOSITO DE EFECTIVO", "DEPOSITO EFECTIVO PRACTIC", "DEPCHEQUES DE OTRO BANCO").Split('¬').ToList();


        public bool Search(string data, string busqueda)
        {
            if (string.IsNullOrEmpty(busqueda))
                return false;
            var obj = SearchStringSimilitud(data, new string[] { busqueda });
            return obj != null;
        }


        public object SearchStringSimilitud(string data, string[] busqueda)
        {
            try
            {
                data = UtilityIO.NormalizaNombre(data);
                if (data.StartsWith("SPEI RECIBIDO") || data.StartsWith("TEF RECIBIDO"))
                    return null;

                foreach (var item in listaRemove)
                {
                    data = data.Replace(item, "").Trim();
                }
                if (string.IsNullOrEmpty(data))
                    return null;
                if (busqueda.Max(c => c.Length) <= 13)
                {
                 
                   var filteredDataList = busqueda.Where(c => data.ToUpper().Contains(c.ToUpper())).Distinct().ToList();
                    if (filteredDataList.Count == 1)
                        return filteredDataList.FirstOrDefault();

                    List<string> rfcs = RFCExtractor.ExtraerRFCs(data);
                    foreach (var rfc in rfcs)
                    {
                        var filteredDataList2 = busqueda.Where(c => rfc.ToUpper().Contains(c.ToUpper())).Distinct().ToList();
                        if (filteredDataList2.Count == 1)
                            return filteredDataList2.FirstOrDefault();
                        if (rfcs.Count == 1)
                            return rfc;
                    }
                    return null;
                }
                if (data.Contains("REFBNTC") || data.Contains("CTA CLABE"))
                {
                    var filteredDataList = busqueda.Where(c => data.ToUpper().Contains(c.ToUpper())).Distinct().ToList();
                    if (filteredDataList.Count == 1)
                        return filteredDataList.FirstOrDefault();
                }

                var filteredList = busqueda.Where(c=> Common.UtilityString.findSimilarity(c, data)>0.25).Distinct().OrderBy(c=>c).ToList();
                if (filteredList != null)
                {
                    var lastItem = "";
                    double maxIndex = 0;
                    foreach (var item in filteredList.OrderByDescending(c => c.Length))
                    {
                        if (string.IsNullOrEmpty(item))
                            continue;

                        var currentItem = item.Trim();
                        var currentItem2 = item.Trim();
                        double maxSimilarity = 0;
                        do
                        {
                            double similarity = Common.UtilityString.findSimilarity(currentItem, data);
                            maxSimilarity = similarity;
                            if (maxIndex < similarity && similarity >= 0.85)
                            {
                                maxIndex = similarity;
                                lastItem = currentItem;
                            }
                            currentItem = currentItem.Substring(1);

                            similarity = Common.UtilityString.findSimilarity(currentItem2, data);
                            if (similarity > maxSimilarity)
                                maxSimilarity = similarity;
                            if (maxIndex < similarity && similarity >= 0.85)
                            {
                                maxIndex = similarity;
                                lastItem = currentItem2;
                            }
                            currentItem2 = currentItem2.Substring(0, currentItem2.Length - 1);
                        }
                        while (currentItem.Length > 6 && maxSimilarity>0.5);


                        var currentData = data;
                        var currentData2 = data;
                        do
                        {
                            double similarity = Common.UtilityString.findSimilarity(item, currentData);
                            maxSimilarity = similarity;

                            if (maxIndex < similarity && similarity >= 0.85)
                            {
                                maxIndex = similarity;
                                lastItem = item;
                            }
                            currentData = currentData.Substring(1);

                            similarity = Common.UtilityString.findSimilarity(item, currentData2);
                            if (similarity > maxSimilarity)
                                maxSimilarity = similarity;
                            if (maxIndex < similarity && similarity >= 0.85)
                            {
                                maxIndex = similarity;
                                lastItem = item;
                            }
                            currentData2 = currentData2.Substring(0, currentData2.Length - 1);
                        }
                        while (currentData.Length > 6 && maxSimilarity>0.46);
                        if (maxIndex == 1)
                            return SearchData(data, lastItem) as String;
                    }
                    if (lastItem != "")
                        return SearchData(data, lastItem) as String;
                }
            }
            catch
            {
            }
            return null;
        }
    }
    public enum TipoBanco
    {
        AFIRME = 1,
        BAZ = 2,
        BBVA = 3,
        INTERCAM = 4,
        SANTANDER = 5,
        BANCOPPEL = 6,
        BANORTE = 7
    }
    public enum TipoMovimiento
    {
        Cobranza,
        Tesoreria,
        RCE,
        CALIFICATE,
        AVISAME,
        TraspasoPAYPAL,
        Tarjeta,
        Credencial,
        Empleados,
        CajaChica,
        DevolucionProv,
        Intereses,
        FondeoMas,
        FondeoMenos,
        RetiroInversion,
        Inversion,
        PagoProveedor,
        PagoImpuesto,
        CompraUSD,
        PagoDevuelto,
        Comision,
        IVAComision,
        ISRRetenido,
        Devolucion,
        ChequeDevuelto,
        Nomina,
        Finiquito,
        Bonos,
        SinReconocer,
        ErrorSPEI,
        Pensiones,
        OpenPay,
        AvisameMuchoMas,
        Compensacion,
        TransferenciaSPEI,
        CobranzaOP,
        TPV
    }

}
