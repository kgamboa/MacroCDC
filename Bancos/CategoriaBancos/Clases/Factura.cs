using Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{
    internal class Factura
    {

        public Factura()
        {
        }

        public Factura(Factura item)
        {
            this.RFC = item.RFC;
            this.IVA = item.IVA;
            this.Refactura = item.Refactura;
            this.Refacturacion = item.Refacturacion;
            this.EstatusFactura = item.EstatusFactura;
            this.Empresa = item.Empresa;
            this.Estatus = item.Estatus;
            this.Fecha = item.Fecha;
            this.Folio = item.Folio;
            this.Pago = item.Pago;
            this.FolioSust = item.FolioSust;
            this.ImporteNeto = item.ImporteNeto;
            this.Total = item.Total;
            this.TipoDocumento = item.TipoDocumento;
            this.MetodoPago = item.MetodoPago;
            this.Moneda = item.Moneda;
            this.Neto = item.Neto;
            this.NumOtorgante = item.NumOtorgante;
            this.NumSAP = item.NumSAP;
            this.RazonSocial = item.RazonSocial;
            this.RazonSocialEmpresa = item.RazonSocialEmpresa;
            //this.UUID = item.UUID;
            this.Comensado = item.Comensado;
        }

        public MetodoIdentificacion EstatusFactura { get; internal set; }
        public DateTime Fecha { get; internal set; }
        public string RazonSocial { get; internal set; }
        public string Folio { get; internal set; }
        public string FolioSust { get; internal set; }
        //public string UUID { get; internal set; }
        public string RFC { get; internal set; }
        public decimal Diff { get; internal set; }
        public decimal Total { get; internal set; }
        public string Empresa { get; internal set; }
        public int Estatus { get; internal set; }
        public decimal ImporteNeto { get; internal set; }
        public string Moneda { get; internal set; }
        public decimal Neto { get; }
        public decimal IVA { get; internal set; }
        public string Refactura { get; }
        public string Refacturacion { get; set; }
        public decimal Descuento { get; internal set; }
        public string NumSAP { get; internal set; }
        public string NumOtorgante { get; internal set; }
        public string Referencias { get; internal set; }
        public string MetodoPago { get; internal set; }
        public decimal Pago { get; internal set; }
        public string Comensado { get; internal set; }
        public string NumDocumento { get; internal set; }
        public string TipoDocumento { get; internal set; }

        public double FolioSustNumerico
        {
            get
            {
                if (!string.IsNullOrEmpty(FolioSust))
                {
                    var folio = FolioSust.Replace("INVC", "").Replace("INVL", "").Replace("INVI", "");
                    if (double.TryParse(folio, out double result))
                        return result;
                }
                return 0;
            }
        }
        public double FolioNumerico
        {
            get
            {
                var folio = Folio.Replace("INVC", "").Replace("INVL", "").Replace("INVI", "");
                if (double.TryParse(folio, out double result))
                    return result;
                return 0;
            }
        }

        public object RazonSocialEmpresa { get; }
        public string NRazonSocial { get { return UtilityIO.NormalizaNombre(RazonSocial); } }

        public string Observaciones { get; internal set; }
        public OpenPayLiquidacion OpenPayLiquidacion { get; internal set; }
        public OpenPay OpenPay { get; internal set; }
        public int Source { get; internal set; }
        public int ConteoMetodo
        {
            get
            {
                switch (EstatusFactura)
                {
                    case MetodoIdentificacion.NoIdentificado:
                    case MetodoIdentificacion.RazonSocial:
                    case MetodoIdentificacion.SumaRazonSocialFacturas:
                    case MetodoIdentificacion.FolioFactura:
                    case MetodoIdentificacion.SumaFacturas:
                    case MetodoIdentificacion.FacturaMesCorriente:
                    case MetodoIdentificacion.ImporteAprox:
                    case MetodoIdentificacion.Compensada:
                    case MetodoIdentificacion.Posible:
                        return (int)EstatusFactura;
                }
                return 1;
            }
        }

        public string FolioFiscal { get; internal set; }

        public override string ToString()
        {
            return string.Join(",", "Empresa:" + Empresa, "Fecha:" + Fecha, "Folio" + Folio, "RFC:" + RFC, "Razon:" + RazonSocial, "ImporteNeto", ImporteNeto, "Total:" + Total, "Pago:" + Pago, "DIFF:" + Diff, "Estatus:" + EstatusFactura, "Tipo:" + TipoDocumento);
        }
    }
}
