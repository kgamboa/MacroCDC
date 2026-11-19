using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{
    internal class Poliza
    {
        public DateTime Fecha { get; internal set; }

        /// <summary>
        /// Tipo de Poliza
        /// Ejemplo: Egreso, GTOS, F-
        /// </summary>
        public string Facturacion { get; internal set; }
        public decimal ImporteNeto
        {

            get
            {
                if (Movimientos != null)
                {
                    return Movimientos.Sum(c => c.ImporteNeto);
                }
                return 0;
            }
        }
        public decimal Impuesto
        {

            get
            {
                if (Movimientos != null)
                {
                    return Movimientos.Sum(c => c.Impuesto);
                }
                return 0;
            }
        }
        public string SocioComercial { get; internal set; }
        public List<PolizaDetalle> Movimientos { get; internal set; }
        private string moneda;
        public string Moneda
        {
            get
            {
                if (moneda != null && moneda.Trim().Length == 3)
                    return moneda;
                return "MXN";
            }
            set
            {
                moneda = value;
            }
        }

        /// <summary>
        /// Tipo de Documento
        /// Ejemplo: SA, KP, SP, SC, ST, SF, SE
        /// </summary>
        public string TipoDocumento
        {
            get
            {
                if (string.IsNullOrEmpty(tipoDocumento))
                    tipoDocumento = "SA";
                return tipoDocumento;
            }
            set { tipoDocumento = value; }
        }
        private string tipoDocumento;

        /// <summary>
        /// Descripcion Poliza
        ///  
        /// Ejemplo: Ingresos Enero2025
        ///          COMISIONES MARZO 2025   
        ///          NOMINA
        /// </summary>
        public string Descripcion
        {
            get
            {
                var monthStr = this.Fecha.ToString("MMM").ToUpper().Replace(".","");
                return this.description.ToUpper() +" "+ monthStr + this.Fecha.ToString("yy"); //Descripcion Poliza
            }
            set
            {
                if (value != null)
                {
                    this.description = value;
                }
            }
        }
        private string description;

        public bool Ingreso { get; internal set; }
        private string ctasMayor;
        public string CtasMayor
        {
            get
            {
                if (string.IsNullOrEmpty(this.ctasMayor))
                    return "D";
                return ctasMayor;
            }
            set
            {
                ctasMayor = value;
            }
        }

        /// <summary>
        /// Sociedad
        /// #5
        /// </summary>
        public string Sociedad { get; internal set; }

        /// <summary>
        /// Referencia
        /// $8
        /// </summary>
        public string Referencia { get; internal set; }
        public string Catalog { get; internal set; }
        public dynamic DescripcionReferencia { get; internal set; }
    }
}
