using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{
    internal class PolizaDetalle
    {
        /// <summary>
        /// Centro de costos
        /// #4
        /// </summary>
        public string CentroCostos { get; internal set; }

        /// <summary>
        /// Importe Neto
        /// #5
        /// </summary>
        public decimal ImporteNeto { get; internal set; }
        /// <summary>
        /// Importe IVA
        /// #6
        /// </summary>
        public string CodigoImpuesto { get; internal set; }
        /// <summary>
        /// Impuesto
        /// #7
        /// </summary>
        public decimal Impuesto { get; internal set; }
        /// <summary>
        /// SOcio Comercial
        /// #8
        /// </summary>
        public string SocioComercial { get; internal set; }

        /// <summary>
        /// Area Funcional
        /// #10
        /// </summary>
        public string AreaFuncional { get; internal set; }
        public TipoDetalle TipoMovimiento { get; internal set; }
        public string  Descripcion { get; internal set; }
    }
    public enum TipoDetalle
    {
        Abono = 1,
        Cargo = 2
    }
}