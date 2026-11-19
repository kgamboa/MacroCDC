using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{

    public class BBVA
    {
        //Cuenta beneficiaria	Titular cuenta	Número banco receptor	Cuenta ordenante	Cuenta CLABE ordenante	Divisa	Titular cuenta ordenante	RFC ordenante	Fecha de pago	Número de movimiento	Código de leyenda	Concepto	Debe-Haber	Importe	Efectivo real	Referencia númerica	Referencia ampliada	Forma de pago

        public string ReferenciaAmpliada { get; internal set; }
        public string Otorgante { get; internal set; }
        public decimal ImporteNeto { get; internal set; }
        public string RFC { get; internal set; }
        public string CuentaOrdenante { get; internal set; }
        public string CuentaClabeOrdenante { get; internal set; }
        public string CodigoLeyenda { get; internal set; }
        public string Referencia { get; internal set; }
        public string Cuenta { get; internal set; }
    }
}
