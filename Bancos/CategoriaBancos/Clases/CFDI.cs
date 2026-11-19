using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{

    public class CFDI
    {
        //Clave producto	Fecha de pago	Forma de pago	Moneda	Tipo de cambio	Monto	Número de folio de operación	Movimiento en cuenta	RFC Banco ordenante	Nombre banco ordenante	RFC ordenante	Nombre ordenante	Cuenta ordenante	RFC Banco beneficiario	Cuenta beneficiaria	Tipo de cadena de pago	Certificado del pago	Cadena original	Sello de pago	Folio de operación	Cuenta
        public string CuentaOrdenante { get; internal set; }
        public string RFCOrdenante { get; internal set; }
        public string Cuenta { get; internal set; }

        public override string ToString()
        {
            return string.Join(",", "CuentaOrdenante", CuentaOrdenante, "Cuenta", Cuenta, "RFCOrdenante", RFCOrdenante);
        }

    }
}
