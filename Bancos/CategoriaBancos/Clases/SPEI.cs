using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
namespace TesoreriaCobranza
{

    public class SPEI
    {
        //        Fecha Referencia numerica Concepto de codigo de leyenda   Referencia Concepto de pago    Importe Saldo   Banco ordenante Nombre ordenante    Cuenta ordenante    Banco beneficiario  Nombre beneficiario Cuenta beneficiario Clave de rastreo Estado del pago Motivo de devolucion
        public string Referencia { get; internal set; }
        public string Otorgante { get; internal set; }
        public dynamic ImporteNeto { get; internal set; }
        public string RFC { get; internal set; }
        public dynamic CuentaOrdenante { get; internal set; }
        public string Cuenta { get; internal set; }
    }
}
