using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesoreriaCobranza
{

    public class RefBBVA
    {
        //public string CuentaOrdenante { get; internal set; }
        public string Ordenante { get; internal set; }
        public string ReferenciaBBVA { get; internal set; }
        public string NumOtorgante { get; internal set; }

        public override string ToString()
        {
            return string.Join(",", "NumOtorgante", NumOtorgante,"Ordenante", Ordenante, "ReferenciaBBVA", ReferenciaBBVA);
        }
    }
}
