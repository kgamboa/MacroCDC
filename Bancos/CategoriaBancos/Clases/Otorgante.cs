using Common;

namespace TesoreriaCobranza
{
    internal class Otorgante
    {
        public string RFC { get; internal set; }
        public string RazonSocial { get; internal set; }
        public double ClienteID { get; internal set; }
        public string Empresa { get; internal set; }
        public string NumSAP { get; internal set; }
        public bool MesVencido { get; internal set; }
        public string NRazonSocial { get { return UtilityIO.NormalizaNombre(RazonSocial); }}
    }
}