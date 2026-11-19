using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public class UtilityDuration
    {

        public DateTime FechaEjecucion { get; set; }
        public int DuracionHoras { get; set; }
        public int DuracionMinutos { get; set; }
        public int DuracionSegundos { get; set; }

        public UtilityDuration()
        {
            this.FechaEjecucion = DateTime.Now;
        }
        public string ObtieneDuracion()
        {
            //DURACION QUERY
            this.DuracionHoras = (DateTime.Now - this.FechaEjecucion).Hours;
            this.DuracionMinutos = (DateTime.Now - this.FechaEjecucion).Minutes;
            this.DuracionSegundos = (DateTime.Now - this.FechaEjecucion).Seconds;

            string Duracion = this.DuracionHoras > 0 ? this.DuracionHoras + " Horas " : "";
            Duracion += this.DuracionMinutos > 0 ? this.DuracionMinutos + " Minutos " : "";
            Duracion += this.DuracionSegundos + " Segundos";
            return Duracion;
        }
    }
}
