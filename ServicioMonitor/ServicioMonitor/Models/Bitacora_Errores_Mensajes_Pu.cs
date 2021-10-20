

using System;

namespace ServicioMonitor.Models
{
    public class Bitacora_Errores_Mensajes_Pu
    {
        public int consecutivo { get; set; }
        public DateTime fecha_hora { get; set; }
        public decimal error_numero { get; set; }
        public string error_descripcion { get; set; }
        public string aplicacion { get; set; }
    }
}
