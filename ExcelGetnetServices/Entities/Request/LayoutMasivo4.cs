using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGetnetServices.Entities.Request
{
    public class LayoutMasivo4
    {
        public string fec_ini { get; set; }
        public string fec_fin { get; set; }
        public string id_proveedor { get; set; }
        public string status_servicio { get; set; }
        public string id_zona { get; set; }
        public int id_proyecto { get; set; }
        public int id_usuario { get; set; }
        public int serie { get; set; }
        public string fec_ini_cierre { get; set; }
        public string fec_fin_cierre { get; set; }
        public string id_servicio { get; set; }
        public string id_falla { get; set; }
    }
}
