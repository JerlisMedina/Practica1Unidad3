using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pract1Unidad3.Models
{
    public class Empleado
    {
        public int EmpleadoID { get; set; }
        public string Nombres { get; set; }
        public string Apellidos { get; set; }
        public Nullable<DateTime> Fecha_Ingreso { get; set; }
        public ICollection<Registro> Registros { get; set; }
    }
}
