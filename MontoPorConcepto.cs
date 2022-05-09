using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class MontoPorConcepto
    {
        private String concepto;
        private String nombreDeEmpleado;
        private String rut;
        private String mesDeProceso;
        private int monto;

        public MontoPorConcepto()
        {
        }

        public MontoPorConcepto(string concepto, string nombreDeEmpleado, string rut, string mesDeProceso, int monto)
        {
            this.Concepto = concepto;
            this.NombreDeEmpleado = nombreDeEmpleado;
            this.Rut = rut;
            this.MesDeProceso = mesDeProceso;
            this.Monto = monto;
        }

        public string Concepto { get => concepto; set => concepto = value; }
        public string NombreDeEmpleado { get => nombreDeEmpleado; set => nombreDeEmpleado = value; }
        public string Rut { get => rut; set => rut = value; }
        public string MesDeProceso { get => mesDeProceso; set => mesDeProceso = value; }
        public int Monto { get => monto; set => monto = value; }
    }
}
