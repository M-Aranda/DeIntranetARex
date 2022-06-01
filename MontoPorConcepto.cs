using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class MontoPorConcepto
    {
        private String empleado;//rut
        private String nombre;
        private String apellidoPaterno;
        private String apellidoMaterno;
        private String empresa;
        private String cargo;
        private String centroCosto;
        private String fechaProceso;
        private String concepto;
        private int monto;


        public MontoPorConcepto()
        {
        }

        public MontoPorConcepto(string empleado, string nombre, string apellidoPaterno, string apellidoMaterno, string empresa, string cargo, string centroCosto, string fechaProceso, string concepto, int monto)
        {
            this.Empleado = empleado;
            this.Nombre = nombre;
            this.ApellidoPaterno = apellidoPaterno;
            this.ApellidoMaterno = apellidoMaterno;
            this.Empresa = empresa;
            this.Cargo = cargo;
            this.CentroCosto = centroCosto;
            this.FechaProceso = fechaProceso;
            this.Concepto = concepto;
            this.Monto = monto;
        }

        public string Empleado { get => empleado; set => empleado = value; }
        public string Nombre { get => nombre; set => nombre = value; }
        public string ApellidoPaterno { get => apellidoPaterno; set => apellidoPaterno = value; }
        public string ApellidoMaterno { get => apellidoMaterno; set => apellidoMaterno = value; }
        public string Empresa { get => empresa; set => empresa = value; }
        public string Cargo { get => cargo; set => cargo = value; }
        public string CentroCosto { get => centroCosto; set => centroCosto = value; }
        public string FechaProceso { get => fechaProceso; set => fechaProceso = value; }
        public string Concepto { get => concepto; set => concepto = value; }
        public int Monto { get => monto; set => monto = value; }
    }
}
