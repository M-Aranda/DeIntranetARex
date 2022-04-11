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
        private String empleado;
        private String id;
        private String contrato;
        private int monto;


        public MontoPorConcepto(string concepto, string empleado, string id, string contrato, int monto)
        {
            this.Concepto = concepto;
            this.Empleado = empleado;
            this.Id = id;
            this.Contrato = contrato;
            this.Monto = monto;
        }

        public MontoPorConcepto()
        {
        }

        public string Concepto { get => concepto; set => concepto = value; }
        public string Empleado { get => empleado; set => empleado = value; }
        public string Id { get => id; set => id = value; }
        public string Contrato { get => contrato; set => contrato = value; }
        public int Monto { get => monto; set => monto = value; }



    }
}
