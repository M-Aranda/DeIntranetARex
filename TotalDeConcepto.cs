using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class TotalDeConcepto
    {
        private String nombreDeConcepto;
        private String centroDeConcepto;
        private int montoTotalDeConcepto;

        public TotalDeConcepto()
        {
        }

        public TotalDeConcepto(string nombreDeConcepto, string centroDeConcepto, int montoTotalDeConcepto)
        {
            this.NombreDeConcepto = nombreDeConcepto;
            this.CentroDeConcepto = centroDeConcepto;
            this.MontoTotalDeConcepto = montoTotalDeConcepto;
        }

        public string NombreDeConcepto { get => nombreDeConcepto; set => nombreDeConcepto = value; }
        public string CentroDeConcepto { get => centroDeConcepto; set => centroDeConcepto = value; }
        public int MontoTotalDeConcepto { get => montoTotalDeConcepto; set => montoTotalDeConcepto = value; }
    }
}
