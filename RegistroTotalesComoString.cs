using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class RegistroTotalesComoString
    {

        private String centro;
        private String cantidadDeConductoresActivos;
        private String cantidadDeConductoresDeLicencia;
        private String cantidadDeAyudantesActivos;
        private String cantidadDeAyudantesDeLicencia;
        private String cantidadDeApoyosActivos;
        private String cantidadDeApoyosDeLicencia;
        private String totalConductores;
        private String totalAyudantes;
        private String totalApoyos;
        private String totalDotacion;// o sea todos los trabajadores
        private String totalRemuneracionesConductores;
        private String totalRemuneracionesAyudantes;
        private String totalRemuneracionesOtros;
        private String totalRemuneracionesDeTodosLosTrabajadores;

        public RegistroTotalesComoString()
        {
        }

        public RegistroTotalesComoString(RegistroDeTotales r)
        {
            this.Centro = r.Centro;
            this.CantidadDeConductoresActivos = r.CantidadDeConductoresActivos.ToString();
            this.CantidadDeConductoresDeLicencia = r.CantidadDeConductoresDeLicencia.ToString();
            this.CantidadDeAyudantesActivos = r.CantidadDeAyudantesActivos.ToString();
            this.CantidadDeAyudantesDeLicencia = r.CantidadDeAyudantesDeLicencia.ToString();
            this.CantidadDeApoyosActivos = r.CantidadDeApoyosActivos.ToString();
            this.CantidadDeApoyosDeLicencia = r.CantidadDeApoyosDeLicencia.ToString();
            this.TotalConductores = r.TotalConductores.ToString();
            this.TotalAyudantes = r.TotalAyudantes.ToString();
            this.TotalApoyos = r.TotalApoyos.ToString();
            this.TotalDotacion = r.TotalDotacion.ToString();
            this.TotalRemuneracionesConductores = r.TotalRemuneracionesConductores.ToString();
            this.TotalRemuneracionesAyudantes = r.TotalRemuneracionesAyudantes.ToString();
            this.TotalRemuneracionesOtros = r.TotalRemuneracionesOtros.ToString();
            this.TotalRemuneracionesDeTodosLosTrabajadores = r.TotalRemuneracionesDeTodosLosTrabajadores.ToString();
        }

        public RegistroTotalesComoString(string centro, string cantidadDeConductoresActivos, string cantidadDeConductoresDeLicencia, string cantidadDeAyudantesActivos, string cantidadDeAyudantesDeLicencia, string cantidadDeApoyosActivos, string cantidadDeApoyosDeLicencia, string totalConductores, string totalAyudantes, string totalApoyos, string totalDotacion, string totalRemuneracionesConductores, string totalRemuneracionesAyudantes, string totalRemuneracionesOtros, string totalRemuneracionesDeTodosLosTrabajadores)
        {
            this.Centro = centro;
            this.CantidadDeConductoresActivos = cantidadDeConductoresActivos;
            this.CantidadDeConductoresDeLicencia = cantidadDeConductoresDeLicencia;
            this.CantidadDeAyudantesActivos = cantidadDeAyudantesActivos;
            this.CantidadDeAyudantesDeLicencia = cantidadDeAyudantesDeLicencia;
            this.CantidadDeApoyosActivos = cantidadDeApoyosActivos;
            this.CantidadDeApoyosDeLicencia = cantidadDeApoyosDeLicencia;
            this.TotalConductores = totalConductores;
            this.TotalAyudantes = totalAyudantes;
            this.TotalApoyos = totalApoyos;
            this.TotalDotacion = totalDotacion;
            this.TotalRemuneracionesConductores = totalRemuneracionesConductores;
            this.TotalRemuneracionesAyudantes = totalRemuneracionesAyudantes;
            this.TotalRemuneracionesOtros = totalRemuneracionesOtros;
            this.TotalRemuneracionesDeTodosLosTrabajadores = totalRemuneracionesDeTodosLosTrabajadores;
        }

        public string Centro { get => centro; set => centro = value; }
        public string CantidadDeConductoresActivos { get => cantidadDeConductoresActivos; set => cantidadDeConductoresActivos = value; }
        public string CantidadDeConductoresDeLicencia { get => cantidadDeConductoresDeLicencia; set => cantidadDeConductoresDeLicencia = value; }
        public string CantidadDeAyudantesActivos { get => cantidadDeAyudantesActivos; set => cantidadDeAyudantesActivos = value; }
        public string CantidadDeAyudantesDeLicencia { get => cantidadDeAyudantesDeLicencia; set => cantidadDeAyudantesDeLicencia = value; }
        public string CantidadDeApoyosActivos { get => cantidadDeApoyosActivos; set => cantidadDeApoyosActivos = value; }
        public string CantidadDeApoyosDeLicencia { get => cantidadDeApoyosDeLicencia; set => cantidadDeApoyosDeLicencia = value; }
        public string TotalConductores { get => totalConductores; set => totalConductores = value; }
        public string TotalAyudantes { get => totalAyudantes; set => totalAyudantes = value; }
        public string TotalApoyos { get => totalApoyos; set => totalApoyos = value; }
        public string TotalDotacion { get => totalDotacion; set => totalDotacion = value; }
        public string TotalRemuneracionesConductores { get => totalRemuneracionesConductores; set => totalRemuneracionesConductores = value; }
        public string TotalRemuneracionesAyudantes { get => totalRemuneracionesAyudantes; set => totalRemuneracionesAyudantes = value; }
        public string TotalRemuneracionesOtros { get => totalRemuneracionesOtros; set => totalRemuneracionesOtros = value; }
        public string TotalRemuneracionesDeTodosLosTrabajadores { get => totalRemuneracionesDeTodosLosTrabajadores; set => totalRemuneracionesDeTodosLosTrabajadores = value; }
    }
}
