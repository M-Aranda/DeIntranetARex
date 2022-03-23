using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class RegistroDeTotales
    {

        private String centro;
        private int cantidadDeConductoresActivos;
        private int cantidadDeConductoresDeLicencia;
        private int cantidadDeAyudantesActivos;
        private int cantidadDeAyudantesDeLicencia;
        private int cantidadDeApoyosActivos;
        private int cantidadDeApoyosDeLicencia;
        private int totalConductores;
        private int totalAyudantes;
        private int totalApoyos;
        private int totalDotacion;// o sea todos los trabajadores
        private int totalRemuneracionesConductores;
        private int totalRemuneracionesAyudantes;
        private int totalRemuneracionesOtros;
        private int totalRemuneracionesDeTodosLosTrabajadores;

        public RegistroDeTotales()
        {
           
        }

        public RegistroDeTotales(string centro, int cantidadDeConductoresActivos, int cantidadDeConductoresDeLicencia, int cantidadDeAyudantesActivos, int cantidadDeAyudantesDeLicencia, int cantidadDeApoyosActivos, int cantidadDeApoyosDeLicencia, int totalConductores, int totalAyudantes, int totalApoyos, int totalDotacion, int totalRemuneracionesConductores, int totalRemuneracionesAyudantes, int totalRemuneracionesOtros, int totalRemuneracionesDeTodosLosTrabajadores)
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
        public int CantidadDeConductoresActivos { get => cantidadDeConductoresActivos; set => cantidadDeConductoresActivos = value; }
        public int CantidadDeConductoresDeLicencia { get => cantidadDeConductoresDeLicencia; set => cantidadDeConductoresDeLicencia = value; }
        public int CantidadDeAyudantesActivos { get => cantidadDeAyudantesActivos; set => cantidadDeAyudantesActivos = value; }
        public int CantidadDeAyudantesDeLicencia { get => cantidadDeAyudantesDeLicencia; set => cantidadDeAyudantesDeLicencia = value; }
        public int CantidadDeApoyosActivos { get => cantidadDeApoyosActivos; set => cantidadDeApoyosActivos = value; }
        public int CantidadDeApoyosDeLicencia { get => cantidadDeApoyosDeLicencia; set => cantidadDeApoyosDeLicencia = value; }
        public int TotalConductores { get => totalConductores; set => totalConductores = value; }
        public int TotalAyudantes { get => totalAyudantes; set => totalAyudantes = value; }
        public int TotalApoyos { get => totalApoyos; set => totalApoyos = value; }
        public int TotalDotacion { get => totalDotacion; set => totalDotacion = value; }
        public int TotalRemuneracionesConductores { get => totalRemuneracionesConductores; set => totalRemuneracionesConductores = value; }
        public int TotalRemuneracionesAyudantes { get => totalRemuneracionesAyudantes; set => totalRemuneracionesAyudantes = value; }
        public int TotalRemuneracionesOtros { get => totalRemuneracionesOtros; set => totalRemuneracionesOtros = value; }
        public int TotalRemuneracionesDeTodosLosTrabajadores { get => totalRemuneracionesDeTodosLosTrabajadores; set => totalRemuneracionesDeTodosLosTrabajadores = value; }
    }
}
