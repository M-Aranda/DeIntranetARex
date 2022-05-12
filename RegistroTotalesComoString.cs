using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class RegistroTotalesComoString
    {

        private String proceso;
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
        private int totalR;
        private int totalBonoTiempoEsperaR;
        private int totalBonoEstacionalR;
        private int totalBtnLR;

        private int totalBonoSobreEsfuerzoR;
        private int totalViaticoAhorroR;

        private int totalRemuneracionesDeTodosLosTrabajadores;



        private String cantidadDeConductoresActivosComoString;
        private String cantidadDeConductoresDeLicenciaComoString;
        private String cantidadDeAyudantesActivosComoString;
        private String cantidadDeAyudantesDeLicenciaComoString;
        private String cantidadDeApoyosActivosComoString;
        private String cantidadDeApoyosDeLicenciaComoString;
        private String totalConductoresComoString;
        private String totalAyudantesComoString;
        private String totalApoyosComoString;
        private String totalDotacionComoString;// o sea todos los trabajadores
        private String totalRemuneracionesConductoresComoString;
        private String totalRemuneracionesAyudantesComoString;
        private String totalRemuneracionesOtrosComoString;
        private String totalBonoTiempoEsperaRComoString;
        private String totalBonoEstacionalRComoString;
        private String totalBtnLRComoString;
        private String totalRComoString;

        private String totalBonoSobreEsfuerzoRComoString;
        private String totalViaticoAhorroRComoString;

        private String totalRemuneracionesDeTodosLosTrabajadoresComoString;


        public RegistroTotalesComoString()
        {
        }

        public RegistroTotalesComoString(RegistroDeTotales r)
        {


            this.Centro = r.Centro;
            this.CantidadDeConductoresActivos = r.CantidadDeConductoresActivos;//.ToString();
            this.CantidadDeConductoresDeLicencia = r.CantidadDeConductoresDeLicencia;//.ToString();
            this.CantidadDeAyudantesActivos = r.CantidadDeAyudantesActivos;//.ToString();
            this.CantidadDeAyudantesDeLicencia = r.CantidadDeAyudantesDeLicencia;//.ToString();
            this.CantidadDeApoyosActivos = r.CantidadDeApoyosActivos;//.ToString();
            this.CantidadDeApoyosDeLicencia = r.CantidadDeApoyosDeLicencia;//.ToString();
            this.TotalConductores = r.TotalConductores;//.ToString();
            this.TotalAyudantes = r.TotalAyudantes;//.ToString();
            this.TotalApoyos = r.TotalApoyos;//.ToString();
            this.TotalDotacion = r.TotalDotacion;//.ToString();
            this.TotalRemuneracionesConductores = r.TotalRemuneracionesConductores;//.ToString();
            this.TotalRemuneracionesAyudantes = r.TotalRemuneracionesAyudantes;//.ToString();
            this.TotalRemuneracionesOtros = r.TotalRemuneracionesOtros;//.ToString();
            this.TotalBonoTiempoEsperaR = r.TotalBonoTiempoEsperaR;
            this.TotalBonoEstacionalR = r.TotalBonoEstacionalR;
            this.TotalBtnLR = r.TotalBtnLR;

            this.totalBonoSobreEsfuerzoR = r.TotalBonoSobreEsfuerzoR;
            this.totalViaticoAhorroR = r.TotalViaticoAhorroR;

            this.TotalR = r.TotalR;
            this.TotalRemuneracionesDeTodosLosTrabajadores = r.TotalRemuneracionesDeTodosLosTrabajadores;//.ToString();
        }


        public RegistroTotalesComoString(RegistroDeTotales r, String vaParaElExcel)
        {

            this.proceso = "";
            this.Centro = r.Centro;
            this.cantidadDeConductoresActivosComoString = "";// r.CantidadDeConductoresActivos.ToString();
            this.cantidadDeConductoresDeLicenciaComoString = r.CantidadDeConductoresDeLicencia.ToString();
            this.cantidadDeAyudantesActivosComoString = r.CantidadDeAyudantesActivos.ToString();
            this.cantidadDeAyudantesDeLicenciaComoString = r.CantidadDeAyudantesDeLicencia.ToString();
            this.cantidadDeApoyosActivosComoString = r.CantidadDeApoyosActivos.ToString();
            this.cantidadDeApoyosDeLicenciaComoString = r.CantidadDeApoyosDeLicencia.ToString();
            this.totalConductoresComoString = r.TotalConductores.ToString();
            this.totalAyudantesComoString = r.TotalAyudantes.ToString();
            this.totalApoyosComoString = r.TotalApoyos.ToString();
            this.totalDotacionComoString = r.TotalDotacion.ToString();
            this.totalRemuneracionesConductoresComoString = r.TotalRemuneracionesConductores.ToString();
            this.totalRemuneracionesAyudantesComoString = r.TotalRemuneracionesAyudantes.ToString();
            this.totalRemuneracionesOtrosComoString = r.TotalRemuneracionesOtros.ToString();
            this.TotalBonoTiempoEsperaRComoString = r.TotalBonoTiempoEsperaR.ToString();
            this.TotalBonoEstacionalRComoString = r.TotalBonoEstacionalR.ToString();
            this.TotalBtnLRComoString = r.TotalBtnLR.ToString();
            this.totalRComoString = r.TotalR.ToString();
            this.TotalBonoSobreEsfuerzoRComoString= r.TotalBonoSobreEsfuerzoR.ToString();
            this.TotalViaticoAhorroRComoString = r.TotalViaticoAhorroR.ToString(); 


            this.totalRemuneracionesDeTodosLosTrabajadoresComoString = r.TotalRemuneracionesDeTodosLosTrabajadores.ToString();
        }

        public RegistroTotalesComoString(String procesoEnLugarDeFecha)        
        {

            this.Centro = procesoEnLugarDeFecha;

        }

        public RegistroTotalesComoString(string centro, int cantidadDeConductoresActivos, int cantidadDeConductoresDeLicencia, int cantidadDeAyudantesActivos, int cantidadDeAyudantesDeLicencia, int cantidadDeApoyosActivos, int cantidadDeApoyosDeLicencia, int totalConductores, int totalAyudantes, int totalApoyos, int totalDotacion, int totalRemuneracionesConductores, int totalRemuneracionesAyudantes, int totalRemuneracionesOtros, int totalRemuneracionesDeTodosLosTrabajadores)
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

        public int TotalR { get => totalR; set => totalR = value; }
        public string TotalRComoString { get => totalRComoString; set => totalRComoString = value; }


        public int TotalBonoTiempoEsperaR { get => totalBonoTiempoEsperaR; set => totalBonoTiempoEsperaR = value; }

        public string TotalBonoTiempoEsperaRComoString { get => totalBonoTiempoEsperaRComoString; set => totalBonoTiempoEsperaRComoString = value; }

        public int TotalBonoEstacionalR { get => totalBonoEstacionalR; set => totalBonoEstacionalR = value; }

        public string TotalBonoEstacionalRComoString { get => totalBonoEstacionalRComoString; set => totalBonoEstacionalRComoString = value; }

        public int TotalBtnLR { get => totalBtnLR; set => totalBtnLR = value; }

        public string TotalBtnLRComoString { get => totalBtnLRComoString; set => totalBtnLRComoString = value; }
        public int TotalBonoSobreEsfuerzoR { get => totalBonoSobreEsfuerzoR; set => totalBonoSobreEsfuerzoR = value; }
        public int TotalViaticoAhorroR { get => totalViaticoAhorroR; set => totalViaticoAhorroR = value; }
        public string TotalBonoSobreEsfuerzoRComoString { get => totalBonoSobreEsfuerzoRComoString; set => totalBonoSobreEsfuerzoRComoString = value; }
        public string TotalViaticoAhorroRComoString { get => totalViaticoAhorroRComoString; set => totalViaticoAhorroRComoString = value; }
    }
}
