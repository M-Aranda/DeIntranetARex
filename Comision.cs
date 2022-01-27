using System;
using System.Collections.Generic;
using System.Text;

namespace DeIntranetARex
{
    class Comision
    {

        //ayudantes tiene 26 columnas
        //conductores tiene 23


        private String plantilla;//el rut
        private String contrato;
        private String concepto;
        private String valor;
        private String origen;
        private String objeto;
        private String periodoDePago;
        private String fechaDeInicio;
        private String fechaDeTermino;
        private String institucion;
        private String datoAdicional;
        private String comentario;
        private String valorPorDefecto;
        private String accion;



        public Comision(string plantilla, string contrato, string concepto, string valor, string origen, string objeto, string periodoDePago, string fechaDeInicio, string fechaDeTermino, string institucion, string datoAdicional, string comentario, string valorPorDefecto, string accion)
        {
            this.Plantilla = plantilla;
            this.Contrato = contrato;
            this.Concepto = concepto;
            this.Valor = valor;
            this.Origen = origen;
            this.Objeto = objeto;
            this.PeriodoDePago = periodoDePago;
            this.FechaDeInicio = fechaDeInicio;
            this.FechaDeTermino = fechaDeTermino;
            this.Institucion = institucion;
            this.DatoAdicional = datoAdicional;
            this.Comentario = comentario;
            this.ValorPorDefecto = valorPorDefecto;
            this.Accion = accion;
        }

        public Comision()
        {
            
        }

        public string Plantilla { get => plantilla; set => plantilla = value; }
        public string Contrato { get => contrato; set => contrato = value; }
        public string Concepto { get => concepto; set => concepto = value; }
        public string Valor { get => valor; set => valor = value; }
        public string Origen { get => origen; set => origen = value; }
        public string Objeto { get => objeto; set => objeto = value; }
        public string PeriodoDePago { get => periodoDePago; set => periodoDePago = value; }
        public string FechaDeInicio { get => fechaDeInicio; set => fechaDeInicio = value; }
        public string FechaDeTermino { get => fechaDeTermino; set => fechaDeTermino = value; }
        public string Institucion { get => institucion; set => institucion = value; }
        public string DatoAdicional { get => datoAdicional; set => datoAdicional = value; }
        public string Comentario { get => comentario; set => comentario = value; }
        public string ValorPorDefecto { get => valorPorDefecto; set => valorPorDefecto = value; }
        public string Accion { get => accion; set => accion = value; }
    }
}
