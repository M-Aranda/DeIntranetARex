using System;
using System.Collections.Generic;
using System.Text;

namespace DeIntranetARex
{
    class Ausencia
    {
        private String empleado;
        private String contratos;
        private String tipo;
        private String fechaInicio;
        private String fechaTermino;
        private String diasDeAusencia;
        private String descripcion;
        private String medioDia;
        private String enviaMailSupervisor;
        private String numeroDeLicencia;
        private String diasAPagar;
        private String noRebaja;
        private String fechaDeCalculo;
        private String fechaDeAplicacion;
        private String goceSueldo;
        private String tipoDePermiso;
        private String nombreDeEmpleado;

        //el archivo de ayudantes entrante tiene 17 columnas, pero el de conductores 16
        //para los efectos subir fallas en masa a manager, solo necesito el rut, el dia y el tipo de ausencia (falla o permiso)

        public Ausencia(string empleado, string contratos, string tipo, string fechaInicio, string fechaTermino, string diasDeAusencia, string descripcion, string medioDia, string enviaMailSupervisor, string numeroDeLicencia, string diasAPagar, string noRebaja, string fechaDeCalculo, string fechaDeAplicacion, string goceSueldo, string tipoDePermiso, string nombreDeEmpleado)
        {
            this.Empleado = empleado;
            this.Contratos = contratos;
            this.Tipo = tipo;
            this.FechaInicio = fechaInicio;
            this.FechaTermino = fechaTermino;
            this.DiasDeAusencia = diasDeAusencia;
            this.Descripcion = descripcion;
            this.MedioDia = medioDia;
            this.EnviaMailSupervisor = enviaMailSupervisor;
            this.NumeroDeLicencia = numeroDeLicencia;
            this.DiasAPagar = diasAPagar;
            this.NoRebaja = noRebaja;
            this.FechaDeCalculo = fechaDeCalculo;
            this.FechaDeAplicacion = fechaDeAplicacion;
            this.GoceSueldo = goceSueldo;
            this.TipoDePermiso = tipoDePermiso;
            this.nombreDeEmpleado = nombreDeEmpleado;
        }

        public Ausencia()
        {

        }

        public string Empleado { get => empleado; set => empleado = value; }
        public string Contratos { get => contratos; set => contratos = value; }
        public string Tipo { get => tipo; set => tipo = value; }
        public string FechaInicio { get => fechaInicio; set => fechaInicio = value; }
        public string FechaTermino { get => fechaTermino; set => fechaTermino = value; }
        public string DiasDeAusencia { get => diasDeAusencia; set => diasDeAusencia = value; }
        public string Descripcion { get => descripcion; set => descripcion = value; }
        public string MedioDia { get => medioDia; set => medioDia = value; }
        public string EnviaMailSupervisor { get => enviaMailSupervisor; set => enviaMailSupervisor = value; }
        public string NumeroDeLicencia { get => numeroDeLicencia; set => numeroDeLicencia = value; }
        public string DiasAPagar { get => diasAPagar; set => diasAPagar = value; }
        public string NoRebaja { get => noRebaja; set => noRebaja = value; }
        public string FechaDeCalculo { get => fechaDeCalculo; set => fechaDeCalculo = value; }
        public string FechaDeAplicacion { get => fechaDeAplicacion; set => fechaDeAplicacion = value; }
        public string GoceSueldo { get => goceSueldo; set => goceSueldo = value; }
        public string TipoDePermiso { get => tipoDePermiso; set => tipoDePermiso = value; }

        public string NombreDeEmpleado { get => nombreDeEmpleado; set => nombreDeEmpleado = value; }




    }
}