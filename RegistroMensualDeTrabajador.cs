using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeIntranetARex
{
    internal class RegistroMensualDeTrabajador
    {

        private String empleado;
        private String nombre;
        private String apellidoPate;
        private String apellidoMate;
        private String fechaNaci;
        private String nombre_empresa;
        private String nombre_cargo;
        private String nombre_centro_costo;
        private String proceso;
        private String imponible_sin_tope;
        private String total_exento;

        public RegistroMensualDeTrabajador()
        {
        }

        public RegistroMensualDeTrabajador(string empleado, string nombre, string apellidoPate, string apellidoMate, string fechaNaci, string nombre_empresa, string nombre_cargo, string nombre_centro_costo, string proceso, string imponible_sin_tope, string total_exento)
        {
            this.Empleado = empleado;
            this.Nombre = nombre;
            this.ApellidoPate = apellidoPate;
            this.ApellidoMate = apellidoMate;
            this.FechaNaci = fechaNaci;
            this.Nombre_empresa = nombre_empresa;
            this.Nombre_cargo = nombre_cargo;
            this.Nombre_centro_costo = nombre_centro_costo;
            this.Proceso = proceso;
            this.Imponible_sin_tope = imponible_sin_tope;
            this.Total_exento = total_exento;
        }

        public string Empleado { get => empleado; set => empleado = value; }
        public string Nombre { get => nombre; set => nombre = value; }
        public string ApellidoPate { get => apellidoPate; set => apellidoPate = value; }
        public string ApellidoMate { get => apellidoMate; set => apellidoMate = value; }
        public string FechaNaci { get => fechaNaci; set => fechaNaci = value; }
        public string Nombre_empresa { get => nombre_empresa; set => nombre_empresa = value; }
        public string Nombre_cargo { get => nombre_cargo; set => nombre_cargo = value; }
        public string Nombre_centro_costo { get => nombre_centro_costo; set => nombre_centro_costo = value; }
        public string Proceso { get => proceso; set => proceso = value; }
        public string Imponible_sin_tope { get => imponible_sin_tope; set => imponible_sin_tope = value; }
        public string Total_exento { get => total_exento; set => total_exento = value; }
    }
}
