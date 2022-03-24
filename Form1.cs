using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Data;
using System.Data.SqlClient;
using Windows.Storage;
namespace DeIntranetARex
{
    public partial class Form1 : Form

    {
        private bool esArchivoDeAyudantes = true;

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }

            List<Ausencia> ausencias = leerExcelDeFallos(sFileName);

            
            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";


            var archivo = new FileInfo(downloads + @"\Asistencias.xlsx");

            SaveExcelFileAusencia(ausencias, archivo);

            MessageBox.Show("Archivo Excel de asistencias creado en carpeta de descargas!");




        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }

 

            List<Comision> comisiones = leerExcelDeComisiones(sFileName);

            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";
            String tipoDeArchivo = "";
            if (esArchivoDeAyudantes == true)
            {
                tipoDeArchivo = "ayudantes";
            }else if(esArchivoDeAyudantes == false)
            {
                tipoDeArchivo = "conductores";
            }


            var archivo = new FileInfo(downloads + @"\Comisiones de "+tipoDeArchivo+".xlsx");

            SaveExcelFileComision(comisiones, archivo);

            MessageBox.Show("Archivo Excel de comisiones de "+ tipoDeArchivo + " creado en carpeta de descargas!");
        }






        private static async Task SaveExcelFileAusencia(List<Ausencia> asistencias, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Asistencias");

            var range = ws.Cells["A1"].LoadFromCollection(asistencias, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }



        private static async Task SaveExcelFileComision(List<Comision> comisiones, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Comisiones");

            var range = ws.Cells["A1"].LoadFromCollection(comisiones, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static async Task SaveExcelFileRegistroDeTotales(List<RegistroTotalesComoString> registrosDeTotales, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Registros de totales");

            var range = ws.Cells["A1"].LoadFromCollection(registrosDeTotales, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }


        private List<Ausencia> leerExcelDeFallos(string FilePath)
        {
            List<Ausencia> ausencias = new List<Ausencia>();     
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
               Ausencia encabezado = new Ausencia();
                encabezado.Empleado = "Empleado";
                encabezado.Contratos = "Contratos";
                encabezado.Tipo = "Tipo";
                encabezado.FechaInicio = "Fecha Inicio";
                encabezado.FechaTermino = "Fecha Término";
                encabezado.DiasDeAusencia = "Dias de ausencia";
                encabezado.Descripcion = "Descripción";
                encabezado.MedioDia = "Medio día";
                encabezado.EnviaMailSupervisor = "Envia mail supervisor";
                encabezado.NumeroDeLicencia = "Número de licencia";
                encabezado.DiasAPagar = "Días a pagar";
                encabezado.NoRebaja = "No rebaja";
                encabezado.FechaDeCalculo = "Fecha Cálculo";
                encabezado.FechaDeAplicacion = "Fecha Aplicación";
                encabezado.GoceSueldo = "Goce sueldo";
                encabezado.TipoDePermiso = "Tipo Permiso";
                encabezado.NombreDeEmpleado = "Nombre de empleado";

                ausencias.Add(encabezado);    


                for (int row = 1; row <= rowCount; row++)
                {

                    Ausencia a = new Ausencia();
                    a.Tipo = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    a.Empleado = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    a.FechaInicio = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    a.FechaInicio = alterarFormatoDeFecha(a.FechaInicio);

                    switch (a.Tipo)
                    {
                        case "Falla":
                            a.Descripcion = "Falla";
                            a.Tipo = "F";
                            a.TipoDePermiso = "";
                            break;
                        case "Permiso":
                            a.Descripcion = "Permiso";
                            a.Tipo = "P";
                            a.TipoDePermiso = "N";
                            break;
                        default:
                            a.Descripcion = "Falla";
                            a.Tipo = "F";
                            a.TipoDePermiso = "";
                            break;
                    }




                    a.FechaTermino = a.FechaInicio;
                    a.Contratos = "1";
                    
                    a.DiasDeAusencia = "1";
                    a.MedioDia = "";
                    a.EnviaMailSupervisor = "N";
                    a.NumeroDeLicencia = "";
                    a.DiasAPagar = "0";
                    a.NoRebaja = "N";
                    a.FechaDeCalculo = "";
                    a.FechaDeAplicacion = "";
                    a.GoceSueldo = "N";
                    a.NombreDeEmpleado = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    


                    for (int col = 1; col <= colCount; col++)
                    {
                       // Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                    }

                    if (a.Tipo != "Estado")
                    {
                        ausencias.Add(a);
                    }
                        
      
                    
                }
            }

            return ausencias;
        
        }




        private List<Comision>  leerExcelDeComisiones(string FilePath)
        {

            List<Comision> comisiones = new List<Comision>();
            List<Comision>comisionesTemporales = new List<Comision>();

            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                //si el titulo de la quinta columna es Cajas v1, entonces es un archivo de ayudantes, si no, es de Conductores

                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count


                if (worksheet.Cells[1, 5].Value?.ToString().Trim()== "Cajas v1")
                {
                    esArchivoDeAyudantes = true;
                }
                else
                {
                    esArchivoDeAyudantes = false;
                }

                for (int row = 1; row <= rowCount; row++)
                {


                    Comision c = new Comision();
                    String numerosRut = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String digitoVerificador = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    String rut = numerosRut + "-" + digitoVerificador;
                    //del Excel que manda el Francisco Cornejo, lo que se necesita es el rut, y el valor del concepto
                    //se puede determinar que concepto es segun la posicion de la columna y el tipo de archivo  (comision de ayudantes
                    //o comision de conductores)
                    c.Plantilla = rut;
                    c.Contrato = "1";
                    c.Origen = "M";
                    c.Objeto = "";
                    c.PeriodoDePago = "M";
                    c.FechaDeInicio = "";
                    c.FechaDeTermino = "";
                    c.Institucion = "";
                    c.DatoAdicional = "";
                    c.Comentario = "";
                    c.ValorPorDefecto = "";
                    c.Accion = "";

                    if (esArchivoDeAyudantes==true)
                    {
                        //es archivo de ayudantes
                        //los posibles conceptos son estos:
                        //Comisión V1 == comisionMi columna 14 (SE CAMBIO)
                        //Comisión V2 == COMISDAVUELTA columna 15 
                        //Cajas Fijas == CAJASF columna 17
                        //Semana Corrida == semanaCorr columna 18
                        //Innovacion == BONOINNOV columna 19
                        //Clientes == VIATIVISITA columna 20
                        //Dotación == BONODOT columna 21
                        //Bonificado == BOCARGBONI columna 22
                        //Bono Asistencia == BonoAsis columna 24
                        //Recargue == VIATICOREC columna 25
                        Comision comisionPrimeraVuelta = retornarComisionConConcepto(c, "COMISION", worksheet.Cells[row, 14].Value?.ToString().Trim());//comisionMi
                        Comision comisionSegundaVuelta = retornarComisionConConcepto(c, "COMISDAVUELT", worksheet.Cells[row, 15].Value?.ToString().Trim());//COMISDAVUELTA
                        Comision comisionCajasFijas = retornarComisionConConcepto(c, "CAJASF", worksheet.Cells[row, 17].Value?.ToString().Trim());
                        Comision comisionSemanaCorrida = retornarComisionConConcepto(c, "semanaCorr", worksheet.Cells[row, 18].Value?.ToString().Trim());
                        Comision comisionInnovacion = retornarComisionConConcepto(c, "BONOINNOV", worksheet.Cells[row, 19].Value?.ToString().Trim());
                        Comision comisionClientes = retornarComisionConConcepto(c, "VIATIVISITA", worksheet.Cells[row, 20].Value?.ToString().Trim());
                        Comision comisionDotacion = retornarComisionConConcepto(c, "BONODOT", worksheet.Cells[row, 21].Value?.ToString().Trim());
                        Comision comisionBonificado = retornarComisionConConcepto(c, "BOCARGBONI", worksheet.Cells[row, 22].Value?.ToString().Trim());
                        Comision comisionBonoAsistencia = retornarComisionConConcepto(c, "BonoAsis", worksheet.Cells[row, 24].Value?.ToString().Trim());
                        Comision comisionRecargue = retornarComisionConConcepto(c, "VIATICOREC", worksheet.Cells[row, 25].Value?.ToString().Trim());

                        comisionesTemporales.Add(comisionPrimeraVuelta);
                        comisionesTemporales.Add(comisionSegundaVuelta);
                        comisionesTemporales.Add(comisionCajasFijas);
                        comisionesTemporales.Add(comisionSemanaCorrida);
                        comisionesTemporales.Add(comisionInnovacion);
                        comisionesTemporales.Add(comisionClientes);
                        comisionesTemporales.Add(comisionDotacion);
                        comisionesTemporales.Add(comisionBonificado);
                        comisionesTemporales.Add(comisionBonoAsistencia);
                        comisionesTemporales.Add(comisionRecargue);

                    }
                    else
                    {
                        //es archivo de conductores
                        //los posibles conceptos son estos:
                        //Comisión == comisionMi columna 12
                        //Cajas Fijas == CAJASF columna 13
                        //Cli. 10p. == VIATIVISITA columna 14
                        //Semana Corrida == semanaCorr columna 16
                        //Asig.Cajas == asigPerdCajaMi columna 17
                        //Dotación == BONODOT columna 18
                        //Bonificado == BOCARGBONI columna 19
                        //Bono por Caja == VIATICOEXTCAJ columna 20
                        //Bono Asistencia == BonoAsis columna 21
                        //Recargue == VIATICOREC columna 22


                        Comision comisionSimple = retornarComisionConConcepto(c, "COMISION", worksheet.Cells[row, 12].Value?.ToString().Trim());//comisionMi
                        Comision comisionCajasFijas = retornarComisionConConcepto(c, "CAJASF", worksheet.Cells[row, 13].Value?.ToString().Trim());
                        Comision comisionCli10p = retornarComisionConConcepto(c, "VIATIVISITA", worksheet.Cells[row, 14].Value?.ToString().Trim());
                        Comision comisionSemanaCorrida = retornarComisionConConcepto(c, "semanaCorr", worksheet.Cells[row, 16].Value?.ToString().Trim());
                        Comision comisionAsignacionDeCajas = retornarComisionConConcepto(c, "asigPerdCajaMi", worksheet.Cells[row, 17].Value?.ToString().Trim());
                        Comision comisionDotacion = retornarComisionConConcepto(c, "BONODOT", worksheet.Cells[row, 18].Value?.ToString().Trim());
                        Comision comisionBonificado = retornarComisionConConcepto(c, "BOCARGBONI", worksheet.Cells[row, 19].Value?.ToString().Trim());
                        Comision comisionBonoPorCaja = retornarComisionConConcepto(c, "VIATICOEXTCAJ", worksheet.Cells[row, 20].Value?.ToString().Trim());
                        Comision comisionBonoAsistencia = retornarComisionConConcepto(c, "BonoAsis", worksheet.Cells[row, 21].Value?.ToString().Trim());
                        Comision comisionRecargue = retornarComisionConConcepto(c, "VIATICOREC", worksheet.Cells[row, 22].Value?.ToString().Trim());


                        comisionesTemporales.Add(comisionSimple);
                        comisionesTemporales.Add(comisionCajasFijas);
                        comisionesTemporales.Add(comisionCli10p);
                        comisionesTemporales.Add(comisionSemanaCorrida);
                        comisionesTemporales.Add(comisionAsignacionDeCajas);
                        comisionesTemporales.Add(comisionDotacion);
                        comisionesTemporales.Add(comisionBonificado);
                        comisionesTemporales.Add(comisionBonoPorCaja);
                        comisionesTemporales.Add(comisionBonoAsistencia);
                        comisionesTemporales.Add(comisionRecargue);


                    }

                    //Comisiones de conductores


                    //for (int col = 1; col <= colCount; col++)
                    //{
                    //    Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value ?.ToString().Trim());
                    //}

                    foreach (var item in comisionesTemporales)
                    {

                        

                        if (item.Plantilla!="Rut-Dv" && item.Valor!="0" && String.IsNullOrEmpty(item.Valor)==false)
                        {                        
                                comisiones.Add(item);
                                                
                        }
                    }
                    
                    
                }
            }

            comisiones = comisiones.Distinct().ToList();

            return comisiones;

        }

        private List<RegistroMensualDeTrabajador> leerExcelDeRegistroDeTrabajadores(string FilePath)
        {
            List<RegistroMensualDeTrabajador> registrosMensualesDeTrabajadores = new List<RegistroMensualDeTrabajador>();
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count



                for (int row = 1; row <= rowCount; row++)
                {
                    String columnaDeNombre = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    if (columnaDeNombre!="" && columnaDeNombre!="nombre")
                    {
                        RegistroMensualDeTrabajador r = new RegistroMensualDeTrabajador();
                        r.Empleado = worksheet.Cells[row, 1].Value?.ToString().Trim();
                        r.Nombre = worksheet.Cells[row, 2].Value?.ToString().Trim();
                        r.ApellidoPate = worksheet.Cells[row, 3].Value?.ToString().Trim();
                        r.ApellidoMate = worksheet.Cells[row, 4].Value?.ToString().Trim();
                        r.FechaNaci = worksheet.Cells[row, 5].Value?.ToString().Trim();
                        r.Nombre_empresa = worksheet.Cells[row, 6].Value?.ToString().Trim();
                        r.Nombre_cargo = worksheet.Cells[row, 7].Value?.ToString().Trim();
                        r.Nombre_centro_costo = worksheet.Cells[row, 8].Value?.ToString().Trim();
                        r.Proceso = worksheet.Cells[row, 9].Value?.ToString().Trim();
                        r.Imponible_sin_tope = worksheet.Cells[row, 10].Value?.ToString().Trim();
                        r.Total_exento = worksheet.Cells[row, 11].Value?.ToString().Trim();

                        registrosMensualesDeTrabajadores.Add(r);
                    }
                    


                }
            }

            return registrosMensualesDeTrabajadores;

        }


        private String alterarFormatoDeFecha(String fechaATransformar)
        {

            String fechaNueva = "";
            if (fechaATransformar != "Fecha")
            {
                fechaNueva = "";

                string[] words = fechaATransformar.Split('-');
                //0 anio
                //1 mes
                //2 dia
                fechaNueva = words[2] + "-" + words[1] + "-" + words[0];

                
            }

            return fechaNueva;
        }


        private Comision retornarComisionConConcepto(Comision comisionAAfectar, String concepto, String valorDeConcepto)
        {
            Comision comision = new Comision(); 
            comision.Plantilla =comisionAAfectar.Plantilla;
            comision.Contrato = comisionAAfectar.Contrato;
            comision.Origen = comisionAAfectar.Origen;
            comision.Objeto = comisionAAfectar.Objeto;
            comision.PeriodoDePago = comisionAAfectar.PeriodoDePago;
            comision.FechaDeInicio = comisionAAfectar.FechaDeInicio;
            comision.FechaDeTermino = comisionAAfectar.FechaDeTermino;
            comision.Institucion = comisionAAfectar.Institucion;
            comision.DatoAdicional = comisionAAfectar.DatoAdicional;
            comision.Comentario = comisionAAfectar.Comentario;
            comision.ValorPorDefecto = comisionAAfectar.ValorPorDefecto;
            comision.Accion = comisionAAfectar.Accion;

            comision.Concepto = concepto;
            comision.Valor = valorDeConcepto;

            return comision;
        }

        private List<RegistroTotalesComoString> procesarRegistrosMensuales(List<RegistroMensualDeTrabajador> registrosMensualesDeTrabajadores)
        {

            List<RegistroTotalesComoString> listadoDeRegistrosDeTotales = new List<RegistroTotalesComoString>();

            List<String> procesos = new List<String>();

            for (int i = 2022; i < 2023; i++)
            {
                for (int j = 1; j < 13; j++)
                {
                    if (j<10)
                    {
                        procesos.Add(i+"-0"+j);
                    }
                    else
                    {
                        procesos.Add(i + "-" + j);
                    }
                }
            }


            foreach (var procesoActual in procesos)
            {
                
            

            //List <RegistroTotalesComoString> listadoDeRegistrosDeTotales = new List<RegistroTotalesComoString>();

            //un registro de totales por centro

            RegistroDeTotales registroProceso = new RegistroDeTotales(procesoActual, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroEspacio = new RegistroDeTotales("", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroCurico = new RegistroDeTotales("Curico",0,0,0,0,0,0,0,0,0,0,0,0,0,0);
            RegistroDeTotales registroInterplanta = new RegistroDeTotales("Interplanta", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroRancagua = new RegistroDeTotales("Rancagua", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroTaller = new RegistroDeTotales("Taller", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);// no se a donde es taller
            RegistroDeTotales registroMelipilla = new RegistroDeTotales("Melipilla", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroSanAntonio = new RegistroDeTotales("San Antonio", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroIllapel = new RegistroDeTotales("Illapel", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroSantiago = new RegistroDeTotales("Santiago", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroMovilizadores = new RegistroDeTotales("Movilizadores", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroAdministracion = new RegistroDeTotales("Administracion", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroEmprendedores = new RegistroDeTotales("Emprendedores", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroEspacio2 = new RegistroDeTotales("", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            RegistroDeTotales registroEspacio3 = new RegistroDeTotales("", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

              



            foreach (var item in registrosMensualesDeTrabajadores)
            {

                    if (item.Proceso == procesoActual)
                    {

                        if (item.Nombre_centro_costo == "CURICO" || item.Nombre_centro_costo == "CURICO E2")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroCurico.TotalAyudantes = registroCurico.TotalAyudantes + 1;
                                    registroCurico.TotalDotacion = registroCurico.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroCurico.CantidadDeAyudantesDeLicencia = registroCurico.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroCurico.CantidadDeAyudantesActivos = registroCurico.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroCurico.TotalRemuneracionesAyudantes = registroCurico.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroCurico.TotalAyudantes = registroCurico.TotalAyudantes + 1;
                                    registroCurico.TotalDotacion = registroCurico.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroCurico.CantidadDeAyudantesDeLicencia = registroCurico.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroCurico.CantidadDeAyudantesActivos = registroCurico.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroCurico.TotalRemuneracionesAyudantes = registroCurico.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroCurico.TotalConductores = registroCurico.TotalConductores + 1;
                                    registroCurico.TotalDotacion = registroCurico.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroCurico.CantidadDeConductoresDeLicencia = registroCurico.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroCurico.CantidadDeConductoresActivos = registroCurico.CantidadDeConductoresActivos + 1;

                                    }

                                    registroCurico.TotalRemuneracionesConductores = registroCurico.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroCurico.TotalConductores = registroCurico.TotalConductores + 1;
                                    registroCurico.TotalDotacion = registroCurico.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroCurico.CantidadDeConductoresDeLicencia = registroCurico.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroCurico.CantidadDeConductoresActivos = registroCurico.CantidadDeConductoresActivos + 1;

                                    }

                                    registroCurico.TotalRemuneracionesConductores = registroCurico.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                               
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroCurico.TotalApoyos = registroCurico.TotalApoyos + 1;
                                    registroCurico.TotalDotacion = registroCurico.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroCurico.CantidadDeApoyosDeLicencia = registroCurico.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroCurico.CantidadDeApoyosActivos = registroCurico.CantidadDeApoyosActivos + 1;

                                    }

                                    registroCurico.TotalRemuneracionesOtros = registroCurico.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

                        }


                        if (item.Nombre_centro_costo == "ILLAPEL")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroIllapel.TotalAyudantes = registroIllapel.TotalAyudantes + 1;
                                    registroIllapel.TotalDotacion = registroIllapel.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroIllapel.CantidadDeAyudantesDeLicencia = registroIllapel.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroIllapel.CantidadDeAyudantesActivos = registroIllapel.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroIllapel.TotalRemuneracionesAyudantes = registroIllapel.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroIllapel.TotalAyudantes = registroIllapel.TotalAyudantes + 1;
                                    registroIllapel.TotalDotacion = registroIllapel.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroIllapel.CantidadDeAyudantesDeLicencia = registroIllapel.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroIllapel.CantidadDeAyudantesActivos = registroIllapel.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroIllapel.TotalRemuneracionesAyudantes = registroIllapel.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroIllapel.TotalConductores = registroIllapel.TotalConductores + 1;
                                    registroIllapel.TotalDotacion = registroIllapel.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroIllapel.CantidadDeConductoresDeLicencia = registroIllapel.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroIllapel.CantidadDeConductoresActivos = registroIllapel.CantidadDeConductoresActivos + 1;

                                    }

                                    registroIllapel.TotalRemuneracionesConductores = registroIllapel.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroIllapel.TotalConductores = registroIllapel.TotalConductores + 1;
                                    registroIllapel.TotalDotacion = registroIllapel.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroIllapel.CantidadDeConductoresDeLicencia = registroIllapel.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroIllapel.CantidadDeConductoresActivos = registroIllapel.CantidadDeConductoresActivos + 1;

                                    }

                                    registroIllapel.TotalRemuneracionesConductores = registroIllapel.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                              
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroIllapel.TotalApoyos = registroIllapel.TotalApoyos + 1;
                                    registroIllapel.TotalDotacion = registroIllapel.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroIllapel.CantidadDeApoyosDeLicencia = registroIllapel.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroIllapel.CantidadDeApoyosActivos = registroIllapel.CantidadDeApoyosActivos + 1;

                                    }

                                    registroIllapel.TotalRemuneracionesOtros = registroIllapel.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }


                        }
                        if (item.Nombre_centro_costo == "INTERPLANTA" || item.Nombre_centro_costo == "INTERPLANTA E2")
                        {
                            

                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroInterplanta.TotalDotacion = registroInterplanta.TotalDotacion + 1;
                                    registroInterplanta.TotalAyudantes = registroInterplanta.TotalAyudantes + 1;
                                    
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroInterplanta.CantidadDeAyudantesDeLicencia = registroInterplanta.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroInterplanta.CantidadDeAyudantesActivos = registroInterplanta.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroInterplanta.TotalRemuneracionesAyudantes = registroInterplanta.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroInterplanta.TotalDotacion = registroInterplanta.TotalDotacion + 1;
                                    registroInterplanta.TotalAyudantes = registroInterplanta.TotalAyudantes + 1;
                                   
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroInterplanta.CantidadDeAyudantesDeLicencia = registroInterplanta.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroInterplanta.CantidadDeAyudantesActivos = registroInterplanta.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroInterplanta.TotalRemuneracionesAyudantes = registroInterplanta.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroInterplanta.TotalDotacion = registroInterplanta.TotalDotacion + 1;
                                    registroInterplanta.TotalConductores = registroInterplanta.TotalConductores + 1;
                                   
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroInterplanta.CantidadDeConductoresDeLicencia = registroInterplanta.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroInterplanta.CantidadDeConductoresActivos = registroInterplanta.CantidadDeConductoresActivos + 1;

                                    }

                                    registroInterplanta.TotalRemuneracionesConductores = registroInterplanta.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroInterplanta.TotalDotacion = registroInterplanta.TotalDotacion + 1;
                                    registroInterplanta.TotalConductores = registroInterplanta.TotalConductores + 1;
                                  
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroInterplanta.CantidadDeConductoresDeLicencia = registroInterplanta.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroInterplanta.CantidadDeConductoresActivos = registroInterplanta.CantidadDeConductoresActivos + 1;

                                    }

                                    registroInterplanta.TotalRemuneracionesConductores = registroInterplanta.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                              
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroInterplanta.TotalDotacion = registroInterplanta.TotalDotacion + 1;
                                    registroInterplanta.TotalApoyos = registroInterplanta.TotalApoyos + 1;

                                    Console.WriteLine(item.Nombre);

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroInterplanta.CantidadDeApoyosDeLicencia = registroInterplanta.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroInterplanta.CantidadDeApoyosActivos = registroInterplanta.CantidadDeApoyosActivos + 1;

                                    }

                                    registroInterplanta.TotalRemuneracionesOtros = registroInterplanta.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

                         


                        }
                        if (item.Nombre_centro_costo == "MELIPILLA" || item.Nombre_centro_costo == "MELIPILLA E2")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroMelipilla.TotalAyudantes = registroMelipilla.TotalAyudantes + 1;
                                    registroMelipilla.TotalDotacion = registroMelipilla.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMelipilla.CantidadDeAyudantesDeLicencia = registroMelipilla.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMelipilla.CantidadDeAyudantesActivos = registroMelipilla.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroMelipilla.TotalRemuneracionesAyudantes = registroMelipilla.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroMelipilla.TotalAyudantes = registroMelipilla.TotalAyudantes + 1;
                                    registroMelipilla.TotalDotacion = registroMelipilla.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMelipilla.CantidadDeAyudantesDeLicencia = registroMelipilla.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMelipilla.CantidadDeAyudantesActivos = registroMelipilla.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroMelipilla.TotalRemuneracionesAyudantes = registroMelipilla.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroMelipilla.TotalConductores = registroMelipilla.TotalConductores + 1;
                                    registroMelipilla.TotalDotacion = registroMelipilla.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMelipilla.CantidadDeConductoresDeLicencia = registroMelipilla.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMelipilla.CantidadDeConductoresActivos = registroMelipilla.CantidadDeConductoresActivos + 1;

                                    }

                                    registroMelipilla.TotalRemuneracionesConductores = registroMelipilla.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroMelipilla.TotalConductores = registroMelipilla.TotalConductores + 1;
                                    registroMelipilla.TotalDotacion = registroMelipilla.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMelipilla.CantidadDeConductoresDeLicencia = registroMelipilla.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMelipilla.CantidadDeConductoresActivos = registroMelipilla.CantidadDeConductoresActivos + 1;

                                    }

                                    registroMelipilla.TotalRemuneracionesConductores = registroMelipilla.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                              
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroMelipilla.TotalApoyos = registroMelipilla.TotalApoyos + 1;
                                    registroMelipilla.TotalDotacion = registroMelipilla.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMelipilla.CantidadDeApoyosDeLicencia = registroMelipilla.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMelipilla.CantidadDeApoyosActivos = registroMelipilla.CantidadDeApoyosActivos + 1;

                                    }

                                    registroMelipilla.TotalRemuneracionesOtros = registroMelipilla.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

         

                        }
                        if (item.Nombre_centro_costo == "RANCAGUA" || item.Nombre_centro_costo == "RANCAGUA  E2")
                        {
                            //Console.WriteLine("Empleado de Rancagua en el proceso"+item.Proceso);

                            registroRancagua.TotalDotacion = registroRancagua.TotalDotacion + 1;

                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroRancagua.TotalAyudantes = registroRancagua.TotalAyudantes + 1;
                                   
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroRancagua.CantidadDeAyudantesDeLicencia = registroRancagua.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroRancagua.CantidadDeAyudantesActivos = registroRancagua.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroRancagua.TotalRemuneracionesAyudantes = registroRancagua.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroRancagua.TotalAyudantes = registroRancagua.TotalAyudantes + 1;
                                   
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroRancagua.CantidadDeAyudantesDeLicencia = registroRancagua.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroRancagua.CantidadDeAyudantesActivos = registroRancagua.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroRancagua.TotalRemuneracionesAyudantes = registroRancagua.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroRancagua.TotalConductores = registroRancagua.TotalConductores + 1;
                                    
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroRancagua.CantidadDeConductoresDeLicencia = registroRancagua.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroRancagua.CantidadDeConductoresActivos = registroRancagua.CantidadDeConductoresActivos + 1;

                                    }

                                    registroRancagua.TotalRemuneracionesConductores = registroRancagua.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroRancagua.TotalConductores = registroRancagua.TotalConductores + 1;
                                    
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroRancagua.CantidadDeConductoresDeLicencia = registroRancagua.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroRancagua.CantidadDeConductoresActivos = registroRancagua.CantidadDeConductoresActivos + 1;

                                    }

                                    registroRancagua.TotalRemuneracionesConductores = registroRancagua.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                             
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroRancagua.TotalApoyos = registroRancagua.TotalApoyos + 1;
                                    
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroRancagua.CantidadDeApoyosDeLicencia = registroRancagua.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroRancagua.CantidadDeApoyosActivos = registroRancagua.CantidadDeApoyosActivos + 1;

                                    }

                                    registroRancagua.TotalRemuneracionesOtros = registroRancagua.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "SAN ANTONIO" || item.Nombre_centro_costo == "SAN ANTONIO E2")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroSanAntonio.TotalAyudantes = registroSanAntonio.TotalAyudantes + 1;
                                    registroSanAntonio.TotalDotacion = registroSanAntonio.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSanAntonio.CantidadDeAyudantesDeLicencia = registroSanAntonio.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSanAntonio.CantidadDeAyudantesActivos = registroSanAntonio.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroSanAntonio.TotalRemuneracionesAyudantes = registroSanAntonio.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroSanAntonio.TotalAyudantes = registroSanAntonio.TotalAyudantes + 1;
                                    registroSanAntonio.TotalDotacion = registroSanAntonio.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSanAntonio.CantidadDeAyudantesDeLicencia = registroSanAntonio.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSanAntonio.CantidadDeAyudantesActivos = registroSanAntonio.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroSanAntonio.TotalRemuneracionesAyudantes = registroSanAntonio.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroSanAntonio.TotalConductores = registroSanAntonio.TotalConductores + 1;
                                    registroSanAntonio.TotalDotacion = registroSanAntonio.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSanAntonio.CantidadDeConductoresDeLicencia = registroSanAntonio.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSanAntonio.CantidadDeConductoresActivos = registroSanAntonio.CantidadDeConductoresActivos + 1;

                                    }

                                    registroSanAntonio.TotalRemuneracionesConductores = registroSanAntonio.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroSanAntonio.TotalConductores = registroSanAntonio.TotalConductores + 1;
                                    registroSanAntonio.TotalDotacion = registroSanAntonio.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSanAntonio.CantidadDeConductoresDeLicencia = registroSanAntonio.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSanAntonio.CantidadDeConductoresActivos = registroSanAntonio.CantidadDeConductoresActivos + 1;

                                    }

                                    registroSanAntonio.TotalRemuneracionesConductores = registroSanAntonio.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                           
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroSanAntonio.TotalApoyos = registroSanAntonio.TotalApoyos + 1;
                                    registroSanAntonio.TotalDotacion = registroSanAntonio.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSanAntonio.CantidadDeApoyosDeLicencia = registroSanAntonio.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSanAntonio.CantidadDeApoyosActivos = registroSanAntonio.CantidadDeApoyosActivos + 1;

                                    }

                                    registroSanAntonio.TotalRemuneracionesOtros = registroSanAntonio.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "SANTIAGO" || item.Nombre_centro_costo == "SANTIAGO E2")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroSantiago.TotalAyudantes = registroSantiago.TotalAyudantes + 1;
                                    registroSantiago.TotalDotacion = registroSantiago.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSantiago.CantidadDeAyudantesDeLicencia = registroSantiago.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSantiago.CantidadDeAyudantesActivos = registroSantiago.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroSantiago.TotalAyudantes = registroSantiago.TotalAyudantes + 1;
                                    registroSantiago.TotalDotacion = registroSantiago.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSantiago.CantidadDeAyudantesDeLicencia = registroSantiago.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSantiago.CantidadDeAyudantesActivos = registroSantiago.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroSantiago.TotalConductores = registroSantiago.TotalConductores + 1;
                                    registroSantiago.TotalDotacion = registroSantiago.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSantiago.CantidadDeConductoresDeLicencia = registroSantiago.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSantiago.CantidadDeConductoresActivos = registroSantiago.CantidadDeConductoresActivos + 1;

                                    }

                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroSantiago.TotalConductores = registroSantiago.TotalConductores + 1;
                                    registroSantiago.TotalDotacion = registroSantiago.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSantiago.CantidadDeConductoresDeLicencia = registroSantiago.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSantiago.CantidadDeConductoresActivos = registroSantiago.CantidadDeConductoresActivos + 1;

                                    }

                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "MOVILIZADOR":
                                    registroMovilizadores.TotalApoyos = registroMovilizadores.TotalApoyos + 1;
                                    registroMovilizadores.TotalDotacion = registroMovilizadores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMovilizadores.CantidadDeApoyosDeLicencia = registroMovilizadores.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMovilizadores.CantidadDeApoyosActivos = registroMovilizadores.CantidadDeApoyosActivos + 1;

                                    }

                                    registroMovilizadores.TotalRemuneracionesOtros = registroMovilizadores.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroSantiago.TotalApoyos = registroSantiago.TotalApoyos + 1;
                                    registroSantiago.TotalDotacion = registroSantiago.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroSantiago.CantidadDeApoyosDeLicencia = registroSantiago.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroSantiago.CantidadDeApoyosActivos = registroSantiago.CantidadDeApoyosActivos + 1;

                                    }

                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "CENTRAL" || item.Nombre_centro_costo == "CENTRAL E2")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroAdministracion.TotalAyudantes = registroAdministracion.TotalAyudantes + 1;
                                    registroAdministracion.TotalDotacion = registroAdministracion.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroAdministracion.CantidadDeAyudantesDeLicencia = registroAdministracion.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroAdministracion.CantidadDeAyudantesActivos = registroAdministracion.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroAdministracion.TotalRemuneracionesAyudantes = registroAdministracion.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroAdministracion.TotalAyudantes = registroAdministracion.TotalAyudantes + 1;
                                    registroAdministracion.TotalDotacion = registroAdministracion.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroAdministracion.CantidadDeAyudantesDeLicencia = registroAdministracion.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroAdministracion.CantidadDeAyudantesActivos = registroAdministracion.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroAdministracion.TotalRemuneracionesAyudantes = registroAdministracion.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroAdministracion.TotalConductores = registroAdministracion.TotalConductores + 1;
                                    registroAdministracion.TotalDotacion = registroAdministracion.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroAdministracion.CantidadDeConductoresDeLicencia = registroAdministracion.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroAdministracion.CantidadDeConductoresActivos = registroAdministracion.CantidadDeConductoresActivos + 1;

                                    }

                                    registroAdministracion.TotalRemuneracionesConductores = registroAdministracion.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroAdministracion.TotalConductores = registroAdministracion.TotalConductores + 1;
                                    registroAdministracion.TotalDotacion = registroAdministracion.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroAdministracion.CantidadDeConductoresDeLicencia = registroAdministracion.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroAdministracion.CantidadDeConductoresActivos = registroAdministracion.CantidadDeConductoresActivos + 1;

                                    }

                                    registroAdministracion.TotalRemuneracionesConductores = registroAdministracion.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                            
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroAdministracion.TotalApoyos = registroAdministracion.TotalApoyos + 1;
                                    registroAdministracion.TotalDotacion = registroAdministracion.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroAdministracion.CantidadDeApoyosDeLicencia = registroAdministracion.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroAdministracion.CantidadDeApoyosActivos = registroAdministracion.CantidadDeApoyosActivos + 1;

                                    }

                                    registroAdministracion.TotalRemuneracionesOtros = registroAdministracion.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "EMPRENDEDORES")
                        {
                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroEmprendedores.TotalAyudantes = registroEmprendedores.TotalAyudantes + 1;
                                    registroEmprendedores.TotalDotacion = registroEmprendedores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroEmprendedores.CantidadDeAyudantesDeLicencia = registroEmprendedores.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroEmprendedores.CantidadDeAyudantesActivos = registroEmprendedores.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroEmprendedores.TotalRemuneracionesAyudantes = registroEmprendedores.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroEmprendedores.TotalAyudantes = registroEmprendedores.TotalAyudantes + 1;
                                    registroEmprendedores.TotalDotacion = registroEmprendedores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroEmprendedores.CantidadDeAyudantesDeLicencia = registroEmprendedores.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroEmprendedores.CantidadDeAyudantesActivos = registroEmprendedores.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroEmprendedores.TotalRemuneracionesAyudantes = registroEmprendedores.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO":
                                    registroEmprendedores.TotalConductores = registroEmprendedores.TotalConductores + 1;
                                    registroEmprendedores.TotalDotacion = registroEmprendedores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroEmprendedores.CantidadDeConductoresDeLicencia = registroEmprendedores.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroEmprendedores.CantidadDeConductoresActivos = registroEmprendedores.CantidadDeConductoresActivos + 1;

                                    }

                                    registroEmprendedores.TotalRemuneracionesConductores = registroEmprendedores.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroEmprendedores.TotalConductores = registroEmprendedores.TotalConductores + 1;
                                    registroEmprendedores.TotalDotacion = registroEmprendedores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroEmprendedores.CantidadDeConductoresDeLicencia = registroEmprendedores.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroEmprendedores.CantidadDeConductoresActivos = registroEmprendedores.CantidadDeConductoresActivos + 1;

                                    }

                                    registroEmprendedores.TotalRemuneracionesConductores = registroEmprendedores.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    break;
                             
                                case "MECANICO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    //restarDeRancagua

                                    break;
                                case "NOCHERO":
                                    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                    }

                                    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    //restarDeRancagua

                                    break;
                                default:
                                    registroEmprendedores.TotalApoyos = registroEmprendedores.TotalApoyos + 1;
                                    registroEmprendedores.TotalDotacion = registroEmprendedores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroEmprendedores.CantidadDeApoyosDeLicencia = registroEmprendedores.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroEmprendedores.CantidadDeApoyosActivos = registroEmprendedores.CantidadDeApoyosActivos + 1;

                                    }

                                    registroEmprendedores.TotalRemuneracionesOtros = registroEmprendedores.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));

                                    break;
                            }


                        }
                    }
                
            }


            RegistroTotalesComoString registroProcesoComoString = new RegistroTotalesComoString(registroProceso.Centro);
            RegistroTotalesComoString registroEspacioComoString = new RegistroTotalesComoString(registroEspacio.Centro);
            RegistroTotalesComoString registroCuricoComoString = new RegistroTotalesComoString(registroCurico);
            RegistroTotalesComoString registroInterplantaComoString = new RegistroTotalesComoString(registroInterplanta);
            RegistroTotalesComoString registroRancaguaComoString = new RegistroTotalesComoString(registroRancagua);
            RegistroTotalesComoString registroTallerComoString = new RegistroTotalesComoString(registroTaller);
            RegistroTotalesComoString registroMelipillaComoString = new RegistroTotalesComoString(registroMelipilla);
            RegistroTotalesComoString registroSanAntonioComoString = new RegistroTotalesComoString(registroSanAntonio);
            RegistroTotalesComoString registroIllapelComoString = new RegistroTotalesComoString(registroIllapel);
            RegistroTotalesComoString registroSantiagoComoString = new RegistroTotalesComoString(registroSantiago);
            RegistroTotalesComoString registroMovilizadoresComoString = new RegistroTotalesComoString(registroMovilizadores);
            RegistroTotalesComoString registroAdministracionComoString = new RegistroTotalesComoString(registroAdministracion);
            RegistroTotalesComoString registroEmprendedoresComoString = new RegistroTotalesComoString(registroEmprendedores);
            RegistroTotalesComoString registroEspacioComoString2 = new RegistroTotalesComoString(registroEspacio2.Centro);
            RegistroTotalesComoString registroEspacioComoString3 = new RegistroTotalesComoString(registroEspacio3.Centro);


            listadoDeRegistrosDeTotales.Add(registroProcesoComoString);
            listadoDeRegistrosDeTotales.Add(registroEspacioComoString);
            listadoDeRegistrosDeTotales.Add(registroCuricoComoString);
            listadoDeRegistrosDeTotales.Add(registroInterplantaComoString);
            listadoDeRegistrosDeTotales.Add(registroRancaguaComoString);
            listadoDeRegistrosDeTotales.Add(registroTallerComoString);
            listadoDeRegistrosDeTotales.Add(registroMelipillaComoString);
            listadoDeRegistrosDeTotales.Add(registroSanAntonioComoString);
            listadoDeRegistrosDeTotales.Add(registroIllapelComoString);
            listadoDeRegistrosDeTotales.Add(registroSantiagoComoString);
            listadoDeRegistrosDeTotales.Add(registroMovilizadoresComoString);
            listadoDeRegistrosDeTotales.Add(registroEmprendedoresComoString);
            listadoDeRegistrosDeTotales.Add(registroAdministracionComoString);
            listadoDeRegistrosDeTotales.Add(registroEspacioComoString2);
            listadoDeRegistrosDeTotales.Add(registroEspacioComoString3);

            }

            return listadoDeRegistrosDeTotales;



        }

        private void mostrarAyuda()
        {
            MessageBox.Show("para subir asistencias a rex: * recibir excel de Francisco * copiar los datos que vienen filtrados en el excel, a un excel nuevo que tenga la cabecera(ese excel se descarga de rex) * Guardar el nuevo excel con los registros copiados como formato CSV * Enviar a las de remuneraciones para que ellas hagan la carga");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            mostrarAyuda();
        }

        private void button4_Click(object sender, EventArgs e)
        {


            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }



            List<RegistroMensualDeTrabajador> registros = leerExcelDeRegistroDeTrabajadores(sFileName);

            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";
            
           List<RegistroTotalesComoString> registrosDeTotales = procesarRegistrosMensuales(registros);


            var archivo = new FileInfo(downloads + @"\Registro de totales.xlsx");

            SaveExcelFileRegistroDeTotales(registrosDeTotales, archivo);

            MessageBox.Show("Archivo Excel de de totales creado en carpeta de descargas!");


        }
    }
}
