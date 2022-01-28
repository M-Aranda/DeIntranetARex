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
                        //Comisión V1 == ComisionMi columna 14
                        //Comisión V2 == COMISDAVUELTA columna 15
                        //Cajas Fijas == CAJASF columna 17
                        //Semana Corrida == semanaCorr columna 18
                        //Innovacion == BONOINNOV columna 19
                        //Clientes == VIATIVISITA columna 20
                        //Dotación == BONODOT columna 21
                        //Bonificado == BOCARGBONI columna 22
                        //Bono Asistencia == BonoAsis columna 24
                        //Recargue == VIATICOREC columna 25
                        Comision comisionPrimeraVuelta = retornarComisionConConcepto(c, "ComisionMi", worksheet.Cells[row, 14].Value?.ToString().Trim());
                        Comision comisionSegundaVuelta = retornarComisionConConcepto(c, "COMISDAVUELTA", worksheet.Cells[row, 15].Value?.ToString().Trim());
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
                        //Semana Corrida == semanacorr columna 16
                        //Asig.Cajas == asigPerdCajaMi columna 17
                        //Dotación == BONODOT columna 18
                        //Bonificado == BOCARGBONI columna 19
                        //Bono por Caja == VIATICOEXTCAJA columna 20
                        //Bono Asistencia == BonoAsis columna 21
                        //Recargue == VIATICOREC columna 22


                        Comision comisionSimple = retornarComisionConConcepto(c, "ComisionMi", worksheet.Cells[row, 12].Value?.ToString().Trim());
                        Comision comisionCajasFijas = retornarComisionConConcepto(c, "CAJASF", worksheet.Cells[row, 13].Value?.ToString().Trim());
                        Comision comisionCli10p = retornarComisionConConcepto(c, "VIATIVISITA", worksheet.Cells[row, 14].Value?.ToString().Trim());
                        Comision comisionSemanaCorrida = retornarComisionConConcepto(c, "semanacorr", worksheet.Cells[row, 16].Value?.ToString().Trim());
                        Comision comisionAsignacionDeCajas = retornarComisionConConcepto(c, "asigPerdCajaMi", worksheet.Cells[row, 17].Value?.ToString().Trim());
                        Comision comisionDotacion = retornarComisionConConcepto(c, "BONODOT", worksheet.Cells[row, 18].Value?.ToString().Trim());
                        Comision comisionBonificado = retornarComisionConConcepto(c, "BOCARGBONI", worksheet.Cells[row, 19].Value?.ToString().Trim());
                        Comision comisionBonoPorCaja = retornarComisionConConcepto(c, "VIATICOEXTCAJA", worksheet.Cells[row, 20].Value?.ToString().Trim());
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

            return comisiones;

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



    }
}
