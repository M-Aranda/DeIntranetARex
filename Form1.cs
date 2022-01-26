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

            List<Asistencia> asistencias = leerExcelDeFallos(sFileName);





            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";


            var archivo = new FileInfo(downloads + @"\Asistencias.xlsx");

            SaveExcelFileAsistencia(asistencias, archivo);

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

            readXLS2(sFileName);





            List<Comision> comisiones = new List<Comision>();

            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";


            var archivo = new FileInfo(downloads + @"\Comisiones.xlsx");

            SaveExcelFileComision(comisiones, archivo);

            MessageBox.Show("Archivo Excel de comisiones creado en carpeta de descargas!");
        }






        private static async Task SaveExcelFileAsistencia(List<Asistencia> asistencias, FileInfo file)
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


        private List<Asistencia> leerExcelDeFallos(string FilePath)
        {
            List<Asistencia> asistencias = new List<Asistencia>();     
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
               Asistencia encabezado = new Asistencia();
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

                asistencias.Add(encabezado);    


                for (int row = 1; row <= rowCount; row++)
                {

                    Asistencia a = new Asistencia();
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
                        asistencias.Add(a);
                    }
                        
      
                    
                }
            }

            return asistencias;
        
        }




        public void readXLS2(string FilePath)
        {
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                    }
                }
            }
        }


        private String alterarFormatoDeFecha(String fechaATransformar)
        {
            String fechaNueva = "";

            string[] words = fechaATransformar.Split('-');
            //0 anio
            //1 mes
            //2 dia
            fechaNueva = words[2]+"-"+words[1] + "-"+words[0];

            return fechaNueva;
        }



    }
}
