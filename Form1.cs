using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Windows.Storage;
using System.IO;
using System.Runtime.InteropServices;

namespace DeIntranetARex
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
         

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





            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@sFileName);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            // barra de progreso
            //pBar1.Visible = true;
            //pBar1.Minimum = 1;
            //pBar1.Maximum = rowCount;
            //pBar1.Value = 1;
            //pBar1.Step = 1;

            string TipoArchivo = "";
            // leer los títulos
            // Archivo Conductores
            // Estado	Centro	Rut	Nombre	Fecha	Dia	Feriado	Retiro.Env	Corte	Carga	Bonificado	Cajas V1	Cajas V2	Sindicato	Empresa	Pallets

            //if (xlRange.Cells[1, 1].Value2.ToString() == "Estado" &&
            //    xlRange.Cells[1, 2].Value2.ToString() == "Centro" &&
            //    xlRange.Cells[1, 3].Value2.ToString() == "Rut" &&
            //    xlRange.Cells[1, 4].Value2.ToString() == "Nombre" &&
            //    xlRange.Cells[1, 5].Value2.ToString() == "Fecha" &&
            //    xlRange.Cells[1, 6].Value2.ToString() == "Dia" &&
            //    xlRange.Cells[1, 7].Value2.ToString() == "Feriado" &&
            //    xlRange.Cells[1, 8].Value2.ToString() == "Retiro.Env" &&
            //    xlRange.Cells[1, 9].Value2.ToString() == "Corte" &&
            //    xlRange.Cells[1, 10].Value2.ToString() == "Carga" &&
            //    xlRange.Cells[1, 11].Value2.ToString() == "Bonificado" &&
            //    xlRange.Cells[1, 12].Value2.ToString() == "Cajas V1" &&
            //    xlRange.Cells[1, 13].Value2.ToString() == "Cajas V2" &&
            //    xlRange.Cells[1, 14].Value2.ToString() == "Sindicato" &&
            //    xlRange.Cells[1, 15].Value2.ToString() == "Empresa" &&
            //    xlRange.Cells[1, 16].Value2.ToString() == "Pallets")
            //{
            //    // corresponde a archivo de pago para conductores
            //    TipoArchivo = "Conductores";
            //}
            //// Estado	Centro	Rut	Nombre	Fecha	Dia	Feriado	Retiro.Env	Corte	Carga	Bonificado	Cajas V1	Cajas V2	Patente	Pallets	Sindicato	Empresa
            //if (xlRange.Cells[1, 1].Value2.ToString() == "Estado" &&
            //    xlRange.Cells[1, 2].Value2.ToString() == "Centro" &&
            //    xlRange.Cells[1, 3].Value2.ToString() == "Rut" &&
            //    xlRange.Cells[1, 4].Value2.ToString() == "Nombre" &&
            //    xlRange.Cells[1, 5].Value2.ToString() == "Fecha" &&
            //    xlRange.Cells[1, 6].Value2.ToString() == "Dia" &&
            //    xlRange.Cells[1, 7].Value2.ToString() == "Feriado" &&
            //    xlRange.Cells[1, 8].Value2.ToString() == "Retiro.Env" &&
            //    xlRange.Cells[1, 9].Value2.ToString() == "Corte" &&
            //    xlRange.Cells[1, 10].Value2.ToString() == "Carga" &&
            //    xlRange.Cells[1, 11].Value2.ToString() == "Bonificado" &&
            //    xlRange.Cells[1, 12].Value2.ToString() == "Cajas V1" &&
            //    xlRange.Cells[1, 13].Value2.ToString() == "Cajas V2" &&
            //    xlRange.Cells[1, 14].Value2.ToString() == "Patente" &&
            //    xlRange.Cells[1, 15].Value2.ToString() == "Pallets" &&
            //    xlRange.Cells[1, 16].Value2.ToString() == "Sindicato" &&
            //    xlRange.Cells[1, 17].Value2.ToString() == "Empresa")
            //{
            //    TipoArchivo = "Ayudantes";
            //}

            //if (TipoArchivo == "Conductores" || TipoArchivo == "Ayudantes")
            //{
            //    DialogResult respuesta = MessageBox.Show("Archivo de " + TipoArchivo + ", ¿desea continuar?", "Confirmación", MessageBoxButtons.YesNo);

            //    if (respuesta == DialogResult.Yes)
            //    {
            //        if (TipoArchivo == "Conductores")
            //        {
            //            for (int i = 2; i <= rowCount; i++)
            //            {
            //                ArchivoConductor a = new ArchivoConductor();

            //                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
            //                {
            //                    a.rut = xlRange.Cells[i, 3].Value2.ToString();
            //                    a.rut = a.rut.Substring(0, a.rut.Length - 2);
            //                }
            //                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
            //                {
            //                    a.tipo = xlRange.Cells[i, 1].Value2.ToString();
            //                }
            //                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
            //                {
            //                    a.fecha = xlRange.Cells[i, 5].Value2.ToString();
            //                }
            //                if (xlRange.Cells[i, 15] != null && xlRange.Cells[i, 15].Value2 != null)
            //                {
            //                    a.empresa = xlRange.Cells[i, 15].Value2.ToString();
            //                }
            //                if (a.tipo == "FALLA" || a.tipo == "PERMISO")
            //                {
            //                    a.GuardarEnTranstecnia(mesNumero(cmbMes.Text));
            //                }
            //                pBar1.PerformStep();

            //            }

            //        }
            //        if (TipoArchivo == "Ayudantes")
            //        {
            //            for (int i = 2; i <= rowCount; i++)
            //            {
            //                ArchivoAyudante a = new ArchivoAyudante();

            //                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
            //                {
            //                    a.rut = xlRange.Cells[i, 3].Value2.ToString();
            //                    a.rut = a.rut.Substring(0, a.rut.Length - 2);
            //                }
            //                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
            //                {
            //                    a.tipo = xlRange.Cells[i, 1].Value2.ToString();
            //                }
            //                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
            //                {
            //                    a.fecha = xlRange.Cells[i, 5].Value2.ToString();
            //                }
            //                if (xlRange.Cells[i, 17] != null && xlRange.Cells[i, 17].Value2 != null)
            //                {
            //                    a.empresa = xlRange.Cells[i, 17].Value2.ToString();
            //                }
            //                if (a.tipo == "Falla" || a.tipo == "Permiso")
            //                {
            //                    a.GuardarEnTranstecnia(mesNumero(cmbMes.Text));
            //                }
            //                pBar1.PerformStep();
            //            }
            //        }
            //    }
                else
                {
                    MessageBox.Show("No se ha cargado el archivo", "Confirmación");
                }
            }
            else
            {
                MessageBox.Show("El archivo no tiene el formato esperado", "Error", MessageBoxButtons.OK);
            }


            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);





            List<Asistencia> asistencias = new List<Asistencia>();

            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";



            var archivo = new FileInfo(downloads+ @"\Asistencias.xlsx");

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





    }
}











