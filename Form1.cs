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
        private string mesDeBonos = "";
 

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

            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                    break;
                }else
                {
                    MessageBox.Show("No se seleccionó nada, proceso terminado.");
                    System.Environment.Exit(0);
                }

            }

            try
            {

           
            List<Ausencia> ausencias = leerExcelDeFallos(sFileName);

            
            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";


            var archivo = new FileInfo(downloads + @"\Asistencias.xlsx");

            SaveExcelFileAusencia(ausencias, archivo);


            MessageBox.Show("Archivo Excel de asistencias creado en carpeta de descargas!");

            }
            catch (Exception)
            {
                MessageBox.Show("Revisar Excel de carga. Cerrando programa");
                System.Environment.Exit(0);
                throw;
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                    break;
                }
                else
                {
                    MessageBox.Show("No se seleccionó nada, proceso terminado.");
                    System.Environment.Exit(0);
                }

            }

            try
            {
        
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
            catch (Exception)
            {
                MessageBox.Show("Revisar Excel de carga. Cerrando programa");
                System.Environment.Exit(0);
                throw;
            }
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

            //  var range = ws.Cells["A1"].LoadFromCollection(registrosDeTotales, true);
            //var range = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + 1 + "C" + 1, 0, 0)].LoadFromCollection(registrosDeTotales, true);

            // 28/03/2022, Antonio Alonso pidio aplicar un formato al Excel

            // hay que hacer un sub cuadro resumen con la siguiente informacion:
            //interplantas (remuneraciones totales)
            //movilizadores (remuneraciones totales)
            //interplantas (remuneraciones totales)

            //tradicional (remuneraciones totales) a su vez se divide en 3 más:
            //directos (ayudantes y conductores), administracion (los que trabajan en el centro de administración)
            //e indirectos (todos los que NO sean conductores, ayudantes o sean de administracion)


            ////fijar  estilo de letra
            //ws.Cells["B1"].Style.Font.Bold = true;


            ////fijar color de fondo de ciertas celdas

            ////ws.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            ///
            //12/10/2022
            // se agrega columna para Bono Especial Temporada R


            int fila1 = 1;
            int fila2 = 2;
            int fila3 = 3;
            int fila4 = 4;
            int fila5 = 5;
            int fila6 = 6;
            int fila7 = 7;
            int fila8 = 8;
            int fila9 = 9;
            int fila10 = 10;
            int fila11 = 11;
            int fila12 = 12;
            int fila13 = 13;
            int fila14 = 14;
            int fila15 = 15;

            int columna1 = 1;
            int columna2 = 2;
            int columna3 = 3;
            int columna4 = 4;
            int columna5 = 5;
            int columna6 = 6;
            int columna7 = 7;
            int columna8 = 8;
            int columna9 = 9;
            int columna10 = 10;
            int columna11 = 11;
            int columna12 = 12;
            int columna13 = 13;
            int columna14 = 14;
            int columna15 = 15;



            //iterar por un 12 meses
            for (int i = 1; i < 13; i++)
            {

              

                //fijar titulo de columnas
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 2, 0, 0)].Value = "Cantidad de conductores activos";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 3, 0, 0)].Value = "Conductores de licencia";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 4, 0, 0)].Value = "Ayudantes activos";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 5, 0, 0)].Value = "Ayudantes de licencia";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 6, 0, 0)].Value = "Apoyos activos";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 7, 0, 0)].Value = "Apoyos de licencia";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 8, 0, 0)].Value = "Total de conductores";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 9, 0, 0)].Value = "Total de ayudantes";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 10, 0, 0)].Value = "Total de apoyos";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 11, 0, 0)].Value = "Total de trabajadores";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 12, 0, 0)].Value = "$ Conductores";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 13, 0, 0)].Value = "$ Ayudantes";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 14, 0, 0)].Value = "$ Apoyos";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 15, 0, 0)].Value = "Total";
                //titulos de bonos
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 16, 0, 0)].Value = "Total bono tiempo de espera";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 17, 0, 0)].Value = "Total bono estacional";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 18, 0, 0)].Value = "Total Btn I R";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 19, 0, 0)].Value = "Total bono sobre esfuerzo R";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 20, 0, 0)].Value = "Total viático R";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 21, 0, 0)].Value = "Total bono compensatorio R";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 22, 0, 0)].Value = "Bono Especial Temporada R";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 23, 0, 0)].Value = "Total a Recuperar";

                //titulos en negrita
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 2, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 3, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 4, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 5, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 6, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 7, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 8, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 9, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 10, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 11, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 12, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 13, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 14, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 15, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 16, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 17, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 18, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 19, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 20, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 21, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 22, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 23, 0, 0)].Style.Font.Bold = true;

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 1, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 1, 0, 0)].Style.Font.Bold = true;



                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila1 - 1].Centro;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila3 + "C" + 1, 0, 0)].Style.Font.Italic = true;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila2].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila3].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila4].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila5].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila6].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila7].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila8].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila9].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila10].Centro;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 1, 0, 0)].Value = registrosDeTotales[fila11].Centro;




                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila2].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila3].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila4].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila5].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila6].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila7].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila8].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila9].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila10].CantidadDeConductoresActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 2, 0, 0)].Value = registrosDeTotales[fila11].CantidadDeConductoresActivos;

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila2].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila3].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila4].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila5].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila6].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila7].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila8].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila9].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila10].CantidadDeConductoresDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 3, 0, 0)].Value = registrosDeTotales[fila11].CantidadDeConductoresDeLicencia;

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila2].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila3].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila4].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila5].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila6].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila7].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila8].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila9].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila10].CantidadDeAyudantesActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 4, 0, 0)].Value = registrosDeTotales[fila11].CantidadDeAyudantesActivos;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila2].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila3].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila4].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila5].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila6].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila7].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila8].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila9].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila10].CantidadDeAyudantesDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 5, 0, 0)].Value = registrosDeTotales[fila11].CantidadDeAyudantesDeLicencia;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila2].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila3].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila4].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila5].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila6].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila7].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila8].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila9].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila10].CantidadDeApoyosActivos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 6, 0, 0)].Value = registrosDeTotales[fila11].CantidadDeApoyosActivos;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila2].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila3].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila4].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila5].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila6].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila7].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila8].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila9].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila10].CantidadDeApoyosDeLicencia;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 7, 0, 0)].Value = registrosDeTotales[fila11].CantidadDeApoyosDeLicencia;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila2].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila3].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila4].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila5].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila6].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila7].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila8].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila9].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila10].TotalConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 8, 0, 0)].Value = registrosDeTotales[fila11].TotalConductores;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila2].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila3].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila4].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila5].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila6].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila7].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila8].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila9].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila10].TotalAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 9, 0, 0)].Value = registrosDeTotales[fila11].TotalAyudantes;

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila2].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila3].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila4].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila5].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila6].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila7].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila8].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila9].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila10].TotalApoyos;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 10, 0, 0)].Value = registrosDeTotales[fila11].TotalApoyos;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila2].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila3].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila4].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila5].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila6].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila7].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila8].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila9].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila10].TotalDotacion;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 11, 0, 0)].Value = registrosDeTotales[fila11].TotalDotacion;


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila2].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila3].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila4].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila5].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila6].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila7].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila8].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila9].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila10].TotalRemuneracionesConductores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 12, 0, 0)].Value = registrosDeTotales[fila11].TotalRemuneracionesConductores;



                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila2].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila3].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila4].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila5].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila6].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila7].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila8].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila9].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila10].TotalRemuneracionesAyudantes;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 13, 0, 0)].Value = registrosDeTotales[fila11].TotalRemuneracionesAyudantes;



                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila2].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila3].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila4].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila5].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila6].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila7].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila8].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila9].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila10].TotalRemuneracionesOtros;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 14, 0, 0)].Value = registrosDeTotales[fila11].TotalRemuneracionesOtros;



                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila2].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila3].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila4].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila5].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila6].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila7].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila8].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila9].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila10].TotalRemuneracionesDeTodosLosTrabajadores;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 15, 0, 0)].Value = registrosDeTotales[fila11].TotalRemuneracionesDeTodosLosTrabajadores;

                //valores de bonos
  
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila2].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila3].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila4].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila5].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila6].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila7].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila8].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila9].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila10].TotalBonoTiempoEsperaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 16, 0, 0)].Value = registrosDeTotales[fila11].TotalBonoTiempoEsperaR;

              
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila2].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila3].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila4].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila5].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila6].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila7].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila8].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila9].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila10].TotalBonoEstacionalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 17, 0, 0)].Value = registrosDeTotales[fila11].TotalBonoEstacionalR;

               
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila2].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila3].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila4].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila5].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila6].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila7].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila8].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila9].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila10].TotalBtnLR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 18, 0, 0)].Value = registrosDeTotales[fila11].TotalBtnLR;

              
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila2].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila3].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila4].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila5].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila6].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila7].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila8].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila9].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila10].TotalBonoSobreEsfuerzoR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 19, 0, 0)].Value = registrosDeTotales[fila11].TotalBonoSobreEsfuerzoR;

               
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila2].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila3].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila4].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila5].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila6].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila7].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila8].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila9].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila10].TotalViaticoAhorroR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 20, 0, 0)].Value = registrosDeTotales[fila11].TotalViaticoAhorroR;

            
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila2].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila3].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila4].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila5].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila6].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila7].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila8].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila9].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila10].TotalBonoCompensatorioR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 21, 0, 0)].Value = registrosDeTotales[fila11].TotalBonoCompensatorioR;

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila2].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila3].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila4].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila5].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila6].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila7].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila8].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila9].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila10].TotalBonoEspecialTemporadaR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 22, 0, 0)].Value = registrosDeTotales[fila11].TotalBonoEspecialTemporadaR;


                //total a recuperar de los bonos


                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila2].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila3].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila4].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila5].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila6].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila7].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila8].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila9].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila10].TotalR;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 23, 0, 0)].Value = registrosDeTotales[fila11].TotalR;


                // agregar  bordes finos a la tabla
                ws.Cells["A" + fila3 + ":W" + fila13].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells["A" + fila3 + ":W" + fila13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells["A" + fila3 + ":W" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells["A" + fila3 + ":W" + fila13].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // agregar  bordes gruesos a la tabla
                ws.Cells["B" + fila4 + ":W" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["W" + fila4 + ":W" + fila13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila13 + ":W" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila4 + ":B" + fila13].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;


                ws.Cells["L" + fila4 + ":W" + fila13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                ws.Cells["Y" + fila5 + ":AB" + fila10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                ws.Cells["Z" + fila11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";


                //sección resumen de modelos

                //titulos
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 25, 0, 0)].Value = "RESUMEN DE MODELOS";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 26, 0, 0)].Value = "Total";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 27, 0, 0)].Value = "Por Cobrar a CCU";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 28, 0, 0)].Value = "Total Mes";

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 25, 0, 0)].Value = "INTERPLANTAS";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 25, 0, 0)].Value = "MOVILIZADORES";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 25, 0, 0)].Value = "EMPRENDEDORES";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 25, 0, 0)].Value = "DIRECTOS";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 25, 0, 0)].Value = "INDIRECTOS";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 25, 0, 0)].Value = "ADMINISTRACION";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 25, 0, 0)].Value = "TOTAL $";
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 25, 0, 0)].Value = "TOTAL TRABAJADORES";


                //formateo de cuadro sub resumen
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 25, 0, 0)].Style.Font.Bold = true;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 25, 0, 0)].Style.Font.Bold = true;

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 25, 0, 0)].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 25, 0, 0)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 25, 0, 0)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 25, 0, 0)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;


                ws.Cells["Y" + fila5 + ":AB" + fila5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["AB" + fila5 + ":AB" + fila10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Y" + fila10 + ":AB" + fila10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Y" + fila5 + ":Y" + fila10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                //valores
                //total de trabajadores
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 11, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 11, 0, 0)].Value.ToString());

                //valores de remuneraciones
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 15, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 15, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 15, 0, 0)].Value.ToString());

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 12, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 12, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 12, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 12, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 12, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 12, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 13, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 13, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 13, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 13, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 13, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 13, 0, 0)].Value.ToString())
                    ;//directos (suma de todas las remuneraciones de conductores y ayudantes)

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 14, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 14, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 14, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 14, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 14, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 14, 0, 0)].Value.ToString());
                    //indirectos (suma de todas las remuneraciones de los que NO SON conductores y ayudantes)

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 15, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 26, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 26, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 26, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 26, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 26, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 26, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 26, 0, 0)].Value.ToString());//totales


                //Por cobrar a CCU

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 27, 0, 0)].Value = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 23, 0, 0)].Value;//interplantas
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 27, 0, 0)].Value = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila11 + "C" + 23, 0, 0)].Value;//movilizadores
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 27, 0, 0)].Value = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila12 + "C" + 23, 0, 0)].Value;//emprendedores
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 27, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila4 + "C" + 23, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 23, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 23, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 23, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 23, 0, 0)].Value.ToString())
                    + int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 23, 0, 0)].Value.ToString());//directos, filas 4, 6 ,7, 8, 9 y  10 
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 27, 0, 0)].Value = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 23, 0, 0)].Value;//indirectos
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 27, 0, 0)].Value = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila13 + "C" + 23, 0, 0)].Value;//administracion



                //Totales del mes (se supone que lo que se cobra a CCU se recupera)

                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 28, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 26, 0, 0)].Value.ToString())- int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila5 + "C" + 27, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 28, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 26, 0, 0)].Value.ToString()) - int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila6 + "C" + 27, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 28, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 26, 0, 0)].Value.ToString()) - int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila7 + "C" + 27, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 28, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 26, 0, 0)].Value.ToString()) - int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila8 + "C" + 27, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 28, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 26, 0, 0)].Value.ToString()) - int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila9 + "C" + 27, 0, 0)].Value.ToString());
                ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 28, 0, 0)].Value = int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 26, 0, 0)].Value.ToString()) - int.Parse(ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + fila10 + "C" + 27, 0, 0)].Value.ToString());


                //incrementar contadores
                fila1 = fila1 + 14;
                fila2 = fila2 + 14;
                fila3 = fila3 + 14;
                fila4 = fila4 + 14;
                fila5 = fila5 + 14;
                fila6 = fila6 + 14;
                fila7 = fila7 + 14;
                fila8 = fila8 + 14;
                fila9 = fila9 + 14;
                fila10 = fila10 + 14;
                fila11 = fila11 + 14;
                fila12 = fila12 + 14;
                fila13 = fila13 + 14;
                fila14 = fila14 + 14;
                fila15 = fila15 + 14;

            }




           

            //coordenadas para autoajustar columnas
            var range = ws.Cells[ExcelCellBase.TranslateFromR1C1("R" + 1 + "C" + 1 + ":R" + 15 + "C" + 27, 0, 0)];

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
                    Boolean ingresoValido = false;

                    Ausencia a = new Ausencia();
                    a.Tipo = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    a.Empleado = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    a.FechaInicio = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    a.FechaInicio = alterarFormatoDeFecha(a.FechaInicio);

                    if (a.Tipo.ToUpper()=="FALLA" || a.Tipo.ToUpper() == "PERMISO" || a.Tipo.ToUpper() == "ESTADO")
                    {
                        ingresoValido = true;
                    }

                    switch (a.Tipo)
                    {
                        case "Falla":
                            a.Descripcion = "Falla";
                            a.Tipo = "F";
                            a.TipoDePermiso = "";
                            break;
                        case "FALLA":
                            a.Descripcion = "Falla";
                            a.Tipo = "F";
                            a.TipoDePermiso = "";
                            break;
                        case "Permiso":
                            a.Descripcion = "Permiso";
                            a.Tipo = "P";
                            a.TipoDePermiso = "N";
                            break;
                        case "PERMISO":
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


//                    Roberto Fernández,( 15523665 - 5)
//Jorge Henríquez(16493501 - 9),
//Christopher Figueroa(18200697 - 1) y
// Fernando López




                if (a.Tipo != "Estado" && a.Empleado!= "15523665-5" && a.Empleado != "16493501-9" && a.Empleado != "18200697-1" && a.Empleado != "Rut"  && ingresoValido)
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
                        //05-05-2022 se agrega comision de BONODISPONIBILIDAD, para los ayudantes
                        Comision comisionBonoDisponibilidad = retornarComisionConConcepto(c, "BONODISPONIBILIDAD", worksheet.Cells[row, 27].Value?.ToString().Trim());


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
                        comisionesTemporales.Add(comisionBonoDisponibilidad);

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



        private List<MontoPorConcepto> leerHojaDeConceptos(string FilePath, int hoja)
        {
            List<MontoPorConcepto> listadoDeMontosPorConceptos= new List<MontoPorConcepto>();
            
            //listado de nombresDeConceptos
            List<String> nombresDeConceptos = new List<String>();

            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[hoja];
                int colCount = worksheet.Dimension.End.Column; 
                int rowCount = worksheet.Dimension.End.Row;

       

                for (int row = 3; row <= rowCount; row++)//row solia ser 1
                {         

                    String columnaDeNombre = worksheet.Cells[row, 9].Value?.ToString().Trim();

                    if (columnaDeNombre!="" && columnaDeNombre!=null)
                    {

                    
                    char ultimoCaracter = columnaDeNombre[columnaDeNombre.Length - 1];
                    String ultimaLetra = ultimoCaracter.ToString();

                        //modificar aqui para considerar conceptos extra a restar
                        List<String> listadoDeConceptosARestar = new List<string>();

                        listadoDeConceptosARestar.Add("Aporte a CCAF");
                        listadoDeConceptosARestar.Add("Asignacion Familiar Retroactiva");
                        listadoDeConceptosARestar.Add("Cargas Familiares Invalidas");
                        listadoDeConceptosARestar.Add("Cargas Familiares Maternales");
                        listadoDeConceptosARestar.Add("Cargas Familiares Simples");
                        listadoDeConceptosARestar.Add("Desc Dif Cargas Familiares");
                        listadoDeConceptosARestar.Add("Reintegro Cargas Familiares");
                        //agregado el 12 de Octubre del 2022
                        listadoDeConceptosARestar.Add("Bono Especial Temporada R");


                        if (ultimaLetra=="R" || listadoDeConceptosARestar.Contains(columnaDeNombre))//columnaDeNombre== "Aporte a CCAF")//Agregar solo los montos por concepto cuyo nombre de concepto termine en R,  sea Aporte a CCAF
                            //o una de las asignaciones familiares que se restan
                    {
                        MontoPorConcepto mpc = new MontoPorConcepto();
                        mpc.Concepto = worksheet.Cells[row, 9].Value?.ToString().Trim();//nombre del concepto
                        mpc.Nombre = worksheet.Cells[row, 2].Value?.ToString().Trim();//nombre
                        mpc.Empleado = worksheet.Cells[row, 1].Value?.ToString().Trim();//rut
                        mpc.FechaProceso = worksheet.Cells[row, 8].Value?.ToString().Trim();//fechaProceso
                        mpc.Monto = int.Parse(worksheet.Cells[row, 10].Value?.ToString().Trim());//monto

                        mpc.ApellidoPaterno = worksheet.Cells[row, 3].Value?.ToString().Trim();//apellido Ppaterno
                        mpc.ApellidoMaterno = worksheet.Cells[row, 4].Value?.ToString().Trim();//apellido materno
                        mpc.CentroCosto = worksheet.Cells[row, 7].Value?.ToString().Trim();//centro de costo
                        mpc.Cargo = worksheet.Cells[row, 6].Value?.ToString().Trim();//cargo del trabajador
                        mpc.Empresa = worksheet.Cells[row, 5].Value?.ToString().Trim();//empresa del trabajador


                       listadoDeMontosPorConceptos.Add(mpc);

                        nombresDeConceptos.Add(worksheet.Cells[row, 9].Value?.ToString().Trim());
                        }

                    }

                }

                nombresDeConceptos= nombresDeConceptos.Distinct().ToList();

                foreach (var item in nombresDeConceptos)
                {
                    Console.WriteLine(item);
                }            
            }

            return listadoDeMontosPorConceptos;
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
                        r.Total_aportes= worksheet.Cells[row, 12].Value?.ToString().Trim();

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

        private List<RegistroTotalesComoString> procesarRegistrosMensuales(List<RegistroMensualDeTrabajador> registrosMensualesDeTrabajadores, string FilePath)
        {

            List<RegistroTotalesComoString> listadoDeRegistrosDeTotales = new List<RegistroTotalesComoString>();

            List<String> procesos = new List<String>();

            //agregado el 02/11/2022, para asegurar la continuidad a lo largo de los anios, con los conceptos ya establecidos
            DateTime dt = DateTime.Now;         
            int anioActual = dt.Year + 1;

            //esta seccion determina por cuantos anios se itera
            for (int i = 2022; i < anioActual; i++)//for (int i = 2022; i < 2023; i++)
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

            //Leer todos los conceptos UNA VEZ
            List<MontoPorConcepto> listadoDeConceptosEnMasa = leerHojaDeConceptos(FilePath, 1);
            List<String> nombresDeConceptos = new List<String>();
            foreach (var item in listadoDeConceptosEnMasa)
            {
                nombresDeConceptos.Add(item.Concepto);
              
            }

            List<String> nuevaListaDeNombresDeConceptos = nombresDeConceptos.Distinct().ToList<String>();

            int contadorConceptos = 0;
            foreach (var item in nuevaListaDeNombresDeConceptos)
            {
                contadorConceptos++;
            }

            //Cantidad de conceptos en listado (que terminan en R o son de Aporte a CCAF)
            Console.WriteLine(contadorConceptos.ToString());


            List<MontoPorConcepto> listadoDeConceptosDeCuadro = new List<MontoPorConcepto>();


            foreach (var procesoActual in procesos)
            {

                List<TotalDeConcepto> totalesDeConceptos = new List<TotalDeConcepto>();

                foreach (var item in nuevaListaDeNombresDeConceptos)
                {
                    TotalDeConcepto conceptoAAgregarDeCurico = new TotalDeConcepto(item, "Curico", 0);
                    TotalDeConcepto conceptoAAgregarDeInterplanta = new TotalDeConcepto(item, "Interplanta", 0);
                    TotalDeConcepto conceptoAAgregarDeRancagua = new TotalDeConcepto(item, "Rancagua", 0);
                    TotalDeConcepto conceptoAAgregarDeMelipilla = new TotalDeConcepto(item, "Melipilla", 0);
                    TotalDeConcepto conceptoAAgregarDeSanAntonio = new TotalDeConcepto(item, "San Antonio", 0);
                    TotalDeConcepto conceptoAAgregarDeIllapel = new TotalDeConcepto(item, "Illapel", 0);
                    TotalDeConcepto conceptoAAgregarDeSantiago = new TotalDeConcepto(item, "Santiago", 0);
                    TotalDeConcepto conceptoAAgregarDeMovilizadores = new TotalDeConcepto(item, "Movilizadores", 0);
                    TotalDeConcepto conceptoAAgregarDeAdministracion = new TotalDeConcepto(item, "Administracion", 0);
                    TotalDeConcepto conceptoAAgregarDeEmprendedores = new TotalDeConcepto(item, "Emprendedores", 0);

                    totalesDeConceptos.Add(conceptoAAgregarDeCurico);
                    totalesDeConceptos.Add(conceptoAAgregarDeInterplanta);
                    totalesDeConceptos.Add(conceptoAAgregarDeRancagua);
                    totalesDeConceptos.Add(conceptoAAgregarDeMelipilla);
                    totalesDeConceptos.Add(conceptoAAgregarDeSanAntonio);
                    totalesDeConceptos.Add(conceptoAAgregarDeIllapel);
                    totalesDeConceptos.Add(conceptoAAgregarDeSantiago);
                    totalesDeConceptos.Add(conceptoAAgregarDeMovilizadores);
                    totalesDeConceptos.Add(conceptoAAgregarDeAdministracion);
                    totalesDeConceptos.Add(conceptoAAgregarDeEmprendedores);

                }




            RegistroDeTotales registroProceso = new RegistroDeTotales(procesoActual, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroEspacio = new RegistroDeTotales("", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 , listadoDeConceptosDeCuadro);
            RegistroDeTotales registroCurico = new RegistroDeTotales("Curico",0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0, 0, 0, 0, 0, 0, 0,0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroInterplanta = new RegistroDeTotales("Interplanta", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroRancagua = new RegistroDeTotales("Rancagua", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroTaller = new RegistroDeTotales("Taller", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, listadoDeConceptosDeCuadro);// taller serian todos los trabajadores que sean nocheros o mecanicos, independiente del centro 
            RegistroDeTotales registroMelipilla = new RegistroDeTotales("Melipilla", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroSanAntonio = new RegistroDeTotales("San Antonio", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroIllapel = new RegistroDeTotales("Illapel", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroSantiago = new RegistroDeTotales("Santiago", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,listadoDeConceptosDeCuadro);
            RegistroDeTotales registroMovilizadores = new RegistroDeTotales("Movilizadores", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroAdministracion = new RegistroDeTotales("Administracion", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroEmprendedores = new RegistroDeTotales("Emprendedores", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroEspacio2 = new RegistroDeTotales("", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, listadoDeConceptosDeCuadro);
            RegistroDeTotales registroEspacio3 = new RegistroDeTotales("", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,listadoDeConceptosDeCuadro);

                //listado de conceptos a restar

                List<String> listadoDeConceptosARestar = new List<string>();

   
                listadoDeConceptosARestar.Add("Aporte a CCAF");
                listadoDeConceptosARestar.Add("Asignacion Familiar Retroactiva");
                listadoDeConceptosARestar.Add("Cargas Familiares Invalidas");
                listadoDeConceptosARestar.Add("Cargas Familiares Maternales");
                listadoDeConceptosARestar.Add("Cargas Familiares Simples");
                listadoDeConceptosARestar.Add("Desc Dif Cargas Familiares");
                listadoDeConceptosARestar.Add("Reintegro Cargas Familiares");




                foreach (var item in registrosMensualesDeTrabajadores)
            {

                    if (item.Proceso == procesoActual)
                    {                    

                       if (item.Nombre_centro_costo == "CURICO" || item.Nombre_centro_costo == "CURICO E2")
                        {

                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {


                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto== "Bono Tiempo Espera R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalBonoTiempoEsperaR = registroCurico.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;                           
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalBonoEstacionalR = registroCurico.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalBtnLR = registroCurico.TotalBtnLR + mpcSincoFlet.Monto;
                                } //11/5/2022 se agregan 2 bonos: "Bono Sobre Esfuerzo R" y  "VIATICO POR AHORRO" (Viatico Ahorro R)
                                  //SEGUIR DESDE AQUI
                                  else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalBonoSobreEsfuerzoR = registroCurico.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalViaticoAhorroR = registroCurico.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                    
                                }
                                // 02/06/2022 se agrega otro bono: "Bono compensatorio R"
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalBonoCompensatorioR = registroCurico.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                // 12/10/2022 se agrega otro bono: "Bono Especial Temporada R" TotalBonoEspecialTemporadaR
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroCurico.TotalR = registroCurico.TotalR + mpcSincoFlet.Monto;
                                    registroCurico.TotalBonoEspecialTemporadaR = registroCurico.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/08/2022, resulta que hay más de un concepto que debe restarse. Además de "Aporte a CCAF", hay que
                                //restar "Asignacion Familiar Retroactiva", "Cargas Familiares Invalidas", "Cargas Familiares Maternales", "Cargas Familiares Simples", "Desc Dif Cargas Familiares", "Reintegro Cargas Familiares"

                                //inicio de estructura para manejar conceptos a restar

                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroCurico.TotalRemuneracionesAyudantes = registroCurico.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroCurico.TotalRemuneracionesAyudantes = registroCurico.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroCurico.TotalRemuneracionesConductores = registroCurico.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroCurico.TotalRemuneracionesConductores = registroCurico.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroCurico.TotalRemuneracionesOtros = registroCurico.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos




                            }



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

                                    
                                    registroCurico.TotalRemuneracionesAyudantes = registroCurico.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) +int.Parse(item.Total_aportes));      

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

                                    registroCurico.TotalRemuneracionesAyudantes = registroCurico.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroCurico.TotalRemuneracionesConductores = registroCurico.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroCurico.TotalRemuneracionesConductores = registroCurico.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroCurico.TotalRemuneracionesOtros = registroCurico.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroCurico.TotalRemuneracionesDeTodosLosTrabajadores = registroCurico.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }

                        }


                        if (item.Nombre_centro_costo == "ILLAPEL")
                        {
                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalBonoTiempoEsperaR = registroIllapel.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalBonoEstacionalR = registroIllapel.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalBtnLR = registroIllapel.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalBonoSobreEsfuerzoR = registroIllapel.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalViaticoAhorroR = registroIllapel.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalBonoCompensatorioR = registroIllapel.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroIllapel.TotalR = registroIllapel.TotalR + mpcSincoFlet.Monto;
                                    registroIllapel.TotalBonoEspecialTemporadaR = registroIllapel.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroIllapel.TotalRemuneracionesAyudantes = registroIllapel.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroIllapel.TotalRemuneracionesAyudantes = registroIllapel.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroIllapel.TotalRemuneracionesConductores = registroIllapel.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroIllapel.TotalRemuneracionesConductores = registroIllapel.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroIllapel.TotalRemuneracionesOtros = registroIllapel.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos

                            }


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

                                    registroIllapel.TotalRemuneracionesAyudantes = registroIllapel.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroIllapel.TotalRemuneracionesAyudantes = registroIllapel.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroIllapel.TotalRemuneracionesConductores = registroIllapel.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroIllapel.TotalRemuneracionesConductores = registroIllapel.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroIllapel.TotalRemuneracionesOtros = registroIllapel.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores = registroIllapel.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }


                        }
                        if (item.Nombre_centro_costo == "INTERPLANTA" || item.Nombre_centro_costo == "INTERPLANTA E2")
                        {

                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalBonoTiempoEsperaR = registroInterplanta.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalBonoEstacionalR = registroInterplanta.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalBtnLR = registroInterplanta.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalBonoSobreEsfuerzoR = registroInterplanta.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalViaticoAhorroR = registroInterplanta.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalBonoCompensatorioR = registroInterplanta.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroInterplanta.TotalR = registroInterplanta.TotalR + mpcSincoFlet.Monto;
                                    registroInterplanta.TotalBonoEspecialTemporadaR = registroInterplanta.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroInterplanta.TotalRemuneracionesAyudantes = registroInterplanta.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroInterplanta.TotalRemuneracionesAyudantes = registroInterplanta.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroInterplanta.TotalRemuneracionesConductores = registroInterplanta.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroInterplanta.TotalRemuneracionesConductores = registroInterplanta.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroInterplanta.TotalRemuneracionesOtros = registroInterplanta.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos

                            }



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

                                    registroInterplanta.TotalRemuneracionesAyudantes = registroInterplanta.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroInterplanta.TotalRemuneracionesAyudantes = registroInterplanta.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroInterplanta.TotalRemuneracionesConductores = registroInterplanta.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroInterplanta.TotalRemuneracionesConductores = registroInterplanta.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    break;
                              
                              
                                default:
                                    registroInterplanta.TotalDotacion = registroInterplanta.TotalDotacion + 1;
                                    registroInterplanta.TotalApoyos = registroInterplanta.TotalApoyos + 1;

                                    

                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroInterplanta.CantidadDeApoyosDeLicencia = registroInterplanta.CantidadDeApoyosDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroInterplanta.CantidadDeApoyosActivos = registroInterplanta.CantidadDeApoyosActivos + 1;

                                    }

                                    registroInterplanta.TotalRemuneracionesOtros = registroInterplanta.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores = registroInterplanta.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }

                         


                        }
                        if (item.Nombre_centro_costo == "MELIPILLA" || item.Nombre_centro_costo == "MELIPILLA E2")
                        {

                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalBonoTiempoEsperaR = registroMelipilla.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalBonoEstacionalR = registroMelipilla.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalBtnLR = registroMelipilla.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalBonoSobreEsfuerzoR = registroMelipilla.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalViaticoAhorroR = registroMelipilla.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalBonoCompensatorioR = registroMelipilla.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroMelipilla.TotalR = registroMelipilla.TotalR + mpcSincoFlet.Monto;
                                    registroMelipilla.TotalBonoEspecialTemporadaR = registroMelipilla.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroMelipilla.TotalRemuneracionesAyudantes = registroMelipilla.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroMelipilla.TotalRemuneracionesAyudantes = registroMelipilla.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroMelipilla.TotalRemuneracionesConductores = registroMelipilla.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroMelipilla.TotalRemuneracionesConductores = registroMelipilla.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroMelipilla.TotalRemuneracionesOtros = registroMelipilla.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos

                            }



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

                                    registroMelipilla.TotalRemuneracionesAyudantes = registroMelipilla.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroMelipilla.TotalRemuneracionesAyudantes = registroMelipilla.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroMelipilla.TotalRemuneracionesConductores = registroMelipilla.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroMelipilla.TotalRemuneracionesConductores = registroMelipilla.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroMelipilla.TotalRemuneracionesOtros = registroMelipilla.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores = registroMelipilla.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }

         

                        }
                        if (item.Nombre_centro_costo == "RANCAGUA" || item.Nombre_centro_costo == "RANCAGUA  E2")
                        {

                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso==mpcSincoFlet.FechaProceso &&  item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalBonoTiempoEsperaR = registroRancagua.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalBonoEstacionalR = registroRancagua.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalBtnLR = registroRancagua.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalBonoSobreEsfuerzoR = registroRancagua.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalViaticoAhorroR = registroRancagua.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalBonoCompensatorioR = registroRancagua.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroRancagua.TotalR = registroRancagua.TotalR + mpcSincoFlet.Monto;
                                    registroRancagua.TotalBonoEspecialTemporadaR = registroRancagua.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroRancagua.TotalRemuneracionesAyudantes = registroRancagua.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroRancagua.TotalRemuneracionesAyudantes = registroRancagua.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroRancagua.TotalRemuneracionesConductores = registroRancagua.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroRancagua.TotalRemuneracionesConductores = registroRancagua.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroRancagua.TotalRemuneracionesOtros = registroRancagua.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos

                            }




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

                                    registroRancagua.TotalRemuneracionesAyudantes = registroRancagua.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroRancagua.TotalRemuneracionesAyudantes = registroRancagua.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroRancagua.TotalRemuneracionesConductores = registroRancagua.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroRancagua.TotalRemuneracionesConductores = registroRancagua.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    break;
                             
                              //el trabajador esta en rancagua y su cargo es jefe de mantencion = se asigna a taller
                              //actualizacion 28/03/2022, Antonio Alonso solicita tratar al jefe de mantencion de Rancagua como un administrativo más
                                //case "JEFE MANTENCION":
                                //    registroTaller.TotalApoyos = registroTaller.TotalApoyos + 1;
                                //    registroTaller.TotalDotacion = registroTaller.TotalDotacion + 1;

                                //    if (item.Imponible_sin_tope == "0")
                                //    {
                                //        registroTaller.CantidadDeApoyosDeLicencia = registroTaller.CantidadDeApoyosDeLicencia + 1;

                                //    }
                                //    else
                                //    {
                                //        registroTaller.CantidadDeApoyosActivos = registroTaller.CantidadDeApoyosActivos + 1;

                                //    }

                                //    registroTaller.TotalRemuneracionesOtros = registroTaller.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                //    registroTaller.TotalRemuneracionesDeTodosLosTrabajadores = registroTaller.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                  
                                //    break;
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

                                    registroRancagua.TotalRemuneracionesOtros = registroRancagua.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores = registroRancagua.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "SAN ANTONIO" || item.Nombre_centro_costo == "SAN ANTONIO E2")
                        {
                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalBonoTiempoEsperaR = registroSanAntonio.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalBonoEstacionalR = registroSanAntonio.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalBtnLR = registroSanAntonio.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalBonoSobreEsfuerzoR = registroSanAntonio.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalViaticoAhorroR = registroSanAntonio.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalBonoCompensatorioR = registroSanAntonio.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroSanAntonio.TotalR = registroSanAntonio.TotalR + mpcSincoFlet.Monto;
                                    registroSanAntonio.TotalBonoEspecialTemporadaR = registroSanAntonio.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroSanAntonio.TotalRemuneracionesAyudantes = registroSanAntonio.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroSanAntonio.TotalRemuneracionesAyudantes = registroSanAntonio.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroSanAntonio.TotalRemuneracionesConductores = registroSanAntonio.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroSanAntonio.TotalRemuneracionesConductores = registroSanAntonio.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroSanAntonio.TotalRemuneracionesOtros = registroSanAntonio.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos


                            }


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

                                    registroSanAntonio.TotalRemuneracionesAyudantes = registroSanAntonio.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroSanAntonio.TotalRemuneracionesAyudantes = registroSanAntonio.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroSanAntonio.TotalRemuneracionesConductores = registroSanAntonio.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroSanAntonio.TotalRemuneracionesConductores = registroSanAntonio.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroSanAntonio.TotalRemuneracionesOtros = registroSanAntonio.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }

                        }


                        if (item.Nombre_centro_costo == "MOVILIZADORES" || item.Nombre_centro_costo == "MOVILIZADORES E2")
                        {
                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalBonoTiempoEsperaR = registroMovilizadores.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalBonoEstacionalR = registroMovilizadores.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalBtnLR = registroMovilizadores.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalBonoSobreEsfuerzoR = registroMovilizadores.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalViaticoAhorroR = registroMovilizadores.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalBonoCompensatorioR = registroMovilizadores.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                    registroMovilizadores.TotalBonoEspecialTemporadaR = registroMovilizadores.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }

                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroMovilizadores.TotalRemuneracionesAyudantes = registroMovilizadores.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroMovilizadores.TotalRemuneracionesAyudantes = registroMovilizadores.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroSanAntonio.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroMovilizadores.TotalRemuneracionesConductores = registroMovilizadores.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroMovilizadores.TotalRemuneracionesConductores = registroMovilizadores.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroMovilizadores.TotalRemuneracionesOtros = registroMovilizadores.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos


                            }


                            switch (item.Nombre_cargo)
                            {
                                case "AYUDANTE CHOFER":
                                    registroMovilizadores.TotalAyudantes = registroMovilizadores.TotalAyudantes + 1;
                                    registroMovilizadores.TotalDotacion = registroMovilizadores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMovilizadores.CantidadDeAyudantesDeLicencia = registroMovilizadores.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMovilizadores.CantidadDeAyudantesActivos = registroMovilizadores.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroMovilizadores.TotalRemuneracionesAyudantes = registroMovilizadores.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                                case "AYUDANTE CHOFER E2":
                                    registroMovilizadores.TotalAyudantes = registroMovilizadores.TotalAyudantes + 1;
                                    registroMovilizadores.TotalDotacion = registroMovilizadores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMovilizadores.CantidadDeAyudantesDeLicencia = registroMovilizadores.CantidadDeAyudantesDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMovilizadores.CantidadDeAyudantesActivos = registroMovilizadores.CantidadDeAyudantesActivos + 1;

                                    }

                                    registroMovilizadores.TotalRemuneracionesAyudantes = registroMovilizadores.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    break;
                                case "CHOFER PORTEO":
                                    registroMovilizadores.TotalConductores = registroMovilizadores.TotalConductores + 1;
                                    registroMovilizadores.TotalDotacion = registroMovilizadores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMovilizadores.CantidadDeConductoresDeLicencia = registroMovilizadores.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMovilizadores.CantidadDeConductoresActivos = registroMovilizadores.CantidadDeConductoresActivos + 1;

                                    }

                                    registroMovilizadores.TotalRemuneracionesConductores = registroMovilizadores.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    break;
                                case "CHOFER PORTEO E2":
                                    registroMovilizadores.TotalConductores = registroMovilizadores.TotalConductores + 1;
                                    registroMovilizadores.TotalDotacion = registroMovilizadores.TotalDotacion + 1;
                                    if (item.Imponible_sin_tope == "0")
                                    {
                                        registroMovilizadores.CantidadDeConductoresDeLicencia = registroMovilizadores.CantidadDeConductoresDeLicencia + 1;

                                    }
                                    else
                                    {
                                        registroMovilizadores.CantidadDeConductoresActivos = registroMovilizadores.CantidadDeConductoresActivos + 1;

                                    }

                                    registroMovilizadores.TotalRemuneracionesConductores = registroMovilizadores.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    break;


                                default:
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

                                    registroMovilizadores.TotalRemuneracionesOtros = registroMovilizadores.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                                    {
                                        if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoTiempoEsperaR = registroSantiago.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEstacionalR = registroSantiago.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBtnLR = registroSantiago.TotalBtnLR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoSobreEsfuerzoR = registroSantiago.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                            
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalViaticoAhorroR = registroSantiago.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoCompensatorioR = registroSantiago.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEspecialTemporadaR = registroSantiago.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                        }
                                        // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                        // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                        //inicio de estructura para manejar conceptos a restar
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                        {
                                            switch (item.Nombre_cargo)
                                            {
                                                case "AYUDANTE CHOFER":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "AYUDANTE CHOFER E2":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO E2":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                default:
                                                    //son apoyos
                                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                            }
                                        }
                                        //fin de estructura para manejar resta de conceptos

                                    }


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

                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                                    {
                                        if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoTiempoEsperaR = registroSantiago.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEstacionalR = registroSantiago.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBtnLR = registroSantiago.TotalBtnLR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoSobreEsfuerzoR = registroSantiago.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalViaticoAhorroR = registroSantiago.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoCompensatorioR = registroSantiago.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEspecialTemporadaR = registroSantiago.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                        }
                                        // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                        // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                        //inicio de estructura para manejar conceptos a restar
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                        {
                                            switch (item.Nombre_cargo)
                                            {
                                                case "AYUDANTE CHOFER":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "AYUDANTE CHOFER E2":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO E2":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                default:
                                                    //son apoyos
                                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                            }
                                        }
                                        //fin de estructura para manejar resta de conceptos

                                    }


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

                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                                    {
                                        if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoTiempoEsperaR = registroSantiago.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEstacionalR = registroSantiago.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBtnLR = registroSantiago.TotalBtnLR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoSobreEsfuerzoR = registroSantiago.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalViaticoAhorroR = registroSantiago.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoCompensatorioR = registroSantiago.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEspecialTemporadaR = registroSantiago.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                        }
                                        // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                        // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                        //inicio de estructura para manejar conceptos a restar
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                        {
                                            switch (item.Nombre_cargo)
                                            {
                                                case "AYUDANTE CHOFER":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "AYUDANTE CHOFER E2":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO E2":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                default:
                                                    //son apoyos
                                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                            }
                                        }
                                        //fin de estructura para manejar resta de conceptos

                                    }


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

                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                                    {
                                        if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoTiempoEsperaR = registroSantiago.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEstacionalR = registroSantiago.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBtnLR = registroSantiago.TotalBtnLR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoSobreEsfuerzoR = registroSantiago.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalViaticoAhorroR = registroSantiago.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoCompensatorioR = registroSantiago.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEspecialTemporadaR = registroSantiago.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                        }
                                        // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                        // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                        //inicio de estructura para manejar conceptos a restar
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                        {
                                            switch (item.Nombre_cargo)
                                            {
                                                case "AYUDANTE CHOFER":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "AYUDANTE CHOFER E2":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO E2":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                default:
                                                    //son apoyos
                                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                            }
                                        }
                                        //fin de estructura para manejar resta de conceptos

                                    }


                                    break;
                                    //el trabajador esta en santiago y es un movilizador = se asigna a movilizadores
                                    //a partir del 10/08/2022 debiese haber 2 centros aparte de movilizadores,
                                    //el nombre de esos centros es MOVILIZADORES Y MOVILIZADORES E2
                                //case "MOVILIZADOR":
                                //    registroMovilizadores.TotalApoyos = registroMovilizadores.TotalApoyos + 1;
                                //    registroMovilizadores.TotalDotacion = registroMovilizadores.TotalDotacion + 1;
                                //    if (item.Imponible_sin_tope == "0")
                                //    {
                                //        registroMovilizadores.CantidadDeApoyosDeLicencia = registroMovilizadores.CantidadDeApoyosDeLicencia + 1;

                                //    }
                                //    else
                                //    {
                                //        registroMovilizadores.CantidadDeApoyosActivos = registroMovilizadores.CantidadDeApoyosActivos + 1;

                                //    }

                                //    registroMovilizadores.TotalRemuneracionesOtros = registroMovilizadores.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                //    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                //    foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                                //    {
                                //        if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                //        {
                                //            registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                //            registroMovilizadores.TotalBonoTiempoEsperaR = registroMovilizadores.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                //        }
                                //        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                //        {
                                //            registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                //            registroMovilizadores.TotalBonoEstacionalR = registroMovilizadores.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                //        }
                                //        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                //        {
                                //            registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                //            registroMovilizadores.TotalBtnLR = registroMovilizadores.TotalBtnLR + mpcSincoFlet.Monto;
                                //        }
                                //        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                //        {
                                //            registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                //            registroMovilizadores.TotalBonoSobreEsfuerzoR = registroMovilizadores.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                //        }
                                //        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                //        {
                                //            registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                //            registroMovilizadores.TotalViaticoAhorroR = registroMovilizadores.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                //        }
                                //        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                //        {
                                //            registroMovilizadores.TotalR = registroMovilizadores.TotalR + mpcSincoFlet.Monto;
                                //            registroMovilizadores.TotalBonoCompensatorioR = registroMovilizadores.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                //        }

                                //        // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                //        // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //        //inicio de estructura para manejar conceptos a restar
                                //        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                //        {
                                //            switch (item.Nombre_cargo)
                                //            {
                                //                case "AYUDANTE CHOFER":
                                //                    registroMovilizadores.TotalRemuneracionesAyudantes = registroMovilizadores.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                //                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                //                    break;
                                //                case "AYUDANTE CHOFER E2":
                                //                    registroMovilizadores.TotalRemuneracionesAyudantes = registroMovilizadores.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                //                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                //                    break;
                                //                case "CHOFER PORTEO":
                                //                    registroMovilizadores.TotalRemuneracionesConductores = registroMovilizadores.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                //                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                //                    break;
                                //                case "CHOFER PORTEO E2":
                                //                    registroMovilizadores.TotalRemuneracionesConductores = registroMovilizadores.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                //                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                //                    break;
                                //                default:
                                //                    //son apoyos
                                //                    registroMovilizadores.TotalRemuneracionesOtros = registroMovilizadores.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                //                    registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores = registroMovilizadores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                //                    break;
                                //            }
                                //        }
                                //        //fin de estructura para manejar resta de conceptos

                                //    }



                                //    break;
                               
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

                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                                    {
                                        if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoTiempoEsperaR = registroSantiago.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEstacionalR = registroSantiago.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBtnLR = registroSantiago.TotalBtnLR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoSobreEsfuerzoR = registroSantiago.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalViaticoAhorroR = registroSantiago.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoCompensatorioR = registroSantiago.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                        }
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                        {
                                            registroSantiago.TotalR = registroSantiago.TotalR + mpcSincoFlet.Monto;
                                            registroSantiago.TotalBonoEspecialTemporadaR = registroSantiago.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                        }
                                        // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                        // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                        //inicio de estructura para manejar conceptos a restar
                                        else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                        {
                                            switch (item.Nombre_cargo)
                                            {
                                                case "AYUDANTE CHOFER":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "AYUDANTE CHOFER E2":
                                                    registroSantiago.TotalRemuneracionesAyudantes = registroSantiago.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                case "CHOFER PORTEO E2":
                                                    registroSantiago.TotalRemuneracionesConductores = registroSantiago.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                                default:
                                                    //son apoyos
                                                    registroSantiago.TotalRemuneracionesOtros = registroSantiago.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                                    registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores = registroSantiago.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                                    break;
                                            }
                                        }
                                        //fin de estructura para manejar resta de conceptos

                                    }


                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "CENTRAL" || item.Nombre_centro_costo == "CENTRAL E2")
                        {

                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalBonoTiempoEsperaR = registroAdministracion.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalBonoEstacionalR = registroAdministracion.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalBtnLR = registroAdministracion.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalBonoSobreEsfuerzoR = registroAdministracion.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalViaticoAhorroR = registroAdministracion.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalBonoCompensatorioR = registroAdministracion.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroAdministracion.TotalR = registroAdministracion.TotalR + mpcSincoFlet.Monto;
                                    registroAdministracion.TotalBonoEspecialTemporadaR = registroAdministracion.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }
                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroAdministracion.TotalRemuneracionesAyudantes = registroAdministracion.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroAdministracion.TotalRemuneracionesAyudantes = registroAdministracion.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroAdministracion.TotalRemuneracionesConductores = registroAdministracion.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroAdministracion.TotalRemuneracionesConductores = registroAdministracion.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroAdministracion.TotalRemuneracionesOtros = registroAdministracion.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos

                            }



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

                                    registroAdministracion.TotalRemuneracionesAyudantes = registroAdministracion.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroAdministracion.TotalRemuneracionesAyudantes = registroAdministracion.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroAdministracion.TotalRemuneracionesConductores = registroAdministracion.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroAdministracion.TotalRemuneracionesConductores = registroAdministracion.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    break;
                                    
                                    //el trabajador esta en central y es nochero = se asigna a taller
                                    //actualizacion 28/03/2022, Antonio Alonso solicita que nocheros de central pasen a ser del centro de Administracion
                                case "NOCHERO":
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

                                    registroAdministracion.TotalRemuneracionesOtros = registroAdministracion.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                            

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

                                    registroAdministracion.TotalRemuneracionesOtros = registroAdministracion.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }

                        }
                        if (item.Nombre_centro_costo == "EMPRENDEDOR")//es EMPRENDEDOR, NO EMPRENDEDORES
                        {

                            foreach (var mpcSincoFlet in listadoDeConceptosEnMasa)
                            {
                                if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Tiempo Espera R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalBonoTiempoEsperaR = registroEmprendedores.TotalBonoTiempoEsperaR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono estacional R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalBonoEstacionalR = registroEmprendedores.TotalBonoEstacionalR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Btn I R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalBtnLR = registroEmprendedores.TotalBtnLR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Sobre Esfuerzo R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalBonoSobreEsfuerzoR = registroEmprendedores.TotalBonoSobreEsfuerzoR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "VIATICO POR AHORRO R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalViaticoAhorroR = registroEmprendedores.TotalViaticoAhorroR + mpcSincoFlet.Monto;
                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Compensatorio R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalBonoCompensatorioR = registroEmprendedores.TotalBonoCompensatorioR + mpcSincoFlet.Monto;

                                }
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && mpcSincoFlet.Concepto == "Bono Especial Temporada R")
                                {
                                    registroEmprendedores.TotalR = registroEmprendedores.TotalR + mpcSincoFlet.Monto;
                                    registroEmprendedores.TotalBonoEspecialTemporadaR = registroEmprendedores.TotalBonoEspecialTemporadaR + mpcSincoFlet.Monto;

                                }
                                // 08/07/2022 se agrega concepto a restarse "Aporte a CCAF"
                                // 08/09/2022 se modifica concepto a restarse; en vez de ser uno solo, ahora es un listado de conceptos


                                //inicio de estructura para manejar conceptos a restar
                                else if (item.Proceso == mpcSincoFlet.FechaProceso && item.Empleado == mpcSincoFlet.Empleado && listadoDeConceptosARestar.Contains(mpcSincoFlet.Concepto))
                                {
                                    switch (item.Nombre_cargo)
                                    {
                                        case "AYUDANTE CHOFER":
                                            registroEmprendedores.TotalRemuneracionesAyudantes = registroEmprendedores.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "AYUDANTE CHOFER E2":
                                            registroEmprendedores.TotalRemuneracionesAyudantes = registroEmprendedores.TotalRemuneracionesAyudantes - mpcSincoFlet.Monto;
                                            registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO":
                                            registroEmprendedores.TotalRemuneracionesConductores = registroEmprendedores.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        case "CHOFER PORTEO E2":
                                            registroEmprendedores.TotalRemuneracionesConductores = registroEmprendedores.TotalRemuneracionesConductores - mpcSincoFlet.Monto;
                                            registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                        default:
                                            //son apoyos
                                            registroEmprendedores.TotalRemuneracionesOtros = registroEmprendedores.TotalRemuneracionesOtros - mpcSincoFlet.Monto;
                                            registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores - mpcSincoFlet.Monto;
                                            break;
                                    }
                                }
                                //fin de estructura para manejar resta de conceptos

                            }


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

                                    registroEmprendedores.TotalRemuneracionesAyudantes = registroEmprendedores.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

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

                                    registroEmprendedores.TotalRemuneracionesAyudantes = registroEmprendedores.TotalRemuneracionesAyudantes + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroEmprendedores.TotalRemuneracionesConductores = registroEmprendedores.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroEmprendedores.TotalRemuneracionesConductores = registroEmprendedores.TotalRemuneracionesConductores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
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

                                    registroEmprendedores.TotalRemuneracionesOtros = registroEmprendedores.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));
                                    registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores = registroEmprendedores.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento) + int.Parse(item.Total_aportes));

                                    break;
                            }


                        }

                        

                    }

                    
                
            }

                //hasta aqui llegan los datos
               
            
            RegistroTotalesComoString registroProcesoComoString = new RegistroTotalesComoString(registroProceso, "titulo");
            RegistroTotalesComoString registroEspacioComoString = new RegistroTotalesComoString(registroEspacio, "titulo");
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
            RegistroTotalesComoString registroEspacioComoString2 = new RegistroTotalesComoString(registroEspacio2, "titulo");
            RegistroTotalesComoString registroEspacioComoString3 = new RegistroTotalesComoString(registroEspacio3, "titulo");

           

            listadoDeRegistrosDeTotales.Add(registroProcesoComoString);
            listadoDeRegistrosDeTotales.Add(registroEspacioComoString);
            listadoDeRegistrosDeTotales.Add(registroCuricoComoString);
            listadoDeRegistrosDeTotales.Add(registroInterplantaComoString);
            listadoDeRegistrosDeTotales.Add(registroRancaguaComoString);
           // 28/03/2022, registro de taller se considera innecesario, por lo que se quita del listado
           // listadoDeRegistrosDeTotales.Add(registroTallerComoString);
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
            MessageBox.Show("para subir asistencias a rex: * recibir excel de Francisco * Ejecutar programa, seleccionar asistencia o comisiones dependiendo del archivo cargado * Excel generado se crea en descargas (Asistencia o Comisiones de Ayudante o de trabajador), luego eliminar la primera fila de cada Excel (puede que a veces haya que convertir de formato XLSX o XLS a CSV) * Enviar Excels generados al personal de remuneraciones para que ellas hagan la carga.", "Sobre la subida a Rex");
            
            MessageBox.Show("Transformar registros a totales sigue la siguiente lógica: se toma el archivo excel de base, se filtra primero por mes y luego por Centro. Los montos y totales para cada centro se obtienen con esos 2 filtros, salvo 2 excpeciones. La primera es si un trabajador de SANTIAGO o SANTIAGO E2 es un movilizador, en cuyo caso se asigna al centro de movilizadores. La segunda es cuando el trabajador de central es un nochero, en cuyo caso se asigna a administración.", "Sobre el registro de totales, parte 1");
            MessageBox.Show("Desde Mayo del 2022, el programa también es capaz de filtrar valores de conceptos (sólo los solicitados por Eliana Valdes).", "Sobre el registro de totales, parte 2");
            MessageBox.Show("Programa creado por Marcelo Andrés Aranda Tatto, bajo ordenes de Antonio Alonso.", "Sobre el programa");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            mostrarAyuda();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            MessageBox.Show("Debe cargar un Excel con 2 hojas: la primera con TODAS las remuneraciones," +
                " y la segunda con TODOS los conceptos. Vale decir, desde Enero del 2022 a la fecha");

            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                    break;
                }
                else
                {
                    MessageBox.Show("Revisar Excel de carga. Cerrando programa");
                    System.Environment.Exit(0);
                }

            }

            try
            {
                
            List<RegistroMensualDeTrabajador> registros = leerExcelDeRegistroDeTrabajadores(sFileName);

            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";
            
           List<RegistroTotalesComoString> registrosDeTotales = procesarRegistrosMensuales(registros, sFileName);


            var archivo = new FileInfo(downloads + @"\Registro de montos totales.xlsx");

            SaveExcelFileRegistroDeTotales(registrosDeTotales, archivo);

            MessageBox.Show("Archivo Excel llamado Registro de montos totales, creado en carpeta de descargas!");

            }
            catch (Exception)
            {
                MessageBox.Show("Ocurrió un error. Cerrando programa");
                System.Environment.Exit(0);
                throw;
            }



        }
    }
}
