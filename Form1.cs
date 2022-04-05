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

            // 28/03/2022, Antonio Alonso pidio aplicar un formato al Excel

            // hay que hacer un sub cuadro resumen con la siguiente informacion:
            //interplantas (remuneraciones totales)
            //movilizadores (remuneraciones totales)
            //interplantas (remuneraciones totales)

            //tradicional (remuneraciones totales) a su vez se divide en 3 más:
            //directos (ayudantes y conductores), administracion (los que trabajan en el centro de administración)
            //e indirectos (todos los que NO sean conductores, ayudantes o sean de administracion)


            //fijar titulo de columnas
            //ws.Cells["B1"].Value = "Cantidad de conductores activos";
            //ws.Cells["C1"].Value = "Conductores de licencia";
            //ws.Cells["D1"].Value = "Ayudantes activos";
            //ws.Cells["E1"].Value = "Ayudantes de licencia";
            //ws.Cells["F1"].Value = "Apoyos activos";
            //ws.Cells["G1"].Value = "Apoyos de licencia";
            //ws.Cells["H1"].Value = "Total de conductores";
            //ws.Cells["I1"].Value = "Total de ayudantes";
            //ws.Cells["J1"].Value = "Total de apoyos";
            //ws.Cells["K1"].Value = "Total de trabajadores";
            //ws.Cells["L1"].Value = "$ Conductores";
            //ws.Cells["M1"].Value = "$ Ayudantes";
            //ws.Cells["N1"].Value = "$ Apoyos";
            //ws.Cells["O1"].Value = "Total";

            // ws.Cells["A1"].Value = "";
            ws.Cells["B1"].Value = "";
            ws.Cells["C1"].Value = "";
            ws.Cells["D1"].Value = "";
            ws.Cells["E1"].Value = "";
            ws.Cells["F1"].Value = "";
            ws.Cells["G1"].Value = "";
            ws.Cells["H1"].Value = "";
            ws.Cells["I1"].Value = "";
            ws.Cells["J1"].Value = "";
            ws.Cells["K1"].Value = "";
            ws.Cells["L1"].Value = "";
            ws.Cells["M1"].Value = "";
            ws.Cells["N1"].Value = "";
            ws.Cells["O1"].Value = "";





            //fijar  estilo de letra
            ws.Cells["B1"].Style.Font.Bold = true;
            ws.Cells["C1"].Style.Font.Bold = true;
            ws.Cells["D1"].Style.Font.Bold = true;
            ws.Cells["E1"].Style.Font.Bold = true;
            ws.Cells["F1"].Style.Font.Bold = true;
            ws.Cells["G1"].Style.Font.Bold = true;
            ws.Cells["H1"].Style.Font.Bold = true;
            ws.Cells["I1"].Style.Font.Bold = true;
            ws.Cells["J1"].Style.Font.Bold = true;
            ws.Cells["K1"].Style.Font.Bold = true;
            ws.Cells["L1"].Style.Font.Bold = true;
            ws.Cells["M1"].Style.Font.Bold = true;
            ws.Cells["N1"].Style.Font.Bold = true;
            ws.Cells["O1"].Style.Font.Bold = true;

            //fijar color de fondo de ciertas celdas
            //ws.Cells["B1"].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
            //ws.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["D1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["F1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["G1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["H1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["I1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["J1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["K1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["L1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["M1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["N1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            //ws.Cells["O1"].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);





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



            //quitar 0's sobrantes
            for (int i = 1; i < 13; i++)
            {

                
                
                // agregar  bordes a tabla
                ws.Cells["A"+fila2+":O"+fila13].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells["A" + fila2 + ":O" + fila13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells["A" + fila2 + ":O" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells["A" + fila2 + ":O" + fila13].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                ws.Cells["H" + fila2 + ":N" + fila2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                //Arriba
                ws.Cells["B" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["C" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["D" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["E" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["F" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["G" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["H" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["I" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["J" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["K" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["L" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["M" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["N" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                ////izquierda
                ws.Cells["B" + fila4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila7].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila9].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila11].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila12].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["B" + fila13].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                ////abajo
                ws.Cells["B" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["C" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["D" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["E" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["F" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["G" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["H" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["I" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["J" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["K" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["L" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["M" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["N" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila13].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;


                ////derecha
                ws.Cells["O" + fila4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila9].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila12].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["O" + fila13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;


                ws.Cells["B" + fila2].Value = "";
                ws.Cells["B" + fila3].Value = "";
                ws.Cells["B" + fila14].Value = "";
                ws.Cells["B" + fila15].Value = "";
                ws.Cells["C"+fila2].Value = "";
                ws.Cells["C"+fila3].Value = "";
                ws.Cells["C"+fila14].Value = "";
                ws.Cells["C"+fila15].Value = "";
                ws.Cells["D"+fila2].Value = "";
                ws.Cells["D"+fila3].Value = "";
                ws.Cells["D"+fila14].Value = "";
                ws.Cells["D"+fila15].Value = "";
                ws.Cells["E"+fila2].Value = "";
                ws.Cells["E"+fila3].Value = "";
                ws.Cells["E"+fila14].Value = "";
                ws.Cells["E"+fila15].Value = "";
                ws.Cells["F"+fila2].Value = "";
                ws.Cells["F"+fila3].Value = "";
                ws.Cells["F"+fila14].Value = "";
                ws.Cells["F"+fila15].Value = "";
                ws.Cells["G"+fila2].Value = "";
                ws.Cells["G"+fila3].Value = "";
                ws.Cells["G"+fila14].Value = "";
                ws.Cells["G"+fila15].Value = "";
                ws.Cells["H"+fila2].Value = "";
                ws.Cells["H"+fila3].Value = "";
                ws.Cells["H"+fila14].Value = "";
                ws.Cells["H"+fila15].Value = "";
                ws.Cells["I"+fila2].Value = "";
                ws.Cells["I"+fila3].Value = "";
                ws.Cells["I"+fila14].Value = "";
                ws.Cells["I"+fila15].Value = "";
                ws.Cells["J"+fila2].Value = "";
                ws.Cells["J"+fila3].Value = "";
                ws.Cells["J"+fila14].Value = "";
                ws.Cells["J"+fila15].Value = "";
                ws.Cells["K"+fila2].Value = "";
                ws.Cells["K"+fila3].Value = "";
                ws.Cells["K"+fila14].Value = "";
                ws.Cells["K"+fila15].Value = "";
                ws.Cells["L"+fila2].Value = "";
                ws.Cells["L"+fila3].Value = "";
                ws.Cells["L"+fila14].Value = "";
                ws.Cells["L"+fila15].Value = "";
                ws.Cells["M"+fila2].Value = "";
                ws.Cells["M"+fila3].Value = "";
                ws.Cells["M"+fila14].Value = "";
                ws.Cells["M"+fila15].Value = "";
                ws.Cells["N"+fila2].Value = "";
                ws.Cells["N"+fila3].Value = "";
                ws.Cells["N"+fila14].Value = "";
                ws.Cells["N"+fila15].Value = "";
                ws.Cells["O"+fila2].Value = "";
                ws.Cells["O"+fila3].Value = "";
                ws.Cells["O"+fila14].Value = "";
                ws.Cells["O"+fila15].Value = "";
                //agrego esto

                //agregar titulos en cada proceso
                ws.Cells["B" + fila2].Value = "# de Conductores";
                ws.Cells["B" + fila3].Value = "Activos";
                ws.Cells["C" + fila3].Value = "De licencia";

                ws.Cells["D" + fila2].Value = "# de Ayudantes";
                ws.Cells["D" + fila3].Value = "Activos";
                ws.Cells["E" +fila3].Value = "De licencia";

                ws.Cells["F" + fila2].Value = "# de Apoyos";
                ws.Cells["F" + fila3].Value = "Activos";
                ws.Cells["G" + fila3].Value = "De licencia";

                //H, I, J, K
                ws.Cells["H" + fila2].Value = "Total conductores";
                ws.Cells["I" + fila2].Value = "Total ayudantes";
                ws.Cells["J" + fila2].Value = "Total apoyos";
                ws.Cells["K" + fila2].Value = "Total dotación";

                ws.Cells["L" + fila2].Value = "$ Concuctores";
                ws.Cells["M" + fila2].Value = "$ Ayudantes";
                ws.Cells["N" + fila2].Value = "$ Otros";
                ws.Cells["O" + fila2].Value = "Total";


                ws.Cells["B" + fila2].Style.Font.Bold = true;
                ws.Cells["D" + fila2].Style.Font.Bold = true;
                ws.Cells["F" + fila2].Style.Font.Bold = true;
                ws.Cells["H" + fila2].Style.Font.Bold = true;
                ws.Cells["I" + fila2].Style.Font.Bold = true;
                ws.Cells["J" + fila2].Style.Font.Bold = true;
                ws.Cells["K" + fila2].Style.Font.Bold = true;
                ws.Cells["L" + fila2].Style.Font.Bold = true;
                ws.Cells["M" + fila2].Style.Font.Bold = true;
                ws.Cells["N" + fila2].Style.Font.Bold = true;
                ws.Cells["O" + fila2].Style.Font.Bold = true;

                ws.Cells["B" + fila3].Style.Font.Italic = true;
                ws.Cells["C" + fila3].Style.Font.Italic = true;
                ws.Cells["D" + fila3].Style.Font.Italic = true;
                ws.Cells["E" + fila3].Style.Font.Italic = true;
                ws.Cells["F" + fila3].Style.Font.Italic = true;
                ws.Cells["G" + fila3].Style.Font.Italic = true;

                ws.Cells["B" + fila2 + ":C" + fila2].Merge = true;
                ws.Cells["D" + fila2 + ":E" + fila2].Merge = true;
                ws.Cells["F" + fila2 + ":G" + fila2].Merge = true;



                //totales al pie de la tabla
                ws.Cells["A" + fila14].Value = "Totales";
                ws.Cells["B"+fila14].Value = int.Parse(ws.Cells["B"+fila4].Value.ToString()) + int.Parse(ws.Cells["B"+fila5].Value.ToString()) + int.Parse(ws.Cells["B"+fila6].Value.ToString()) + int.Parse(ws.Cells["B"+fila7].Value.ToString()) + int.Parse(ws.Cells["B"+fila8].Value.ToString()) + int.Parse(ws.Cells["B"+fila9].Value.ToString()) + int.Parse(ws.Cells["B"+fila10].Value.ToString()) +  int.Parse(ws.Cells["B"+fila11].Value.ToString()) + int.Parse(ws.Cells["B"+fila12].Value.ToString()) + int.Parse(ws.Cells["B"+fila13].Value.ToString());
                ws.Cells["C" + fila14].Value = int.Parse(ws.Cells["C" + fila4].Value.ToString()) + int.Parse(ws.Cells["C" + fila5].Value.ToString()) + int.Parse(ws.Cells["C" + fila6].Value.ToString()) + int.Parse(ws.Cells["C" + fila7].Value.ToString()) + int.Parse(ws.Cells["C" + fila8].Value.ToString()) + int.Parse(ws.Cells["C" + fila9].Value.ToString()) + int.Parse(ws.Cells["C" + fila10].Value.ToString()) + int.Parse(ws.Cells["C" + fila11].Value.ToString()) + int.Parse(ws.Cells["C" + fila12].Value.ToString()) + int.Parse(ws.Cells["C" + fila13].Value.ToString());
                ws.Cells["D" + fila14].Value = int.Parse(ws.Cells["D" + fila4].Value.ToString()) + int.Parse(ws.Cells["D" + fila5].Value.ToString()) + int.Parse(ws.Cells["D" + fila6].Value.ToString()) + int.Parse(ws.Cells["D" + fila7].Value.ToString()) + int.Parse(ws.Cells["D" + fila8].Value.ToString()) + int.Parse(ws.Cells["D" + fila9].Value.ToString()) + int.Parse(ws.Cells["D" + fila10].Value.ToString()) + int.Parse(ws.Cells["D" + fila11].Value.ToString()) + int.Parse(ws.Cells["D" + fila12].Value.ToString()) + int.Parse(ws.Cells["D" + fila13].Value.ToString());
                ws.Cells["E" + fila14].Value = int.Parse(ws.Cells["E" + fila4].Value.ToString()) + int.Parse(ws.Cells["E" + fila5].Value.ToString()) + int.Parse(ws.Cells["E" + fila6].Value.ToString()) + int.Parse(ws.Cells["E" + fila7].Value.ToString()) + int.Parse(ws.Cells["E" + fila8].Value.ToString()) + int.Parse(ws.Cells["E" + fila9].Value.ToString()) + int.Parse(ws.Cells["E" + fila10].Value.ToString()) + int.Parse(ws.Cells["E" + fila11].Value.ToString()) + int.Parse(ws.Cells["E" + fila12].Value.ToString()) + int.Parse(ws.Cells["E" + fila13].Value.ToString());
                ws.Cells["F" + fila14].Value = int.Parse(ws.Cells["F" + fila4].Value.ToString()) + int.Parse(ws.Cells["F" + fila5].Value.ToString()) + int.Parse(ws.Cells["F" + fila6].Value.ToString()) + int.Parse(ws.Cells["F" + fila7].Value.ToString()) + int.Parse(ws.Cells["F" + fila8].Value.ToString()) + int.Parse(ws.Cells["F" + fila9].Value.ToString()) + int.Parse(ws.Cells["F" + fila10].Value.ToString()) + int.Parse(ws.Cells["F" + fila11].Value.ToString()) + int.Parse(ws.Cells["F" + fila12].Value.ToString()) + int.Parse(ws.Cells["F" + fila13].Value.ToString());
                ws.Cells["G" + fila14].Value = int.Parse(ws.Cells["G" + fila4].Value.ToString()) + int.Parse(ws.Cells["G" + fila5].Value.ToString()) + int.Parse(ws.Cells["G" + fila6].Value.ToString()) + int.Parse(ws.Cells["G" + fila7].Value.ToString()) + int.Parse(ws.Cells["G" + fila8].Value.ToString()) + int.Parse(ws.Cells["G" + fila9].Value.ToString()) + int.Parse(ws.Cells["G" + fila10].Value.ToString()) + int.Parse(ws.Cells["G" + fila11].Value.ToString()) + int.Parse(ws.Cells["G" + fila12].Value.ToString()) + int.Parse(ws.Cells["G" + fila13].Value.ToString());
                
                ws.Cells["O" + fila14].Value = int.Parse(ws.Cells["O" + fila4].Value.ToString()) + int.Parse(ws.Cells["O" + fila5].Value.ToString()) + int.Parse(ws.Cells["O" + fila6].Value.ToString()) + int.Parse(ws.Cells["O" + fila7].Value.ToString()) + int.Parse(ws.Cells["O" + fila8].Value.ToString()) + int.Parse(ws.Cells["O" + fila9].Value.ToString()) + int.Parse(ws.Cells["O" + fila10].Value.ToString()) + int.Parse(ws.Cells["O" + fila11].Value.ToString()) + int.Parse(ws.Cells["O" + fila12].Value.ToString()) + int.Parse(ws.Cells["O" + fila13].Value.ToString());

                //cuadro sub resumen
                ws.Cells["Q" + fila4].Value = "RESUMEN DE MODELOS";

                ws.Cells["Q" + fila4].Style.Border.Left.Style= OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;


                ws.Cells["Q" + fila5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila7].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila9].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["Q" + fila10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                ws.Cells["R" + fila5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila9].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                ws.Cells["R" + fila10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;


                // titulos de cuadro resumen
                ws.Cells["Q" + fila5].Value = "INTERPLANTAS";
                ws.Cells["Q" + fila5].Style.Font.Bold = true;
                ws.Cells["Q" + fila6].Value = "MOVILIZADORES";
                ws.Cells["Q" + fila6].Style.Font.Bold = true;
                ws.Cells["Q" + fila7].Value = "EMPRENDEDORES";
                ws.Cells["Q" + fila7].Style.Font.Bold = true;
                ws.Cells["Q" + fila8].Value = "DIRECTOS";
                ws.Cells["Q" + fila8].Style.Font.Bold = true;
                ws.Cells["Q" + fila9].Value = "INDIRECTOS";
                ws.Cells["Q" + fila9].Style.Font.Bold = true;
                ws.Cells["Q" + fila10].Value = "ADMINISTRACION";
                ws.Cells["Q" + fila10].Style.Font.Bold = true;
                ws.Cells["Q" + fila11].Value = "$ TOTAL ";
                ws.Cells["Q" + fila11].Style.Font.Bold = true;
                ws.Cells["Q" + fila12].Value = "TOTAL TRABAJADORES";
                ws.Cells["Q" + fila12].Style.Font.Bold = true;


                //valores de cuadro resumen (interplanta, movilizadores y emprendedores)
                ws.Cells["R" + fila5].Value = ws.Cells["O" + fila5].Value;
                ws.Cells["R" + fila6].Value = ws.Cells["O" + fila11].Value;
                ws.Cells["R" + fila7].Value = ws.Cells["O" + fila12].Value;




                int remuneracionesDirectosCurico = int.Parse(ws.Cells["L"+fila4].Value.ToString()) + int.Parse(ws.Cells["M" + fila4].Value.ToString());
                int remuneracionesDirectosRancagua = int.Parse(ws.Cells["L" + fila6].Value.ToString()) + int.Parse(ws.Cells["M" + fila6].Value.ToString());
                int remuneracionesDirectosMelipilla = int.Parse(ws.Cells["L" + fila7].Value.ToString()) + int.Parse(ws.Cells["M" + fila7].Value.ToString());
                int remuneracionesDirectosSanAntonio = int.Parse(ws.Cells["L" + fila8].Value.ToString()) + int.Parse(ws.Cells["M" + fila8].Value.ToString());
                int remuneracionesDirectosIllapel = int.Parse(ws.Cells["L" + fila9].Value.ToString()) + int.Parse(ws.Cells["M" + fila9].Value.ToString());
                int remuneracionesDirectosSantiago = int.Parse(ws.Cells["L" + fila10].Value.ToString()) + int.Parse(ws.Cells["M" + fila10].Value.ToString());

                int remuneracionesIndirectosCurico = int.Parse(ws.Cells["N" + fila4].Value.ToString());
                int remuneracionesIndirectosRancagua = int.Parse(ws.Cells["N" + fila6].Value.ToString());
                int remuneracionesIndirectosMelipilla = int.Parse(ws.Cells["N" + fila7].Value.ToString());
                int remuneracionesIndirectosSanAntonio = int.Parse(ws.Cells["N" + fila8].Value.ToString());
                int remuneracionesIndirectosIllapel = int.Parse(ws.Cells["N" + fila9].Value.ToString());
                int remuneracionesIndirectosSantiago = int.Parse(ws.Cells["N" + fila10].Value.ToString());


                int sumaDeDirectos = remuneracionesDirectosCurico + remuneracionesDirectosRancagua + remuneracionesDirectosMelipilla + remuneracionesDirectosSanAntonio + remuneracionesDirectosIllapel + remuneracionesDirectosSantiago;
                int sumaDeIndirectos = remuneracionesIndirectosCurico + remuneracionesIndirectosRancagua + remuneracionesIndirectosMelipilla + remuneracionesIndirectosSanAntonio + remuneracionesIndirectosIllapel + remuneracionesIndirectosSantiago;


                //valores de modelo tradicional (directos e indirectos)
                ws.Cells["R" + fila8].Value = sumaDeDirectos;
                ws.Cells["R" + fila9].Value = sumaDeIndirectos;
                //valor de cuadro resumen (administracion)
                ws.Cells["R" + fila10].Value = ws.Cells["O" + fila13].Value;

                //valor de celda de totales
                ws.Cells["R" + fila11].Value = sumaDeDirectos + sumaDeIndirectos + int.Parse(ws.Cells["R" + fila10].Value.ToString()) + int.Parse(ws.Cells["R" + fila5].Value.ToString()) + int.Parse(ws.Cells["R" + fila6].Value.ToString()) + int.Parse(ws.Cells["R" + fila7].Value.ToString());


                //formatear como numero usando comas como separadores de decimales
                ws.Cells["L" + fila4 + ":O" + fila14].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                ws.Cells["R" + fila5 + ":R" + fila11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                //total de trabajadores

                int conductoresActivos = int.Parse(ws.Cells["B" + fila4].Value.ToString()) + int.Parse(ws.Cells["B" + fila5].Value.ToString()) + int.Parse(ws.Cells["B" + fila6].Value.ToString()) + int.Parse(ws.Cells["B" + fila7].Value.ToString()) + int.Parse(ws.Cells["B" + fila8].Value.ToString()) + int.Parse(ws.Cells["B" + fila9].Value.ToString()) + int.Parse(ws.Cells["B" + fila10].Value.ToString()) + int.Parse(ws.Cells["B" + fila11].Value.ToString()) + int.Parse(ws.Cells["B" + fila12].Value.ToString()) + int.Parse(ws.Cells["B" + fila13].Value.ToString());
                int conductoresDeLicencia = int.Parse(ws.Cells["C" + fila4].Value.ToString()) + int.Parse(ws.Cells["C" + fila5].Value.ToString()) + int.Parse(ws.Cells["C" + fila6].Value.ToString()) + int.Parse(ws.Cells["C" + fila7].Value.ToString()) + int.Parse(ws.Cells["C" + fila8].Value.ToString()) + int.Parse(ws.Cells["C" + fila9].Value.ToString()) + int.Parse(ws.Cells["C" + fila10].Value.ToString()) + int.Parse(ws.Cells["C" + fila11].Value.ToString()) + int.Parse(ws.Cells["C" + fila12].Value.ToString()) + int.Parse(ws.Cells["C" + fila13].Value.ToString());
                int ayudantesActivos = int.Parse(ws.Cells["D" + fila4].Value.ToString()) + int.Parse(ws.Cells["D" + fila5].Value.ToString()) + int.Parse(ws.Cells["D" + fila6].Value.ToString()) + int.Parse(ws.Cells["D" + fila7].Value.ToString()) + int.Parse(ws.Cells["D" + fila8].Value.ToString()) + int.Parse(ws.Cells["D" + fila9].Value.ToString()) + int.Parse(ws.Cells["D" + fila10].Value.ToString()) + int.Parse(ws.Cells["D" + fila11].Value.ToString()) + int.Parse(ws.Cells["D" + fila12].Value.ToString()) + int.Parse(ws.Cells["D" + fila13].Value.ToString());
                int ayudantesDeLicencia = int.Parse(ws.Cells["E" + fila4].Value.ToString()) + int.Parse(ws.Cells["E" + fila5].Value.ToString()) + int.Parse(ws.Cells["E" + fila6].Value.ToString()) + int.Parse(ws.Cells["E" + fila7].Value.ToString()) + int.Parse(ws.Cells["E" + fila8].Value.ToString()) + int.Parse(ws.Cells["E" + fila9].Value.ToString()) + int.Parse(ws.Cells["E" + fila10].Value.ToString()) + int.Parse(ws.Cells["E" + fila11].Value.ToString()) + int.Parse(ws.Cells["E" + fila12].Value.ToString()) + int.Parse(ws.Cells["E" + fila13].Value.ToString());
                int apoyosActivos = int.Parse(ws.Cells["F" + fila4].Value.ToString()) + int.Parse(ws.Cells["F" + fila5].Value.ToString()) + int.Parse(ws.Cells["F" + fila6].Value.ToString()) + int.Parse(ws.Cells["F" + fila7].Value.ToString()) + int.Parse(ws.Cells["F" + fila8].Value.ToString()) + int.Parse(ws.Cells["F" + fila9].Value.ToString()) + int.Parse(ws.Cells["F" + fila10].Value.ToString()) + int.Parse(ws.Cells["F" + fila11].Value.ToString()) + int.Parse(ws.Cells["F" + fila12].Value.ToString()) + int.Parse(ws.Cells["F" + fila13].Value.ToString());
                int apoyosDeLicencia = int.Parse(ws.Cells["G" + fila4].Value.ToString()) + int.Parse(ws.Cells["G" + fila5].Value.ToString()) + int.Parse(ws.Cells["G" + fila6].Value.ToString()) + int.Parse(ws.Cells["G" + fila7].Value.ToString()) + int.Parse(ws.Cells["G" + fila8].Value.ToString()) + int.Parse(ws.Cells["G" + fila9].Value.ToString()) + int.Parse(ws.Cells["G" + fila10].Value.ToString()) + int.Parse(ws.Cells["G" + fila11].Value.ToString()) + int.Parse(ws.Cells["G" + fila12].Value.ToString()) + int.Parse(ws.Cells["G" + fila13].Value.ToString());
                ws.Cells["R" + fila12].Value = conductoresActivos + conductoresDeLicencia + ayudantesActivos + ayudantesDeLicencia + apoyosActivos + apoyosDeLicencia;
                
                //formatear celdas de centros y de Totales en negrita
                ws.Cells["A"+fila4+":A"+fila14].Style.Font.Bold = true;

                //quitar bordes hacia la derecha de la segunda fila de cada tabla (aparentemente no funciona)
                ws.Cells["H"+fila3+":N"+fila3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                ws.Cells["H" + fila3 + ":O" + fila3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                ws.Cells["A"+fila2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A"+fila2].Style.Fill.BackgroundColor.SetColor(Color.Yellow);


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
            RegistroDeTotales registroTaller = new RegistroDeTotales("Taller", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);// taller serian todos los trabajadores que sean nocheros o mecanicos, independiente del centro 
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
                                    //el trabajador esta en santiago y es un movilizador = se asigna a movilizadores
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

                                    registroAdministracion.TotalRemuneracionesOtros = registroAdministracion.TotalRemuneracionesOtros + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                                    registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores = registroAdministracion.TotalRemuneracionesDeTodosLosTrabajadores + (int.Parse(item.Imponible_sin_tope) + int.Parse(item.Total_exento));
                            

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
            MessageBox.Show("para subir asistencias a rex: * recibir excel de Francisco * copiar los datos que vienen filtrados en el excel, a un excel nuevo que tenga la cabecera(ese excel se descarga de rex) * Guardar el nuevo excel con los registros copiados como formato CSV * Enviar a las de remuneraciones para que ellas hagan la carga.", "Sobre la subida a Rex");
            MessageBox.Show("Transformar registros a totales sigue la siguiente lógica: se toma el archivo excel de base, se filtra primero por mes y luego por Centro. Los montos y totales para cada centro se obtienen con esos 2 filtros, salvo 2 excpeciones. La primera es si un trabajador de SANTIAGO o SANTIAGO E2 es un movilizador, en cuyo caso se asigna al centro de movilizadores. La segunda es cuando el trabajador de central es un nochero, en cuyo caso se asigna a administración.", "Sobre el registro de totales");
            MessageBox.Show("Programa creado por Marcelo Andrés Aranda Tatto, bajo ordenes de Antonio Alonso.", "Sobre el programa");

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
