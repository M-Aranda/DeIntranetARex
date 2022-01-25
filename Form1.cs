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

using System.IO;
using System.Runtime.InteropServices;
using Windows.Storage;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;

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
            List<Asistencia> asistencias = new List<Asistencia>();

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

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@sFileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //for (int i = 1; i <= rowCount; i++)
            //{
            //    Asistencia a = new Asistencia();

            //    for (int j = 1; j <= colCount; j++)
            //    {
                   
            //        if (j == 1 && xlRange.Cells[i, j].Value2!=null )
            //        {
            //            a.Empleado = xlRange.Cells[i, j].Value2.ToString();
            //        }
            //        asistencias.Add(a);
            //    }

            //}





            

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
