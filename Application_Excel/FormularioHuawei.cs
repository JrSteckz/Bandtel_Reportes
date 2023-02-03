using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Reportes
{
    public partial class FormularioHuawei : Form
    {
        public class ExcelRow
        {
            public string Id_Nodo { get; set; }
            public string P_N { get; set; }
            public string DESCRIPCION { get; set; }
            public string S_N { get; set; }
            public string CANT { get; set; }
            public string UND { get; set; }
            public string NODO { get; set; }
            public string CONFIGURACION { get; set; }
            public string ENLACE { get; set; }
            public string FREC { get; set; }
            public string TIPO { get; set; }

        }
        public FormularioHuawei()
        {
            InitializeComponent();
        }

        private void FormularioHuawei_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string Id_Nodo = txtIdNodo.Text;
            string filePath = "";

            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File";
            theDialog.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*|Excel (*.xlsm)|*.xlsm";
            DialogResult res = theDialog.ShowDialog();
            filePath = theDialog.FileName;
            //
             // Crear una conexión a la hoja de cálculo
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+filePath+ ";Extended Properties=\"\"Excel 8.0;HDR=YES;''";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();
            //
            List<ExcelRow> rows = new List<ExcelRow>();
            Application excelApp = new Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(@filePath);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets["ENVIOS"];
            int lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row - 3;

            // Crear un comando para leer un rango de celdas específico
            OleDbCommand command = new OleDbCommand("SELECT * FROM [ENVIOS$A1:C"+ lastRow+"]", connection);

            // Ejecutar el comando y crear un DataReader
            OleDbDataReader reader = command.ExecuteReader();

            // Recorrer las filas de los datos leídos
            while (reader.Read())
            {
                // Obtener los valores de cada columna
                string col1 = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                string col2 = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
                string col3 = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);

                // Procesar los valores de las columnas
                Console.WriteLine("Columna 1: " + col1 + " Columna 2: " + col2 + " Columna 3: " + col3);
            }

            // Cerrar la conexión y el DataReader
            reader.Close();
            connection.Close();
            //

            //List<ExcelRow> rows = new List<ExcelRow>();
            //Application excelApp = new Application();
            //Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(@filePath);
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets["ENVIOS"];
            //// Obtener la última fila utilizada en la hoja
            //int lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row - 3;

            //// Recorrer las filas en la hoja
            //for (int i = 4; i <= lastRow; i++)
            //{
            //    // Crear un objeto para representar la fila actual
            //    ExcelRow row = new ExcelRow();

            //    row.Id_Nodo = worksheet.Cells[i, "F"].Value != null ? worksheet.Cells[i, "F"].Value2.ToString() : "";
            //    row.P_N = worksheet.Cells[i, "A"].Value != null ? worksheet.Cells[i, "A"].Value2.ToString() : "";
            //    row.DESCRIPCION = worksheet.Cells[i, "B"].Value != null ? worksheet.Cells[i, "B"].Value2.ToString() : "";
            //    row.S_N = worksheet.Cells[i, "C"].Value != null ? worksheet.Cells[i, "C"].Value2.ToString() : "";
            //    row.CANT = worksheet.Cells[i, "D"].Value != null ? worksheet.Cells[i, "D"].Value2.ToString() : "";
            //    row.UND = worksheet.Cells[i, "E"].Value != null ? worksheet.Cells[i, "E"].Value2.ToString() : "";
            //    row.NODO = worksheet.Cells[i, "G"].Value != null ? worksheet.Cells[i, "G"].Value2.ToString() : "";
            //    row.CONFIGURACION = worksheet.Cells[i, "H"].Value != null ? worksheet.Cells[i, "H"].Value2.ToString() : "";
            //    row.ENLACE = worksheet.Cells[i, "I"].Value != null ? worksheet.Cells[i, "I"].Value2.ToString() : "";
            //    row.FREC = worksheet.Cells[i, "J"].Value != null ? worksheet.Cells[i, "J"].Value2.ToString() : "";
            //    row.TIPO = worksheet.Cells[i, "K"].Value != null ? worksheet.Cells[i, "K"].Value2.ToString() : "";

            //    // Agregar el objeto a la lista
            //    rows.Add(row);
            //}
            //workbook.Close();
            // excelApp.Quit();
            //// Agrupar los objetos basados en la columna "A" (el nombre de nodo)
            //var groupedRows = rows.Where(r => r.Id_Nodo == Id_Nodo)
            //                      .GroupBy(r => r.Id_Nodo)
            //                      .Select(g => new
            //                      {
            //                          P_N = g.Select(r => r.P_N),
            //                          DESCRIPCION = g.Select(r => r.DESCRIPCION),
            //                          S_N = g.Select(r => r.S_N),
            //                          CANT = g.Select(r => r.CANT),
            //                          UND = g.Select(r => r.UND),
            //                          NODO = g.Select(r => r.NODO),
            //                          CONFIGURACION = g.Select(r => r.CONFIGURACION),
            //                          ENLACE = g.Select(r => r.ENLACE),
            //                          FREC = g.Select(r => r.FREC),
            //                          TIPO = g.Select(r => r.TIPO)
            //                      });
            //Console.ReadKey();

        }
    }


}

