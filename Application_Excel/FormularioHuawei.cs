using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application_Excel;
using ExcelLibrary.BinaryFileFormat;
using ExcelLibrary.SpreadSheet;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Reportes
{
    public partial class FormularioHuawei : Form
    {
        public string UbicacionGuardado;
        public string Id_Nodo = "";
        public string filePath = "";

        public FormularioHuawei()
        {
            InitializeComponent();
        }

        private void FormularioHuawei_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var form2 = new FormularioProgressBar();
            form2.Show();
            //
            Id_Nodo = txtIdNodo.Text;
            UbicacionGuardado = txtGuardado.Text;
            //
            FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            stream.Close();
            //
            ISheet sheet = workbook.GetSheet("ENVIOS");
            int cuenta = sheet.LastRowNum - 2;
            int contador = 0;
            Dictionary<string, int> ItemsAgrupados = new Dictionary<string, int>();
            form2.Instalacion(1);
            for (int i = 3; i <= cuenta; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    // Verificar si la primera celda contiene el nodo especificado
                    ICell cell = row.GetCell(5);
                    if (cell != null && cell.ToString() == Id_Nodo)
                    {
                        contador++;
                        string resultado = row.GetCell(8).ToString();
                        if (ItemsAgrupados.ContainsKey(resultado))
                        {
                            ItemsAgrupados[resultado]++;
                        }
                        else
                        {
                            ItemsAgrupados[resultado] = 1;
                        }
                    }
                }

            }
            form2.Instalacion(2);
            // Crear una nueva aplicación de Excel
            Application excel = new Application();
            Workbook workbook2 = excel.Workbooks.Add();
            //
            FileStream streamm = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook workbookk = new XSSFWorkbook(streamm);
            streamm.Close();
            ISheet sheett = workbookk.GetSheet("ENVIOS");
            //
            form2.Instalacion(3);
            foreach (KeyValuePair<string, int> item in ItemsAgrupados)
            {
                int contador1 = 3;
                //
                Worksheet worksheet2 = workbook2.Worksheets.Add();
                worksheet2.Name = item.Key;
                //
                worksheet2.PageSetup.FitToPagesWide = 1;
                worksheet2.PageSetup.FitToPagesTall = 1;
                worksheet2.PageSetup.Zoom = false;
                //
                worksheet2.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                worksheet2.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                worksheet2.PageSetup.Zoom = 60;
                //
                Excel.Range range = worksheet2.Range["A1:K2"];
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                Excel.Range range2 = worksheet2.Range["A2:K2"];
                range2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                //
                worksheet2.Columns[1].ColumnWidth = 11.2;
                worksheet2.Columns[2].ColumnWidth = 77;
                worksheet2.Columns[3].ColumnWidth = 21.5;
                worksheet2.Columns[4].ColumnWidth = 6.6;
                worksheet2.Columns[5].ColumnWidth = 6.0;
                worksheet2.Columns[6].ColumnWidth = 13.6;
                worksheet2.Columns[7].ColumnWidth = 16.6;
                worksheet2.Columns[8].ColumnWidth = 18.3;
                worksheet2.Columns[9].ColumnWidth = 21.6;
                worksheet2.Columns[10].ColumnWidth = 6.0;
                worksheet2.Columns[11].ColumnWidth = 10.8;
                //
                worksheet2.Cells[1, 2].Value = Id_Nodo;
                //
                worksheet2.Cells[2, 1].Value = "P/N";
                worksheet2.Cells[2, 2].Value = "DESCRIPCION";
                worksheet2.Cells[2, 3].Value = "S/N";
                worksheet2.Cells[2, 4].Value = "CANT";
                worksheet2.Cells[2, 5].Value = "UND";
                worksheet2.Cells[2, 6].Value = "COD NODO";
                worksheet2.Cells[2, 7].Value = "NODO";
                worksheet2.Cells[2, 8].Value = "CONFIGURACION";
                worksheet2.Cells[2, 9].Value = "ENLACE";
                worksheet2.Cells[2, 10].Value = "FREC";
                worksheet2.Cells[2, 11].Value = "TIPO";
                //
                form2.Instalacion(5);
                for (int i = 3; i <= cuenta; i++)
                {
                    IRow roww = sheett.GetRow(i);
                    if (roww != null)
                    {
                        // Verificar si la primera celda contiene el nodo especificado
                        ICell celll = roww.GetCell(8);
                        ICell cellll = roww.GetCell(5);
                        if (celll != null && celll.ToString() == item.Key.ToString() && cellll.ToString() == Id_Nodo)
                        {
                            worksheet2.Cells[contador1, 1].Value = roww.GetCell(0).ToString();
                            worksheet2.Cells[contador1, 2].Value = roww.GetCell(1).ToString();
                            worksheet2.Cells[contador1, 3].Value = roww.GetCell(2).ToString();
                            worksheet2.Cells[contador1, 4].Value = roww.GetCell(3).ToString();
                            worksheet2.Cells[contador1, 5].Value = roww.GetCell(4).ToString();
                            worksheet2.Cells[contador1, 6].Value = roww.GetCell(5).ToString();
                            worksheet2.Cells[contador1, 7].Value = roww.GetCell(6).ToString();
                            worksheet2.Cells[contador1, 8].Value = roww.GetCell(7).ToString();
                            worksheet2.Cells[contador1, 9].Value = roww.GetCell(8).ToString();
                            worksheet2.Cells[contador1, 10].Value = roww.GetCell(9).ToString();
                            worksheet2.Cells[contador1, 11].Value = roww.GetCell(10).ToString();
                            //
                            contador1++;
                        }
                    }
                }
                Excel.Range range3 = worksheet2.Range["A3:K" + contador1];
                range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                //                
            }
            form2.Instalacion(10);
            string nombreexcel = txtNombre.Text;
            workbook2.SaveAs(UbicacionGuardado + @"\" + nombreexcel + ".xlsx");
            workbook2.Close();
            excel.Quit();
            form2.Close();
            MessageBox.Show("Listo");
            Process.Start(UbicacionGuardado + @"\"+ nombreexcel + @".xlsx");
        }
        private void btnGuardado_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                UbicacionGuardado = dialog.FileName;
                txtGuardado.Text = UbicacionGuardado;
            }
            if (txtGuardado.Text != "" && txtIdNodo.Text != "")
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }
        private void btnatos_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File";
            theDialog.Filter = "Excel (*.xlsm)|*.xlsm|Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            DialogResult res = theDialog.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                filePath = theDialog.FileName;
                txtDatos.Text = filePath;
            }
        }
    }
}

