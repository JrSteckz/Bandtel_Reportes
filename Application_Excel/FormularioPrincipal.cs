﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using Microsoft.WindowsAPICodePack.Dialogs;
using ExcelLibrary.BinaryFileFormat;
using MS.WindowsAPICodePack.Internal;
using static Application_Excel.FormularioPrincipal;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;
using System.Collections;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Drawing.Imaging;
using ExifLib;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.Excel.Shape;

namespace Application_Excel
{

    public partial class FormularioPrincipal : Form
    {

        public string URL_Guardado = "";
        public string URL_Plantilla = "";
        public string URL_Imagenes = "";
        //
        public int WorkSheet5;
        //
        public Excel.Application oXL;
        public Excel.Workbook xlWBook;
        public Excel.Worksheet xlWSheet;
        public Excel.Range xlRange;
        //
        public string CodigoDigi = "";
        public string CodigoIntermedio = "";
        public string Formato = ".png";
        //
        public string Columna_General = "";
        public string Fila_General = "";
        //
        public string Formato_2 = "";
        public string Formato_5 = "";
        public string Formato_6A = "";
        public string Formato_6B = "";
        public string Formato_8 = "";
        public string Formato_9 = "";
        //
        public string Codigo_2 = "";
        public string Codigo_5 = "";
        public string Codigo_6A = "";
        public string Codigo_6B = "";
        public string Codigo_8 = "";
        public string Codigo_9 = "";
        //
        public string Formato_Default_2 = ".jpeg";
        public string Formato_Default_5 = ".png";
        public string Formato_Default_6A = ".png";
        public string Formato_Default_6B = ".png";
        public string Formato_Default_8 = ".jpeg";
        public string Formato_Default_9 = ".jpeg";
        //
        public string Codigo_Default_2 = "NODO_";
        public string Codigo_Default_5 = "NODO_";
        public string Codigo_Default_6A = "6_A_";
        public string Codigo_Default_6B = "6_B_";
        public string Codigo_Default_8 = "A8_";
        public string Codigo_Default_9 = "B9_";
        //
        public int IndicadordeTamaño = 0;
        //
        Excel.Range RangoWidth;

        public class CodigoNumeracion
        {
            public int grupo { get; set; }
            public int numeracion { get; set; }
            public string dimension { get; set; }
        }
        public class DetalleImagen
        {
            public string imagen { get; set; }
            public int dimension { get; set; }
        }
        public FormularioPrincipal()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            txtConfiFormato2.Text = ".jpeg";
            txtConfiFormato5.Text = ".png";
            txtConfiFormato6A.Text = ".png";
            txtConfiFormato6B.Text = ".png";
            txtConfiFormato8.Text = ".jpeg";
            txtConfiFormato9.Text = ".jpeg";

            //
            txtConfiCodigo2.Text = "NODO_";
            txtConfiCodigo5.Text = "NODO_";
            txtConfiCodigo6A.Text = "6_A_";
            txtConfiCodigo6B.Text = "6_B_";
            txtConfiCodigo8.Text = "A8_";
            txtConfiCodigo9.Text = "B9_";

            //
            txtURL.Enabled = false;
            txtUbicacionPlantilla.Enabled = false;
            txtImagenes.Enabled = false;
            btnGenerar.Enabled = false;

        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            try
            {
                Formato_2 = txtConfiFormato2.Text;
                Formato_5 = txtConfiFormato5.Text;
                Formato_6A = txtConfiFormato6A.Text;
                Formato_6B = txtConfiFormato6B.Text;
                Formato_8 = txtConfiFormato8.Text;
                Formato_9 = txtConfiFormato9.Text;
                //
                Codigo_2 = txtConfiCodigo2.Text;
                Codigo_5 = txtConfiCodigo5.Text;
                Codigo_6A = txtConfiCodigo6A.Text;
                Codigo_6B = txtConfiCodigo6B.Text;
                Codigo_8 = txtConfiCodigo8.Text;
                Codigo_9 = txtConfiCodigo9.Text;
                //
                if (txtNombreExcel.Text.Trim().Length <= 1)
                {
                    MessageBox.Show("Faltan llenar Datos");
                }
                else
                {
                    var form2 = new FormularioProgressBar();
                    form2.Show();
                    //Designar Excel Plantilla
                    oXL = new Excel.Application();
                    var Link = URL_Plantilla;
                    xlWBook = oXL.Workbooks.Open(@Link);
                    xlWSheet = oXL.ActiveSheet as Excel.Worksheet;
                    xlRange = xlWSheet.UsedRange;

                    //Hoja de Trabajo 2
                    xlWSheet = (Excel.Worksheet)xlWBook.Sheets["2.INFORMACION GENERAL"];
                    xlWSheet.Select(Type.Missing);
                    form2.Instalacion(1);
                    Insertarprimera();
                    form2.Instalacion(2);

                    //Hoja de Trabajo 5
                    xlWSheet = (Excel.Worksheet)xlWBook.Sheets["5.Pruebas de Interferencia"];
                    xlWSheet.Select(Type.Missing);
                    form2.Instalacion(3);
                    Insertarsegunda();
                    form2.Instalacion(4);

                    //Detectar Hoja de Trabajo 6
                    xlWSheet = (Excel.Worksheet)oXL.Worksheets["6.Configuración y Mediciones "];
                    xlWSheet.Select(Type.Missing);
                    InsertarFila(1);
                    form2.Instalacion(5);
                    InsertarFila(2);
                    form2.Instalacion(6);

                    //Detectar Hoja de Trabajo 8_A
                    xlWSheet = (Excel.Worksheet)oXL.Worksheets["8.Rep Fot_NODO 1"];
                    xlWSheet.Select(Type.Missing);
                    InsertarFila2(1);
                    form2.Instalacion(8);

                    ////Detectar Hoja de Trabajo 9_B
                    xlWSheet = (Excel.Worksheet)oXL.Worksheets["9.Rep Fot_NODO 2"];
                    xlWSheet.Select(Type.Missing);
                    InsertarFila2(2);

                    //Guardar Excel
                    string Lugar_Guardado = txtURL.Text + @"\";
                    string NombreExcel = txtNombreExcel.Text + @".xlsx";
                    xlWBook.SaveAs(Lugar_Guardado + NombreExcel);
                    xlWBook.Close(true, Type.Missing, Type.Missing);
                    oXL.Quit();
                    form2.Instalacion(10);
                    MessageBox.Show("Archivo Guardado");
                    Process.Start(Lugar_Guardado + NombreExcel);
                    form2.Close();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void BtnBuscadorGuardado_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                URL_Guardado = dialog.FileName;
                txtURL.Text = URL_Guardado;
            }
            if (txtURL.Text != "" && txtUbicacionPlantilla.Text != "" && txtUbicacionPlantilla.Text != "")
            {
                btnGenerar.Enabled = true;
            }
            else
            {
                btnGenerar.Enabled = false;
            }
        }
        private void BtnBuscadorPlantilla_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File";
            theDialog.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            DialogResult res = theDialog.ShowDialog();
            //
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                URL_Plantilla = theDialog.FileName;
                txtUbicacionPlantilla.Text = URL_Plantilla;

            }
            if (txtURL.Text != "" && txtUbicacionPlantilla.Text != "" && txtUbicacionPlantilla.Text != "")
            {
                btnGenerar.Enabled = true;
            }
            else
            {
                btnGenerar.Enabled = false;
            }
        }
        private void BtnBuscadorImagenes_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                URL_Imagenes = dialog.FileName;
                txtImagenes.Text = URL_Imagenes;
            }
            if (txtURL.Text != "" && txtUbicacionPlantilla.Text != "" && txtUbicacionPlantilla.Text != "")
            {
                btnGenerar.Enabled = true;
            }
            else
            {
                btnGenerar.Enabled = false;
            }
        }
        void Insertarprimera()
        {
            Columna_General = "";
            Fila_General = "";
            Formato = Formato_2;
            CodigoDigi = Codigo_2;
            //
            //MessageBox.Show(CodigoDigi + "-" + Codigo_2);
            //
            string Direccion_Informacion_Gemeral = @URL_Imagenes + @"\2.Informacion_General\";
            //
            if (Directory.Exists(Direccion_Informacion_Gemeral))
            {
                string[] Informacion_General = Directory.GetFiles(Direccion_Informacion_Gemeral, "*" + Formato);
                int cantidad_Informacion_General = Informacion_General.Length;
                string NombreImg2 = null;
                //

                String[] Codigo = new String[100];
                String[] Numeracion = new String[100];
                String[] strlist = new String[100];
                String[] separador = { Direccion_Informacion_Gemeral, CodigoDigi, Formato };
                //

                if (cantidad_Informacion_General == 0 || cantidad_Informacion_General == null)
                {
                    MessageBox.Show("No hay contenido en la Carpeta 2.Informacion_General");
                }
                else
                {
                    int Rang_colum = 27;
                    int Rang_row = 47;
                    //
                    for (int i = 0; i <= cantidad_Informacion_General - 1; i++)
                    {
                        string dir2 = Informacion_General[i];
                        NombreImg2 = Path.GetFileName(dir2);
                        //
                        strlist = NombreImg2.Split(separador, separador.Length, StringSplitOptions.RemoveEmptyEntries);
                        Codigo[i] = strlist[0];
                    }
                    int contador = Int32.Parse(Codigo.Max());
                    //
                    for (int cant_var = 1; cant_var <= contador; cant_var++)
                    {
                        string curFile = Direccion_Informacion_Gemeral + CodigoDigi + cant_var + Formato;

                        if ((cant_var % 2) == 0)
                        {
                            //Asignar Rango
                            CalcularTamanoImagen(curFile);
                            if (IndicadordeTamaño == 1)
                            {
                                Rang_colum = 43;
                                Rang_row = 31;
                                Columna_General = "J";
                                Fila_General = "P";
                            }
                            else if (IndicadordeTamaño == 2)
                            {
                                Rang_colum = 27;
                                Rang_row = 47;
                                Columna_General = "K";
                                Fila_General = "O";
                            }

                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                            //Insertar imagenes
                            if (File.Exists(curFile))
                            {

                                Shape shape = xlWSheet.Shapes.AddPicture(curFile,
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                            }
                        }
                        else
                        {
                            CalcularTamanoImagen(curFile);
                            if (IndicadordeTamaño == 1)
                            {
                                Rang_colum = 43;
                                Rang_row = 31;
                                Columna_General = "B";
                                Fila_General = "H";
                            }
                            else if (IndicadordeTamaño == 2)
                            {
                                Rang_colum = 27;
                                Rang_row = 47;
                                Columna_General = "C";
                                Fila_General = "G";
                            }

                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                            //Insertar imagenes
                            if (File.Exists(curFile))
                            {
                                xlWSheet.Shapes.AddPicture(curFile,
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                            }
                        }
                    }
                }
            }
            else if (!Directory.Exists(Direccion_Informacion_Gemeral))
            {
                MessageBox.Show("La Carpeta " + Direccion_Informacion_Gemeral + " no existe");
            }



        }
        void Insertarsegunda()
        {
            Columna_General = "D";
            Fila_General = "Q";
            Formato = Formato_5;
            CodigoDigi = Codigo_5;
            //
            //MessageBox.Show(CodigoDigi + "-" + Formato);
            //
            string Direccion_Informacion_Gemeral = @URL_Imagenes + @"\5.Pruebas de Interferencia\";
            //
            if (Directory.Exists(Direccion_Informacion_Gemeral))
            {
                string[] Informacion_General = Directory.GetFiles(Direccion_Informacion_Gemeral, "*" + Formato);
                int cantidad_Informacion_General = Informacion_General.Length;
                string NombreImg2 = null;
                //

                String[] Codigo = new String[100];
                String[] Numeracion = new String[100];
                String[] strlist = new String[100];
                String[] separador = { Direccion_Informacion_Gemeral, CodigoDigi, Formato };
                //

                if (cantidad_Informacion_General == 0 || cantidad_Informacion_General == null)
                {
                    MessageBox.Show("No hay contenido en la Carpeta 5.Pruebas de Interferencia");
                }
                else
                {
                    int Rang_colum = 15;
                    int Rang_row = 34;
                    //
                    for (int i = 0; i <= cantidad_Informacion_General - 1; i++)
                    {
                        string dir2 = Informacion_General[i];
                        NombreImg2 = Path.GetFileName(dir2);
                        //
                        strlist = NombreImg2.Split(separador, separador.Length, StringSplitOptions.RemoveEmptyEntries);
                        Codigo[i] = strlist[0];
                    }
                    int contador = Int32.Parse(Codigo.Max());
                    //
                    for (int cant_var = 1; cant_var <= contador; cant_var++)
                    {
                        string curFile = Direccion_Informacion_Gemeral + CodigoDigi + cant_var + Formato;

                        CalcularTamanoImagen(curFile);
                        if (IndicadordeTamaño == 1)
                        {

                            Columna_General = "D";
                            Fila_General = "Q";
                        }
                        else if (IndicadordeTamaño == 2)
                        {

                            Columna_General = "N";
                            Fila_General = "G";
                        }
                        //Asignar Rango
                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                        //Insertar imagenes
                        if (File.Exists(curFile))
                        {
                            xlWSheet.Shapes.AddPicture(curFile,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                            float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));



                            Rang_colum += 23;
                            Rang_row += 23;
                        }
                    }
                }
            }
            else if (!Directory.Exists(Direccion_Informacion_Gemeral))
            {
                MessageBox.Show("La Carpeta " + Direccion_Informacion_Gemeral + " no existe");
            }
        }
        void InsertarFila(int RangoFila1)
        {
            Columna_General = "";
            Fila_General = "";
            CodigoIntermedio = "_";
            int numerador = 0;
            if (RangoFila1 == 1)
            {
                Formato = Formato_6A;
                numerador = 1;
                CodigoDigi = Codigo_6A;

            }
            else if (RangoFila1 == 2)
            {
                Formato = Formato_6B;
                numerador = 2;
                CodigoDigi = Codigo_6B;
            }
            //MessageBox.Show(CodigoDigi + "-" + Formato);
            string Direccion_Configuracion_Mediciones_A = URL_Imagenes + @"\6.Configuracion_Mediciones\NODO" + numerador + @"\";
            //
            if (Directory.Exists(Direccion_Configuracion_Mediciones_A))
            {
                string[] Configuracion_Mediciones_A = Directory.GetFiles(Direccion_Configuracion_Mediciones_A, "*" + Formato);
                int cantidad_Configuracion_Mediciones_A = Configuracion_Mediciones_A.Length;
                string NombreImgConfiguracion_A = null;
                //
                String[] CodigoConfiguracion_A = new String[100];
                String[] NumeracionConfiguracion_A = new String[100];
                String[] strlistConfiguracion_A = new String[100];
                String[] separadorConfiguracion_A = { Direccion_Configuracion_Mediciones_A, CodigoDigi, CodigoIntermedio, Formato };
                //
                List<CodigoNumeracion> codigoNumeracion = new List<CodigoNumeracion>();
                //

                if (cantidad_Configuracion_Mediciones_A == 0 || cantidad_Configuracion_Mediciones_A == null)
                {
                    MessageBox.Show("No hay contenido en la Carpeta 6.Configuracion_Mediciones");
                }
                else
                {
                    for (int i = 0; i <= cantidad_Configuracion_Mediciones_A - 1; i++)
                    {
                        string dir2 = Configuracion_Mediciones_A[i];
                        NombreImgConfiguracion_A = Path.GetFileName(dir2);
                        //
                        strlistConfiguracion_A = NombreImgConfiguracion_A.Split(separadorConfiguracion_A, separadorConfiguracion_A.Length, StringSplitOptions.RemoveEmptyEntries);
                        CodigoNumeracion tes = new CodigoNumeracion();
                        tes.grupo = Int32.Parse(strlistConfiguracion_A[0]);
                        tes.numeracion = Int32.Parse(strlistConfiguracion_A[1]);
                        codigoNumeracion.Add(tes);
                        //
                        CodigoConfiguracion_A[i] = strlistConfiguracion_A[0];
                        NumeracionConfiguracion_A[i] = strlistConfiguracion_A[1];
                    }
                    var cantidadmaxima = codigoNumeracion.Max(x => x.grupo);
                    int Rang_colum = 50;
                    int Rang_row = 63;
                    int aumento = 16;
                    //Bucle de insertado de imagenes
                    for (int cant_var = 1; cant_var <= cantidadmaxima; cant_var++)
                    {
                        //Asignar Rango
                        codigoNumeracion.OrderBy(x => x.grupo).ThenBy(y => y.grupo);
                        var ordenado = codigoNumeracion.Where(x => x.grupo == cant_var);
                        var cantidadcodigo = 0;
                        foreach (var value in ordenado)
                        {
                            cantidadcodigo = value.numeracion;
                        }
                        switch (cantidadcodigo)
                        {
                            case 0:
                                Rang_colum += aumento;
                                Rang_row += aumento;
                                break;
                            case 1:
                                for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                {
                                    string curFile = Direccion_Configuracion_Mediciones_A + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                    if (File.Exists(curFile))
                                    {
                                        if (RangoFila1 == 1)
                                        {
                                            CalcularTamanoImagen(curFile);
                                            if (IndicadordeTamaño == 1)
                                            {

                                                Columna_General = "B";
                                                Fila_General = "J";
                                            }
                                            else if (IndicadordeTamaño == 2)
                                            {

                                                Columna_General = "D";
                                                Fila_General = "H";
                                            }
                                        }
                                        else if (RangoFila1 == 2)
                                        {
                                            CalcularTamanoImagen(curFile);
                                            if (IndicadordeTamaño == 1)
                                            {
                                                Columna_General = "L";
                                                Fila_General = "T";
                                            }
                                            else if (IndicadordeTamaño == 2)
                                            {

                                                Columna_General = "N";
                                                Fila_General = "R";
                                            }
                                        }
                                        //
                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        //
                                        xlWSheet.Shapes.AddPicture(curFile,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                    }
                                }
                                Rang_colum += aumento;
                                Rang_row += aumento;
                                break;
                            case 2:
                                for (int numeracionciclo = 1; numeracionciclo <= 2; numeracionciclo++)
                                {
                                    //Insertar imagenes
                                    string curFile = Direccion_Configuracion_Mediciones_A + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                    //
                                    if (File.Exists(curFile))
                                    {
                                        if (RangoFila1 == 1)
                                        {
                                            CalcularTamanoImagen(curFile);
                                            if (IndicadordeTamaño == 1)
                                            {
                                                Columna_General = "B";
                                                Fila_General = "J";
                                            }
                                            else if (IndicadordeTamaño == 2)
                                            {
                                                Columna_General = "D";
                                                Fila_General = "H";
                                            }
                                        }
                                        else if (RangoFila1 == 2)
                                        {
                                            CalcularTamanoImagen(curFile);
                                            if (IndicadordeTamaño == 1)
                                            {
                                                Columna_General = "L";
                                                Fila_General = "T";
                                            }
                                            else if (IndicadordeTamaño == 2)
                                            {
                                                Columna_General = "N";
                                                Fila_General = "R";
                                            }
                                        }
                                        //
                                        if (numeracionciclo == 2)
                                        {
                                            Rang_colum += 14;
                                            Rang_row += 14;
                                        }
                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        //
                                        xlWSheet.Shapes.AddPicture(curFile,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                    }
                                }
                                Rang_colum += aumento;
                                Rang_row += aumento;
                                break;
                        }
                    }
                }
            }
            else if (!Directory.Exists(Direccion_Configuracion_Mediciones_A))
            {
                MessageBox.Show("La Carpeta " + Direccion_Configuracion_Mediciones_A + " no existe");
            }
        }
        void InsertarFila2(int RangoFila1)
        {

            CodigoIntermedio = "_";
            int numerador = 0;
            string asignador = "";
            if (RangoFila1 == 1)
            {
                Formato = Formato_8;
                numerador = 8;
                asignador = "A";
                CodigoDigi = Codigo_8;
            }
            else if (RangoFila1 == 2)
            {
                Formato = Formato_9;
                numerador = 9;
                asignador = "B";
                CodigoDigi = Codigo_9;
            }

            //MessageBox.Show(CodigoDigi + "-" + Formato);
            //Funcion Reporte Fotografico
            string Direccion_Configuracion_Mediciones_A = URL_Imagenes + @"\" + numerador + ".Reporte_Fotografico_" + asignador + @"\";
            //
            if (Directory.Exists(Direccion_Configuracion_Mediciones_A))
            {
                string[] Configuracion_Mediciones_A = Directory.GetFiles(Direccion_Configuracion_Mediciones_A, "*" + Formato);
                int cantidad_Configuracion_Mediciones_A = Configuracion_Mediciones_A.Length;
                string NombreImgConfiguracion_A = null;
                //

                String[] CodigoConfiguracion_A = new String[100];
                String[] NumeracionConfiguracion_A = new String[100];
                String[] strlistConfiguracion_A = new String[100];
                String[] separadorConfiguracion_A = { Direccion_Configuracion_Mediciones_A, CodigoDigi, CodigoIntermedio, Formato };
                //Funcion Serie de Equipos
                //string Direccion_Configuracion_Mediciones_B = URL_Imagenes + @"\" + numerador + @".Reporte_Fotografico_" + asignador + @"\2.Serie_Equipos";
                //string[] Configuracion_Mediciones_B = Directory.GetFiles(Direccion_Configuracion_Mediciones_B, "*" + Formato);
                //int cantidad_Configuracion_Mediciones_B = Configuracion_Mediciones_B.Length;
                //string NombreImgConfiguracion_B = null;

                //String[] CodigoConfiguracion_B = new String[100];
                //String[] NumeracionConfiguracion_B = new String[100];
                //String[] strlistConfiguracion_B = new String[100];
                //String[] separadorConfiguracion_B = { @URL_Imagenes + @"\" + numerador + @".Reporte_Fotografico_" + asignador + @"\2.Serie_Equipos\", "C_", "_", Formato };
                //Listas
                List<CodigoNumeracion> codigoNumeracion = new List<CodigoNumeracion>();
                List<CodigoNumeracion> codigoejemplo = new List<CodigoNumeracion>();
                //

                if (cantidad_Configuracion_Mediciones_A == 0 || cantidad_Configuracion_Mediciones_A == null)
                {
                    MessageBox.Show("No hay contenido en la Carpeta " + numerador + ".Reporte_Fotografico_" + asignador);
                }
                else
                {
                    for (int i = 0; i <= cantidad_Configuracion_Mediciones_A - 1; i++)
                    {
                        string dir2 = Configuracion_Mediciones_A[i];
                        NombreImgConfiguracion_A = Path.GetFileName(dir2);
                        //
                        strlistConfiguracion_A = NombreImgConfiguracion_A.Split(separadorConfiguracion_A, separadorConfiguracion_A.Length, StringSplitOptions.RemoveEmptyEntries);
                        CodigoNumeracion tes = new CodigoNumeracion();
                        tes.grupo = Int32.Parse(strlistConfiguracion_A[0]);
                        tes.numeracion = Int32.Parse(strlistConfiguracion_A[1]);
                        codigoNumeracion.Add(tes);
                        //
                        CodigoConfiguracion_A[i] = strlistConfiguracion_A[0];
                        NumeracionConfiguracion_A[i] = strlistConfiguracion_A[1];
                    }
                    var cantidadmaxima = codigoNumeracion.Max(x => x.grupo);
                    int Rang_colum = 11;
                    int Rang_row = 22;
                    int aumento = 16;
                    int distribucion = 0;
                    //Bucle de insertado de imagenes
                    for (int cant_var = 1; cant_var <= cantidadmaxima; cant_var++)
                    {
                        //Asignar Rango
                        codigoNumeracion.OrderBy(x => x.grupo).ThenBy(y => y.grupo);
                        var ordenado = codigoNumeracion.Where(x => x.grupo == cant_var);
                        var cantidadcodigo = 0;
                        foreach (var value in ordenado)
                        {
                            cantidadcodigo = value.numeracion;
                        }
                        //
                        if ((cant_var % 2) == 0)
                        {
                            distribucion = 2;
                        }
                        else
                        {
                            distribucion = 1;
                        }
                        //
                        if (cant_var == 36)
                        {
                            aumento = 18;
                        }
                        if (cant_var >= 37)
                        {
                            aumento = 16;
                        }
                        //
                        switch (cantidadcodigo)
                        {
                            //
                            case 0:
                                if (distribucion == 2)
                                {
                                    Rang_colum += aumento;
                                    Rang_row += aumento;
                                }
                                break;
                            //
                            case 1:
                                for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                {
                                    string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                    if (File.Exists(curFile))
                                    {
                                        if (distribucion == 1)
                                        {
                                            CalcularTamanoImagen(curFile);
                                            if (IndicadordeTamaño == 1)
                                            {
                                                Columna_General = "C";
                                                Fila_General = "I";
                                            }
                                            else if (IndicadordeTamaño == 2)
                                            {
                                                Columna_General = "D";
                                                Fila_General = "H";
                                            }
                                        }
                                        else if (distribucion == 2)
                                        {
                                            CalcularTamanoImagen(curFile);

                                            if (IndicadordeTamaño == 1)
                                            {
                                                Columna_General = "N";
                                                Fila_General = "T";
                                            }
                                            else if (IndicadordeTamaño == 2)
                                            {
                                                Columna_General = "O";
                                                Fila_General = "S";
                                            }
                                        }
                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        xlWSheet.Shapes.AddPicture(curFile,
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));


                                    }
                                }
                                if (distribucion == 2)
                                {
                                    Rang_colum += aumento;
                                    Rang_row += aumento;
                                }
                                break;
                            //
                            case 2:
                                int contadorcondicional = 0;
                                List<DetalleImagen> detalleImagensalto = new List<DetalleImagen>();
                                List<DetalleImagen> detalleImagensancho = new List<DetalleImagen>();
                                for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                {
                                    DetalleImagen detalleImagenalto = new DetalleImagen();
                                    DetalleImagen detalleImagenancho = new DetalleImagen();
                                    string curFile1 = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                    if (File.Exists(curFile1))
                                    {
                                        CalcularTamanoImagen(curFile1);
                                        if (IndicadordeTamaño == 1)
                                        {
                                            detalleImagenancho.imagen = curFile1;
                                            detalleImagenancho.dimension = IndicadordeTamaño;
                                            detalleImagensancho.Add(detalleImagenancho);
                                        }
                                        else if (IndicadordeTamaño == 2)
                                        {
                                            detalleImagenalto.imagen = curFile1;
                                            detalleImagenalto.dimension = IndicadordeTamaño;
                                            detalleImagensalto.Add(detalleImagenalto);
                                        }
                                    }
                                }
                                if (detalleImagensalto.Count == 2 && detalleImagensancho.Count == 0)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                        if (File.Exists(curFile))
                                        {
                                            //
                                            contadorcondicional++;
                                            switch (contadorcondicional)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "C";
                                                        Fila_General = "E";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "N";
                                                        Fila_General = "P";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "G";
                                                        Fila_General = "I";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "R";
                                                        Fila_General = "T";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }
                                else if (detalleImagensalto.Count == 0 && detalleImagensancho.Count == 2)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                        if (File.Exists(curFile))
                                        {
                                            //
                                            contadorcondicional++;
                                            switch (contadorcondicional)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "F";
                                                        Rang_row = Rang_row - 3;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_row = Rang_row + 3;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "Q";
                                                        Rang_row = Rang_row - 3;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_row = Rang_row + 3;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "F";
                                                        Fila_General = "J";
                                                        Rang_colum = Rang_colum + 3;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 3;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "Q";
                                                        Fila_General = "U";
                                                        Rang_colum = Rang_colum + 3;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 3;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }
                                else if (detalleImagensalto.Count == 1 && detalleImagensancho.Count == 1)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;

                                        if (File.Exists(curFile))
                                        {
                                            //
                                            contadorcondicional++;
                                            switch (contadorcondicional)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "F";
                                                        Rang_colum = Rang_colum + 2;
                                                        Rang_row = Rang_row - 2;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 2;
                                                        Rang_row = Rang_row + 2;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "Q";
                                                        Rang_colum = Rang_colum + 2;
                                                        Rang_row = Rang_row - 2;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 2;
                                                        Rang_row = Rang_row + 2;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "G";
                                                        Fila_General = "I";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "R";
                                                        Fila_General = "T";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }

                                if (distribucion == 2)
                                {
                                    Rang_colum += aumento;
                                    Rang_row += aumento;
                                }
                                break;
                            //
                            case 3:
                                int contadorcondicional2 = 0;
                                List<DetalleImagen> detalleImagensalto1 = new List<DetalleImagen>();
                                List<DetalleImagen> detalleImagensancho1 = new List<DetalleImagen>();
                                String[] detalleImagensalto4 = new String[100];
                                String[] detalleImagensancho4 = new String[100];
                                int conteo1 = 0;
                                int conteo2 = 0;
                                for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                {
                                    DetalleImagen detalleImagenalto = new DetalleImagen();
                                    DetalleImagen detalleImagenancho = new DetalleImagen();
                                    string curFile1 = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                    if (File.Exists(curFile1))
                                    {
                                        CalcularTamanoImagen(curFile1);
                                        if (IndicadordeTamaño == 1)
                                        {
                                            detalleImagenancho.imagen = curFile1;
                                            detalleImagenancho.dimension = IndicadordeTamaño;
                                            detalleImagensancho1.Add(detalleImagenancho);
                                            detalleImagensancho4[conteo1] = curFile1;
                                            conteo1++;
                                        }
                                        else if (IndicadordeTamaño == 2)
                                        {
                                            detalleImagenalto.imagen = curFile1;
                                            detalleImagenalto.dimension = IndicadordeTamaño;
                                            detalleImagensalto1.Add(detalleImagenalto);
                                            detalleImagensalto4[conteo2] = curFile1;
                                            conteo2++;
                                        }
                                    }
                                }
                                //
                                if (detalleImagensalto1.Count == 3 && detalleImagensancho1.Count == 0)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;

                                        if (File.Exists(curFile))
                                        {
                                            contadorcondicional2++;
                                            switch (contadorcondicional2)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "D";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "O";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "E";
                                                        Fila_General = "G";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "P";
                                                        Fila_General = "R";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "H";
                                                        Fila_General = "J";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "S";
                                                        Fila_General = "U";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                default:
                                                    break;
                                                    detalleImagensancho4 = null;
                                                    detalleImagensalto4 = null;
                                            }
                                        }
                                    }
                                }
                                else if (detalleImagensalto1.Count == 0 && detalleImagensancho1.Count == 3)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                        if (File.Exists(curFile))
                                        {
                                            contadorcondicional2++;
                                            switch (contadorcondicional2)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "E";
                                                        Rang_row = Rang_row - 6;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_row = Rang_row + 6;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "P";
                                                        Rang_row = Rang_row - 6;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_row = Rang_row + 6;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "E";
                                                        Rang_colum = Rang_colum + 6;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 6;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "P";
                                                        Rang_colum = Rang_colum + 6;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 6;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "F";
                                                        Fila_General = "J";
                                                        Rang_colum = Rang_colum + 2;
                                                        Rang_row = Rang_row - 2;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 2;
                                                        Rang_row = Rang_row + 2;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "Q";
                                                        Fila_General = "U";
                                                        Rang_colum = Rang_colum + 2;
                                                        Rang_row = Rang_row - 2;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 2;
                                                        Rang_row = Rang_row + 2;
                                                        xlWSheet.Shapes.AddPicture(curFile,
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }
                                else if (detalleImagensalto1.Count == 2 && detalleImagensancho1.Count == 1)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                        if (File.Exists(curFile))
                                        {
                                            contadorcondicional2++;
                                            switch (contadorcondicional2)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "D";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensalto4[0].ToString(),
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "O";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensalto4[0].ToString(),
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "E";
                                                        Fila_General = "G";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensalto4[1],
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "P";
                                                        Fila_General = "R";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensalto4[1],
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "H";
                                                        Fila_General = "J";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensancho4[0],
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "S";
                                                        Fila_General = "U";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensancho4[0],
                                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }
                                else if (detalleImagensalto1.Count == 1 && detalleImagensancho1.Count == 2)
                                {
                                    for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                                    {
                                        //Insertar imagenes
                                        string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                                        if (File.Exists(curFile))
                                        {
                                            contadorcondicional2++;
                                            switch (contadorcondicional2)
                                            {
                                                case 1:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "B";
                                                        Fila_General = "D";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensancho4[0].ToString(),
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "O";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensancho4[0].ToString(),
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "E";
                                                        Fila_General = "G";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensancho4[1],
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "P";
                                                        Fila_General = "R";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensancho4[1],
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "H";
                                                        Fila_General = "J";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensalto4[0],
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "S";
                                                        Fila_General = "U";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        xlWSheet.Shapes.AddPicture(detalleImagensalto4[0],
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                                    }
                                                    detalleImagensancho4 = null;
                                                    detalleImagensalto4 = null;
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }
                                if (distribucion == 2)
                                {
                                    Rang_colum += aumento;
                                    Rang_row += aumento;
                                }
                                break;
                        }
                    }
                }
            }
            else if (!Directory.Exists(Direccion_Configuracion_Mediciones_A))
            {
                MessageBox.Show("La Carpeta " + Direccion_Configuracion_Mediciones_A + " no existe");
            }

        }
        void CalcularTamanoImagen(string enlace)
        {
            if (File.Exists(enlace))
            {
                var image = new Bitmap(enlace);
                PropertyItem propertie = image.PropertyItems.FirstOrDefault(p => p.Id == 274);
                if (propertie != null)
                {
                    int orientation = propertie.Value[0];
                    if (orientation == 6)
                        image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                    if (orientation == 8)
                        image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                }

                int alto = image.Height;
                int ancho = image.Width;

                if (alto < ancho)
                {
                    IndicadordeTamaño = 1;
                }
                if (ancho < alto)
                {
                    IndicadordeTamaño = 2;
                }
                image = null;
            }
        }
        private void checkFormato2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna2.Checked)
            {
                txtConfiFormato2.Enabled = true;
                txtConfiFormato2.Text = "";
            }
            else if (!checkColumna2.Checked)
            {
                txtConfiFormato2.Enabled = false;
                txtConfiFormato2.Text = Formato_Default_2;
            }
        }
        private void checkCodigo2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo2.Checked)
            {
                txtConfiCodigo2.Enabled = true;
                txtConfiCodigo2.Text = "";
            }
            else if (!checkCodigo2.Checked)
            {
                txtConfiCodigo2.Enabled = false;
                txtConfiCodigo2.Text = Codigo_Default_2;
            }
        }
        private void checkFormato5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna5.Checked)
            {
                txtConfiFormato5.Enabled = true;
                txtConfiFormato5.Text = "";
            }
            else if (!checkColumna5.Checked)
            {
                txtConfiFormato5.Enabled = false;
                txtConfiFormato5.Text = Formato_Default_5;
            }
        }
        private void checkCodigo5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo5.Checked)
            {
                txtConfiCodigo5.Enabled = true;
                txtConfiCodigo5.Text = "";

            }
            else if (!checkCodigo5.Checked)
            {
                txtConfiCodigo5.Enabled = false;
                txtConfiCodigo5.Text = Codigo_Default_5;

            }
        }
        private void checkFormato6A_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna6A.Checked)
            {
                txtConfiFormato6A.Enabled = true;
                txtConfiFormato6A.Text = "";
            }
            else if (!checkColumna6A.Checked)
            {
                txtConfiFormato6A.Enabled = false;
                txtConfiFormato6A.Text = Formato_Default_6A;
            }
        }
        private void checkCodigo6A_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo6A.Checked)
            {
                txtConfiCodigo6A.Enabled = true;
                txtConfiCodigo6A.Text = "";
            }
            else if (!checkCodigo6A.Checked)
            {
                txtConfiCodigo6A.Enabled = false;
                txtConfiCodigo6A.Text = Codigo_Default_6A;

            }
        }
        private void checkFormato6B_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna6B.Checked)
            {
                txtConfiFormato6B.Enabled = true;
                txtConfiFormato6B.Text = "";
            }
            else if (!checkColumna6B.Checked)
            {
                txtConfiFormato6B.Enabled = false;
                txtConfiFormato6B.Text = Formato_Default_6B;
            }
        }
        private void checkCodigo6B_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo6B.Checked)
            {
                txtConfiCodigo6B.Enabled = true;
                txtConfiCodigo6B.Text = "";
            }
            else if (!checkCodigo6A.Checked)
            {
                txtConfiCodigo6B.Enabled = false;
                txtConfiCodigo6B.Text = Codigo_Default_6B;

            }
        }
        private void checkFormato8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna8.Checked)
            {
                txtConfiFormato8.Enabled = true;
                txtConfiFormato8.Text = "";
            }
            else if (!checkColumna8.Checked)
            {
                txtConfiFormato8.Enabled = false;
                txtConfiFormato8.Text = Formato_Default_8;
            }
        }
        private void checkCodigo8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo8.Checked)
            {
                txtConfiCodigo8.Enabled = true;
                txtConfiCodigo8.Text = "";
            }
            else if (!checkCodigo8.Checked)
            {
                txtConfiCodigo8.Enabled = false;
                txtConfiCodigo8.Text = Codigo_Default_8;
            }
        }
        private void checkFormato9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna9.Checked)
            {
                txtConfiFormato9.Enabled = true;
                txtConfiFormato9.Text = "";
            }
            else if (!checkColumna9.Checked)
            {
                txtConfiFormato9.Enabled = false;
                txtConfiFormato9.Text = Formato_Default_9;
            }
        }
        private void checkCodigo9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo9.Checked)
            {
                txtConfiCodigo9.Enabled = true;
                txtConfiCodigo9.Text = "";
            }
            else if (!checkCodigo9.Checked)
            {
                txtConfiCodigo9.Enabled = false;
                txtConfiCodigo9.Text = Codigo_Default_9;
            }
        }
        private void btnCrearFolder_Click(object sender, EventArgs e)
        {
            string folderPath;
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                folderPath = dialog.FileName + @"\Reportes de Instalación";
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                    string folderPath1 = folderPath + @"\2.Informacion_General";
                    Directory.CreateDirectory(folderPath1);
                    string folderPath2 = folderPath + @"\5.Pruebas de Interferencia";
                    Directory.CreateDirectory(folderPath2);
                    string folderPath3 = folderPath + @"\6.Configuracion_Mediciones";
                    Directory.CreateDirectory(folderPath3);
                    string folderPath31 = folderPath3 + @"\NODO1";
                    Directory.CreateDirectory(folderPath31);
                    string folderPath32 = folderPath3 + @"\NODO2";
                    Directory.CreateDirectory(folderPath32);
                    string folderPath4 = folderPath + @"\8.Reporte_Fotografico_A";
                    Directory.CreateDirectory(folderPath4);
                    string folderPath41 = folderPath4 + @"\1.Reporte_Fotografico";
                    Directory.CreateDirectory(folderPath41);
                    string folderPath5 = folderPath + @"\9.Reporte_Fotografico_B";
                    Directory.CreateDirectory(folderPath5);
                    string folderPath51 = folderPath5 + @"\1.Reporte_Fotografico";
                    Directory.CreateDirectory(folderPath51);
                    //
                    MessageBox.Show("Folder Estructurado Creado");
                    //
                    Process.Start(@folderPath);
                }
                else if (Directory.Exists(folderPath))
                {
                    MessageBox.Show("Folder ya existente");
                    Process.Start(@folderPath);

                }
            }
        }
    }
}