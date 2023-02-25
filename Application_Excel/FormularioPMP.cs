using Application_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Reportes
{
    public partial class FormularioPMP : Form
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
        Excel.Range RangoWidth;

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
        public string Formato_8 = "";
        public string Formato_9 = "";
        public string Formato_10 = "";
        public string Formato_11 = "";
        public string Formato_12 = "";
        public string Formato_13 = "";
        //
        public string Codigo_2 = "";
        public string Codigo_5 = "";
        public string Codigo_6A = "";
        public string Codigo_8 = "";
        public string Codigo_9 = "";
        public string Codigo_10 = "";
        public string Codigo_11 = "";
        public string Codigo_12 = "";
        public string Codigo_13 = "";
        //
        public string Formato_Default_2 = ".jpeg";
        public string Formato_Default_5 = ".png";
        public string Formato_Default_6A = ".png";
        public string Formato_Default_8 = ".jpeg";
        public string Formato_Default_9 = ".jpeg";
        public string Formato_Default_10 = ".jpeg";
        public string Formato_Default_11 = ".jpeg";
        public string Formato_Default_12 = ".jpeg";
        public string Formato_Default_13 = ".jpeg";
        //
        public string Codigo_Default_2 = "NODO_";
        public string Codigo_Default_5 = "S_";
        public string Codigo_Default_6A = "6_A_";
        public string Codigo_Default_8 = "A8_";
        public string Codigo_Default_9 = "A8_";
        public string Codigo_Default_10 = "A8_";
        public string Codigo_Default_11 = "A8_";
        public string Codigo_Default_12 = "A8_";
        public string Codigo_Default_13 = "A8_";
        //
        public int IndicadordeTamaño = 0;
        //
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
        public FormularioPMP()
        {
            InitializeComponent();
        }
        private void FormularioPMP_Load(object sender, EventArgs e)
        {
            txtConfiFormato2.Text = ".jpeg";
            txtConfiFormato5.Text = ".png";
            txtConfiFormato6A.Text = ".png";
            txtConfiFormato8.Text = ".jpeg";
            txtConfiFormato9.Text = ".jpeg";
            txtConfiFormato10.Text = ".jpeg";
            txtConfiFormato11.Text = ".jpeg";
            txtConfiFormato12.Text = ".jpeg";
            txtConfiFormato13.Text = ".jpeg";
            //
            txtConfiCodigo2.Text = "NODO_";
            txtConfiCodigo5.Text = "NODO_";
            txtConfiCodigo6A.Text = "6_A_";
            txtConfiCodigo8.Text = "A8_";
            txtConfiCodigo9.Text = "A8_";
            txtConfiCodigo10.Text = "A8_";
            txtConfiCodigo11.Text = "A8_";
            txtConfiCodigo12.Text = "A8_";
            txtConfiCodigo13.Text = "A8_";
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
                Formato_8 = txtConfiFormato8.Text;
                Formato_9 = txtConfiFormato9.Text;
                Formato_10 = txtConfiFormato10.Text;
                Formato_11 = txtConfiFormato11.Text;
                Formato_12 = txtConfiFormato12.Text;
                Formato_13 = txtConfiFormato13.Text;
                //
                Codigo_2 = txtConfiCodigo2.Text;
                Codigo_5 = txtConfiCodigo5.Text;
                Codigo_6A = txtConfiCodigo6A.Text;
                Codigo_8 = txtConfiCodigo8.Text;
                Codigo_9 = txtConfiCodigo9.Text;
                Codigo_10 = txtConfiCodigo10.Text;
                Codigo_11 = txtConfiCodigo11.Text;
                Codigo_12 = txtConfiCodigo12.Text;
                Codigo_13 = txtConfiCodigo13.Text;
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
                    string carpeta6 = "";
                    for (int i = 1; i <= 6; i++)
                    {
                        carpeta6 = "SECTOR_" + i.ToString();
                        //
                        string Nombrec6 = URL_Imagenes + @"\6.Configuracion\" + carpeta6;
                        if (Directory.Exists(Nombrec6))
                        {
                            string[] Cuenta6 = Directory.GetFiles(Nombrec6, "*");
                            int cuentaconfiguracion = Cuenta6.Length;
                            //
                            if (cuentaconfiguracion != 0)
                            {
                                xlWSheet = (Excel.Worksheet)oXL.Worksheets["6.Configuración"];
                                xlWSheet.Select(Type.Missing);
                                InsertarFila(i);
                            }
                        }
                    }
                    form2.Instalacion(6);
                    //Detectar Hoja de Trabajo 8
                    bool sheetExists8 = false;
                    string carpeta8 = "";
                    int Rangofila8 = 0;
                    foreach (Worksheet sheet in xlWBook.Sheets)
                    {
                        switch (sheet.Name)
                        {
                            case "8.Reporte fotográfico S1":
                                sheetExists8 = true;
                                carpeta8 = "8.Reporte_Fotografico_S1";
                                Rangofila8 = 1;
                                break;
                            case "8.Reporte fotográfico S2":
                                sheetExists8 = true;
                                carpeta8 = "8.Reporte_Fotografico_S2";
                                Rangofila8 = 2;
                                break;
                            case "8.Reporte fotográfico S3":
                                sheetExists8 = true;
                                carpeta8 = "8.Reporte_Fotografico_S4";
                                Rangofila8 = 3;
                                break;
                            case "8.Reporte fotográfico S4":
                                sheetExists8 = true;
                                carpeta8 = "8.Reporte_Fotografico_S4";
                                Rangofila8 = 4;
                                break;
                            case "8.Reporte fotográfico S5":
                                sheetExists8 = true;
                                carpeta8 = "8.Reporte_Fotografico_S5";
                                Rangofila8 = 5;
                                break;
                            case "8.Reporte fotográfico S6":
                                sheetExists8 = true;
                                carpeta8 = "8.Reporte_Fotografico_S6";
                                Rangofila8 = 6;
                                break;
                        }
                        string Nombrec8 = URL_Imagenes + @"\8.Reporte_Fotografico\" + carpeta8;
                        if (Directory.Exists(Nombrec8))
                        {
                            string[] Cuenta8 = Directory.GetFiles(Nombrec8, "*");
                            int cuentaconfiguracion = Cuenta8.Length;

                            if (sheetExists8 == true)
                            {
                                if (cuentaconfiguracion != 0)
                                {
                                    form2.Instalacion(8);
                                    //Detectar Hoja de Trabajo 8_A
                                    xlWSheet = (Excel.Worksheet)oXL.Worksheets[sheet.Name];
                                    xlWSheet.Select(Type.Missing);
                                    InsertarFila2(Rangofila8);
                                }
                            }
                        }

                        sheetExists8 = false;
                    }
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
                if (cantidad_Informacion_General == 0)
                {
                    MessageBox.Show("No hay contenido en la Carpeta 2.Informacion_General");
                }
                else
                {
                    int Rang_colum = 27;
                    int Rang_row = 47;
                    //
                    int conteofor = 0;
                    try
                    {
                        for (int i = 0; i <= cantidad_Informacion_General - 1; i++)
                        {
                            string dir2 = Informacion_General[i];
                            NombreImg2 = Path.GetFileName(dir2);
                            strlist = NombreImg2.Split(separador, separador.Length, StringSplitOptions.RemoveEmptyEntries);
                            Codigo[i] = strlist[0];
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Problema con la imagen: " + Informacion_General[conteofor], "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    int contador = Int32.Parse(Codigo.Max());
                    //
                    for (int cant_var = 1; cant_var <= contador; cant_var++)
                    {
                        string curFile = Direccion_Informacion_Gemeral + CodigoDigi + cant_var + Formato;

                        if ((cant_var % 2) == 0)
                        {
                            //
                        }
                        else
                        {
                            CalcularTamanoImagen(curFile);
                            if (IndicadordeTamaño == 1)
                            {
                                Rang_colum = 27;
                                Rang_row = 43;
                                Columna_General = "C";
                                Fila_General = "H";
                                RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                                //Insertar imagenes
                                if (File.Exists(curFile))
                                {
                                    ConvertirImagenJPEG_PNG(curFile);
                                }
                            }
                            else if (IndicadordeTamaño == 2)
                            {
                                Rang_colum = 24;
                                Rang_row = 47;
                                Columna_General = "D";
                                Fila_General = "G";
                                RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                                //Insertar imagenes
                                if (File.Exists(curFile))
                                {
                                    ConvertirImagenJPEG_PNG(curFile);
                                }
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
            string Direccion_Informacion_Gemeral = @URL_Imagenes + @"\5.Pruebas de Interferencia\";
            //
            if (Directory.Exists(Direccion_Informacion_Gemeral))
            {
                string[] Informacion_General = Directory.GetFiles(Direccion_Informacion_Gemeral, "*" + Formato);
                int cantidad_Informacion_General = Informacion_General.Length;
                string NombreImg2 = null;
                String[] Codigo = new String[100];
                String[] Numeracion = new String[100];
                String[] strlist = new String[100];
                String[] separador = { Direccion_Informacion_Gemeral, CodigoDigi, Formato };
                //
                if (cantidad_Informacion_General == 0)
                {
                    MessageBox.Show("No hay contenido en la Carpeta 5.Pruebas de Interferencia");
                }
                else
                {
                    int Rang_colum = 21;
                    int Rang_row = 35;
                    //
                    int aumento = 17;
                    int conteofor = 0;
                    try
                    {
                        for (int i = 0; i <= cantidad_Informacion_General - 1; i++)
                        {
                            string dir2 = Informacion_General[i];
                            NombreImg2 = Path.GetFileName(dir2);
                            strlist = NombreImg2.Split(separador, separador.Length, StringSplitOptions.RemoveEmptyEntries);
                            Codigo[i] = strlist[0];
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Problema con la imagen: " + Informacion_General[conteofor], "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    int contador = Int32.Parse(Codigo.Max());
                    for (int cant_var = 1; cant_var <= contador; cant_var++)
                    {
                        string curFile = Direccion_Informacion_Gemeral + CodigoDigi + cant_var + Formato;

                        //Insertar imagenes
                        if (File.Exists(curFile))
                        {
                            CalcularTamanoImagen(curFile);
                            if (IndicadordeTamaño == 1)
                            {
                                Columna_General = "C";
                                Fila_General = "F";
                            }
                            else if (IndicadordeTamaño == 2)
                            {
                                Columna_General = "D";
                                Fila_General = "E";
                            }
                            //Asignar Rango
                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                            //
                            ConvertirImagenJPEG_PNG(curFile);
                            Rang_colum += aumento;
                            Rang_row += aumento;
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
            string asignador = "";
            Formato = Formato_6A;
            CodigoDigi = Codigo_6A;
            //
            int Rang_colum = 0;
            int Rang_row = 0;
            if (RangoFila1 == 1)
            {
                asignador = "SECTOR_1";
                Rang_colum = 35;
                Rang_row = 48;
            }
            else if (RangoFila1 == 2)
            {
                asignador = "SECTOR_2";
                Rang_colum = 117;
                Rang_row = 130;
            }
            else if (RangoFila1 == 3)
            {
                asignador = "SECTOR_3";
                Rang_colum = 199;
                Rang_row = 212;
            }
            else if (RangoFila1 == 4)
            {
                asignador = "SECTOR_4";
                Rang_colum = 281;
                Rang_row = 294;
            }
            else if (RangoFila1 == 5)
            {
                asignador = "SECTOR_5";
                Rang_colum = 363;
                Rang_row = 376;
            }
            else if (RangoFila1 == 6)
            {
                asignador = "SECTOR_6";
                Rang_colum = 449;
                Rang_row = 450;
            }
            string Direccion_Configuracion_Mediciones_A = URL_Imagenes + @"\6.Configuracion\" + asignador + @"\";
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
                if (cantidad_Configuracion_Mediciones_A == 0)
                {
                    MessageBox.Show("No hay contenido en la Carpeta 6.Configuracion");
                }
                else
                {
                    int conteofor = 0;
                    try
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
                    }
                    catch
                    {
                        MessageBox.Show("Problema con la imagen: " + Configuracion_Mediciones_A[conteofor], "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var cantidadmaxima = codigoNumeracion.Max(x => x.grupo);
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
                        if ((cant_var % 2) == 0)
                        {
                            string curFile = Direccion_Configuracion_Mediciones_A + CodigoDigi + cant_var + CodigoIntermedio + "1" + Formato;

                            if (File.Exists(curFile))
                            {
                                CalcularTamanoImagen(curFile);
                                if (IndicadordeTamaño == 1)
                                {
                                    //Rang_colum += 2;
                                    //Rang_row -= 2;
                                    Columna_General = "F";
                                    Fila_General = "I";
                                    //
                                    RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                    //
                                    ConvertirImagenJPEG_PNG(curFile);
                                    //Rang_colum -= 2;
                                    //Rang_row += 2;
                                }
                                else if (IndicadordeTamaño == 2)
                                {
                                    Columna_General = "G";
                                    Fila_General = "H";
                                    //
                                    RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                    //
                                    ConvertirImagenJPEG_PNG(curFile);
                                }
                            }
                            Rang_colum += aumento;
                            Rang_row += aumento;
                        }
                        else
                        {
                            string curFile = Direccion_Configuracion_Mediciones_A + CodigoDigi + cant_var + CodigoIntermedio + "1" + Formato;
                            if (File.Exists(curFile))
                            {
                                CalcularTamanoImagen(curFile);
                                if (IndicadordeTamaño == 1)
                                {
                                    //Rang_colum += 2;
                                    //Rang_row -= 2;
                                    Columna_General = "B";
                                    Fila_General = "E";
                                    //
                                    RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                    //
                                    ConvertirImagenJPEG_PNG(curFile);
                                    //Rang_colum -= 2;
                                    //Rang_row += 2;
                                }
                                else if (IndicadordeTamaño == 2)
                                {
                                    Columna_General = "C";
                                    Fila_General = "D";
                                    //
                                    RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                    //
                                    ConvertirImagenJPEG_PNG(curFile);
                                }
                            }
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
            int numerador = 8;
            string asignador = "";
            if (RangoFila1 == 1)
            {
                Formato = Formato_8;
                asignador = "S1";
                CodigoDigi = Codigo_8;
            }
            else if (RangoFila1 == 2)
            {
                Formato = Formato_9;
                asignador = "S2";
                CodigoDigi = Codigo_9;
            }
            else if (RangoFila1 == 3)
            {
                Formato = Formato_10;
                asignador = "S3";
                CodigoDigi = Codigo_10;
            }
            else if (RangoFila1 == 4)
            {
                Formato = Formato_11;
                asignador = "S4";
                CodigoDigi = Codigo_11;
            }
            else if (RangoFila1 == 5)
            {
                Formato = Formato_12;
                asignador = "S5";
                CodigoDigi = Codigo_12;
            }
            else if (RangoFila1 == 6)
            {
                Formato = Formato_13;
                asignador = "S6";
                CodigoDigi = Codigo_13;
            }
            //Funcion Reporte Fotografico
            string Direccion_Configuracion_Mediciones_A = URL_Imagenes + @"\8.Reporte_Fotografico\" + numerador + ".Reporte_Fotografico_" + asignador + @"\";
            //
            if (Directory.Exists(Direccion_Configuracion_Mediciones_A))
            {
                //
                string[] Configuracion_Mediciones_A = Directory.GetFiles(Direccion_Configuracion_Mediciones_A, "*" + Formato);
                int cantidad_Configuracion_Mediciones_A = Configuracion_Mediciones_A.Length;
                string NombreImgConfiguracion_A = null;
                //
                String[] CodigoConfiguracion_A = new String[100];
                String[] NumeracionConfiguracion_A = new String[100];
                String[] strlistConfiguracion_A = new String[100];
                String[] separadorConfiguracion_A = { Direccion_Configuracion_Mediciones_A, CodigoDigi, CodigoIntermedio, Formato };
                //Listas
                List<CodigoNumeracion> codigoNumeracion = new List<CodigoNumeracion>();
                List<CodigoNumeracion> codigoejemplo = new List<CodigoNumeracion>();
                //
                if (cantidad_Configuracion_Mediciones_A == 0)
                {
                    MessageBox.Show("No hay contenido en la Carpeta " + numerador + ".Reporte_Fotografico_" + asignador);
                }
                else
                {
                    int conteofor = 0;
                    try
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
                    }
                    catch
                    {
                        MessageBox.Show("Problema con la imagen: " + Configuracion_Mediciones_A[conteofor], "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        if (cant_var == 10)
                        {
                            aumento = 17;
                        }
                        if (cant_var >= 11 && cant_var < 22)
                        {
                            aumento = 16;
                        }
                        //
                        if (cant_var == 22)
                        {
                            aumento = 17;
                        }
                        if (cant_var >= 23 && cant_var < 26)
                        {
                            aumento = 16;
                        }
                        //
                        if (cant_var == 26)
                        {
                            aumento = 18;
                        }
                        if (cant_var >= 27)
                        {
                            aumento = 16;
                        }
                        //
                        switch (cantidadcodigo)
                        {
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

                                        ConvertirImagenJPEG_PNG(curFile);
                                    }
                                }
                                if (distribucion == 2)
                                {
                                    Rang_colum += aumento;
                                    Rang_row += aumento;
                                }
                                break;
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "N";
                                                        Fila_General = "P";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "G";
                                                        Fila_General = "I";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "R";
                                                        Fila_General = "T";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "Q";
                                                        Rang_row = Rang_row - 3;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_row = Rang_row + 3;
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "Q";
                                                        Fila_General = "U";
                                                        Rang_colum = Rang_colum + 3;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 3;
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "G";
                                                        Fila_General = "I";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "R";
                                                        Fila_General = "T";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "O";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "E";
                                                        Fila_General = "G";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "P";
                                                        Fila_General = "R";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "H";
                                                        Fila_General = "J";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "S";
                                                        Fila_General = "U";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    break;
                                                default:
                                                    break;
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "P";
                                                        Rang_row = Rang_row - 6;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_row = Rang_row + 6;
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "P";
                                                        Rang_colum = Rang_colum + 6;
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        Rang_colum = Rang_colum - 6;
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(curFile);
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
                                                        ConvertirImagenJPEG_PNG(detalleImagensalto4[0].ToString());
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "O";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensalto4[0].ToString());
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "E";
                                                        Fila_General = "G";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensalto4[1].ToString());
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "P";
                                                        Fila_General = "R";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensalto4[1].ToString());
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "H";
                                                        Fila_General = "J";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensancho4[0].ToString());
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "S";
                                                        Fila_General = "U";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensancho4[0].ToString());
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
                                                        ConvertirImagenJPEG_PNG(detalleImagensancho4[0].ToString());
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "M";
                                                        Fila_General = "O";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensancho4[0].ToString());
                                                    }
                                                    break;
                                                case 2:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "E";
                                                        Fila_General = "G";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensancho4[1].ToString());
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "P";
                                                        Fila_General = "R";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensancho4[1].ToString());
                                                    }
                                                    break;
                                                case 3:
                                                    if (distribucion == 1)
                                                    {
                                                        Columna_General = "H";
                                                        Fila_General = "J";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensalto4[0].ToString());
                                                    }
                                                    else if (distribucion == 2)
                                                    {
                                                        Columna_General = "S";
                                                        Fila_General = "U";
                                                        RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                                        ConvertirImagenJPEG_PNG(detalleImagensalto4[0].ToString());
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
        void ConvertirImagenJPEG_PNG(string ruta)
        {
            //// Load the JPEG image from a file
            //var jpegStream = File.OpenRead(ruta);
            //var jpegImage = Image.FromStream(jpegStream);

            //// Convert the JPEG image to PNG format with higher resolution
            //var pngStream = new MemoryStream();
            //var pngEncoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Png.Guid);
            //var pngEncoderParams = new EncoderParameters(1);
            //pngEncoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)100);
            //jpegImage.Save(pngStream, pngEncoder, pngEncoderParams);

            //// Save the PNG image to a temporary file
            //var tempFileName = Path.GetTempFileName();
            //File.WriteAllBytes(tempFileName, pngStream.ToArray());

            //// Add the image to the Excel worksheet
            //xlWSheet.Shapes.AddPicture(tempFileName,
            //                           Microsoft.Office.Core.MsoTriState.msoTrue,
            //                           Microsoft.Office.Core.MsoTriState.msoTrue,
            //                           float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
            //                           float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));

            //// Delete the temporary file
            //File.Delete(tempFileName);

            xlWSheet.Shapes.AddPicture(ruta,
                                       Microsoft.Office.Core.MsoTriState.msoTrue,
                                       Microsoft.Office.Core.MsoTriState.msoTrue,
                                       float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                       float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));

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
        private void checkFormato10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna10.Checked)
            {
                txtConfiFormato10.Enabled = true;
                txtConfiFormato10.Text = "";
            }
            else if (!checkColumna10.Checked)
            {
                txtConfiFormato10.Enabled = false;
                txtConfiFormato10.Text = Formato_Default_10;
            }
        }
        private void checkCodigo10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo10.Checked)
            {
                txtConfiCodigo10.Enabled = true;
                txtConfiCodigo10.Text = "";
            }
            else if (!checkCodigo9.Checked)
            {
                txtConfiCodigo10.Enabled = false;
                txtConfiCodigo10.Text = Codigo_Default_10;
            }
        }
        private void checkFormato11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna11.Checked)
            {
                txtConfiFormato11.Enabled = true;
                txtConfiFormato11.Text = "";
            }
            else if (!checkColumna11.Checked)
            {
                txtConfiFormato11.Enabled = false;
                txtConfiFormato11.Text = Formato_Default_11;
            }
        }
        private void checkCodigo11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo11.Checked)
            {
                txtConfiCodigo11.Enabled = true;
                txtConfiCodigo11.Text = "";
            }
            else if (!checkCodigo9.Checked)
            {
                txtConfiCodigo11.Enabled = false;
                txtConfiCodigo11.Text = Codigo_Default_11;
            }
        }
        private void checkFormato12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna12.Checked)
            {
                txtConfiFormato12.Enabled = true;
                txtConfiFormato12.Text = "";
            }
            else if (!checkColumna12.Checked)
            {
                txtConfiFormato12.Enabled = false;
                txtConfiFormato12.Text = Formato_Default_12;
            }
        }
        private void checkCodigo12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo12.Checked)
            {
                txtConfiCodigo12.Enabled = true;
                txtConfiCodigo12.Text = "";
            }
            else if (!checkCodigo9.Checked)
            {
                txtConfiCodigo12.Enabled = false;
                txtConfiCodigo12.Text = Codigo_Default_12;
            }
        }
        private void checkFormato13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna13.Checked)
            {
                txtConfiFormato13.Enabled = true;
                txtConfiFormato13.Text = "";
            }
            else if (!checkColumna13.Checked)
            {
                txtConfiFormato13.Enabled = false;
                txtConfiFormato13.Text = Formato_Default_13;
            }
        }
        private void checkCodigo13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo13.Checked)
            {
                txtConfiCodigo13.Enabled = true;
                txtConfiCodigo13.Text = "";
            }
            else if (!checkCodigo9.Checked)
            {
                txtConfiCodigo13.Enabled = false;
                txtConfiCodigo13.Text = Codigo_Default_13;
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
                    string folderPath3 = folderPath + @"\6.Configuracion";
                    Directory.CreateDirectory(folderPath3);
                    string folderPath31 = folderPath3 + @"\SECTOR_1";
                    Directory.CreateDirectory(folderPath31);
                    string folderPath32 = folderPath3 + @"\SECTOR_2";
                    Directory.CreateDirectory(folderPath32);
                    string folderPath33 = folderPath3 + @"\SECTOR_3";
                    Directory.CreateDirectory(folderPath33);
                    string folderPath34 = folderPath3 + @"\SECTOR_4";
                    Directory.CreateDirectory(folderPath34);
                    string folderPath35 = folderPath3 + @"\SECTOR_5";
                    Directory.CreateDirectory(folderPath35);
                    string folderPath36 = folderPath3 + @"\SECTOR_6";
                    Directory.CreateDirectory(folderPath36);
                    string folderPath4 = folderPath + @"\8.Reporte_Fotografico";
                    Directory.CreateDirectory(folderPath4);
                    string folderPath41 = folderPath4 + @"\8.Reporte_Fotografico_S1";
                    Directory.CreateDirectory(folderPath41);
                    string folderPath42 = folderPath4 + @"\8.Reporte_Fotografico_S2";
                    Directory.CreateDirectory(folderPath42);
                    string folderPath43 = folderPath4 + @"\8.Reporte_Fotografico_S3";
                    Directory.CreateDirectory(folderPath43);
                    string folderPath44 = folderPath4 + @"\8.Reporte_Fotografico_S4";
                    Directory.CreateDirectory(folderPath44);
                    string folderPath45 = folderPath4 + @"\8.Reporte_Fotografico_S5";
                    Directory.CreateDirectory(folderPath45);
                    string folderPath46 = folderPath4 + @"\8.Reporte_Fotografico_S6";
                    Directory.CreateDirectory(folderPath46);
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