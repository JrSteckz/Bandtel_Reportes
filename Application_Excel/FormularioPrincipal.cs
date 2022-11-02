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
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;

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
        public int Columna_2 = 0;
        public int Fila_2 = 0;
        public int Columna_6 = 0;
        public int Fila_6 = 0;
        public int Columna_8 = 0;
        public int Fila_8 = 0;
        public int Columna_9 = 0;
        public int Fila_9 = 0;
        //
        public string Columna_Default_2 = "";
        public string Fila_Default_2 = "";
        public string Columna_Default_6 = "";
        public string Fila_Default_6 = "";
        public string Columna_Default_8 = "";
        public string Fila_Default_8 = "";
        public string Columna_Default_9 = "";
        public string Fila_Default_9 = "";
        //
        Excel.Range RangoWidth;
        public class CodigoNumeracion
        {
            public int grupo { get; set; }
            public int numeracion { get; set; }
        }
        public class CodigoCantidad
        {
            public int grupo { get; set; }
            public int cantidad { get; set; }
        }
        public FormularioPrincipal()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            txtConfiColumna2.Text = "27";
            txtConfiColumna5.Text = "";
            txtConfiColumna6.Text = "50";
            txtConfiColumna8.Text = "11";
            txtConfiColumna9.Text = "11";
            //
            txtConfiFila2.Text = "47";
            txtConfiFila5.Text = "";
            txtConfiFila6.Text = "63";
            txtConfiFila8.Text = "22";
            txtConfiFila9.Text = "22";
            //
            txtConfiCodigo2.Text = "";
            txtConfiCodigo5.Text = "";
            txtConfiCodigo6.Text = "";
            txtConfiCodigo8.Text = "";
            txtConfiCodigo9.Text = "";
            //
            //txtURL.Enabled = false;
            //txtNombreExcel.Enabled = false;
            txtURL.Text = @"C:\Users\estef\Desktop";
            URL_Guardado = txtURL.Text;
            txtImagenes.Text = @"C:\Users\estef\Desktop\Reportes de Instalación";
            URL_Imagenes = txtImagenes.Text;
            txtUbicacionPlantilla.Text = @"C:\Users\estef\Desktop\COD-NODO_Reporte_de_Instalacion_PTP_22_09_2022_V3.xlsx";
            URL_Plantilla = txtUbicacionPlantilla.Text;

            btnGenerar.Enabled = true;
        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            var form2 = new FormularioProgressBar();
            form2.Show();
            if (txtNombreExcel.Text.Trim().Length <= 1)
            {
                MessageBox.Show("Faltan llenar Datos");
            }
            else
            {
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
                Insertarprimera();
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
                form2.Instalacion(10);

                //Guardar Excel
                string Lugar_Guardado = txtURL.Text + @"\";
                string NombreExcel = txtNombreExcel.Text + @".xlsx";
                xlWBook.SaveAs(Lugar_Guardado + NombreExcel);
                xlWBook.Close(true, Type.Missing, Type.Missing);
                oXL.Quit();
                MessageBox.Show("Archivo Guardado");
                Process.Start(Lugar_Guardado + NombreExcel);
                form2.Close();

            }
        }
        private void BtnBuscadorGuardado_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            DialogResult res = dlg.ShowDialog();
            

            if (res == System.Windows.Forms.DialogResult.OK)
            {
                URL_Guardado = dlg.SelectedPath;
                txtURL.Text = URL_Guardado;
            }

            //txtNombreExcel.Enabled = true;
            btnGenerar.Enabled = true;
        }
        private void BtnBuscadorPlantilla_Click(object sender, EventArgs e)
        {
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File";
            theDialog.Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            DialogResult res = theDialog.ShowDialog();

            if (res == System.Windows.Forms.DialogResult.OK)
            {
                URL_Plantilla = theDialog.FileName;
                txtUbicacionPlantilla.Text = URL_Plantilla;

            }
            //txtUbicacionPlantilla.Enabled = false;
        }
        private void BtnBuscadorImagenes_Click(object sender, EventArgs e)
        { 
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            DialogResult res = dlg.ShowDialog();

            if (res == System.Windows.Forms.DialogResult.OK)
            {
                URL_Imagenes = dlg.SelectedPath;
                txtImagenes.Text = URL_Imagenes;

            }

            //txtUbicacionPlantilla.Enabled = false;
        }
        void Insertarprimera()
        {
            Columna_General = "K";
            Fila_General = "O";
            Formato = ".jpeg";
            CodigoDigi = "NODO_";
            //
            string Direccion_Informacion_Gemeral = @URL_Imagenes + @"\1.Informacion_General\";
            string[] Informacion_General = Directory.GetFiles(Direccion_Informacion_Gemeral, "*" + Formato);
            int cantidad_Informacion_General = Informacion_General.Length;
            string NombreImg2 = null;
            //
            String[] Codigo = new String[100];
            String[] Numeracion = new String[100];
            String[] strlist = new String[100];
            String[] separador = { Direccion_Informacion_Gemeral, CodigoDigi, Formato };
            //
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
                    RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                    //Insertar imagenes
                    if (File.Exists(curFile))
                    {
                        xlWSheet.Shapes.AddPicture(Direccion_Informacion_Gemeral + CodigoDigi + cant_var + Formato,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                    }
                }
                else
                {
                    //Asignar Rango
                    RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);

                    //Insertar imagenes
                    if (File.Exists(curFile))
                    {
                        xlWSheet.Shapes.AddPicture(Direccion_Informacion_Gemeral + CodigoDigi + cant_var + Formato,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                    }
                }
            }
        }
        void InsertarFila(int RangoFila1) 
        {
            Columna_General = "";
            Fila_General = "";
            CodigoIntermedio = "_";
            Formato = ".png";
            int numerador = 0;
            if (RangoFila1 == 1)
            {
                numerador = 1;
                CodigoDigi = "6_A_";
            }
            else if (RangoFila1 == 2)
            {
                numerador = 2;
                CodigoDigi = "6_B_";
            }
            string Direccion_Configuracion_Mediciones_A = URL_Imagenes + @"\2.Configuracion_Mediciones\NODO" + numerador;
            string[] Configuracion_Mediciones_A = Directory.GetFiles(Direccion_Configuracion_Mediciones_A, "*" + Formato);
            int cantidad_Configuracion_Mediciones_A = Configuracion_Mediciones_A.Length;
            string NombreImgConfiguracion_A = null;
            //
            String[] CodigoConfiguracion_A = new String[100];
            String[] NumeracionConfiguracion_A = new String[100];
            String[] strlistConfiguracion_A = new String[100];
            String[] separadorConfiguracion_A = { @URL_Imagenes + @"\2.Configuracion_Mediciones\NODO" + numerador + @"\", CodigoDigi, CodigoIntermedio, Formato };
            //
            List<CodigoNumeracion> codigoNumeracion = new List<CodigoNumeracion>();
            List<CodigoCantidad> codigoCantidad = new List<CodigoCantidad>();
            //
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

            foreach (var item in codigoNumeracion)
            {
                CodigoCantidad aes = new CodigoCantidad();
                aes.grupo = item.grupo;
                aes.cantidad = codigoNumeracion.Count(x => x.grupo == item.grupo);
                //
                if (!codigoCantidad.Exists(x => x.grupo == item.grupo))
                {
                    codigoCantidad.Add(aes);
                }
            }
            var cantidadmaxima = codigoNumeracion.Max(x => x.grupo);

            int Rang_colum = 50;
            int Rang_row = 63;
            int aumento = 16;
            //Bucle de insertado de imagenes
            for (int cant_var = 1; cant_var <= cantidadmaxima; cant_var++)
            { 
                //Asignar Rango

                var NumeracionCodigo = codigoCantidad.Where(x => x.grupo == cant_var);
                var cantidadcodigo = 0;
                foreach (var value in NumeracionCodigo)
                {
                    cantidadcodigo = value.cantidad;
                }

                switch (cantidadcodigo)
                {
                    case 0:
                        Rang_colum += aumento;
                        Rang_row += aumento;
                        break;
                    case 1:

                        if (RangoFila1 == 1)
                        {
                            Columna_General = "C";
                            Fila_General = "I";
                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                        }
                        else if (RangoFila1 == 2)
                        {
                            Columna_General = "M";
                            Fila_General = "S";
                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                        }
                        for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                        {
                            string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                            if (File.Exists(curFile))
                            {
                                xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                            }
                        }
                        Rang_colum += aumento;
                        Rang_row += aumento;
                        break;
                    case 2:

                        int contadorcondicional = 0;
                        for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                        {
                            //Insertar imagenes
                            string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                            //
                            if (File.Exists(curFile))
                            {
                                contadorcondicional++;
                                switch (contadorcondicional)
                                {
                                    case 1:
                                        if (RangoFila1 == 1)
                                        {
                                            Columna_General = "B";
                                            Fila_General = "E";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        else if (RangoFila1 == 2)
                                        {
                                            Columna_General = "L";
                                            Fila_General = "O";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    case 2:
                                        if (RangoFila1 == 1)
                                        {
                                            Columna_General = "G";
                                            Fila_General = "J";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        else if (RangoFila1 == 2)
                                        {
                                            Columna_General = "Q";
                                            Fila_General = "T";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }

                        Rang_colum += aumento;
                        Rang_row += aumento;
                        break;
                    default:
                        break;
                }
            }
        }
        void InsertarFila2(int RangoFila1)
        {
            Formato = ".jpeg";
            CodigoIntermedio = "_";
            int numerador = 0;
            string asignador = "";
            if (RangoFila1 == 1)
            {
                numerador = 3;
                asignador = "A";
                CodigoDigi = "A8_";
            }
            else if (RangoFila1 == 2)
            {
                numerador = 4;
                asignador = "B";
                CodigoDigi = "B9_";
            }
            //Funcion Reporte Fotografico
            string Direccion_Configuracion_Mediciones_A = URL_Imagenes + @"\" + numerador + ".Reporte_Fotografico_" + asignador + @"\1.Reporte_Fotografico";
            string[] Configuracion_Mediciones_A = Directory.GetFiles(Direccion_Configuracion_Mediciones_A, "*" + Formato);
            int cantidad_Configuracion_Mediciones_A = Configuracion_Mediciones_A.Length;
            string NombreImgConfiguracion_A = null;
            //
            String[] CodigoConfiguracion_A = new String[100];
            String[] NumeracionConfiguracion_A = new String[100];
            String[] strlistConfiguracion_A = new String[100];
            String[] separadorConfiguracion_A = { @URL_Imagenes + @"\"+ numerador + @".Reporte_Fotografico_" + asignador + @"\1.Reporte_Fotografico\", CodigoDigi, CodigoIntermedio, Formato };
            //Funcion Serie de Equipos
            //string Direccion_Configuracion_Mediciones_B = URL_Imagenes + @"\" + numerador + @".Reporte_Fotografico_" + asignador + @"\2.Serie_Equipos";
            //string[] Configuracion_Mediciones_B = Directory.GetFiles(Direccion_Configuracion_Mediciones_B, "*" + Formato);
            //int cantidad_Configuracion_Mediciones_B = Configuracion_Mediciones_B.Length;
            //string NombreImgConfiguracion_B = null;
            ////
            //String[] CodigoConfiguracion_B = new String[100];
            //String[] NumeracionConfiguracion_B = new String[100];
            //String[] strlistConfiguracion_B = new String[100];
            //String[] separadorConfiguracion_B = { @URL_Imagenes + @"\" + numerador + @".Reporte_Fotografico_" + asignador + @"\2.Serie_Equipos\", "C_", "_", Formato };
            //Listas
            List<CodigoNumeracion> codigoNumeracion = new List<CodigoNumeracion>();
            List<CodigoNumeracion> codigoejemplo = new List<CodigoNumeracion>();
            List<CodigoCantidad> codigoCantidad = new List<CodigoCantidad>();
            //
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
                if(cant_var == 36)
                {
                    aumento = 18;
                }
                if(cant_var >= 37)
                {
                    aumento = 16;
                }
                switch (cantidadcodigo)
                {
                    case -1:
                        if (distribucion == 2)
                        {
                            Rang_colum += aumento;
                            Rang_row += aumento;
                        }
                        break;
                    case 0:
                        if (distribucion == 2)
                        {
                            Rang_colum += aumento;
                            Rang_row += aumento;
                        }
                        break;
                    case 1:
                        if (distribucion == 1)
                        {
                            Columna_General = "C";
                            Fila_General = "I";
                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                        }
                        else if (distribucion == 2)
                        {
                            Columna_General = "N";
                            Fila_General = "T";
                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                        }
                        for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                        {
                            string curFile = Direccion_Configuracion_Mediciones_A + @"\"+ CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                            if (File.Exists(curFile))
                            {
                                xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                            }
                            else
                            {
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
                        for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                        {
                            //Insertar imagenes
                            string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                            //
                            if (File.Exists(curFile))
                            {
                                contadorcondicional++;
                                switch (contadorcondicional)
                                {
                                    case 1:
                                        if (distribucion == 1)
                                        {
                                            Columna_General = "B";
                                            Fila_General = "E";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            Columna_General = "M";
                                            Fila_General = "P";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    case 2:
                                        if (distribucion == 1)
                                        {
                                            Columna_General = "G";
                                            Fila_General = "J";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            Columna_General = "R";
                                            Fila_General = "U";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        if(distribucion == 2)
                        {
                            Rang_colum += aumento;
                            Rang_row += aumento;
                        }
                        break;
                    case 3:
                        int contadorcondicional2 = 0;
                        for (int numeracionciclo = 1; numeracionciclo <= 3; numeracionciclo++)
                        {
                            //Insertar imagenes
                            string curFile = Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato;
                            //
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
                                        }
                                        else if (distribucion == 2)
                                        {
                                            Columna_General = "M";
                                            Fila_General = "O";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    case 2:
                                        if (distribucion == 1)
                                        {
                                            Columna_General = "E";
                                            Fila_General = "G";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            Columna_General = "P";
                                            Fila_General = "R";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    case 3:
                                        if (distribucion == 1)
                                        {
                                            Columna_General = "H";
                                            Fila_General = "J";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            Columna_General = "S";
                                            Fila_General = "U";
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range(Columna_General + Rang_colum, Fila_General + Rang_row);
                                        }
                                        //
                                        xlWSheet.Shapes.AddPicture(Direccion_Configuracion_Mediciones_A + @"\" + CodigoDigi + cant_var + CodigoIntermedio + numeracionciclo + Formato,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                                        break;
                                    default:
                                        break;
                                }
                            }

                        }
                        if (distribucion == 2)
                        {
                            Rang_colum += aumento;
                            Rang_row += aumento;
                        }
                        break;
                    default:

                        break;
                }
            }
        }
        private void checkColumna2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna2.Checked)
            {
                txtConfiColumna2.Enabled = true;
                txtConfiColumna2.Text = "";
                
            }
            else if (!checkColumna2.Checked)
            {
                txtConfiColumna2.Enabled = false;
                Columna_Default_2 = "27";
                txtConfiColumna2.Text = Columna_Default_2;
            }
        }
        private void checkFila2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFila2.Checked)
            {
                txtConfiFila2.Enabled = true;
                txtConfiFila2.Text = "";
            }
            else if (!checkFila2.Checked)
            {
                txtConfiFila2.Enabled = false;
                Fila_Default_2 = "47";
                txtConfiFila2.Text = Fila_Default_2;
            }
        }
        private void checkCodigo2_CheckedChanged(object sender, EventArgs e)
        {
            if(checkCodigo2.Checked)
            {
                txtConfiCodigo2.Enabled = true;
            }
            else if (!checkCodigo2.Checked)
            {
                txtConfiCodigo2.Enabled = false;
            }
        }
        private void checkColumna5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna5.Checked)
            {
                txtConfiColumna5.Enabled = true;
            }
            else if (!checkColumna5.Checked)
            {
                txtConfiColumna5.Enabled = false;
            }
        }
        private void checkFila5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFila5.Checked)
            {
                txtConfiFila5.Enabled = true;
            }
            else if (!checkFila5.Checked)
            {
                txtConfiFila5.Enabled = false;
            }
        }
        private void checkCodigo5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo5.Checked)
            {
                txtConfiCodigo5.Enabled = true;
            }
            else if (!checkCodigo5.Checked)
            {
                txtConfiCodigo5.Enabled = false;
            }
        }
        private void checkColumna6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna6.Checked)
            {
                txtConfiColumna6.Enabled = true;
                txtConfiColumna6.Text = "";
            }
            else if (!checkColumna6.Checked)
            {
                txtConfiColumna6.Enabled = false;
                Columna_Default_6 = "50";
                txtConfiColumna6.Text = Columna_Default_6;
            }
        }
        private void checkFila6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFila6.Checked)
            {
                txtConfiFila6.Enabled = true;
                txtConfiFila6.Text = "";
            }
            else if (!checkFila6.Checked)
            {
                txtConfiFila6.Enabled = false;
                Fila_Default_6 = "63";
                txtConfiFila6.Text = Fila_Default_6;
            }
        }
        private void checkCodigo6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo6.Checked)
            {
                txtConfiCodigo6.Enabled = true;
            }
            else if (!checkCodigo6.Checked)
            {
                txtConfiCodigo6.Enabled = false;
            }
        }
        private void checkColumna8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna8.Checked)
            {
                txtConfiColumna8.Enabled = true;
                txtConfiColumna8.Text = "";
            }
            else if (!checkColumna8.Checked)
            {
                txtConfiColumna8.Enabled = false;
                Columna_Default_8 = "11";
                txtConfiColumna8.Text = Columna_Default_8;
            }
        }
        private void checkFila8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFila8.Checked)
            {
                txtConfiFila8.Enabled = true;
                txtConfiFila8.Text = "";
            }
            else if (!checkFila8.Checked)
            {
                txtConfiFila8.Enabled = false;
                Fila_Default_8 = "22";
                txtConfiFila8.Text = Fila_Default_8;
            }
        }
        private void checkCodigo8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo8.Checked)
            {
                txtConfiCodigo8.Enabled = true;
            }
            else if (!checkCodigo8.Checked)
            {
                txtConfiCodigo8.Enabled = false;
            }
        }
        private void checkColumna9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkColumna9.Checked)
            {
                txtConfiColumna9.Enabled = true;
                txtConfiColumna9.Text = "";
            }
            else if (!checkColumna9.Checked)
            {
                txtConfiColumna9.Enabled = false;
                Columna_Default_9 = "11";
                txtConfiColumna9.Text = Columna_Default_9;
            }
        }
        private void checkFila9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkFila9.Checked)
            {
                txtConfiFila9.Enabled = true;
                txtConfiFila9.Text = "";
            }
            else if (!checkFila9.Checked)
            {
                txtConfiFila9.Enabled = false;
                Fila_Default_9 = "22";
                txtConfiFila9.Text = Fila_Default_9;
            }
        }
        private void checkCodigo9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCodigo9.Checked)
            {
                txtConfiCodigo9.Enabled = true;
            }
            else if (!checkCodigo9.Checked)
            {
                txtConfiCodigo9.Enabled = false;
            }
        }
    }
}
