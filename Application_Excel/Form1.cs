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

namespace Application_Excel
{
    public partial class Form1 : Form
    {
        public string URL_Guardado = "";
        public string URL_Plantilla = "";
        public string URL_Imagenes = "";
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
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
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
                Insertarprimera();

                //Hoja de Trabajo 5
                xlWSheet = (Excel.Worksheet)xlWBook.Sheets["5.Pruebas de Interferencia"];
                xlWSheet.Select(Type.Missing);
                Insertarprimera();

                //Detectar Hoja de Trabajo 6
                xlWSheet = (Excel.Worksheet)oXL.Worksheets["6.Configuración y Mediciones "];
                xlWSheet.Select(Type.Missing);
                InsertarFila(1);
                InsertarFila(2);

                //Detectar Hoja de Trabajo 8_A
                xlWSheet = (Excel.Worksheet)oXL.Worksheets["8.Rep Fot_NODO 1"];
                xlWSheet.Select(Type.Missing);
                InsertarFila2(1);
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
                MessageBox.Show("Archivo Guardado");
                Process.Start(Lugar_Guardado + NombreExcel);
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
            Formato = ".jpeg";
            string Direccion = @URL_Imagenes + @"\1.Informacion_General\";
            string[] Informacion_General = Directory.GetFiles(Direccion, "*" + Formato);
            int cantidad_Informacion_General = Informacion_General.Length;
            string NombreImg2 = null;
            //
            String[] Codigo = new String[100];
            String[] Numeracion = new String[100];
            String[] strlist = new String[100];
            String[] separador = { @URL_Imagenes + @"\1.Informacion_General\", "NODO_", Formato };
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
                string curFile = Direccion + @"NODO_" + cant_var + Formato;

                if ((cant_var % 2) == 0)
                {
                    //Asignar Rango
                    RangoWidth = (Excel.Range)xlWSheet.get_Range("K27", "O47");

                    //Insertar imagenes
                    if (File.Exists(curFile))
                    {
                        xlWSheet.Shapes.AddPicture(URL_Imagenes + @"\1.Informacion_General\" + "NODO_" + cant_var + Formato,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                    }
                    else
                    {
                    }
                }
                else
                {
                    //Asignar Rango
                    RangoWidth = (Excel.Range)xlWSheet.get_Range("C27", "G47");

                    //Insertar imagenes
                    if (File.Exists(curFile))
                    {
                        xlWSheet.Shapes.AddPicture(URL_Imagenes + @"\1.Informacion_General\" + "NODO_" + cant_var + Formato,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        float.Parse(RangoWidth.Left.ToString()), float.Parse(RangoWidth.Top.ToString()),
                        float.Parse(RangoWidth.Width.ToString()), float.Parse(RangoWidth.Height.ToString()));
                    }
                    else
                    {
                    }
                }
            }
        }
        void InsertarFila(int RangoFila1) 
        {
            
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

            int Rang_colum2 = 50;
            int Rang_row2 = 63;
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
                        Rang_colum2 += aumento;
                        Rang_row2 += aumento;
                        break;
                    case 1:

                        if (RangoFila1 == 1)
                        {
                            RangoWidth = (Excel.Range)xlWSheet.get_Range("C" + Rang_colum2, "I" + Rang_row2);
                        }
                        else if (RangoFila1 == 2)
                        {
                            RangoWidth = (Excel.Range)xlWSheet.get_Range("M" + Rang_colum2, "S" + Rang_row2);
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
                            else
                            {
                                //MessageBox.Show("No existe el documento NODO_" + cant_var + "_" + numeracionciclo);
                            }
                        }
                        Rang_colum2 += aumento;
                        Rang_row2 += aumento;
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("B" + Rang_colum2, "E" + Rang_row2);
                                        }
                                        else if (RangoFila1 == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("L" + Rang_colum2, "O" + Rang_row2);
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("G" + Rang_colum2, "J" + Rang_row2);
                                        }
                                        else if (RangoFila1 == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("Q" + Rang_colum2, "T" + Rang_row2);
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
                            else
                            {
                                //MessageBox.Show("No existe el documento NODO_" + cant_var + "_" + numeracionciclo);
                            }
                        }

                        Rang_colum2 += aumento;
                        Rang_row2 += aumento;
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
            int Rang_colum2 = 11;
            int Rang_row2 = 22;
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
                            Rang_colum2 += aumento;
                            Rang_row2 += aumento;
                        }
                        break;
                    case 0:
                        if (distribucion == 2)
                        {
                            Rang_colum2 += aumento;
                            Rang_row2 += aumento;
                        }
                        break;
                    case 1:
                        if (distribucion == 1)
                        {
                            RangoWidth = (Excel.Range)xlWSheet.get_Range("C" + Rang_colum2, "I" + Rang_row2);
                        }
                        else if (distribucion == 2)
                        {
                            RangoWidth = (Excel.Range)xlWSheet.get_Range("N" + Rang_colum2, "T" + Rang_row2);
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
                            Rang_colum2 += aumento;
                            Rang_row2 += aumento;
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("B" + Rang_colum2, "E" + Rang_row2);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("M" + Rang_colum2, "P" + Rang_row2);
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("G" + Rang_colum2, "J" + Rang_row2);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("R" + Rang_colum2, "U" + Rang_row2);
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
                            else
                            {
                            }
                        }
                        if(distribucion == 2)
                        {
                            Rang_colum2 += aumento;
                            Rang_row2 += aumento;
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("B" + Rang_colum2, "D" + Rang_row2);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("M" + Rang_colum2, "O" + Rang_row2);
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("E" + Rang_colum2, "G" + Rang_row2);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("P" + Rang_colum2, "R" + Rang_row2);
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
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("H" + Rang_colum2, "J" + Rang_row2);
                                        }
                                        else if (distribucion == 2)
                                        {
                                            RangoWidth = (Excel.Range)xlWSheet.get_Range("S" + Rang_colum2, "U" + Rang_row2);
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
                            else
                            {
                            }
                        }
                        if (distribucion == 2)
                        {
                            Rang_colum2 += aumento;
                            Rang_row2 += aumento;
                        }
                        break;
                    default:

                        break;
                }
            }
        }
    }
}
