using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Reportes
{
    public partial class Mapa : Form
    {
        public Mapa()
        {
            InitializeComponent();
        }
        //string rutaImagen = "F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\5.Pruebas de Interferencia\\NODO_1.png";
        //libro.SaveAs("C:\\Users\\Mcarrera\\source\\repos\\hola.xlsx");

        private void button1_Click(object sender, EventArgs e)
        {

            //// Crear una instancia de Application
            //Excel.Application excel = new Excel.Application();

            //// Crear un nuevo libro de Excel
            //Excel.Workbook libro = excel.Workbooks.Add();

            //// Obtener la hoja de trabajo activa
            //Excel.Worksheet hoja = libro.ActiveSheet;

            //// Seleccionar el rango de celdas
            //Excel.Range rango = hoja.Range["A1:F9"];

            //// Combinar las celdas
            //rango.Merge();

            //// Obtener las rutas de las imágenes
            //string[] rutasImagenes = new string[]
            //{
            //"F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\8.Reporte_Fotografico\\8.Reporte_Fotografico_S4\\A8_7_1.jpeg",
            //"F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\8.Reporte_Fotografico\\8.Reporte_Fotografico_S4\\A8_9_1.jpeg",
            //"F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\8.Reporte_Fotografico\\8.Reporte_Fotografico_S4\\A8_20_1.jpeg"
            //};

            //// Insertar las imágenes en la celda combinada
            //Excel.Range celdaCombinada = hoja.Range["A1:F9"];
            //celdaCombinada.Select();
            //Excel.Pictures imagenes = hoja.Pictures(System.Reflection.Missing.Value) as Excel.Pictures;

            //float anchoMaximo = (float)celdaCombinada.Width;
            //float alturaMaxima = (float)celdaCombinada.Height;

            //float anchoImagen, alturaImagen;

            //// Ordenar las imágenes por ancho
            //Array.Sort(rutasImagenes, (ruta1, ruta2) =>
            //{
            //    float ancho1 = (float)System.Drawing.Image.FromFile(ruta1).Width;
            //    float ancho2 = (float)System.Drawing.Image.FromFile(ruta2).Width;

            //    return ancho2.CompareTo(ancho1);
            //});

            //// Calcular el número de imágenes a agregar
            //int numImagenes = Math.Min(rutasImagenes.Length, 3);

            //// Calcular la cantidad de imágenes por fila y columna
            //int numColumnas = (int)Math.Ceiling(Math.Sqrt(numImagenes));
            //int numFilas = (int)Math.Ceiling((float)numImagenes / numColumnas);

            //// Calcular el tamaño de cada imagen
            //float anchoImagenes = anchoMaximo / numColumnas;
            //float alturaImagenes = alturaMaxima / numFilas;

            //// Distribuir las imágenes en el rango
            //for (int i = 0; i < numImagenes; i++)
            //{
            //    // Calcular la posición de la imagen en la matriz
            //    int fila = i / numColumnas;
            //    int columna = i % numColumnas;

            //    // Calcular la posición absoluta de la imagen
            //    float posicionTop = (float)celdaCombinada.Top + alturaImagenes * fila;
            //    float posicionLeft = (float)celdaCombinada.Left + anchoImagenes * columna;

            //    // Cargar la imagen desde el archivo
            //    string rutaImagen = rutasImagenes[i];
            //    Excel.Picture imagen = imagenes.Insert(rutaImagen, System.Reflection.Missing.Value);
            //    anchoImagen = (float)imagen.Width;
            //    alturaImagen = (float)imagen.Height;

            //    // Calcular el tamaño de la imagen para que quepa en la celda
            //    if (anchoImagen > anchoImagenes)
            //    {
            //        alturaImagen *= anchoImagenes / anchoImagen;
            //        anchoImagen = anchoImagenes;
            //    }

            //    if (alturaImagen > alturaImagenes)
            //    {
            //        anchoImagen *= alturaImagenes / alturaImagen;
            //        alturaImagen = alturaImagenes;
            //    }

            //    // Ajustar la posición y tamaño de la imagen
            //    imagen.Left = posicionLeft + (anchoImagenes - anchoImagen) / 2;
            //    imagen.Top = posicionTop + (alturaImagenes - alturaImagen) / 2;
            //    imagen.Width = anchoImagen;
            //    imagen.Height = alturaImagen;
            //}

            //// Guardar el libro de Excel y cerrar la aplicación
            //string nombreArchivo = "C:\\Users\\Mcarrera\\source\\repos\\hola.xlsx";
            //libro.SaveAs(nombreArchivo);
            //libro.Close();
            //excel.Quit();






            // Crear una instancia de Application
            Excel.Application excel = new Excel.Application();

            // Crear un nuevo libro de Excel
            Excel.Workbook libro = excel.Workbooks.Add();

            // Obtener la hoja de trabajo activa
            Excel.Worksheet hoja = libro.ActiveSheet;

            // Seleccionar el rango de celdas
            Excel.Range rango = hoja.Range["A1:I9"];

            // Combinar las celdas
            rango.Merge();

            // Obtener las rutas de las imágenes
            string[] rutasImagenes = new string[]
            {
            "F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\8.Reporte_Fotografico\\8.Reporte_Fotografico_S4\\A8_7_1.jpeg",
            "F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\8.Reporte_Fotografico\\8.Reporte_Fotografico_S4\\A8_9_1.jpeg",
            "F:\\Escritorio 20.02.2023\\Entorno de Prueba\\PMP\\Reportes de Instalación\\8.Reporte_Fotografico\\8.Reporte_Fotografico_S4\\A8_20_1.jpeg"
            };

            // Insertar las imágenes en la celda combinada
            Excel.Range celdaCombinada = hoja.Range["A1"];
            celdaCombinada.Select();
            Excel.Pictures imagenes = hoja.Pictures(System.Reflection.Missing.Value) as Excel.Pictures;

            float anchoMaximo = (float)celdaCombinada.Width;
            float alturaMaxima = (float)celdaCombinada.Height;

            float anchoImagen, alturaImagen;

            // Ordenar las imágenes por ancho
            Array.Sort(rutasImagenes, (ruta1, ruta2) =>
            {
                float ancho1 = (float)System.Drawing.Image.FromFile(ruta1).Width;
                float ancho2 = (float)System.Drawing.Image.FromFile(ruta2).Width;

                return ancho2.CompareTo(ancho1);
            });

            // Calcular el número de imágenes a agregar
            int numImagenes = Math.Min(rutasImagenes.Length, 3);

            // Calcular la cantidad de imágenes por fila y columna
            int numColumnas = (int)Math.Ceiling(Math.Sqrt(numImagenes));
            int numFilas = (int)Math.Ceiling((float)numImagenes / numColumnas);

            // Calcular el tamaño de cada imagen
            float anchoImagenes = anchoMaximo / numColumnas;
            float alturaImagenes = alturaMaxima / numFilas;

            // Distribuir las imágenes en el rango
            for (int i = 0; i < numImagenes; i++)
            {
                // Calcular la posición de la imagen en la matriz
                int fila = i / numColumnas;
                int columna = i % numColumnas;

                // Calcular la posición absoluta de la imagen
                float posicionTop = (float)celdaCombinada.Top + alturaImagenes * fila;
                float posicionLeft = (float)celdaCombinada.Left + anchoImagenes * columna;

                // Cargar la imagen desde el archivo
                string rutaImagen = rutasImagenes[i];
                Excel.Picture imagen = imagenes.Insert(rutaImagen, System.Reflection.Missing.Value);
                anchoImagen = (float)imagen.Width;
                alturaImagen = (float)imagen.Height;

                // Calcular el tamaño de la imagen para que quepa en la celda
                if (anchoImagen > anchoImagenes)
                {
                    alturaImagen *= anchoImagenes / anchoImagen;
                    anchoImagen = anchoImagenes;
                }

                if (alturaImagen > alturaImagenes)
                {
                    anchoImagen *= alturaImagenes / alturaImagen;
                    alturaImagen = alturaImagenes;
                }

                // Posicionar la imagen en la celda
                imagen.Left = (float)posicionLeft + ((float)anchoImagenes - anchoImagen) / 2;
                imagen.Top = (float)posicionTop + ((float)alturaImagenes - alturaImagen) / 2;
                imagen.Width = anchoImagen;
                imagen.Height = alturaImagen;
            }

            // Guardar el libro de Excel
            string nombreArchivo = "C:\\Users\\Mcarrera\\source\\repos\\hola.xlsx";

            libro.SaveAs(nombreArchivo);

            // Cerrar el libro de Excel
            libro.Close();

            // Cerrar la aplicación de Excel
            excel.Quit();
            Console.WriteLine("Archivo guardado exitosamente en " + Path.GetFullPath(nombreArchivo));
        }
    }

}
