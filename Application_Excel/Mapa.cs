using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NPOI;
using Newtonsoft.Json;

namespace Reportes
{
    public partial class Mapa : Form
    {
        public Mapa()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var json = File.ReadAllText("coordinates.json");
            var coordinates = JsonConvert.DeserializeObject<LatLng[]>(json);

            var map = new GoogleMaps();
            map.Zoom = 8;
            map.Center = new LatLng(-9.02817, -75.72586);

            var polygon = new Polygon();
            polygon.Paths = coordinates;
            polygon.StrokeColor = Color.Red;
            polygon.StrokeOpacity = 0.8;
            polygon.StrokeWeight = 2;
            polygon.FillColor = Color.Red;

            map.Polygons.Add(polygon);
        }
    }
}
