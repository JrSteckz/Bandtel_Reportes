using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Application_Excel
{
    public partial class FormularioProgressBar : Form
    {
        public FormularioProgressBar()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
        public void Instalacion(int numeracion)
        {
            switch (numeracion)
            {
                case 1:
                    ProgressGenerar.Value = 10;
                    labelProceso.Text = "10%";
                    break;
                case 2:
                    ProgressGenerar.Value = 20;
                    labelProceso.Text = "20%";
                    break;
                case 3:
                    ProgressGenerar.Value = 30;
                    labelProceso.Text = "30%";
                    break;
                case 4:
                    ProgressGenerar.Value = 40;
                    labelProceso.Text = "40%";
                    break;
                case 5:
                    ProgressGenerar.Value = 50;
                    labelProceso.Text = "50%";
                    break;
                case 6:
                    ProgressGenerar.Value = 60;
                    labelProceso.Text = "60%";
                    break;
                case 7:
                    ProgressGenerar.Value = 70;
                    labelProceso.Text = "70%";
                    break;
                case 8:
                    ProgressGenerar.Value = 80;
                    labelProceso.Text = "80%";
                    break;
                case 9:
                    ProgressGenerar.Value = 90;
                    labelProceso.Text = "90%";
                    break;
                case 10:
                    ProgressGenerar.Value = 100;
                    labelProceso.Text = "100%";
                    break;
                default:
                    break;
            }
        }
    }
}
