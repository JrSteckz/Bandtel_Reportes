using Application_Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reportes
{
    public partial class FormularioInicio : Form
    {
        public FormularioInicio()
        {
            InitializeComponent();
        }
    

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            //Code to trigger when the "Yes"-button is pressed.
            FormularioInicio settings = new FormularioInicio();
            this.Close();
            settings.Close();
        }

        private void btnPMP_Click(object sender, EventArgs e)
        {
            FormularioPMP setting = new FormularioPMP();
            setting.ShowDialog();
        }

        private void btnPTP_Click(object sender, EventArgs e)
        {
            FormularioPrincipal setting = new FormularioPrincipal();
            setting.ShowDialog();
        }
    }
}
