namespace Application_Excel
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.Tab_Principal = new System.Windows.Forms.TabControl();
            this.tab_Caratula = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtImagenes = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.txtUbicacionPlantilla = new System.Windows.Forms.TextBox();
            this.BtnBuscador = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_UbicacionPlantilla = new System.Windows.Forms.Button();
            this.txtURL = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.txtNombreExcel = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Tab_Principal.SuspendLayout();
            this.tab_Caratula.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // Tab_Principal
            // 
            this.Tab_Principal.Controls.Add(this.tab_Caratula);
            this.Tab_Principal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Tab_Principal.Location = new System.Drawing.Point(27, 100);
            this.Tab_Principal.Name = "Tab_Principal";
            this.Tab_Principal.SelectedIndex = 0;
            this.Tab_Principal.Size = new System.Drawing.Size(861, 229);
            this.Tab_Principal.TabIndex = 1;
            // 
            // tab_Caratula
            // 
            this.tab_Caratula.BackColor = System.Drawing.Color.LightGray;
            this.tab_Caratula.Controls.Add(this.groupBox1);
            this.tab_Caratula.Location = new System.Drawing.Point(4, 24);
            this.tab_Caratula.Name = "tab_Caratula";
            this.tab_Caratula.Padding = new System.Windows.Forms.Padding(3);
            this.tab_Caratula.Size = new System.Drawing.Size(853, 201);
            this.tab_Caratula.TabIndex = 0;
            this.tab_Caratula.Text = "Caratulá";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtNombreExcel);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtImagenes);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtUbicacionPlantilla);
            this.groupBox1.Controls.Add(this.BtnBuscador);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.btn_UbicacionPlantilla);
            this.groupBox1.Controls.Add(this.txtURL);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(22, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(810, 178);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Inicio";
            // 
            // txtImagenes
            // 
            this.txtImagenes.Location = new System.Drawing.Point(156, 105);
            this.txtImagenes.Name = "txtImagenes";
            this.txtImagenes.Size = new System.Drawing.Size(532, 21);
            this.txtImagenes.TabIndex = 7;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(694, 105);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 24);
            this.button1.TabIndex = 8;
            this.button1.Text = "Abrir";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.BtnBuscadorImagenes_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 107);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(143, 15);
            this.label6.TabIndex = 9;
            this.label6.Text = "Ubicación de Imagenes :";
            // 
            // txtUbicacionPlantilla
            // 
            this.txtUbicacionPlantilla.Location = new System.Drawing.Point(156, 69);
            this.txtUbicacionPlantilla.Name = "txtUbicacionPlantilla";
            this.txtUbicacionPlantilla.Size = new System.Drawing.Size(532, 21);
            this.txtUbicacionPlantilla.TabIndex = 4;
            // 
            // BtnBuscador
            // 
            this.BtnBuscador.Location = new System.Drawing.Point(694, 31);
            this.BtnBuscador.Name = "BtnBuscador";
            this.BtnBuscador.Size = new System.Drawing.Size(100, 24);
            this.BtnBuscador.TabIndex = 1;
            this.BtnBuscador.Text = "Abrir";
            this.BtnBuscador.UseVisualStyleBackColor = true;
            this.BtnBuscador.Click += new System.EventHandler(this.BtnBuscadorGuardado_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 71);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(132, 15);
            this.label5.TabIndex = 6;
            this.label5.Text = "Ubicación de Plantilla :";
            // 
            // btn_UbicacionPlantilla
            // 
            this.btn_UbicacionPlantilla.Location = new System.Drawing.Point(694, 69);
            this.btn_UbicacionPlantilla.Name = "btn_UbicacionPlantilla";
            this.btn_UbicacionPlantilla.Size = new System.Drawing.Size(100, 24);
            this.btn_UbicacionPlantilla.TabIndex = 5;
            this.btn_UbicacionPlantilla.Text = "Abrir";
            this.btn_UbicacionPlantilla.UseVisualStyleBackColor = true;
            this.btn_UbicacionPlantilla.Click += new System.EventHandler(this.BtnBuscadorPlantilla_Click);
            // 
            // txtURL
            // 
            this.txtURL.Location = new System.Drawing.Point(156, 31);
            this.txtURL.Name = "txtURL";
            this.txtURL.Size = new System.Drawing.Size(532, 21);
            this.txtURL.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(140, 15);
            this.label3.TabIndex = 3;
            this.label3.Text = "Ubicación de Guardado ";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(13, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(288, 81);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(328, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(309, 28);
            this.label1.TabIndex = 4;
            this.label1.Text = "REPORTE DE INSTALACION";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
            this.label2.Location = new System.Drawing.Point(314, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(343, 28);
            this.label2.TabIndex = 5;
            this.label2.Text = "RADIO ENLACE MICROONDAS";
            // 
            // btnGenerar
            // 
            this.btnGenerar.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGenerar.Location = new System.Drawing.Point(27, 335);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(861, 47);
            this.btnGenerar.TabIndex = 6;
            this.btnGenerar.Text = "Generar";
            this.btnGenerar.UseVisualStyleBackColor = true;
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // txtNombreExcel
            // 
            this.txtNombreExcel.Location = new System.Drawing.Point(156, 140);
            this.txtNombreExcel.Name = "txtNombreExcel";
            this.txtNombreExcel.Size = new System.Drawing.Size(532, 21);
            this.txtNombreExcel.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(33, 143);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(117, 15);
            this.label4.TabIndex = 11;
            this.label4.Text = "Nombre de Archivo :";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(913, 387);
            this.Controls.Add(this.btnGenerar);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.Tab_Principal);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "REPORTE DE INSTALACION";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Tab_Principal.ResumeLayout(false);
            this.tab_Caratula.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TabControl Tab_Principal;
        private System.Windows.Forms.TabPage tab_Caratula;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button BtnBuscador;
        private System.Windows.Forms.TextBox txtURL;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtUbicacionPlantilla;
        private System.Windows.Forms.Button btn_UbicacionPlantilla;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtImagenes;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtNombreExcel;
        private System.Windows.Forms.Label label4;
    }
}

