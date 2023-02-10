namespace Reportes
{
    partial class FormularioHuawei
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormularioHuawei));
            this.button1 = new System.Windows.Forms.Button();
            this.txtIdNodo = new System.Windows.Forms.TextBox();
            this.txtGuardado = new System.Windows.Forms.TextBox();
            this.btnGuardado = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtNombre = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnatos = new System.Windows.Forms.Button();
            this.txtDatos = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.button1.Location = new System.Drawing.Point(28, 308);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(740, 58);
            this.button1.TabIndex = 0;
            this.button1.Text = "Generar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtIdNodo
            // 
            this.txtIdNodo.Location = new System.Drawing.Point(131, 29);
            this.txtIdNodo.Name = "txtIdNodo";
            this.txtIdNodo.Size = new System.Drawing.Size(479, 20);
            this.txtIdNodo.TabIndex = 1;
            // 
            // txtGuardado
            // 
            this.txtGuardado.Location = new System.Drawing.Point(131, 104);
            this.txtGuardado.Name = "txtGuardado";
            this.txtGuardado.Size = new System.Drawing.Size(479, 20);
            this.txtGuardado.TabIndex = 2;
            // 
            // btnGuardado
            // 
            this.btnGuardado.Location = new System.Drawing.Point(623, 102);
            this.btnGuardado.Name = "btnGuardado";
            this.btnGuardado.Size = new System.Drawing.Size(76, 26);
            this.btnGuardado.TabIndex = 3;
            this.btnGuardado.Text = "Abrir";
            this.btnGuardado.UseVisualStyleBackColor = true;
            this.btnGuardado.Click += new System.EventHandler(this.btnGuardado_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(273, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(310, 28);
            this.label1.TabIndex = 11;
            this.label1.Text = "GESTION DE PACKING LIST";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(28, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(223, 85);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 10;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Gainsboro;
            this.groupBox1.Controls.Add(this.txtNombre);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnatos);
            this.groupBox1.Controls.Add(this.txtDatos);
            this.groupBox1.Controls.Add(this.txtIdNodo);
            this.groupBox1.Controls.Add(this.txtGuardado);
            this.groupBox1.Controls.Add(this.btnGuardado);
            this.groupBox1.Location = new System.Drawing.Point(28, 114);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(740, 185);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Formulario";
            // 
            // txtNombre
            // 
            this.txtNombre.Location = new System.Drawing.Point(131, 137);
            this.txtNombre.Name = "txtNombre";
            this.txtNombre.Size = new System.Drawing.Size(479, 20);
            this.txtNombre.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(36, 140);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Nombre del excel";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 107);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(99, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Lugar de Guardado";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(90, 72);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Datos";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(78, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "ID Nodo";
            // 
            // btnatos
            // 
            this.btnatos.Location = new System.Drawing.Point(624, 69);
            this.btnatos.Name = "btnatos";
            this.btnatos.Size = new System.Drawing.Size(75, 23);
            this.btnatos.TabIndex = 5;
            this.btnatos.Text = "Abrir";
            this.btnatos.UseVisualStyleBackColor = true;
            this.btnatos.Click += new System.EventHandler(this.btnatos_Click);
            // 
            // txtDatos
            // 
            this.txtDatos.Location = new System.Drawing.Point(131, 69);
            this.txtDatos.Name = "txtDatos";
            this.txtDatos.Size = new System.Drawing.Size(479, 20);
            this.txtDatos.TabIndex = 4;
            // 
            // FormularioHuawei
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(800, 378);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pictureBox1);
            this.Name = "FormularioHuawei";
            this.Text = "FormularioHuawei";
            this.Load += new System.EventHandler(this.FormularioHuawei_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtIdNodo;
        private System.Windows.Forms.TextBox txtGuardado;
        private System.Windows.Forms.Button btnGuardado;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtDatos;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnatos;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtNombre;
        private System.Windows.Forms.Label label5;
    }
}