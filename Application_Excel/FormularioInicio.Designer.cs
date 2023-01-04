namespace Reportes
{
    partial class FormularioInicio
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
            this.btnPTP = new System.Windows.Forms.Button();
            this.btnPMP = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnPTP
            // 
            this.btnPTP.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPTP.Location = new System.Drawing.Point(12, 12);
            this.btnPTP.Name = "btnPTP";
            this.btnPTP.Size = new System.Drawing.Size(167, 140);
            this.btnPTP.TabIndex = 0;
            this.btnPTP.Text = "PTP";
            this.btnPTP.UseVisualStyleBackColor = true;
            this.btnPTP.Click += new System.EventHandler(this.btnPTP_Click);
            // 
            // btnPMP
            // 
            this.btnPMP.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.btnPMP.Location = new System.Drawing.Point(197, 12);
            this.btnPMP.Name = "btnPMP";
            this.btnPMP.Size = new System.Drawing.Size(167, 140);
            this.btnPMP.TabIndex = 1;
            this.btnPMP.Text = "PMP";
            this.btnPMP.UseVisualStyleBackColor = true;
            this.btnPMP.Click += new System.EventHandler(this.btnPMP_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.button3.Location = new System.Drawing.Point(381, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(167, 140);
            this.button3.TabIndex = 2;
            this.button3.Text = "?";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCerrar.Location = new System.Drawing.Point(12, 165);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(536, 51);
            this.btnCerrar.TabIndex = 3;
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // FormularioInicio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 228);
            this.ControlBox = false;
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btnPMP);
            this.Controls.Add(this.btnPTP);
            this.Name = "FormularioInicio";
            this.Text = "INICIO";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnPTP;
        private System.Windows.Forms.Button btnPMP;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btnCerrar;
    }
}