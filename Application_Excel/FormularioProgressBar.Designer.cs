namespace Application_Excel
{
    partial class FormularioProgressBar
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormularioProgressBar));
            this.ProgressGenerar = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.labelProceso = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ProgressGenerar
            // 
            this.ProgressGenerar.Location = new System.Drawing.Point(12, 42);
            this.ProgressGenerar.Name = "ProgressGenerar";
            this.ProgressGenerar.Size = new System.Drawing.Size(453, 45);
            this.ProgressGenerar.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(113, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(190, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Implementacion :";
            // 
            // labelProceso
            // 
            this.labelProceso.AutoSize = true;
            this.labelProceso.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold);
            this.labelProceso.Location = new System.Drawing.Point(309, 9);
            this.labelProceso.Name = "labelProceso";
            this.labelProceso.Size = new System.Drawing.Size(0, 25);
            this.labelProceso.TabIndex = 2;
            // 
            // FormularioProgressBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 99);
            this.ControlBox = false;
            this.Controls.Add(this.labelProceso);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ProgressGenerar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormularioProgressBar";
            this.Text = "ProgressBar";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar ProgressGenerar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelProceso;
    }
}