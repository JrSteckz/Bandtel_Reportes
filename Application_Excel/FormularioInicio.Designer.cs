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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormularioInicio));
            this.btnPTP = new System.Windows.Forms.Button();
            this.btnPMP = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnHuawei = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.SuspendLayout();
            // 
            // btnPTP
            // 
            this.btnPTP.Location = new System.Drawing.Point(0, 0);
            this.btnPTP.Name = "btnPTP";
            this.btnPTP.Size = new System.Drawing.Size(75, 23);
            this.btnPTP.TabIndex = 0;
            // 
            // btnPMP
            // 
            this.btnPMP.BackColor = System.Drawing.Color.White;
            this.btnPMP.Font = new System.Drawing.Font("Microsoft Uighur", 24F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPMP.Location = new System.Drawing.Point(244, 126);
            this.btnPMP.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnPMP.Name = "btnPMP";
            this.btnPMP.Size = new System.Drawing.Size(224, 77);
            this.btnPMP.TabIndex = 1;
            this.btnPMP.Text = "PMP";
            this.btnPMP.UseVisualStyleBackColor = false;
            this.btnPMP.Click += new System.EventHandler(this.btnPMP_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.Font = new System.Drawing.Font("Microsoft Uighur", 24F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(12, 126);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(224, 77);
            this.button1.TabIndex = 4;
            this.button1.Text = "PTP";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.btnPTP_Click);
            // 
            // btnHuawei
            // 
            this.btnHuawei.BackColor = System.Drawing.Color.White;
            this.btnHuawei.Font = new System.Drawing.Font("Microsoft Uighur", 24F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnHuawei.Location = new System.Drawing.Point(12, 209);
            this.btnHuawei.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnHuawei.Name = "btnHuawei";
            this.btnHuawei.Size = new System.Drawing.Size(456, 77);
            this.btnHuawei.TabIndex = 5;
            this.btnHuawei.Text = "PACKING LIST";
            this.btnHuawei.UseVisualStyleBackColor = false;
            this.btnHuawei.Click += new System.EventHandler(this.btnHuawei_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Uighur", 24F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(139, 63);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(218, 38);
            this.label1.TabIndex = 6;
            this.label1.Text = "APLICACIONES RF";
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(12, 12);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(106, 89);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox3.TabIndex = 17;
            this.pictureBox3.TabStop = false;
            // 
            // FormularioInicio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.CornflowerBlue;
            this.ClientSize = new System.Drawing.Size(484, 309);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnHuawei);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnPMP);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormularioInicio";
            this.Text = "FORMULARIO INICIAL";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnPTP;
        private System.Windows.Forms.Button btnPMP;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnHuawei;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox3;
    }
}