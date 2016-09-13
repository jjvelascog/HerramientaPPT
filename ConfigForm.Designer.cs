namespace Herramientas
{
    partial class ConfigForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkBlltsAjustarEspaciado = new System.Windows.Forms.CheckBox();
            this.chkBlltsAjustarSangria = new System.Windows.Forms.CheckBox();
            this.chkBlltsAjustarTamanho = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkTakeawayAnimate = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.txtGapV = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtGapH = new System.Windows.Forms.NumericUpDown();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chkSelPreguntar = new System.Windows.Forms.CheckBox();
            this.chkSelTipo = new System.Windows.Forms.CheckBox();
            this.chkSelLinea = new System.Windows.Forms.CheckBox();
            this.chkSelFondo = new System.Windows.Forms.CheckBox();
            this.chkSelTamanio = new System.Windows.Forms.CheckBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.cmbFonts = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtGapV)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtGapH)).BeginInit();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkBlltsAjustarEspaciado);
            this.groupBox1.Controls.Add(this.chkBlltsAjustarSangria);
            this.groupBox1.Controls.Add(this.chkBlltsAjustarTamanho);
            this.groupBox1.Location = new System.Drawing.Point(25, 22);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(357, 97);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Dar formato de bullets";
            // 
            // chkBlltsAjustarEspaciado
            // 
            this.chkBlltsAjustarEspaciado.AutoSize = true;
            this.chkBlltsAjustarEspaciado.Location = new System.Drawing.Point(19, 65);
            this.chkBlltsAjustarEspaciado.Name = "chkBlltsAjustarEspaciado";
            this.chkBlltsAjustarEspaciado.Size = new System.Drawing.Size(166, 17);
            this.chkBlltsAjustarEspaciado.TabIndex = 2;
            this.chkBlltsAjustarEspaciado.Text = "Ajustar espacio entre párrafos";
            this.chkBlltsAjustarEspaciado.UseVisualStyleBackColor = true;
            // 
            // chkBlltsAjustarSangria
            // 
            this.chkBlltsAjustarSangria.AutoSize = true;
            this.chkBlltsAjustarSangria.Location = new System.Drawing.Point(19, 42);
            this.chkBlltsAjustarSangria.Name = "chkBlltsAjustarSangria";
            this.chkBlltsAjustarSangria.Size = new System.Drawing.Size(151, 17);
            this.chkBlltsAjustarSangria.TabIndex = 1;
            this.chkBlltsAjustarSangria.Text = "Ajustar sangría por párrafo";
            this.chkBlltsAjustarSangria.UseVisualStyleBackColor = true;
            // 
            // chkBlltsAjustarTamanho
            // 
            this.chkBlltsAjustarTamanho.AutoSize = true;
            this.chkBlltsAjustarTamanho.Location = new System.Drawing.Point(19, 19);
            this.chkBlltsAjustarTamanho.Name = "chkBlltsAjustarTamanho";
            this.chkBlltsAjustarTamanho.Size = new System.Drawing.Size(134, 17);
            this.chkBlltsAjustarTamanho.TabIndex = 0;
            this.chkBlltsAjustarTamanho.Text = "Ajustar tamaño de letra";
            this.chkBlltsAjustarTamanho.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(264, 390);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(118, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Guardar cambios";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkTakeawayAnimate);
            this.groupBox2.Location = new System.Drawing.Point(25, 125);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(357, 61);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Takeaway";
            // 
            // chkTakeawayAnimate
            // 
            this.chkTakeawayAnimate.AutoSize = true;
            this.chkTakeawayAnimate.Location = new System.Drawing.Point(19, 29);
            this.chkTakeawayAnimate.Name = "chkTakeawayAnimate";
            this.chkTakeawayAnimate.Size = new System.Drawing.Size(162, 17);
            this.chkTakeawayAnimate.TabIndex = 0;
            this.chkTakeawayAnimate.Text = "Agregar animación al insertar";
            this.chkTakeawayAnimate.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.txtGapV);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.txtGapH);
            this.groupBox3.Location = new System.Drawing.Point(25, 192);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(357, 47);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Empalmar objetos";
            // 
            // txtGapV
            // 
            this.txtGapV.DecimalPlaces = 2;
            this.txtGapV.Location = new System.Drawing.Point(252, 18);
            this.txtGapV.Name = "txtGapV";
            this.txtGapV.Size = new System.Drawing.Size(43, 20);
            this.txtGapV.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(181, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Gap vertical";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Gap horizontal";
            // 
            // txtGapH
            // 
            this.txtGapH.DecimalPlaces = 2;
            this.txtGapH.Location = new System.Drawing.Point(97, 19);
            this.txtGapH.Name = "txtGapH";
            this.txtGapH.Size = new System.Drawing.Size(42, 20);
            this.txtGapH.TabIndex = 0;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.chkSelPreguntar);
            this.groupBox4.Controls.Add(this.chkSelTipo);
            this.groupBox4.Controls.Add(this.chkSelLinea);
            this.groupBox4.Controls.Add(this.chkSelFondo);
            this.groupBox4.Controls.Add(this.chkSelTamanio);
            this.groupBox4.Location = new System.Drawing.Point(25, 246);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(357, 72);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Seleccionar objetos similares por...";
            // 
            // chkSelPreguntar
            // 
            this.chkSelPreguntar.AutoSize = true;
            this.chkSelPreguntar.Location = new System.Drawing.Point(47, 43);
            this.chkSelPreguntar.Name = "chkSelPreguntar";
            this.chkSelPreguntar.Size = new System.Drawing.Size(265, 17);
            this.chkSelPreguntar.TabIndex = 4;
            this.chkSelPreguntar.Text = "Preguntar siempre antes de seleccionar los objetos";
            this.chkSelPreguntar.UseVisualStyleBackColor = true;
            // 
            // chkSelTipo
            // 
            this.chkSelTipo.AutoSize = true;
            this.chkSelTipo.Location = new System.Drawing.Point(294, 20);
            this.chkSelTipo.Name = "chkSelTipo";
            this.chkSelTipo.Size = new System.Drawing.Size(47, 17);
            this.chkSelTipo.TabIndex = 3;
            this.chkSelTipo.Text = "Tipo";
            this.chkSelTipo.UseVisualStyleBackColor = true;
            // 
            // chkSelLinea
            // 
            this.chkSelLinea.AutoSize = true;
            this.chkSelLinea.Location = new System.Drawing.Point(198, 20);
            this.chkSelLinea.Name = "chkSelLinea";
            this.chkSelLinea.Size = new System.Drawing.Size(90, 17);
            this.chkSelLinea.TabIndex = 2;
            this.chkSelLinea.Text = "Color de linea";
            this.chkSelLinea.UseVisualStyleBackColor = true;
            // 
            // chkSelFondo
            // 
            this.chkSelFondo.AutoSize = true;
            this.chkSelFondo.Location = new System.Drawing.Point(97, 19);
            this.chkSelFondo.Name = "chkSelFondo";
            this.chkSelFondo.Size = new System.Drawing.Size(95, 17);
            this.chkSelFondo.TabIndex = 1;
            this.chkSelFondo.Text = "Color de fondo";
            this.chkSelFondo.UseVisualStyleBackColor = true;
            // 
            // chkSelTamanio
            // 
            this.chkSelTamanio.AutoSize = true;
            this.chkSelTamanio.Location = new System.Drawing.Point(19, 20);
            this.chkSelTamanio.Name = "chkSelTamanio";
            this.chkSelTamanio.Size = new System.Drawing.Size(65, 17);
            this.chkSelTamanio.TabIndex = 0;
            this.chkSelTamanio.Text = "Tamaño";
            this.chkSelTamanio.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.cmbFonts);
            this.groupBox5.Location = new System.Drawing.Point(25, 324);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(357, 47);
            this.groupBox5.TabIndex = 4;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Fuente de texto (se usará en las formas o textos a crear)";
            // 
            // cmbFonts
            // 
            this.cmbFonts.CausesValidation = false;
            this.cmbFonts.FormattingEnabled = true;
            this.cmbFonts.Items.AddRange(new string[] {
            "Verdana",
            "Calibri",
            "Times New Roman",
            "Arial"});
            this.cmbFonts.Location = new System.Drawing.Point(47, 19);
            this.cmbFonts.Name = "cmbFonts";
            this.cmbFonts.Size = new System.Drawing.Size(265, 21);
            this.cmbFonts.TabIndex = 0;
            // 
            // ConfigForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(417, 425);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Name = "ConfigForm";
            this.Text = "Configuración";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtGapV)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtGapH)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkBlltsAjustarTamanho;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox chkBlltsAjustarSangria;
        private System.Windows.Forms.CheckBox chkBlltsAjustarEspaciado;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkTakeawayAnimate;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.NumericUpDown txtGapV;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown txtGapH;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox chkSelPreguntar;
        private System.Windows.Forms.CheckBox chkSelTipo;
        private System.Windows.Forms.CheckBox chkSelLinea;
        private System.Windows.Forms.CheckBox chkSelFondo;
        private System.Windows.Forms.CheckBox chkSelTamanio;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.ComboBox cmbFonts;
    }
}