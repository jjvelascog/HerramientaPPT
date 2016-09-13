namespace Herramientas
{
    partial class GeekEncontrado
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.imgGeek = new System.Windows.Forms.PictureBox();
            this.lblDescripcion = new System.Windows.Forms.Label();
            this.lblEncontrados = new System.Windows.Forms.Label();
            this.lblNombre = new System.Windows.Forms.Label();
            this.lblMotivo = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.imgGeek)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(46, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(147, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "¡Felicidades!";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(212, 54);
            this.label2.TabIndex = 1;
            this.label2.Text = "¡Encontraste a un\r\nnuevo Geek oculto!";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // imgGeek
            // 
            this.imgGeek.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.imgGeek.Image = global::Herramientas.Properties.Resources.integ3;
            this.imgGeek.Location = new System.Drawing.Point(62, 143);
            this.imgGeek.Name = "imgGeek";
            this.imgGeek.Size = new System.Drawing.Size(112, 117);
            this.imgGeek.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.imgGeek.TabIndex = 2;
            this.imgGeek.TabStop = false;
            // 
            // lblDescripcion
            // 
            this.lblDescripcion.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDescripcion.Location = new System.Drawing.Point(46, 272);
            this.lblDescripcion.Name = "lblDescripcion";
            this.lblDescripcion.Size = new System.Drawing.Size(147, 114);
            this.lblDescripcion.TabIndex = 3;
            this.lblDescripcion.Text = "Ricardo es miembro fundador del equipo Extra Geek. Lideró el desarrollo de la Bar" +
    "ra Matrix y el curso de Macros.";
            this.lblDescripcion.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // lblEncontrados
            // 
            this.lblEncontrados.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEncontrados.Location = new System.Drawing.Point(43, 457);
            this.lblEncontrados.Name = "lblEncontrados";
            this.lblEncontrados.Size = new System.Drawing.Size(150, 35);
            this.lblEncontrados.TabIndex = 4;
            this.lblEncontrados.Text = "¡Ya has encontrado 3 de los 10 Geeks ocultos!";
            this.lblEncontrados.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // lblNombre
            // 
            this.lblNombre.Font = new System.Drawing.Font("Verdana", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNombre.Location = new System.Drawing.Point(43, 117);
            this.lblNombre.Name = "lblNombre";
            this.lblNombre.Size = new System.Drawing.Size(150, 23);
            this.lblNombre.TabIndex = 5;
            this.lblNombre.Text = "Juan Domingo Pau";
            this.lblNombre.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // lblMotivo
            // 
            this.lblMotivo.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMotivo.Location = new System.Drawing.Point(43, 386);
            this.lblMotivo.Name = "lblMotivo";
            this.lblMotivo.Size = new System.Drawing.Size(150, 60);
            this.lblMotivo.TabIndex = 6;
            this.lblMotivo.Text = "Encontraste a este Geek al escribir Extra Geek en un disclaimer.";
            this.lblMotivo.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // GeekEncontrado
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(236, 501);
            this.Controls.Add(this.lblMotivo);
            this.Controls.Add(this.lblNombre);
            this.Controls.Add(this.lblEncontrados);
            this.Controls.Add(this.lblDescripcion);
            this.Controls.Add(this.imgGeek);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "GeekEncontrado";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Encontraste un Geek";
            ((System.ComponentModel.ISupportInitialize)(this.imgGeek)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public void inicializar(string nombreGeek, string imagen, string descripcion, string motivo, int encontrados, int total)
        {
            lblNombre.Text = nombreGeek;
            lblDescripcion.Text = descripcion;
            lblEncontrados.Text = "¡Ya has encontrado " + (encontrados+1) + " de los " + total + " Geeks ocultos!";
            lblMotivo.Text = "Encontraste a este Geek al " + motivo + ".";
            imgGeek.Image = (System.Drawing.Image) global::Herramientas.Properties.Resources.ResourceManager.GetObject(imagen);
        }


        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox imgGeek;
        private System.Windows.Forms.Label lblDescripcion;
        private System.Windows.Forms.Label lblEncontrados;
        private System.Windows.Forms.Label lblNombre;
        private System.Windows.Forms.Label lblMotivo;
    }
}