using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Herramientas
{
    public partial class SeleccionarSimilares : Form
    {

        Clases.metodos csmet = Clases.metodos.Instance;
        private barra_matrix barraMatrix;

        public SeleccionarSimilares()
        {
            InitializeComponent();

            checkBox1.Checked = csmet.CONFIG["SeleccionarSimilares_tamanio"].Equals("YES") ? true : false;
            checkBox2.Checked = csmet.CONFIG["SeleccionarSimilares_fondo"].Equals("YES") ? true : false;
            checkBox3.Checked = csmet.CONFIG["SeleccionarSimilares_linea"].Equals("YES") ? true : false;
            checkBox4.Checked = csmet.CONFIG["SeleccionarSimilares_tipo"].Equals("YES") ? true : false;
        }
        
        public void inicializar(barra_matrix newBarraMatrix)
        {
            barraMatrix = newBarraMatrix;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            barraMatrix.tamanio = checkBox1.Checked;
            barraMatrix.colorFondo = checkBox2.Checked;
            barraMatrix.colorLinea = checkBox3.Checked;
            barraMatrix.tipo = checkBox4.Checked;
            barraMatrix.seleccionarSimilares2();

            if (checkBox5.Checked)
            {

                csmet.CONFIG.Remove("SeleccionarSimilares_tamanio");
                csmet.CONFIG.Add("SeleccionarSimilares_tamanio", (checkBox1.Checked ? "YES" : "NO"));

                csmet.CONFIG.Remove("SeleccionarSimilares_fondo");
                csmet.CONFIG.Add("SeleccionarSimilares_fondo", (checkBox2.Checked ? "YES" : "NO"));

                csmet.CONFIG.Remove("SeleccionarSimilares_linea");
                csmet.CONFIG.Add("SeleccionarSimilares_linea", (checkBox3.Checked ? "YES" : "NO"));

                csmet.CONFIG.Remove("SeleccionarSimilares_tipo");
                csmet.CONFIG.Add("SeleccionarSimilares_tipo", (checkBox4.Checked ? "YES" : "NO"));

                try
                {
                    csmet.guardarConfig();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hubo  un error al intentar guardar la configuración, por favor intente más tarde. Si el problema persiste contacte a su administrador.", "Barra MATRIX", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }



            Close();
        }
    }
}
