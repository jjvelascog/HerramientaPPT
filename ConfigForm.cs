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
    public partial class ConfigForm : Form
    {

        Clases.metodos csmet = Clases.metodos.Instance;

        public ConfigForm()
        {
            InitializeComponent();

            chkBlltsAjustarTamanho.Checked = csmet.CONFIG["Bllts_ajustarTam"].Equals("YES")? true :false;
            chkBlltsAjustarSangria.Checked = csmet.CONFIG["Bllts_ajustarSangria"].Equals("YES") ? true : false;
            chkBlltsAjustarEspaciado.Checked = csmet.CONFIG["Bllts_ajustarEspaciado"].Equals("YES") ? true : false;
            chkTakeawayAnimate.Checked = csmet.CONFIG["Takeaway_animate"].Equals("YES") ? true : false;
            txtGapH.Value = Decimal.Parse(csmet.CONFIG["Empalmar_gapH"]);
            txtGapV.Value = Decimal.Parse(csmet.CONFIG["Empalmar_gapV"]);
            chkSelTamanio.Checked = csmet.CONFIG["SeleccionarSimilares_tamanio"].Equals("YES") ? true : false;
            chkSelFondo.Checked = csmet.CONFIG["SeleccionarSimilares_fondo"].Equals("YES") ? true : false;
            chkSelLinea.Checked = csmet.CONFIG["SeleccionarSimilares_linea"].Equals("YES") ? true : false;
            chkSelTipo.Checked = csmet.CONFIG["SeleccionarSimilares_tipo"].Equals("YES") ? true : false;
            chkSelPreguntar.Checked = csmet.CONFIG["SeleccionarSimilares_preguntar"].Equals("YES") ? true : false;
            cmbFonts.Text = csmet.CONFIG["MainFont"];

        }

        private void button1_Click(object sender, EventArgs e)
        {
            csmet.CONFIG.Remove("Bllts_ajustarTam");
            csmet.CONFIG.Add("Bllts_ajustarTam", (chkBlltsAjustarTamanho.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("Bllts_ajustarSangria");
            csmet.CONFIG.Add("Bllts_ajustarSangria", (chkBlltsAjustarSangria.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("Bllts_ajustarEspaciado");
            csmet.CONFIG.Add("Bllts_ajustarEspaciado", (chkBlltsAjustarEspaciado.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("Takeaway_animate");
            csmet.CONFIG.Add("Takeaway_animate", (chkTakeawayAnimate.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("Empalmar_gapH");
            csmet.CONFIG.Add("Empalmar_gapH", "" + txtGapH.Value);

            csmet.CONFIG.Remove("Empalmar_gapV");
            csmet.CONFIG.Add("Empalmar_gapV", "" + txtGapV.Value);

            decimal gap = txtGapV.Value + txtGapH.Value;
            if (gap > 0 && (csmet.CONFIG["Alineacion_integ15"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
            {
                GeekEncontrado ge = new GeekEncontrado();
                ge.inicializar("Jorge Marín", "integ15", "Jorge se unió al Equipo Extra Geek en 2016 tras resolver exitosamente el Recruiting Brainteaser de las 12 bolas y la balanza.", "configurar Gaps mayores que cero", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                ge.ShowDialog();
                csmet.CONFIG.Remove("Alineacion_integ15");
                csmet.CONFIG.Add("Alineacion_integ15", "1");
                csmet.guardarConfig();

            }

            csmet.CONFIG.Remove("SeleccionarSimilares_tamanio");
            csmet.CONFIG.Add("SeleccionarSimilares_tamanio", (chkSelTamanio.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("SeleccionarSimilares_fondo");
            csmet.CONFIG.Add("SeleccionarSimilares_fondo", (chkSelFondo.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("SeleccionarSimilares_linea");
            csmet.CONFIG.Add("SeleccionarSimilares_linea", (chkSelLinea.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("SeleccionarSimilares_tipo");
            csmet.CONFIG.Add("SeleccionarSimilares_tipo", (chkSelTipo.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("SeleccionarSimilares_preguntar");
            csmet.CONFIG.Add("SeleccionarSimilares_preguntar", (chkSelPreguntar.Checked ? "YES" : "NO"));

            csmet.CONFIG.Remove("MainFont");
            csmet.CONFIG.Add("MainFont", (string)cmbFonts.Text);

            try {
                csmet.guardarConfig();

                MessageBox.Show("Configuración guardada", "Barra MATRIX", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hubo  un error al intentar guardar la configuración, por favor intente más tarde. Si el problema persiste contacte a su administrador.", "Barra MATRIX", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
    }
}
