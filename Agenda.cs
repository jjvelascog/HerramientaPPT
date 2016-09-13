using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

namespace Herramientas
{
    public partial class Agenda : Form
    {
        private Slide agendaSlide;
        private barra_matrix barraMatrix;
        public AgendaClase clase;

        public Agenda()
        {
            InitializeComponent();
        }

        public void inicializar(barra_matrix newBarraMatrix, Slide newAgenda, string titulo, AgendaClase agenda)
        {
            barraMatrix = newBarraMatrix;
            agendaSlide = newAgenda;
            clase = agenda;
            textBox1.Text = titulo;
        }

        public void agregarItem()
        {
            if (listaItems.Items.Count + 1 <= 12)
                listaItems.Items.Add(clase.LABEL_ITEM_AGENDA + (listaItems.Items.Count + 1));
        }

        public void eliminarItem()
        {
            if(listaItems.SelectedIndex == listaItems.Items.Count - 1)
            {
                listaItems.SelectedIndex = listaItems.Items.Count - 2;
            }
            listaItems.Items.RemoveAt(listaItems.Items.Count - 1);

            try { agendaSlide.Shapes[listaItems.SelectedItem].Select(); } catch (Exception e) { }
            
        }

        //Agregar
        private void button1_Click(object sender, EventArgs e)
        {
            if(listaItems.Items.Count + 1 > 12)
                System.Windows.Forms.MessageBox.Show("No se pueden agregar más de 12 elementos");
            else
                clase.agregarItemAgenda(false);
        }

        //Eliminar
        private void button2_Click(object sender, EventArgs e)
        {
            if (listaItems.SelectedIndex == -1)
                MessageBox.Show("Debe seleccionar un ítem de la lista");
            else if(listaItems.Items.Count == 1)
                MessageBox.Show("No puede eliminar todos los items de la lista");
            else
                clase.eliminarItemAgenda(listaItems.SelectedIndex + 1,true);
        }

        //Actualizar
        private void button3_Click(object sender, EventArgs e)
        {
            clase.actualizarAgenda();
            this.Close();
        }

        private void selectionChanged(object sender, EventArgs e)
        {
            try
            {
                agendaSlide.Shapes[listaItems.SelectedItem].Select();
            }
            catch (Exception except)
            {

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            clase.actualizarTituloAgenda(textBox1.Text);
        }
    }
}
