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
    public partial class SeleccionAgenda : Form
    {
        private List<AgendaClase> agendas;
        public string seleccionado;

        public SeleccionAgenda(List<AgendaClase> listaForms)
        {
            InitializeComponent();
            agendas = listaForms;
            for(int i = 0; i < agendas.Count; i++)
            {
                switch(i)
                {
                    case 0:
                        button1.Text = agendas[i].titulo;
                        break;
                    case 1:
                        button2.Text = agendas[i].titulo;
                        break;
                    case 2:
                        button3.Text = agendas[i].titulo;
                        break;
                    case 3:
                        button4.Text = agendas[i].titulo;
                        break;      
                }
            }
            switch(agendas.Count)
            {
                case 1:
                    button2.Visible = button3.Visible = button4.Visible = false;
                    break;
                case 2:
                    button3.Visible = button4.Visible = false;
                    break;
                case 3:
                    button4.Visible = false;
                    break;
            }
        }

        //Método para seleccionar agenda a eliminar
        public SeleccionAgenda(List<AgendaClase> listaForms, string texto)
        {
            InitializeComponent();
            this.Text = texto;
            agendas = listaForms;
            for (int i = 0; i < agendas.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        button1.Text = agendas[i].titulo;
                        break;
                    case 1:
                        button2.Text = agendas[i].titulo;
                        break;
                    case 2:
                        button3.Text = agendas[i].titulo;
                        break;
                    case 3:
                        button4.Text = agendas[i].titulo;
                        break;
                }
            }
            switch (agendas.Count)
            {
                case 1:
                    button2.Visible = button3.Visible = button4.Visible = false;
                    break;
                case 2:
                    button3.Visible = button4.Visible = false;
                    break;
                case 3:
                    button4.Visible = false;
                    break;
            }

            button5.Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Button boton = (Button)sender;
            seleccionado = boton.Text;
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Button boton = (Button)sender;
            seleccionado = boton.Text;
            this.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Button boton = (Button)sender;
            seleccionado = boton.Text;
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Button boton = (Button)sender;
            seleccionado = boton.Text;
            this.Dispose();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Button boton = (Button)sender;
            seleccionado = "Nueva";
            this.Dispose();
        }
    }
}
