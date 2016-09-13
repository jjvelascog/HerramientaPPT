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
    public partial class LinkMessageBox : Form
    {
        public LinkMessageBox()
        {
            InitializeComponent();
        }

        public void inicializar(string link, string titulo)
        {
            linkLabel1.Text = link;
            linkLabel1.Links.Add(0,link.Length,link);
            this.Text = titulo;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
        }
    }
}
