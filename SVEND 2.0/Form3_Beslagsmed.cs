using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SVEND_2._0
{
    public partial class Form3_Beslagsmed : Form
    {
        public Form3_Beslagsmed()
        {
            InitializeComponent();
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //Process.Start("https://www.google.dk");
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
