using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CMI_lab3
{
    public partial class Form2 : Form
    {
        string[] zaver2;

        public Form2(string[] zaver)
        {
            zaver2 = zaver;
            InitializeComponent();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            zaver2[0] = comboBox1.Text;
            zaver2[1] = comboBox2.Text;
            zaver2[2] = comboBox3.Text;
            zaver2[3] = comboBox4.Text;
        }
    }
}
