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
    public partial class Form3 : Form
    {
        int applyDoc;

        public Form3(int num)
        {
            InitializeComponent();

            numericUpDown1.Value = Convert.ToInt32(num);
            applyDoc = num;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            applyDoc = Convert.ToInt32(numericUpDown1.Value);
        }

        public int GetNum()
        {
            return applyDoc;
        }
    }
}
