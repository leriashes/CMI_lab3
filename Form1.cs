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
    public partial class Form1 : Form
    {
        int countRows = 1;

        public Form1()
        {
            InitializeComponent();
        }

        private void DataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            int delta = dataGridView2.Rows.Count - countRows;

            countRows += delta;

            for (; delta > 0; delta--) 
            {
                dataGridView1.Rows.Add();

                dataGridView3.Rows.Add();

                if (tabControl1.Height < 312)
                {
                    dataGridView1.Height += 22;
                    dataGridView2.Height += 22;
                    dataGridView3.Height += 22;

                    tabControl1.Height += 22;
                }
                else if (!vScrollBar1.Visible)
                {
                    vScrollBar1.Visible = true;
                }
            }
        }

        private void DataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[0].Value = dataGridView1.Rows.Count - 1;
        }

    }
}
