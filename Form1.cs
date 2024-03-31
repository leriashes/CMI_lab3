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

        private void AddRows(int delta, DataGridView table, DataGridView table1, DataGridView table2)
        {
            countRows += delta;

            for (; delta > 0; delta--)
            {
                if (tabControl1.Height < 300)
                {
                    label37.Location = new System.Drawing.Point(label37.Location.X, label37.Location.Y + 20);
                    dataGridView4.Location = new System.Drawing.Point(dataGridView4.Location.X, dataGridView4.Location.Y + 20);

                    dataGridView1.Height += 20;
                    dataGridView2.Height += 20;
                    dataGridView3.Height += 20;

                    tabControl1.Height += 20;
                }
                else if (!vScrollBar1.Visible)
                {
                    vScrollBar1.Visible = true;
                }

                table1.Rows.Add();
                table2.Rows.Add();

                if (vScrollBar1.Visible)
                {
                    vScrollBar1.Maximum = dataGridView1.RowCount - 9;
                    vScrollBar1.Value = table.FirstDisplayedScrollingRowIndex + 1;
                }
            }
        }

        private void DataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            int delta = dataGridView2.Rows.Count - countRows;

            AddRows(delta, dataGridView2, dataGridView1, dataGridView3);
        }

        private void DataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[0].Value = dataGridView1.Rows.Count - 1;

            int delta = dataGridView1.Rows.Count - countRows;

            AddRows(delta, dataGridView1, dataGridView2, dataGridView3);
        }

        private void DataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            int delta = dataGridView3.Rows.Count - countRows;

            AddRows(delta, dataGridView3, dataGridView1, dataGridView2);
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int val1 = -1;
                int val2 = -1;

                if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
                {
                    for (int i = 0; i < Column2.Items.Count && val1 < 0 && dataGridView1.Rows[e.RowIndex].Cells[1].Value != null; i++)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == Column2.Items[i].ToString())
                        {
                            val1 = i;
                        }
                    }

                    //for (int i = 0; i < Column3.Items.Count && val2 < 0 && dataGridView1.Rows[e.RowIndex].Cells[2].Value != null; i++)
                    //{
                    //    if (dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString() == Column3.Items[i].ToString())
                    //    {
                    //        val2 = i;
                    //    }
                    //}


                    //if (val1 != val2)
                    //{
                    //    if (e.ColumnIndex == 1)
                    //    {
                    //        dataGridView1.Rows[e.RowIndex].Cells[2].Value = Column3.Items[val1];
                    //    }
                    //    else
                    //    {
                    //        dataGridView1.Rows[e.RowIndex].Cells[1].Value = Column2.Items[val2];
                    //    }
                    //}
                }
            }
        }

        private void DataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int val1 = -1;
                int val2 = -1;

                if (e.ColumnIndex == 0 || e.ColumnIndex == 1)
                {
                    for (int i = 0; i < Column4.Items.Count && val1 < 0 && dataGridView2.Rows[e.RowIndex].Cells[0].Value != null; i++)
                    {
                        if (dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString() == Column4.Items[i].ToString())
                        {
                            val1 = i;
                        }
                    }

                    //for (int i = 0; i < Column5.Items.Count && val2 < 0 && dataGridView2.Rows[e.RowIndex].Cells[1].Value != null; i++)
                    //{
                    //    if (dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString() == Column5.Items[i].ToString())
                    //    {
                    //        val2 = i;
                    //    }
                    //}


                    //if (val1 != val2)
                    //{
                    //    if (e.ColumnIndex == 0)
                    //    {
                    //        dataGridView2.Rows[e.RowIndex].Cells[1].Value = Column5.Items[val1];
                    //    }
                    //    else
                    //    {
                    //        dataGridView2.Rows[e.RowIndex].Cells[0].Value = Column4.Items[val2];
                    //    }
                    //}
                }
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void VScrollBar1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView1.FirstDisplayedScrollingRowIndex = vScrollBar1.Value;
            dataGridView2.FirstDisplayedScrollingRowIndex = vScrollBar1.Value;
            dataGridView3.FirstDisplayedScrollingRowIndex = vScrollBar1.Value;
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label37.Visible = dataGridView4.Visible = tabControl1.SelectedIndex == 1;
        }

        private void DataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            vScrollBar1.Value = dataGridView2.FirstDisplayedScrollingRowIndex;
        }

        private void DataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            vScrollBar1.Value = dataGridView1.FirstDisplayedScrollingRowIndex;
        }

        private void DataGridView3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            vScrollBar1.Value = dataGridView3.FirstDisplayedScrollingRowIndex;
        }
    }
}
