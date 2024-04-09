using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace CMI_lab3
{
    public partial class Form1 : Form
    {
        int countRows = 1;
        int applyDoc = 0;
        string prevValue = "";
        string[] zaver = { "", "", "", ""};

        string[] OKPO = { "2016639571",
            "2141051250",
            "10006475",
            "52078951",
            "65386358"};

        string[] OKDP =
        {
            "10.71",
            "10.72",
            "47.11",
            "47.24",
            "10.71"
        };

        int[] edizm = { };

        string[] kod = { "105",
            "127",
            "162",
            "164",
            "165",
            "169",
            "171"};

        string[] kodOKEI = {
            "163",
            "166",
            "796"};



        public Form1()
        {
            InitializeComponent();
            dateTimePicker3.MinDate = dateTimePicker2.Value;
            dataGridView4.Rows.Add();
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

                    for (int i = 0; i < kod.Length && val2 < 0 && dataGridView1.Rows[e.RowIndex].Cells[2].Value != null; i++)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString() == kod[i])
                        {
                            val2 = i;
                        }
                    }
                    
                    if (val1 != val2)
                    {
                        if (e.ColumnIndex == 1) 
                        {
                            if (val1 >= 0)
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[2].Value = kod[val1];
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[2].Value = "";
                            }
                        }
                        else
                        {
                            if (val2 >= 0)
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[1].Value = Column2.Items[val2];
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[1].Value = "";
                            }
                        }
                    }
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

                    for (int i = 0; i < kodOKEI.Length && val2 < 0 && dataGridView2.Rows[e.RowIndex].Cells[1].Value != null; i++)
                    {
                        if (dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString() == kodOKEI[i])
                        {
                            val2 = i;
                        }
                    }

                    if (val1 != val2)
                    {
                        if (e.ColumnIndex == 0)
                        {
                            if (val1 >= 0)
                            {
                                dataGridView2.Rows[e.RowIndex].Cells[1].Value = kodOKEI[val1];
                            }
                            else
                            {
                                dataGridView2.Rows[e.RowIndex].Cells[1].Value = "";
                            }
                        }
                        else
                        {
                            if (val2 >= 0)
                            {
                                dataGridView2.Rows[e.RowIndex].Cells[0].Value = Column4.Items[val2];
                            }
                            else
                            {
                                dataGridView2.Rows[e.RowIndex].Cells[0].Value = "";
                            }
                        }
                    }
                }
            }

            if (e.ColumnIndex == 3)
            {
                CountSum(e.RowIndex);

                for (int i = 0; i < dataGridView3.ColumnCount; i += 2)
                {
                    CountRes(i);
                }
            }
        }

        private void CountSum(int row) 
        {
            if (row >= 0)
            {
                for (int i = 1; i < dataGridView3.ColumnCount; i += 2)
                {
                    if (dataGridView3.Rows[row].Cells[i - 1].Value == null || dataGridView3.Rows[row].Cells[i - 1].Value.ToString() == "")
                    {
                        dataGridView3.Rows[row].Cells[i].Value = "";
                    }
                    else
                    {
                        int num = Convert.ToInt32(dataGridView3.Rows[row].Cells[i - 1].Value.ToString());

                        if (dataGridView2.Rows[row].Cells[3].Value == null || dataGridView2.Rows[row].Cells[3].Value.ToString() == "")
                        {
                            dataGridView3.Rows[row].Cells[i].Value = "";
                        }
                        else
                        {
                            double cost = Convert.ToDouble(dataGridView2.Rows[row].Cells[3].Value.ToString());

                            dataGridView3.Rows[row].Cells[i].Value = num * cost;
                        }
                    }
                }
            }
        }

        private void CountRes(int col)
        {
            if (dataGridView4.RowCount == 0)
            {
                if (dataGridView4.ColumnCount > 0)
                {
                    dataGridView4.Rows.Add();

                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                        dataGridView4.Rows[0].Cells[i].Value = 12340;
                }
            }
            else
            {
                int sum1 = 0;
                double sum2 = 0;

                for (int i = 0; i < countRows - 1; i++)
                {
                    if (dataGridView3.Rows[i].Cells[col].Value != null && dataGridView3.Rows[i].Cells[col].Value.ToString() != "")
                    {
                        sum1 += Convert.ToInt32(dataGridView3.Rows[i].Cells[col].Value.ToString());
                    }

                    if (dataGridView3.Rows[i].Cells[col + 1].Value != null && dataGridView3.Rows[i].Cells[col + 1].Value.ToString() != "")
                    {
                        sum2 += Convert.ToDouble(dataGridView3.Rows[i].Cells[col + 1].Value.ToString());
                    }
                }

                dataGridView4.Rows[0].Cells[col].Value = sum1;
                dataGridView4.Rows[0].Cells[col + 1].Value = sum2;
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            string[] zfio = new string[4];
            Form2 form2 = new Form2(zfio);
            form2.ShowDialog();

            if (form2.DialogResult == DialogResult.OK)
            {
                for (int i = 0; i < 4; i++)
                {
                    zaver[i] = zfio[i];
                }
            }
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
            dataGridView4.Rows[0].Cells[0].Selected = false;
        }

        private void DataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (vScrollBar1.Visible)
                vScrollBar1.Value = dataGridView2.FirstDisplayedScrollingRowIndex;
        }

        private void DataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (vScrollBar1.Visible)
                vScrollBar1.Value = dataGridView1.FirstDisplayedScrollingRowIndex;
        }

        private void DataGridView3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (vScrollBar1.Visible)
                vScrollBar1.Value = dataGridView3.FirstDisplayedScrollingRowIndex;
        }

        private void DataGridView1_Leave(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++) 
            {
                for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++) 
                {
                    if (dataGridView1.Rows[i].Cells[j].Selected)
                        dataGridView1.Rows[i].Cells[j].Selected = false;
                }
            }
        }

        private void DataGridView2_Leave(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView2.Rows[i].Cells.Count; j++)
                {
                    if (dataGridView2.Rows[i].Cells[j].Selected)
                        dataGridView2.Rows[i].Cells[j].Selected = false;
                }
            }
        }

        private void DataGridView3_Leave(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView3.Rows[i].Cells.Count; j++)
                {
                    if (dataGridView3.Rows[i].Cells[j].Selected)
                        dataGridView3.Rows[i].Cells[j].Selected = false;
                }
            }
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].Selected = dataGridView1.Rows[i].Selected;
            }

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                dataGridView3.Rows[i].Selected = dataGridView1.Rows[i].Selected;
            }
        }

        private void DeleteRow()
        {
            if (dataGridView1.Rows.Count == 10)
            {
                vScrollBar1.Visible = false;
            }
            
            if (tabControl1.Height > 21 && !vScrollBar1.Visible)
            {
                label37.Location = new System.Drawing.Point(label37.Location.X, label37.Location.Y - 20);
                dataGridView4.Location = new System.Drawing.Point(dataGridView4.Location.X, dataGridView4.Location.Y - 20);

                dataGridView1.Height -= 20;
                dataGridView2.Height -= 20;
                dataGridView3.Height -= 20;

                tabControl1.Height -= 20;
            }

            if (vScrollBar1.Visible)
            {
                vScrollBar1.Maximum = dataGridView1.RowCount - 9;
                vScrollBar1.Value = dataGridView1.FirstDisplayedScrollingRowIndex + 1;
            }
        }

        private void DataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (e.Row.Index != dataGridView1.RowCount - 1)
            {
                dataGridView2.Rows.RemoveAt(e.Row.Index);
                dataGridView3.Rows.RemoveAt(e.Row.Index);
                countRows--;

                DeleteRow();
            }
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();

            string filename = "Example.XLS";
            string[] exc = { "A", "D", "N", "Q", "U", "Y", "AC", "AH", "AN", "AX", "BB", "BH", "BL", "BQ", "BU", "BY", "CF" };

            if (countRows - 1 > 8)
                filename = "Example1.XLS";

            workbook.LoadFromFile(filename);

            Worksheet sheet = workbook.Worksheets[0];


            sheet.Range["A7"].Text = comboBox1.Text;
            sheet.Range["A9"].Text = comboBox2.Text;
            sheet.Range["BA13"].Text = textBox5.Text;
            sheet.Range["CA7"].Text = textBox1.Text;
            sheet.Range["CA10"].Text = textBox2.Text;
            sheet.Range["CA11"].Text = comboBox5.Text;

            if (dateTimePicker1.Text != "")
            {
                dateTimePicker1.Format = DateTimePickerFormat.Short;
                sheet.Range["BL13"].Text = dateTimePicker1.Text;
                dateTimePicker1.Format = DateTimePickerFormat.Long;
            }

            if (dateTimePicker2.Text != "")
            {
                string[] words = dateTimePicker2.Text.Split(' ');
                sheet.Range["AI17"].Text = words[0];
                sheet.Range["AL17"].Text = words[1];
                sheet.Range["AR17"].Text = words[2];
            }

            if (dateTimePicker3.Text != "")
            {
                string[] words = dateTimePicker3.Text.Split(' ');
                sheet.Range["BZ17"].Text = words[0];
                sheet.Range["CD17"].Text = words[1];
                sheet.Range["CJ17"].Text = words[2];
            }

            sheet.Range["R32"].Text = zaver[0];
            sheet.Range["BC32"].Text = zaver[1];
            sheet.Range["S34"].Text = zaver[3];
            sheet.Range["AX34"].Text = zaver[2];

            int row = 23;

            for (int i = 0; i < countRows - 1; i++, row++)
            {
                for (int j = 0, k = 0; j < 17; j++, k++)
                {
                    string cell = exc[j] + row.ToString();
                    DataGridView table;

                    if (j < 3)
                    {
                        table = dataGridView1;
                    }
                    else if (j < 7)
                    {
                        table = dataGridView2;
                    }
                    else
                    {
                        table = dataGridView3;
                    }

                    if (table.Rows[i].Cells[k].Value != null && table.Rows[i].Cells[k].Value != "")
                        sheet.Range[cell].Text = table.Rows[i].Cells[k].Value.ToString();
                    else
                        sheet.Range[cell].Text = "X";

                    if (j == 2 || j == 6)
                        k = -1;
                }
            }

            row = 31;

            if (countRows - 1 > 8)
            {
                row = 39;
            }

            for (int i = 0; i < 10; i++)
            {
                string cell = exc[i + 7] + row.ToString();
                
                if (dataGridView4.Rows[0].Cells[i].Value != null)
                    sheet.Range[cell].Text = dataGridView4.Rows[0].Cells[i].Value.ToString();
                else
                    sheet.Range[cell].Text = "X";
            }

            workbook.SaveToFile("Sample.xls");

        }

        private void ComboBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = OKPO[comboBox1.SelectedIndex];
            textBox2.Text = OKDP[comboBox1.SelectedIndex];
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker3.MinDate = dateTimePicker2.Value;
        }

        private void DataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex % 2 == 0)
            {
                CountSum(e.RowIndex);

                CountRes(e.ColumnIndex);
            }
        }

        private void DataGridView4_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView4.Rows[0].Cells[e.ColumnIndex].Selected = false;
        }

        private void DataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void DataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView2.IsCurrentCellDirty)
            {
                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void DataGridView3_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView3.IsCurrentCellDirty)
            {
                dataGridView3.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3(applyDoc);
            form3.ShowDialog();

            if (form3.DialogResult == DialogResult.OK)
            {
                applyDoc = form3.GetNum();
            }
        }
    }
}
