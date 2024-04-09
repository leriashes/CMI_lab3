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

        string[] fio = {
            "Абрамова А.Н.",
            "Бухтоярова А.А.",
            "Ворошина У.Н.",
            "Железняк Т.В.",
            "Заварушкина М.М.",
            "Зайцева Л.И.",
            "Парускин И.Д.",
            "Петрова Е.И.",
            "Полетаев К.А.",
            "Сальникова В.В.",
            "Стоянова А.К.",
            "Топалова М.В.",
            "Хлопченко И.В."};

        int[] job =
        {
            2,
            0,
            0,
            0,
            1,
            3,
            3,
            3,
            2,
            0,
            0,
            2,
            1
        };

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

        private void ComboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < fio.Length; i++) 
            {
                if (fio[i] == comboBox3.Text)
                {
                    comboBox4.SelectedIndex = job[i] + 1;
                    break;
                }
            }
        }

        private void ComboBox4_SelectedValueChanged(object sender, EventArgs e)
        {
            string value = comboBox3.Text;

            int valnum = 0;

            for (int i = 0; i < fio.Length; i++)
            {
                if (fio[i] == value)
                {
                    valnum = job[i] + 1;
                    break;
                }
            }

            if (comboBox4.SelectedItem == null || comboBox4.SelectedIndex == 0)
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                "Абрамова А.Н.",
                "Бухтоярова А.А.",
                "Ворошина У.Н.",
                "Железняк Т.В.",
                "Заварушкина М.М.",
                "Зайцева Л.И.",
                "Парускин И.Д.",
                "Петрова Е.И.",
                "Полетаев К.А.",
                "Сальникова В.В.",
                "Стоянова А.К.",
                "Топалова М.В.",
                "Хлопченко И.В."});

                comboBox3.Text = "";
            }
            else if (comboBox4.SelectedIndex == 1)
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                "Бухтоярова А.А.",
                "Ворошина У.Н.",
                "Железняк Т.В.",
                "Сальникова В.В.",
                "Стоянова А.К."});

                if (valnum == 1)
                    comboBox3.Text = value;
                else
                    comboBox3.Text = "";
            }
            else if (comboBox4.SelectedIndex == 2)
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                "Заварушкина М.М.",
                "Хлопченко И.В."
                });

                if (valnum == 2)
                    comboBox3.Text = value;
                else
                    comboBox3.Text = "";
            }
            else if (comboBox4.SelectedIndex == 3)
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                "Абрамова А.Н.",
                "Полетаев К.А.",
                "Топалова М.В."});

                if (valnum == 3)
                    comboBox3.Text = value;
                else
                    comboBox3.Text = "";
            }
            else if (comboBox4.SelectedIndex == 4)
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                "Зайцева Л.И.",
                "Парускин И.Д.",
                "Петрова Е.И.",});

                if (valnum == 4)
                    comboBox3.Text = value;
                else
                    comboBox3.Text = "";
            }
        }
    }
}
