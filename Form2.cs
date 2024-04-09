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
                    comboBox4.SelectedIndex = job[i];
                    break;
                }
            }
        }
    }
}
