using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
        private Form1 _form1;
        public Form3(Form1 form1)
        {
            InitializeComponent();
            LoadBox();
            _form1 = form1;
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }
        private void LoadBox()
        {

            listBox1.Items.Add("Запрос 1");
            listBox1.Items.Add("Запрос 2");
            listBox1.Items.Add("Запрос 3");
            listBox1.Items.Add("Запрос 4");

        }
        private void button1_Click(object sender, EventArgs e)
        {
           int index =  listBox1.SelectedIndex;
            switch (index)
            {
                case 0:
                    _form1.temp_(1, "C:\\Users\\angerquell\\Desktop\\Новая папка\\Пример шаблона отчета1.docx", "Стоимость выплат");
                    break;
                case 1:
                    _form1.temp_(2, "C:\\Users\\angerquell\\Desktop\\Новая папка\\Пример шаблона отчета2.docx", "Количество заключенных  договоров");
                    break;
                case 2:
                    _form1.temp_(3, "C:\\Users\\angerquell\\Desktop\\Новая папка\\Пример шаблона отчета3.docx", "Стоимость выплат");
                    break;
                case 3:
                    _form1.temp_(4, "C:\\Users\\angerquell\\Desktop\\Новая папка\\Пример шаблона отчета4.docx", "Стоимость выплат");
                    break;
            }
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
