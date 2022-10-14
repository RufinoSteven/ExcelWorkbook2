using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelWorkbook2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void Btn_Registrar_Click(object sender, EventArgs e)
        {

            var sheet1 = Globals.Sheet1;
            var r = sheet1.Range["A1"];
            long f = r.CurrentRegion.Rows.Count + 1;

            sheet1.Cells[f, 1].Value = textBox1.Text;
            sheet1.Cells[f, 2].Value = textBox2.Text;
            sheet1.Cells[f, 3].Value = textBox3.Text;
            sheet1.Cells[f, 4].Value = textBox4.Text;
            sheet1.Cells[f, 5].Value = textBox5.Text;



            this.textBox1.Text = "";
            this.textBox2.Text = "";
            this.textBox3.Text = "";
            this.textBox4.Text = "";
            this.textBox5.Text = "";



            this.textBox1.Select();

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
