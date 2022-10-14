using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbook2
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {

            MessageBox.Show("Bienvenido aqui insertaras los campos requeridos para conectarte al Web Service");

            var sheetActive = (Excel.Worksheet)Globals.ThisWorkbook.ActiveSheet;

            sheetActive.Range["A1"].Value = "User";
            sheetActive.Range["B1"].Value = "Password";
            sheetActive.Range["c1"].Value = "Port";
            sheetActive.Range["D1"].Value = "Protocolo";
            sheetActive.Range["E1"].Value = "Url";




        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {


            MessageBox.Show("Adios");




        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Form miFormulario = new Form1();

            miFormulario.ShowDialog();


        }
    }
}
