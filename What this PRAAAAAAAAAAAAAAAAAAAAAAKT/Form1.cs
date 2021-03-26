using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data;
using MetroFramework.Forms;

namespace What_this_PRAAAAAAAAAAAAAAAAAAAAAAKT
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }

       
        private System.Data.DataTable CreateTable()
        {
           
            System.Data.DataTable dt = new System.Data.DataTable("studotchet");
       
            DataColumn colfiost = new DataColumn("ФИО Студента", typeof(String));
            DataColumn colgr = new DataColumn("Группа", typeof(String));
            DataColumn colpr = new DataColumn("Предмет", typeof(String));
            DataColumn colfioped = new DataColumn("ФИО преподавателя", typeof(String));
            DataColumn colmark = new DataColumn("Оценка", typeof(String));
          
            dt.Columns.Add(colfiost);
            dt.Columns.Add(colgr);
            dt.Columns.Add(colpr);
            dt.Columns.Add(colfioped);
            dt.Columns.Add(colmark);
            
            DataRow row = null;
            row = dt.NewRow();            
            return dt;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = CreateTable();
        }

        private void ExportToExcel()
        {
            Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;
            exApp.Workbooks.Add();
            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;
            workSheet.Cells[1, 1] = "ФИО студента";
            workSheet.Cells[1, 2] = "Группа";
            workSheet.Cells[1, 3] = "Предмет";
            workSheet.Cells[1, 4] = "ФИО преподавателя";
            workSheet.Cells[1, 5] = "Оценка";

            int rowExcel = 2; 
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
             
                workSheet.Cells[rowExcel, "A"] = dataGridView1.Rows[i].Cells["ФИО студента"].Value;
                workSheet.Cells[rowExcel, "B"] = dataGridView1.Rows[i].Cells["Группа"].Value;
                workSheet.Cells[rowExcel, "C"] = dataGridView1.Rows[i].Cells["Предмет"].Value;
                workSheet.Cells[rowExcel, "D"] = dataGridView1.Rows[i].Cells["ФИО преподавателя"].Value;
                workSheet.Cells[rowExcel, "E"] = dataGridView1.Rows[i].Cells["Оценка"].Value;

                ++rowExcel;
            }
            string pathToXmlFile;
            pathToXmlFile = Environment.CurrentDirectory + "\\" + "Отчёт.xlsx";
            workSheet.SaveAs(pathToXmlFile);
           
            exApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel();
            MessageBox.Show("Файл сохранен");
        }
    }
}
