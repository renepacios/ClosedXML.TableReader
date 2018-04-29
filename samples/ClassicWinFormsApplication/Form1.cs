using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using ClosedXML.TableReader.Model;

namespace ClassicWinFormsApplication
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            var s = GetExcelPath();
            var bytes = GetExcel(s);
            var wb = new ClosedXML.Excel.XLWorkbook(new MemoryStream(bytes));




            var l = wb.ReadTable(1,
                new ReadOptions()
                {
                    TitlesInFirstRow = true
                });

            dataGridView1.DataSource = l;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetExcelPath();
        }

        private static byte[] GetExcel(string excelPath)
        {
            byte[] file = System.IO.File.ReadAllBytes(excelPath);

            return file;
        }

        private static string GetExcelPath(string name = "TableWithHeaders.xlsx")
        {
            return Path.Combine(Directory.GetCurrentDirectory(), "Excels", name);
        }
    }
}
