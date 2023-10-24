using OfficeOpenXml;
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

namespace From_GRID_to_ECxel_WinFormsApp_dotnetframework
{
    public partial class Form1 : Form
    {
        public List<Person> People {  get; set; }

        public Form1()
        {
            People = GetPeople();
            InitializeComponent();
        }
        private List<Person> GetPeople()
        {
            var list = new List<Person>();
            list.Add(new Person()
            {
                PersonId = 1,
                Name = "name1",
                Surname = "surname1",
                Email = "emile1"
            });
            list.Add(new Person()
            {
                PersonId = 2,
                Name = "name2",
                Surname = "surname2",
                Email = "emile2"
            });
            list.Add(new Person()
            {
                PersonId = 3,
                Name = "name3",
                Surname = "surname3",
                Email = "emile3"
            });

            return list;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            var people = this.People;

            dataGridView1.DataSource = people;
            dataGridView1.Columns["PersonId"].Visible = false;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured: " + ex.Message + " - " + ex.Source);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Excel package.
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a worksheet to the Excel package.
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Data");

                // Loop through the DataGridView and export data to Excel.
                for (int row = 0; row < dataGridView1.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView1.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 1, col + 1].Value = dataGridView1[col, row].Value;
                    }
                }

                // Save the Excel package to a file.
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Files|*.xlsx";
                saveDialog.FileName = "Data.xlsx";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo excelFile = new FileInfo(saveDialog.FileName);
                    excelPackage.SaveAs(excelFile);
                }
            }

            MessageBox.Show("Data exported to Excel successfully!");
        }

    }
}
