using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace WinFormsExcelApp
{
    public partial class MainForm : Form
    {
        private string excelFilePath = @"C:\Users\egitim2\Desktop\makine ilerleme.xlsx";
        private DataTable dataTable = new DataTable();
        private DataGridView dataGridView1;

        public MainForm()
        {
            InitializeComponent();
            InitializeDataGridView();
            LoadExcelData();
        }

        private void InitializeDataGridView()
        {
            dataGridView1 = new DataGridView
            {
                Size = new System.Drawing.Size(600, 400),
                Location = new System.Drawing.Point(20, 20),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            Controls.Add(dataGridView1);
        }

        private void LoadExcelData()
        {
            if (!File.Exists(excelFilePath))
            {
                MessageBox.Show("Excel dosyası bulunamadı.");
                return;
            }

            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);
                var firstRow = worksheet.FirstRowUsed();  
                var rows = worksheet.RowsUsed().Skip(1);  

                dataTable.Clear();
                dataTable.Columns.Clear();

                
                foreach (var cell in firstRow.CellsUsed())
                {
                    dataTable.Columns.Add(cell.GetValue<string>(), typeof(string));
                }

                dataTable.Columns.Add("Selection", typeof(string)); 

                foreach (var row in rows)
                {
                    var rowData = row.CellsUsed().Select(c => c.GetValue<string>()).ToList();
                    rowData.Add(""); 
                    dataTable.Rows.Add(rowData.ToArray());
                }
            }

            dataGridView1.DataSource = dataTable;
            AddSelectionComboBox();
        }

        private void AddSelectionComboBox()
        {
            if (dataGridView1.Columns.Contains("Selection"))
            {
                return;
            }

            DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn
            {
                HeaderText = "Selection",
                Name = "Selection",
                DataSource = new string[] {"", "e", "q", "y" },
                DataPropertyName = "Selection"
            };

            dataGridView1.Columns.Add(comboBoxColumn);
        }
    }
}
