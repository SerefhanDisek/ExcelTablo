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

            if (dataGridView1 == null)
            {
                dataGridView1 = new DataGridView
                {
                    Dock = DockStyle.Fill, 
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    AllowUserToAddRows = true, 
                    AllowUserToDeleteRows = true, 
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect 
                };

                dataGridView1.CellDoubleClick += DataGridView1_CellDoubleClick; // Çift tıklanınca seçim penceresi açılacak
                this.Controls.Add(dataGridView1);
            }

            LoadExcelData();
        }

        private void LoadExcelData()
        {
            if (!File.Exists(excelFilePath))
            {
                MessageBox.Show("Excel dosyası bulunamadı.");
                return;
            }

            if (dataGridView1 == null)
            {
                MessageBox.Show("HATA: dataGridView1 nesnesi NULL.");
                return;
            }

            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                dataTable.Clear();
                dataTable.Columns.Clear();

                foreach (var cell in rows.First().Cells())
                {
                    dataTable.Columns.Add(cell.GetValue<string>(), typeof(string));
                }

                foreach (var row in rows.Skip(1))
                {
                    var values = row.Cells().Select(c => c.GetValue<string>()).ToArray();
                    dataTable.Rows.Add(values);
                }
            }

            dataGridView1.DataSource = dataTable;
        }

        private void DataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string selectedValue = ShowSelectionDialog();

                if (!string.IsNullOrEmpty(selectedValue))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = selectedValue;
                }
            }
        }

        private string ShowSelectionDialog()
        {
            Form selectionForm = new Form
            {
                Width = 300,
                Height = 200,
                Text = "Seçim Yap",
                StartPosition = FormStartPosition.CenterParent
            };

            ComboBox comboBox = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Items = { "e", "q", "y", "Yeni Değer" } 
            };

            Button okButton = new Button
            {
                Text = "Seç",
                Dock = DockStyle.Bottom
            };

            selectionForm.Controls.Add(comboBox);
            selectionForm.Controls.Add(okButton);

            string selectedValue = null;
            okButton.Click += (s, e) =>
            {
                selectedValue = comboBox.SelectedItem?.ToString();
                selectionForm.Close();
            };

            selectionForm.ShowDialog();
            return selectedValue;
        }

        private void SaveToExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sayfa1");

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cell(1, i + 1).Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cell(i + 2, j + 1).Value = dataTable.Rows[i][j].ToString();
                    }
                }

                workbook.SaveAs(excelFilePath);
            }
        }
    }
}
