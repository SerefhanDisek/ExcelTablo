using System;
using System.Data;
using System.Drawing;
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
        private Button saveButton;

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
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2 
                };

                dataGridView1.CellClick += DataGridView1_CellClick;
                dataGridView1.CellValueChanged += DataGridView1_CellValueChanged; 
                this.Controls.Add(dataGridView1);
            }

            saveButton = new Button
            {
                Text = "Verileri Kaydet",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            saveButton.Click += SaveButton_Click;
            this.Controls.Add(saveButton);

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

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var cell = dataGridView1[e.ColumnIndex, e.RowIndex];
                ShowComboBoxInCell(e.RowIndex, e.ColumnIndex, cell);
            }
        }

        private void ShowComboBoxInCell(int rowIndex, int columnIndex, DataGridViewCell cell)
        {
            var cellRect = dataGridView1.GetCellDisplayRectangle(columnIndex, rowIndex, true);

            if (dataGridView1.Controls.OfType<ComboBox>().Any(c => c.Parent == dataGridView1 && c.Bounds.IntersectsWith(cellRect)))
                return;

            var comboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Items = {"", "e", "q", "y", "Yeni Değer" }, 
                Text = cell.Value?.ToString() ?? string.Empty, 
                Size = new Size(cellRect.Width, 20), 
                Location = new Point(cellRect.Left, cellRect.Top)
            };

            comboBox.SelectedIndexChanged += (sender, args) =>
            {
                cell.Value = comboBox.SelectedItem?.ToString();
                dataGridView1.Controls.Remove(comboBox); 
            };

            dataGridView1.Controls.Add(comboBox);
            comboBox.BringToFront();
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var changedCell = dataGridView1[e.ColumnIndex, e.RowIndex];
                dataTable.Rows[e.RowIndex][e.ColumnIndex] = changedCell.Value.ToString();
            }
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

        private void SaveButton_Click(object sender, EventArgs e)
        {
            SaveToExcel();
            MessageBox.Show("Veriler başarıyla kaydedildi.");
        }
    }
}
