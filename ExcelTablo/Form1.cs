using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Collections.Generic;

namespace WinFormsExcelApp
{
    public partial class MainForm : Form
    {
        private string excelFilePath;
        private DataTable dataTable = new DataTable();
        private DataGridView dataGridView1;
        private Button saveButton;
        private Button exportButton;
        private List<string> temporaryComboBoxItems = new List<string> { "", "e", "q", "y", "Yeni Değer Ekle..." };

        private Stack<DataTable> undoStack = new Stack<DataTable>();
        private Stack<DataTable> redoStack = new Stack<DataTable>();

        private Button undoButton;
        private Button redoButton;

        public MainForm()
        {
            InitializeComponent();
            ShowFileSelectionWindow();

            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Excel dosyası seçilmedi. Program kapanıyor.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0);
            }

            InitializeDataGridView();
            //InitializeButtons();
            LoadExcelData();
        }

        private void InitializeDataGridView()
        {
            dataGridView1 = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                BackgroundColor = Color.WhiteSmoke, 
                RowHeadersVisible = true, 
                GridColor = Color.LightGray, 
                DefaultCellStyle = new DataGridViewCellStyle { SelectionBackColor = Color.LightSkyBlue, SelectionForeColor = Color.White } 
            };
            dataGridView1.CellClick += DataGridView1_CellClick;
            dataGridView1.CellMouseDown += DataGridView1_CellMouseDown;
            this.Controls.Add(dataGridView1);

            AddAddColumnButton();
        }

        private void SaveState()
        {
            var currentState = dataTable.Copy();
            undoStack.Push(currentState);
            redoStack.Clear();  
        }

        private void UndoButton_Click(object sender, EventArgs e)
        {
            if (undoStack.Count > 0)
            {
                var lastState = undoStack.Pop();
                redoStack.Push(dataTable.Copy());
                dataTable = lastState;
                dataGridView1.DataSource = dataTable;
            }
            else
            {
                MessageBox.Show("Geri alınacak işlem yok.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void RedoButton_Click(object sender, EventArgs e)
        {
            if (redoStack.Count > 0)
            {
                var redoState = redoStack.Pop();
                undoStack.Push(dataTable.Copy());
                dataTable = redoState;
                dataGridView1.DataSource = dataTable;
            }
            else
            {
                MessageBox.Show("Yineleyecek işlem yok.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void AddAddColumnButton()
        {
            Panel panel = new Panel
            {
                Width = 120,
                Height = this.ClientSize.Height,
                Dock = DockStyle.Right,
                BackColor = Color.FromArgb(240, 240, 240)
            };
            Button addColumnButton = new Button
            {
                Text = "Sütun Ekle",
                Width = 100,
                Height = 40,
                Location = new Point(10, 10),
                BackColor = Color.FromArgb(100, 149, 237),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            /*Button deleteColumnButton = new Button
            {
                Text = "Sütun Sil",
                Width = 100,
                Height = 40,
                Location = new Point(10, 60),
                BackColor = Color.FromArgb(237, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            Button deleteRowButton = new Button
            {
                Text = "Satır Sil",
                Width = 100,
                Height = 40,
                Location = new Point(10, 110),
                BackColor = Color.FromArgb(237, 64, 64),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };*/

            addColumnButton.Click += AddColumnButton_Click;
            //deleteColumnButton.Click += DeleteColumnButton_Click;
            //deleteRowButton.Click += DeleteRowButton_Click;

            panel.Controls.Add(addColumnButton);
            //panel.Controls.Add(deleteColumnButton);
            //panel.Controls.Add(deleteRowButton);
            this.Controls.Add(panel);

            Button saveButton = new Button
            {
                Text = "Verileri Kaydet",
                Dock = DockStyle.Bottom,
                Width = 90,
                Height = 40,
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            saveButton.Click += SaveButton_Click;
            panel.Controls.Add(saveButton);
            this.Controls.Add(panel);

            Button exportButton = new Button
            {
                Text = "Excel Çıktısı Al",
                Dock = DockStyle.Bottom,
                Width = 80,
                Height = 40,
                BackColor = Color.SeaGreen,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            exportButton.Click += ExportButton_Click;
            panel.Controls.Add(exportButton);
            this.Controls.Add(panel);

            Button undoButton = new Button
            {
                Text = "Geri Al",
                Dock = DockStyle.Bottom,
                Width = 70,
                Height = 40,
                BackColor = Color.FromArgb(255, 99, 71),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            undoButton.Click += UndoButton_Click;
            panel.Controls.Add(undoButton);
            this.Controls.Add(panel);

            Button redoButton = new Button
            {
                Text = "Yinele",
                Dock = DockStyle.Bottom,
                Height = 40,
                Width = 60,
                BackColor = Color.FromArgb(34, 139, 34),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            redoButton.Click += RedoButton_Click;
            panel.Controls.Add(redoButton);
            this.Controls.Add(panel);

        }

        private void AddColumnButton_Click(object sender, EventArgs e)
        {
            string newColumnName = PromptForNewColumnName();
            if (!string.IsNullOrEmpty(newColumnName))
            {
                dataTable.Columns.Add(newColumnName, typeof(string));
            }
            SaveState();
        }

        private void DeleteColumnButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count > 0)
            {
                int columnIndex = dataGridView1.SelectedCells[0].ColumnIndex;
                if (dataTable.Columns.Count > columnIndex)
                {
                    dataTable.Columns.RemoveAt(columnIndex);
                }
            }
            SaveState();
        }

        private void DeleteRowButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count > 0)
            {
                int rowIndex = dataGridView1.SelectedCells[0].RowIndex;
                if (dataTable.Rows.Count > rowIndex)
                {
                    dataTable.Rows.RemoveAt(rowIndex);
                }
            }
        }

        private void DataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    ContextMenuStrip rowContextMenu = new ContextMenuStrip();
                    rowContextMenu.Items.Add("Satırı Sil", null, (s, ev) => DeleteRowAt(e.RowIndex));
                    rowContextMenu.Show(dataGridView1, dataGridView1.PointToClient(MousePosition));
                }
                else if (e.RowIndex == -1)
                {
                    ContextMenuStrip columnContextMenu = new ContextMenuStrip();
                    columnContextMenu.Items.Add("Sütunu Sil", null, (s, ev) => DeleteColumnAt(e.ColumnIndex));
                    columnContextMenu.Show(dataGridView1, dataGridView1.PointToClient(MousePosition));
                }
            }
        }

        private void DeleteRowAt(int rowIndex)
        {
            dataTable.Rows.RemoveAt(rowIndex);
        }

        private void DeleteColumnAt(int columnIndex)
        {
            dataTable.Columns.RemoveAt(columnIndex);
        }

        private string PromptForNewColumnName()
        {
            using (Form inputForm = new Form())
            {
                inputForm.Width = 300;
                inputForm.Height = 150;
                inputForm.Text = "Yeni Sütun Adı";

                Label label = new Label() { Left = 10, Top = 20, Text = "Yeni sütun adını giriniz:" };
                TextBox textBox = new TextBox() { Left = 10, Top = 50, Width = 260 };
                Button okButton = new Button() { Text = "Ekle", Left = 180, Width = 80, Top = 80, DialogResult = DialogResult.OK };

                okButton.Click += (sender, e) => { inputForm.Close(); };
                inputForm.Controls.Add(label);
                inputForm.Controls.Add(textBox);
                inputForm.Controls.Add(okButton);
                inputForm.AcceptButton = okButton;

                return inputForm.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : null;
            }
        }

        private void ShowFileSelectionWindow()
        {
            using (Form fileSelectionForm = new Form())
            {
                fileSelectionForm.Text = "Excel Dosyası Seçin";
                fileSelectionForm.Width = 560;
                fileSelectionForm.Height = 320;

                Button selectFileButton = new Button
                {
                    Text = "Excel Dosyası Seçin",
                    Width = 400,
                    Height = 80,
                    Location = new Point(80, 80),
                    BackColor = Color.FromArgb(255, 159, 64), 
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 12, FontStyle.Bold)
                };

                selectFileButton.Click += (sender, e) =>
                {
                    SelectExcelFile();
                    fileSelectionForm.Close();
                };

                fileSelectionForm.Controls.Add(selectFileButton);
                fileSelectionForm.ShowDialog();
            }
        }

        private void SelectExcelFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Bir Excel dosyası seçin";
                openFileDialog.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                }
            }
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
                var rows = worksheet.RowsUsed();

                dataTable.Clear();
                dataTable.Columns.Clear();

                foreach (var cell in rows.First().Cells())
                {
                    dataTable.Columns.Add(cell.GetValue<string>(), typeof(string));
                }

                foreach (var row in rows.Skip(1))
                {
                    dataTable.Rows.Add(row.Cells().Select(c => c.GetValue<string>()).ToArray());
                }
            }
            SaveState();
            dataGridView1.DataSource = dataTable;
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            var cell = dataGridView1[e.ColumnIndex, e.RowIndex];

            if (e.ColumnIndex == 0)
            {
                dataGridView1.BeginEdit(true);
            }
            else
            {
                ShowComboBoxInCell(e.RowIndex, e.ColumnIndex, cell);
            }
        }

        private void ShowComboBoxInCell(int rowIndex, int columnIndex, DataGridViewCell cell)
        {
            var cellRect = dataGridView1.GetCellDisplayRectangle(columnIndex, rowIndex, true);
            var comboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(cellRect.Width, 20),
                Location = new Point(cellRect.Left, cellRect.Top),
                ForeColor = Color.Black
            };
            comboBox.Items.AddRange(temporaryComboBoxItems.ToArray());
            comboBox.Text = cell.Value?.ToString() ?? string.Empty;

            comboBox.SelectedIndexChanged += (sender, args) =>
            {
                string selectedValue = comboBox.SelectedItem?.ToString();
                if (selectedValue == "Yeni Değer Ekle...")
                {
                    string newValue = PromptForNewValue();
                    if (!string.IsNullOrEmpty(newValue) && !temporaryComboBoxItems.Contains(newValue))
                    {
                        temporaryComboBoxItems.Insert(temporaryComboBoxItems.Count - 1, newValue);
                    }
                    ShowComboBoxInCell(rowIndex, columnIndex, cell);
                }
                else
                {
                    cell.Value = selectedValue;
                    dataGridView1.Controls.Remove(comboBox);
                }
            };
            dataGridView1.Controls.Add(comboBox);
            comboBox.BringToFront();
            comboBox.Focus();
        }

        private string PromptForNewValue()
        {
            using (Form inputForm = new Form())
            {
                inputForm.Width = 300;
                inputForm.Height = 150;
                inputForm.Text = "Yeni Değer Ekle";

                Label label = new Label() { Left = 10, Top = 20, Text = "Yeni değeri giriniz:" };
                TextBox textBox = new TextBox() { Left = 10, Top = 50, Width = 260 };
                Button okButton = new Button() { Text = "Ekle", Left = 180, Width = 80, Top = 80, DialogResult = DialogResult.OK };

                okButton.Click += (sender, e) => { inputForm.Close(); };
                inputForm.Controls.Add(label);
                inputForm.Controls.Add(textBox);
                inputForm.Controls.Add(okButton);
                inputForm.AcceptButton = okButton;

                return inputForm.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : null;
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
            MessageBox.Show("Excel dosyasına kaydedildi.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Dosyası|*.xlsx",
                Title = "Excel Dosyasını Kaydet"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string exportPath = saveFileDialog.FileName;
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
                    workbook.SaveAs(exportPath);
                }
                MessageBox.Show("Veriler başarıyla dışa aktarıldı.");
            }
        }
    }
}
