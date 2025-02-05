﻿using System;
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

        private List<string> comboBoxItems = new List<string> { "", "e", "q", "y", "Yeni Değer Ekle..." };

        public MainForm()
        {
            InitializeComponent();

            SelectExcelFile(); // Kullanıcıdan dosya seçmesini iste

            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Excel dosyası seçilmedi. Program kapanıyor.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Environment.Exit(0); // Kullanıcı dosya seçmezse uygulamayı kapat
            }

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

            exportButton = new Button
            {
                Text = "Excel Çıktısı Al",
                Dock = DockStyle.Bottom,
                Height = 40
            };
            exportButton.Click += ExportButton_Click;
            this.Controls.Add(exportButton);

            LoadExcelData(); // Seçilen dosyayı yükle
        }

        private void SelectExcelFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Bir Excel dosyası seçin";
                openFileDialog.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                openFileDialog.Multiselect = false;

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
                    var values = row.Cells().Select(c => c.GetValue<string>()).ToArray();
                    dataTable.Rows.Add(values);
                }
            }

            dataGridView1.DataSource = dataTable;

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.ReadOnly = column.Index != 0; // İlk sütun düzenlenebilir, diğerleri seçim yapılabilir olmalı
            }
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

            if (dataGridView1.Controls.OfType<ComboBox>().Any(c => c.Parent == dataGridView1 && c.Bounds.IntersectsWith(cellRect)))
                return;

            var comboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(cellRect.Width, 20),
                Location = new Point(cellRect.Left, cellRect.Top)
            };

            comboBox.Items.AddRange(comboBoxItems.ToArray());
            comboBox.Text = cell.Value?.ToString() ?? string.Empty;

            comboBox.SelectedIndexChanged += (sender, args) =>
            {
                string selectedValue = comboBox.SelectedItem?.ToString();

                if (selectedValue == "Yeni Değer Ekle...")
                {
                    string newValue = PromptForNewValue();
                    if (!string.IsNullOrEmpty(newValue) && !comboBoxItems.Contains(newValue))
                    {
                        comboBoxItems.Insert(comboBoxItems.Count - 1, newValue);
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
            MessageBox.Show("Veriler başarıyla kaydedildi.");
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            SaveToExcel();
            MessageBox.Show("Veriler başarıyla dışa aktarıldı.");
        }
    }
}
