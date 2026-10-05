using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelMerger
{
    public partial class MainWindow : Window
    {
        private List<string> selectedFiles = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnSelectFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Multiselect = true
            };

            if (dlg.ShowDialog() == true)
            {
                selectedFiles.Clear();
                selectedFiles.AddRange(dlg.FileNames);

                txtStatus.Text = "Selected files:\n" + string.Join("\n", selectedFiles);
            }
        }

        private async void BtnMerge_Click(object sender, RoutedEventArgs e)
        {
            if (selectedFiles.Count == 0)
            {
                MessageBox.Show("Please select at least one Excel file.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string mode = ((cmbMergeMode.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString()) ?? "";

            SaveFileDialog saveDlg = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "Merged.xlsx"
            };

            if (saveDlg.ShowDialog() != true)
                return;

            try
            {
                progressBar.Value = 0;
                txtProgress.Text = "Progress: 0%";
                txtStatus.Text += "\nStarting merge...";

                if (mode == "Merge into One Workbook")
                {
                    await Task.Run(() => MergeIntoOneWorkbook(saveDlg.FileName));
                }
                else
                {
                    await Task.Run(() => MergeIntoOneSheet(saveDlg.FileName));
                }

                txtStatus.Text += $"\n\nMerged successfully → {saveDlg.FileName}";
                MessageBox.Show("Merge Completed!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                txtStatus.Text += "\n\nError: " + ex.Message;
                MessageBox.Show("Error during merge: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressBar.Value = 100;
                txtProgress.Text = "Progress: 100%";
            }
        }

        private void UpdateProgress(int current, int total)
        {
            Dispatcher.Invoke(() =>
            {
                double percent = (double)current / total * 100;
                progressBar.Value = percent;
                txtProgress.Text = $"Progress: {percent:F0}%";
            });
        }

        private void MergeIntoOneWorkbook(string outputPath)
        {
            int totalSheets = selectedFiles.Sum(f => new XLWorkbook(f).Worksheets.Count);
            int processed = 0;

            using (var finalWb = new XLWorkbook())
            {
                foreach (var file in selectedFiles)
                {
                    using (var tempWb = new XLWorkbook(file))
                    {
                        foreach (var ws in tempWb.Worksheets)
                        {
                            string sheetName = ws.Name;
                            int suffix = 1;

                            while (finalWb.Worksheets.Any(s => s.Name == sheetName))
                            {
                                sheetName = ws.Name + "_" + suffix;
                                suffix++;
                            }

                            ws.CopyTo(finalWb, sheetName);

                            processed++;
                            UpdateProgress(processed, totalSheets);
                        }
                    }
                }

                finalWb.SaveAs(outputPath);
            }
        }

        private void MergeIntoOneSheet(string outputPath)
        {
            int totalSheets = selectedFiles.Sum(f => new XLWorkbook(f).Worksheets.Count);
            int processed = 0;

            using (var finalWb = new XLWorkbook())
            {
                var finalSheet = finalWb.Worksheets.Add("MergedSheet");
                int currentRow = 1;

                foreach (var file in selectedFiles)
                {
                    using (var tempWb = new XLWorkbook(file))
                    {
                        foreach (var ws in tempWb.Worksheets)
                        {
                            var usedRange = ws.RangeUsed();
                            if (usedRange == null) continue;

                            finalSheet.Cell(currentRow, 1).Value = $"Source: {Path.GetFileName(file)} - {ws.Name}";
                            finalSheet.Cell(currentRow, 1).Style.Font.Bold = true;
                            currentRow++;

                            usedRange.CopyTo(finalSheet.Cell(currentRow, 1));
                            currentRow += usedRange.RowCount() + 2;

                            processed++;
                            UpdateProgress(processed, totalSheets);
                        }
                    }
                }

                finalWb.SaveAs(outputPath);
            }
        }
    }
}
