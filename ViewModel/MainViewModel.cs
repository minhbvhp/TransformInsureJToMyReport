﻿using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TransformInsureJToMyReport.ViewModel
{
    internal partial class MainViewModel : ObservableObject
    {
        #region IJNotInReport
        [ObservableProperty]
        private List<string> iJNotInReport;
        private IEnumerable<string> ReadUploadIJNotInReportFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                FileInfo excelFile = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];                    
                    int rowCount = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (worksheet.Cells[row, 1].Value != null)
                        {
                            yield return worksheet.Cells[row, 1].Value.ToString();
                        }
                        else
                        {
                            yield return "Không có dữ liệu";
                        }
                    }
                }                
            }
            else
            {
                yield return "Không có dữ liệu";
            }
        }

        [RelayCommand]
        private async Task UploadIJNotInReport()
        {
            string? fileName = "";
            var dialog = new OpenFileDialog();
            dialog.Filter = "Workbook (*.xlsx)|*.xlsx";
            dialog.Title = "Chọn file liệt kê những đơn thiếu";

            if (dialog.ShowDialog() == true)
            {
                fileName = dialog.FileName;

                if (!String.IsNullOrEmpty(fileName))
                {
                    IJNotInReport = await Task.Run(() => ReadUploadIJNotInReportFile(fileName).ToList());
                }
            }
        }
        #endregion

        #region ExportReport
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(ExportReportCommand))]
        private List<string>? insureJFiles;
        partial void OnInsureJFilesChanged(List<string>? value)
        {
            MessageBox.Show("đã thay đổi");
        }

        private bool CanExportReport()
        {
            return InsureJFiles is not null;
        }

        [RelayCommand(CanExecute = nameof(CanExportReport))]
        private async Task ExportReport()
        {
            foreach (string insureJFile in InsureJFiles)
            {
                if (File.Exists(insureJFile))
                {
                    using (ExcelPackage package = new ExcelPackage(insureJFile))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.End.Row;

                        MessageBox.Show("File có {0} dòng", rowCount.ToString());
                    }
                }
            }
        }

        #endregion




        public MainViewModel()
        {
            
        }
    }
}
