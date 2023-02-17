using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TransformInsureJToMyReport.ViewModel
{
    internal partial class MainViewModel : ObservableObject
    {
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

        [ObservableProperty]
        private List<string>? iJNotInReport;

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

        public MainViewModel()
        {
            
        }
    }
}
