﻿using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace TransformInsureJToMyReport.ViewModel
{
    internal partial class MainViewModel : ObservableObject
    {
        private Dictionary<string, int> _allTitle = new Dictionary<string, int>();
        private List<List<string>> _allMatchDataFetchFromIJFile = new();
        private int GetUsefulCategory(HashSet<string> strings, List<string> substrings)
        {
            int i = 0;
            foreach (var item in strings)
            {
                i++;
                if (substrings.Any(s => item.Contains(s, StringComparison.OrdinalIgnoreCase)))
                {
                    return i;
                }
            }
            throw new ArgumentNullException("Substring not match strings");
        }

        [ObservableProperty]
        private ObservableCollection<string> insureJFiles = new();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RemoveFileCommand))]
        private string selectedInsureJFile;

        #region IJNotInReport
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(ExportReportCommand))]
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
       
        #region AddFile
        [RelayCommand]
        private void AddFile()
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Workbook (*.xlsx)|*.xlsx";
            dialog.Title = "Chọn file trích xuất trực tiếp từ InsureJ";

            if (dialog.ShowDialog() == true)
            {
                if (!InsureJFiles.Contains(dialog.FileName))
                {
                    InsureJFiles.Add(dialog.FileName);
                    ExportReportCommand.NotifyCanExecuteChanged();
                }
            }
        }
        #endregion

        #region RemoveFile
        private bool CanRemove()
            => SelectedInsureJFile != null;

        [RelayCommand(CanExecute = nameof(CanRemove))]
        private void RemoveFile()
        {            
            InsureJFiles.Remove(SelectedInsureJFile);
            ExportReportCommand.NotifyCanExecuteChanged();
        }
        #endregion

        #region ExportReport

        private Dictionary<string, int> CustomTitleColumn(HashSet<string> indicators)
        {
            return new Dictionary<string, int>
            {
                ["Ngày tạo đơn"] = GetUsefulCategory(indicators, new List<string>{"N.Nhập"}),
                ["PKD"] = GetUsefulCategory(indicators, new List<string>{"Đơn vị KD cấp dưới"}),
                ["Sản phẩm bảo hiểm"] = GetUsefulCategory(indicators, new List<string>{"Sản phẩm", "Tên sản phẩm"}),
                ["Trung gian"] = GetUsefulCategory(indicators, new List<string>{"Đại lý", "Bên trung gian"}),
                ["Số tiền bảo hiểm"] = GetUsefulCategory(indicators, new List<string>{"Tổng tiền", "Tổng số tiền"}),
                ["Số đơn"] = GetUsefulCategory(indicators, new List<string>{"Đơn BH"}),
                ["Đồng BH"] = GetUsefulCategory(indicators, new List<string>{"Cty Đồng BH"}),
                ["Ngày cấp đơn"] = GetUsefulCategory(indicators, new List<string>{"N.Nhập"}),
                ["Hiệu lực từ"] = GetUsefulCategory(indicators, new List<string>{"N.Bắt đầu BH", "N.Hiệu lực"}),
                ["Hiệu lực đến"] = GetUsefulCategory(indicators, new List<string>{"N.Hết hiệu lực"}),
                ["Loại tiền"] = GetUsefulCategory(indicators, new List<string>{"Loại tiền BH"}),
                ["ST phải trả"] = GetUsefulCategory(indicators, new List<string>{"Phí PS NET"}),
                ["Hạn thanh toán"] = GetUsefulCategory(indicators, new List<string>{"Ngày đến hạn TT"}),
                ["Ngày ký"] = GetUsefulCategory(indicators, new List<string>{"N.Nhập"}),
                ["Người nhận"] = GetUsefulCategory(indicators, new List<string>{"Nhân viên QLDV"}),

            };
        }

        private bool CanExportReport()
            => InsureJFiles.Any() && IJNotInReport != null && IJNotInReport.Any();

        [RelayCommand(CanExecute = nameof(CanExportReport))]
        private void ExportReport()
        {
            HashSet<string> indicators = new HashSet<string>();
            var titleColumn = new Dictionary<string, int>();

            string? fileName = "";
            var dialog = new SaveFileDialog();
            dialog.Filter = "Workbook (*.xlsx)|*.xlsx";
            dialog.Title = "Xuất báo cáo";

            if (dialog.ShowDialog() == true)
            {
                fileName = dialog.FileName;                
            }

            //Read InsureJ Files one by one
            foreach (string insureJFile in InsureJFiles)
            {
                if (File.Exists(insureJFile))
                {
                    using (ExcelPackage package = new ExcelPackage(insureJFile))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.End.Row;
                        int colCount = worksheet.Dimension.End.Column;

                        for (int col = 1; col < colCount; col++)
                        {
                            if (worksheet.Cells[1, col].Value != null)
                            {
                                indicators.Add(worksheet.Cells[1, col].Value.ToString());
                            }
                        }                        

                        titleColumn = CustomTitleColumn(indicators);

                        //Pull data to _allMatchDataFetchFromIJFile
                        for (int row = 2; row < rowCount; row++)
                        {
                            string policyNumber = worksheet.Cells[row, titleColumn["Số đơn"]].Value.ToString();

                            if (IJNotInReport.Contains(policyNumber))
                            {       
                                var record = new List<string>();
                                foreach (var t in titleColumn)
                                {
                                    string s = worksheet.Cells[row, t.Value].Value.ToString();
                                    record.Add(s);
                                }
                                _allMatchDataFetchFromIJFile.Add(record);
                            }
                        }
                    }                    
                }
            }

            _allTitle = titleColumn;

            //Export report
            if (_allMatchDataFetchFromIJFile.Count > 0 && !string.IsNullOrWhiteSpace(fileName) && fileName.IndexOfAny(Path.GetInvalidPathChars()) < 0)
            {
                var checkingPath = new FileInfo(fileName).DirectoryName;
                var checkingFileName = new FileInfo(fileName).Name;

                FileInfo newFile = new FileInfo(fileName);


                try
                {
                    if (newFile.Exists)
                    {
                        newFile.Delete();
                        newFile = new FileInfo(fileName);
                    }

                    if (checkingFileName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 && !string.IsNullOrWhiteSpace(checkingFileName))
                    {
                        using (ExcelPackage package = new ExcelPackage(fileName))
                        {
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                            package.Workbook.Properties.Author = "Trần Khoa Minh";
                            package.Workbook.Worksheets.Add("Báo cáo");
                            var worksheet = package.Workbook.Worksheets[0];
                            int row = 2;
                            int col = 1;

                            foreach(var title in _allTitle)
                            {
                                worksheet.Cells[1, col].Value = title.Key;
                                col++;
                            }

                            col = 1;

                            foreach (var data in _allMatchDataFetchFromIJFile)
                            {
                                foreach (var element in data)
                                {
                                    worksheet.Cells[row, col].Value = element;
                                    col++;
                                }
                                col = 1;
                                row++;
                            }
                            package.Save();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Đường dẫn không hợp lệ");
                    }
                }
                catch
                {
                    MessageBox.Show("File đang mở", "Lỗi !");
                }                
            }
        }

        #endregion
        

        public MainViewModel()
        {
            
        }
    }
}
