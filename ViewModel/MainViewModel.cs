using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;

namespace TransformInsureJToMyReport.ViewModel
{
    internal partial class MainViewModel : ObservableObject
    {
        private Dictionary<string, int> _allTitle = new Dictionary<string, int>();
        private List<List<string>> _allMatchDataFetchFromIJFile = new();

        private List<List<string>> CustomListConverted(List<List<string>> list)
        {
            foreach (var row in list)
            {
                for (int i = 0; i < row.Count; i++)
                {
                    switch (row[i].ToString())
                    {
                        case "Bảo hiểm cháy nổ bắt buộc":
                            row[i] = "BH Cháy nổ bắt buộc (CNBB) + BH cháy và các RR đặc biệt (FI)";
                            break;

                        case "Bảo hiểm Cháy và rủi ro đặc biệt":
                            row[i] = "BH Cháy nổ bắt buộc (CNBB) + BH cháy và các RR đặc biệt (FI)";
                            break;

                        case "Bảo hiểm máy móc và thiết bị xây dựng":
                            row[i] = "BH Máy móc và thiết bị xây dựng (CPM)";
                            break;

                        case "Bảo hiểm Đổ vỡ máy móc":
                            row[i] = "BH Hư hỏng máy móc (MB)";
                            break;

                        case "Bảo hiểm công trình dân dụng hoàn thành":
                            row[i] = "BH Công trình xây dựng đã hoàn thành (CECR)";
                            break;

                        case "BH trách nhiệm công cộng":
                            row[i] = "BH Trách nhiệm công cộng (PBL)";
                            break;

                        case "Bảo hiểm mọi rủi ro tài sản":
                            row[i] = "BH Mọi rủi ro tài sản (PAR)";
                            break;

                        case "VND":
                            row[i] = "VNĐ";
                            break;

                        case "HP Phòng Bảo hiểm Cháy - Kỹ thuật":
                            row[i] = "P. SỐ 7";
                            break;

                        case "HP Phòng Bảo hiểm Tàu thủy":
                            row[i] = "P. TÀU THỦY";
                            break;

                        case "HP Phòng Bảo hiểm Hàng hóa":
                            row[i] = "P. HÀNG HÓA";
                            break;

                        case "HP Phòng Bảo hiểm số 1":
                            row[i] = "P. SỐ 1";
                            break;

                        case "HP Phòng Bảo hiểm số 2":
                            row[i] = "P. SỐ 2";
                            break;

                        case "HP Phòng Bảo hiểm số 3":
                            row[i] = "P. SỐ 3";
                            break;

                        case "HP Phòng Bảo hiểm số 4":
                            row[i] = "P. SỐ 4";
                            break;

                        case "HP Phòng Bảo hiểm số 5":
                            row[i] = "P. SỐ 5";
                            break;

                        case "HP Phòng Bảo hiểm số 6":
                            row[i] = "P. SỐ 6";
                            break;

                        case "HP Phòng phát triển kênh phân phối (cũ)":
                            row[i] = "P. SỐ 6";
                            break;

                        case "HP Phòng Bảo hiểm số 7":
                            row[i] = "P. SỐ 7";
                            break;

                        case "HP Phòng bảo hiểm số 8":
                            row[i] = "P. SỐ 8";
                            break;

                        case "HP Phòng Bảo hiểm số 9":
                            row[i] = "P. SỐ 9";
                            break;

                        case "HP Phòng Bảo hiểm số 10":
                            row[i] = "P. SỐ 10";
                            break;

                        default:
                            break;
                    }
                }
            }

            return list;
        }
        private int GetUsefulCategory(List<string> strings, List<string> substrings)
        {
            foreach (var item in strings)
            {
                if (substrings.Any(s => item.Contains(s, StringComparison.OrdinalIgnoreCase)))
                {
                    return strings.IndexOf(item) + 1;
                }
            }
            return 1;
        }

        [ObservableProperty]
        private SnackbarMessageQueue _message = new SnackbarMessageQueue(TimeSpan.FromSeconds(2));

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

        private Dictionary<string, int> CustomTitleColumn()
        {
            return new Dictionary<string, int>
            {
                ["Ngày tạo đơn"] = 11,
                ["PKD"] = 2,
                ["Sản phẩm bảo hiểm"] = 5,
                ["Khách hàng"] = 9,
                ["Trung gian"] = 10,
                ["Số tiền bảo hiểm"] = 15,
                ["Số đơn"] = 7,
                ["Đồng BH"] = 24,
                ["Ngày cấp đơn"] = 13,
                ["Hiệu lực từ"] = 12,
                ["Hiệu lực đến"] = 14,
                ["Loại tiền"] = 16,
                ["ST phải trả"] = 21,
                ["Hạn thanh toán"] = 34,
                ["Ngày ký"] = 13,
                ["Người nhận"] = 4,

            };
        }

        private bool CanExportReport()
            => InsureJFiles.Any() && IJNotInReport != null && IJNotInReport.Any();

        [RelayCommand(CanExecute = nameof(CanExportReport))]
        private async Task ExportReport()
        {            
            var titleColumn = new Dictionary<string, int>();
            var tempIJNotInReport = new List<string>(IJNotInReport);

            string? fileName = "";
            var dialog = new SaveFileDialog();
            dialog.Filter = "Workbook (*.xlsx)|*.xlsx";
            dialog.Title = "Xuất báo cáo";

            if (dialog.ShowDialog() == true)
            {
                fileName = dialog.FileName;                
            }

            await Task.Run(() =>
            {
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
                            List<string> indicators = new List<string>();

                            for (int col = 1; col < colCount; col++)
                            {
                                if (worksheet.Cells[4, col].Value != null)
                                {
                                    indicators.Add(worksheet.Cells[4, col].Value.ToString());
                                }
                            }

                            titleColumn = CustomTitleColumn();

                            //Pull data to _allMatchDataFetchFromIJFile
                            for (int row = 5; row < rowCount; row++)
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
                                    tempIJNotInReport.Remove(policyNumber);
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

                    if (newFile.Exists)
                    {
                        try
                        {
                            newFile.Delete();
                        }
                        catch
                        {
                            Message.Enqueue("File đang mở");
                            return;
                        }
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

                            foreach (var title in _allTitle)
                            {
                                worksheet.Cells[1, col].Value = title.Key.ToUpper();
                                col++;
                            }

                            //Format style title
                            worksheet.Row(1).Height = 25;
                            col--;
                            var titleRange = worksheet.Cells[1, 1, 1, col];
                            titleRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            titleRange.Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                            titleRange.Style.Font.Bold = true;
                            titleRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            titleRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            col = 1;

                            //Load data from List to range
                            var customListConverted = CustomListConverted(_allMatchDataFetchFromIJFile);
                            foreach (var data in customListConverted)
                            {
                                foreach (var element in data)
                                {
                                    worksheet.Cells[row, col].Value = element;
                                    col++;
                                }
                                col = 1;
                                row++;
                            }

                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                            package.Save();

                            Message.Enqueue("Xuất báo cáo thành công");
                        }
                    }
                    else
                    {
                        Message.Enqueue("Đường dẫn không hợp lệ");
                    }
                }

            }
            );

            IJNotInReport = tempIJNotInReport;
        }
        #endregion
        

        public MainViewModel()
        {
            
        }
    }
}
