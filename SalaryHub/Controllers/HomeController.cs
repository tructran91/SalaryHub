using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using SalaryHub.Helpers;
using SalaryHub.Models;
using SalaryHub.Services;
using System.Diagnostics;

namespace SalaryHub.Controllers
{
    public class HomeController : Controller
    {
        private readonly ExcelSalaryService _excelService;

        public HomeController(ExcelSalaryService excelService)
        {
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult MonthlyExport()
        {
            return View();
        }

        public IActionResult CumulativeExport()
        {
            return View();
        }

        [HttpGet]
        public async Task<ActionResult> GetCumulativeReport(List<int> months, int year)
        {
            var data = await _excelService.GetCumulativeReport(months, year);

            return PartialView("_PayrollReportRows", data);
        }

        public IActionResult Guideline()
        {
            return View();
        }

        [HttpPost]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        public IActionResult Upload(int month, int year, List<IFormFile> monthlyFiles, List<IFormFile> onceTimeFiles)
        {
            var errors = new List<object>();
            var warnings = new List<object>();
            int totalInserted = 0;

            var allFiles = new List<(IFormFile file, bool isMonthly)>();

            if (monthlyFiles != null && monthlyFiles.Count > 0)
            {
                foreach (var file in monthlyFiles)
                {
                    allFiles.Add((file, true));
                }
            }

            if (onceTimeFiles != null && onceTimeFiles.Count > 0)
            {
                foreach (var file in onceTimeFiles)
                {
                    allFiles.Add((file, false));
                }
            }

            if (allFiles.Count == 0)
            {
                errors.Add(new
                {
                    file = "",
                    message = "Không có file hoặc file bị lỗi"
                });

                return Ok(new
                {
                    success = false,
                    errors
                });
            }

            foreach (var (file, isMonthly) in allFiles)
            {
                using var stream = file.OpenReadStream();

                var result = _excelService.Import(stream, month, year, isMonthly);

                if (!result.Success)
                {
                    foreach (var err in result.Errors)
                    {
                        errors.Add(new
                        {
                            file = file.FileName,
                            message = err
                        });
                    }
                }
                else
                {
                    totalInserted += result.TotalInserted;

                    foreach (var warn in result.Warnings)
                    {
                        warnings.Add(new
                        {
                            file = file.FileName,
                            message = warn
                        });
                    }
                }
            }

            if (errors.Any())
            {
                return Ok(new
                {
                    success = false,
                    errors
                });
            }

            return Ok(new
            {
                success = true,
                totalInserted,
                warnings
            });
        }

        [HttpGet]
        public async Task<IActionResult> GetReport(int month, int year)
        {
            var data = await _excelService.GetMonthlySalaryReport(month, year);

            return PartialView("_PayrollReportRows", data);
        }

        [HttpGet]
        public async Task<IActionResult> ExportExcelMonthly(int month, int year)
        {
            var vm = await _excelService.GetMonthlySalaryReport(month, year);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Bảng lương");

            int totalCols = 4 + vm.IncomeTitles.Count + 1 + vm.PitTitles.Count + 3;

            // ===== ROW 1-3: TRỐNG =====
            // ===== ROW 4: TITLE =====
            ws.Range(4, 1, 4, totalCols).Merge();
            ws.Cell(4, 1).Value = $"TỔNG HỢP THUẾ TNCN THÁNG {month:D2} NĂM {year}";
            ws.Cell(4, 1).Style.Font.Bold = true;
            ws.Cell(4, 1).Style.Font.FontSize = 16;
            ws.Cell(4, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(4, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(4).Height = 30;

            // ===== ROW 5: TRỐNG =====

            // ===== ROW 6: GROUP HEADER (THU NHẬP / THUẾ TNCN) =====
            int col = 1;

            // STT, Mã NV, Họ tên, Phòng ban — rowspan 2 => merge row 6 & 7
            foreach (var label in new[] { "STT", "Mã NV", "HỌ VÀ TÊN", "Phòng ban" })
            {
                ws.Range(6, col, 7, col).Merge();
                ws.Cell(6, col).Value = label;
                StyleGroupCell(ws.Cell(6, col), XLColor.FromHtml("#F4CCCC")); // hồng nhạt
                col++;
            }

            // Group THU NHẬP
            int incomeStartCol = col;
            int incomeColCount = vm.IncomeTitles.Count + 1; // +1 cho "Tổng thu nhập"
            ws.Range(6, incomeStartCol, 6, incomeStartCol + incomeColCount - 1).Merge();
            ws.Cell(6, incomeStartCol).Value = $"THU NHẬP  T{month:D2}/{year}";
            StyleGroupHeader(ws.Cell(6, incomeStartCol), XLColor.FromHtml("#D9EAD3")); // xanh lá nhạt

            // Group THUẾ TNCN
            int pitStartCol = incomeStartCol + incomeColCount;
            int pitColCount = vm.PitTitles.Count;
            ws.Range(6, pitStartCol, 6, pitStartCol + pitColCount - 1).Merge();
            ws.Cell(6, pitStartCol).Value = $"THUẾ TNCN T{month:D2}/{year}";
            StyleGroupHeader(ws.Cell(6, pitStartCol), XLColor.FromHtml("#D9EAD3"));

            // BHXH, Thuế tạm trích, Thuế phải nộp — merge row 6 & 7
            int afterPitCol = pitStartCol + pitColCount;
            foreach (var label in new[] { $"`BHXH\nT{month:D2}/{year}", $"Thuế TNCN\nT{month:D2}/{year} đã\ntạm trích", $"Thuế TNCN\nphải nộp\nT{month:D2}/{year}" })
            {
                ws.Range(6, afterPitCol, 7, afterPitCol).Merge();
                ws.Cell(6, afterPitCol).Value = label;
                StyleGroupCell(ws.Cell(6, afterPitCol), XLColor.FromHtml("#F4CCCC"));
                afterPitCol++;
            }

            // ===== ROW 7: SUB HEADER =====
            col = incomeStartCol;
            foreach (var title in vm.IncomeTitles)
            {
                var displayTitle = TitleGroupHelper.IsGroupSummary(title) ? "Tổng nhóm" : title;
                ws.Cell(7, col).Value = displayTitle;
                StyleSubHeader(ws.Cell(7, col), TitleGroupHelper.GetExcelColor(title));
                if (TitleGroupHelper.IsGroupSummary(title)) ws.Cell(7, col).Style.Font.Italic = true;
                col++;
            }
            ws.Cell(7, col).Value = $"Tổng thu nhập\nchịu thuế\nT{month:D2}/{year}";
            StyleSubHeader(ws.Cell(7, col));
            col++;

            foreach (var title in vm.PitTitles)
            {
                var displayTitle = TitleGroupHelper.IsGroupSummary(title) ? "Tổng nhóm" : title;
                ws.Cell(7, col).Value = displayTitle;
                StyleSubHeader(ws.Cell(7, col), TitleGroupHelper.GetExcelColor(title));
                if (TitleGroupHelper.IsGroupSummary(title)) ws.Cell(7, col).Style.Font.Italic = true;
                col++;
            }

            // Style chung cho toàn bộ header row 6-7
            ws.Range(6, 1, 7, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(6, 1, 7, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(6, 1, 7, totalCols).Style.Alignment.WrapText = true;
            ws.Range(6, 1, 7, totalCols).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(6, 1, 7, totalCols).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(6).Height = 20;
            ws.Row(7).Height = 60;

            // ===== DATA ROWS bắt đầu từ ROW 8 =====
            for (int i = 0; i < vm.Rows.Count; i++)
            {
                var row = vm.Rows[i];
                bool isTotal = row.FullName == "TOTAL";
                int r = i + 8;
                col = 1;

                decimal totalIncome = 0;
                decimal totalPit = 0;

                ws.Cell(r, col++).Value = isTotal ? "" : (i + 1).ToString();
                ws.Cell(r, col++).Value = row.EmployeeCode;
                ws.Cell(r, col++).Value = row.FullName;
                ws.Cell(r, col++).Value = row.Department;

                foreach (var title in vm.IncomeTitles)
                {
                    decimal value = row.Incomes.ContainsKey(title) ? row.Incomes[title] : 0;
                    if (!TitleGroupHelper.ShouldExcludeFromTotal(title))
                    {
                        totalIncome += value;
                    }
                    var dataCell = ws.Cell(r, col++);
                    dataCell.Value = value;
                    if (!isTotal)
                    {
                        dataCell.Style.Fill.BackgroundColor = XLColor.FromHtml(TitleGroupHelper.GetExcelColor(title));
                        if (TitleGroupHelper.IsGroupSummary(title)) dataCell.Style.Font.Bold = true;
                    }
                }
                ws.Cell(r, col++).Value = totalIncome;

                foreach (var title in vm.PitTitles)
                {
                    decimal value = row.Pits.ContainsKey(title) ? row.Pits[title] : 0;
                    if (!TitleGroupHelper.ShouldExcludeFromTotal(title))
                    {
                        totalPit += value;
                    }
                    var dataCell = ws.Cell(r, col++);
                    dataCell.Value = value;
                    if (!isTotal)
                    {
                        dataCell.Style.Fill.BackgroundColor = XLColor.FromHtml(TitleGroupHelper.GetExcelColor(title));
                        if (TitleGroupHelper.IsGroupSummary(title)) dataCell.Style.Font.Bold = true;
                    }
                }

                ws.Cell(r, col++).Value = row.Bhxh;
                ws.Cell(r, col++).Value = totalPit;
                ws.Cell(r, col).Value = totalPit;

                var dataRow = ws.Range(r, 1, r, totalCols);
                dataRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Cột text (STT, Mã NV, Phòng ban) align center
                ws.Range(r, 1, r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // STT
                ws.Range(r, 2, r, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Mã NV
                ws.Range(r, 3, r, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(r, 4, r, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Phòng ban

                // Cột số (thu nhập, thuế, BHXH) align right
                ws.Range(r, incomeStartCol, r, totalCols).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                if (isTotal)
                {
                    dataRow.Style.Font.Bold = true;
                    dataRow.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3CD");
                }
            }

            // ===== FORMAT SỐ =====
            int dataRowCount = vm.Rows.Count + 8;
            ws.Range(8, incomeStartCol, dataRowCount, totalCols)
              .Style.NumberFormat.Format = "#,##0";

            // Align right cho các cột số
            ws.Range(8, incomeStartCol, dataRowCount, totalCols)
              .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            // ===== AUTO FIT =====
            // Cột thông tin cố định
            ws.Column(1).Width = 6;   // STT
            ws.Column(2).Width = 10;  // Mã NV
            ws.Column(3).Width = 25;  // Họ tên
            ws.Column(4).Width = 20;  // Phòng ban

            // Cột số tiền — tính theo vị trí động
            int numberColStart = incomeStartCol;
            int numberColEnd = totalCols;
            for (int c = numberColStart; c <= numberColEnd; c++)
            {
                ws.Column(c).Width = 15; // đủ hiển thị ~10 chữ số "999,999,999"
            }

            // ===== EXPORT =====
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            string fileName = $"Tong hop thue TNCN Thang {month:D2}_{year}.xlsx";
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportCumulativeExcel(List<int> months, int year)
        {
            var vm = await _excelService.GetCumulativeReport(months, year);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Báo cáo lũy kế");

            int totalCols = 4 + vm.IncomeTitles.Count + 1 + vm.PitTitles.Count + 3;

            var sortedMonths = months.OrderBy(m => m).ToList();
            var monthsDisplay = string.Join(", ", sortedMonths.Select(m => $"T{m:D2}"));

            ws.Range(4, 1, 4, totalCols).Merge();
            ws.Cell(4, 1).Value = $"BÁO CÁO LŨY KẾ THUẾ TNCN ({monthsDisplay}) NĂM {year}";
            ws.Cell(4, 1).Style.Font.Bold = true;
            ws.Cell(4, 1).Style.Font.FontSize = 16;
            ws.Cell(4, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(4, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(4).Height = 30;

            int col = 1;

            foreach (var label in new[] { "STT", "Mã NV", "HỌ VÀ TÊN", "Phòng ban" })
            {
                ws.Range(6, col, 7, col).Merge();
                ws.Cell(6, col).Value = label;
                StyleGroupCell(ws.Cell(6, col), XLColor.FromHtml("#F4CCCC"));
                col++;
            }

            int incomeStartCol = col;
            int incomeColCount = vm.IncomeTitles.Count + 1;
            ws.Range(6, incomeStartCol, 6, incomeStartCol + incomeColCount - 1).Merge();
            ws.Cell(6, incomeStartCol).Value = $"THU NHẬP ({monthsDisplay})";
            StyleGroupHeader(ws.Cell(6, incomeStartCol), XLColor.FromHtml("#D9EAD3"));

            int pitStartCol = incomeStartCol + incomeColCount;
            int pitColCount = vm.PitTitles.Count;
            ws.Range(6, pitStartCol, 6, pitStartCol + pitColCount - 1).Merge();
            ws.Cell(6, pitStartCol).Value = $"THUẾ TNCN ({monthsDisplay})";
            StyleGroupHeader(ws.Cell(6, pitStartCol), XLColor.FromHtml("#D9EAD3"));

            int afterPitCol = pitStartCol + pitColCount;
            foreach (var label in new[] { $"BHXH\n({monthsDisplay})", $"Thuế TNCN\n({monthsDisplay}) đã\ntạm trích", $"Thuế TNCN\nphải nộp\n({monthsDisplay})" })
            {
                ws.Range(6, afterPitCol, 7, afterPitCol).Merge();
                ws.Cell(6, afterPitCol).Value = label;
                StyleGroupCell(ws.Cell(6, afterPitCol), XLColor.FromHtml("#F4CCCC"));
                afterPitCol++;
            }

            col = incomeStartCol;
            foreach (var title in vm.IncomeTitles)
            {
                var displayTitle = TitleGroupHelper.IsGroupSummary(title) ? "Tổng nhóm" : title;
                ws.Cell(7, col).Value = displayTitle;
                StyleSubHeader(ws.Cell(7, col), TitleGroupHelper.GetExcelColor(title));
                if (TitleGroupHelper.IsGroupSummary(title)) ws.Cell(7, col).Style.Font.Italic = true;
                col++;
            }
            ws.Cell(7, col).Value = $"Tổng thu nhập\nchịu thuế\n({monthsDisplay}/{year})";
            StyleSubHeader(ws.Cell(7, col));
            col++;

            foreach (var title in vm.PitTitles)
            {
                var displayTitle = TitleGroupHelper.IsGroupSummary(title) ? "Tổng nhóm" : title;
                ws.Cell(7, col).Value = displayTitle;
                StyleSubHeader(ws.Cell(7, col), TitleGroupHelper.GetExcelColor(title));
                if (TitleGroupHelper.IsGroupSummary(title)) ws.Cell(7, col).Style.Font.Italic = true;
                col++;
            }

            ws.Range(6, 1, 7, totalCols).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(6, 1, 7, totalCols).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range(6, 1, 7, totalCols).Style.Alignment.WrapText = true;
            ws.Range(6, 1, 7, totalCols).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(6, 1, 7, totalCols).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(6).Height = 20;
            ws.Row(7).Height = 60;

            for (int i = 0; i < vm.Rows.Count; i++)
            {
                var row = vm.Rows[i];
                bool isTotal = row.FullName == "TOTAL";
                int r = i + 8;
                col = 1;

                decimal totalIncome = 0;
                decimal totalPit = 0;

                ws.Cell(r, col++).Value = isTotal ? "" : (i + 1).ToString();
                ws.Cell(r, col++).Value = row.EmployeeCode;
                ws.Cell(r, col++).Value = row.FullName;
                ws.Cell(r, col++).Value = row.Department;

                foreach (var title in vm.IncomeTitles)
                {
                    decimal value = row.Incomes.ContainsKey(title) ? row.Incomes[title] : 0;
                    if (!TitleGroupHelper.ShouldExcludeFromTotal(title))
                    {
                        totalIncome += value;
                    }
                    var dataCell = ws.Cell(r, col++);
                    dataCell.Value = value;
                    if (!isTotal)
                    {
                        dataCell.Style.Fill.BackgroundColor = XLColor.FromHtml(TitleGroupHelper.GetExcelColor(title));
                        if (TitleGroupHelper.IsGroupSummary(title)) dataCell.Style.Font.Bold = true;
                    }
                }
                ws.Cell(r, col++).Value = totalIncome;

                foreach (var title in vm.PitTitles)
                {
                    decimal value = row.Pits.ContainsKey(title) ? row.Pits[title] : 0;
                    if (!TitleGroupHelper.ShouldExcludeFromTotal(title))
                    {
                        totalPit += value;
                    }
                    var dataCell = ws.Cell(r, col++);
                    dataCell.Value = value;
                    if (!isTotal)
                    {
                        dataCell.Style.Fill.BackgroundColor = XLColor.FromHtml(TitleGroupHelper.GetExcelColor(title));
                        if (TitleGroupHelper.IsGroupSummary(title)) dataCell.Style.Font.Bold = true;
                    }
                }

                ws.Cell(r, col++).Value = row.Bhxh;
                ws.Cell(r, col++).Value = totalPit;
                ws.Cell(r, col).Value = totalPit;

                var dataRow = ws.Range(r, 1, r, totalCols);
                dataRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                ws.Range(r, 1, r, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(r, 2, r, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(r, 4, r, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(r, incomeStartCol, r, totalCols).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                if (isTotal)
                {
                    dataRow.Style.Font.Bold = true;
                    dataRow.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3CD");
                }
            }

            int dataRowCount = vm.Rows.Count + 8;
            ws.Range(8, incomeStartCol, dataRowCount, totalCols)
              .Style.NumberFormat.Format = "#,##0";
            ws.Range(8, incomeStartCol, dataRowCount, totalCols)
              .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            ws.Column(1).Width = 6;
            ws.Column(2).Width = 10;
            ws.Column(3).Width = 25;
            ws.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            ws.Column(4).Width = 20;

            for (int c = incomeStartCol; c <= totalCols; c++)
            {
                ws.Column(c).Width = 15;
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            string fileName = $"Bao cao luy ke Thue TNCN ({monthsDisplay})_{year}.xlsx";
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ===== HELPER METHODS =====
        private void StyleGroupHeader(IXLCell cell, XLColor bgColor)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = bgColor;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private void StyleGroupCell(IXLCell cell, XLColor bgColor)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = bgColor;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
        }

        private void StyleSubHeader(IXLCell cell, string hexColor = "#D9EAD3")
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.Red;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml(hexColor);
            cell.Style.Alignment.WrapText = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
