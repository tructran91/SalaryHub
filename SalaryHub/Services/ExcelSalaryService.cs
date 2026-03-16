using ClosedXML.Excel;
using SalaryHub.Data;
using SalaryHub.Entities;
using SalaryHub.Models;

namespace SalaryHub.Services
{
    public class ExcelSalaryService
    {
        private readonly AppDbContext _context;

        public ExcelSalaryService(AppDbContext context)
        {
            _context = context;
        }

        public ImportResult Import(Stream fileStream, int month, int year)
        {
            var result = new ImportResult();

            using var workbook = new XLWorkbook(fileStream);
            var ws = workbook.Worksheet(1);

            int headerRow = FindHeaderRow(ws);

            if (headerRow == -1)
            {
                result.Errors.Add("Không tìm thấy header 'Mã NV'");
                return result;
            }

            var columnMap = MapColumns(ws, headerRow);

            string title = FindTitle(ws, headerRow);

            if (string.IsNullOrWhiteSpace(title))
            {
                result.Errors.Add("Không tìm thấy title bảng tính");
                return result;
            }

            int lastRow = ws.LastRowUsed().RowNumber();

            int row = headerRow + 1;
            int dataStartRow = -1;

            // tìm dòng nhân viên đầu tiên
            while (row <= lastRow)
            {
                var empCode = ws.Cell(row, columnMap.EmpCode).GetString().Trim();

                if (int.TryParse(empCode, out _))
                {
                    dataStartRow = row;
                    break;
                }

                row++;
            }

            if (dataStartRow == -1)
            {
                result.Errors.Add("Không tìm thấy dữ liệu nhân viên");
                return result;
            }

            row = dataStartRow;

            var empCheck = new HashSet<string>();
            var salaryList = new List<SalaryRecord>();
            var duplicateEmployees = new HashSet<string>();

            var existingUsers = _context.Users
                .ToDictionary(x => x.EmployeeCode, x => x);

            using var transaction = _context.Database.BeginTransaction();

            try
            {
                while (row <= lastRow)
                {
                    var no = ws.Cell(row, columnMap.No).GetString().Trim();
                    var empCode = ws.Cell(row, columnMap.EmpCode).GetString().Trim();

                    // Chỉ xử lý nếu cả STT và Mã NV đều có giá trị số
                    if (!int.TryParse(no, out _) || !int.TryParse(empCode, out _))
                    {
                        row++;
                        continue;
                    }

                    var name = ws.Cell(row, columnMap.Name).GetString().Trim();

                    string department = string.Empty;
                    if (columnMap.Department > 0)
                    {
                        department = ws.Cell(row, columnMap.Department).GetString().Trim();
                    }

                    var incomeCell = ws.Cell(row, columnMap.TaxableIncome);

                    if (incomeCell.IsEmpty())
                    {
                        result.Errors.Add($"Row {row}: thiếu 'Tổng thu nhập chịu thuế'");
                        transaction.Rollback();
                        return result;
                    }

                    decimal taxableIncome = incomeCell.GetValue<decimal>();

                    decimal? pit = null;
                    decimal? bhxh = null;

                    if (columnMap.Pit != -1)
                    {
                        var cell = ws.Cell(row, columnMap.Pit);
                        if (!cell.IsEmpty())
                            pit = cell.GetValue<decimal>();
                    }

                    if (columnMap.Bhxh != -1)
                    {
                        var cell = ws.Cell(row, columnMap.Bhxh);
                        if (!cell.IsEmpty())
                            bhxh = cell.GetValue<decimal>();
                    }

                    if (empCheck.Contains(empCode))
                    {
                        var existingSalary = salaryList.First(s => s.User.EmployeeCode == empCode);
                        existingSalary.TaxableIncome += taxableIncome;
                        existingSalary.Pit = (existingSalary.Pit ?? 0) + (pit ?? 0);
                        existingSalary.Bhxh = (existingSalary.Bhxh ?? 0) + (bhxh ?? 0);

                        duplicateEmployees.Add(empCode);

                        row++;
                        continue;
                    }

                    empCheck.Add(empCode);

                    if (!existingUsers.TryGetValue(empCode, out var user))
                    {
                        user = new User
                        {
                            EmployeeCode = empCode,
                            FullName = name,
                            Department = department,
                            CreatedDate = DateTime.UtcNow
                        };

                        _context.Users.Add(user);
                        existingUsers.Add(empCode, user);
                    }

                    salaryList.Add(new SalaryRecord
                    {
                        User = user,
                        Month = month,
                        Year = year,
                        Title = title,
                        TaxableIncome = taxableIncome,
                        Pit = pit,
                        Bhxh = bhxh,
                        CreatedDate = DateTime.UtcNow
                    });

                    row++;
                }

                //var existingRecords = _context.SalaryRecords
                //    .Where(s => s.Month == month && s.Year == year && s.Title == title)
                //    .ToList();

                //if (existingRecords.Any())
                //{
                //    _context.SalaryRecords.RemoveRange(existingRecords);
                //}

                _context.SalaryRecords.AddRange(salaryList);
                _context.SaveChanges();

                transaction.Commit();

                result.Success = true;
                result.TotalInserted = salaryList.Count;

                foreach (var empCode in duplicateEmployees)
                {
                    result.Warnings.Add($"Nhân viên có mã {empCode} xuất hiện nhiều lần trong file. Các giá trị đã được gộp lại.");
                }

                return result;
            }
            catch (Exception ex)
            {
                transaction.Rollback();

                // Log to SystemLog table using a separate transaction
                try
                {
                    _context.ChangeTracker.Clear();

                    using var logTransaction = _context.Database.BeginTransaction();

                    var systemLog = new SystemLog
                    {
                        ErrorMessage = ex.Message,
                        ErrorInnerMessage = ex.InnerException?.Message,
                        StackTrace = ex.StackTrace,
                        CreatedDate = DateTime.UtcNow
                    };

                    _context.SystemLogs.Add(systemLog);
                    _context.SaveChanges();

                    logTransaction.Commit();
                }
                catch (Exception logEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to log error: {logEx.Message}");
                }

                result.Errors.Add("Có lỗi gì đó rồi, má xui vcl");
                return result;
            }
        }

        private int FindHeaderRow(IXLWorksheet ws)
        {
            foreach (var row in ws.RowsUsed())
            {
                foreach (var cell in row.CellsUsed())
                {
                    if (cell.GetString().ToLower().Trim() == "mã nv")
                        return row.RowNumber();
                }
            }

            return -1;
        }

        private ColumnMap MapColumns(IXLWorksheet ws, int headerRow)
        {
            var map = new ColumnMap();

            var header = ws.Row(headerRow);

            foreach (var cell in header.CellsUsed())
            {
                var text = cell.GetString().ToLower().Trim();

                switch (text)
                {
                    case "tt":
                    case "stt":
                        map.No = cell.Address.ColumnNumber;
                        break;

                    case "mã nv":
                        map.EmpCode = cell.Address.ColumnNumber;
                        break;

                    case "họ tên":
                    case "họ và tên":
                        map.Name = cell.Address.ColumnNumber;
                        break;

                    case "phòng":
                    case "phòng ban":
                    case "phòng/ban":
                        map.Department = cell.Address.ColumnNumber;
                        break;

                    case "tổng thu nhập chịu thuế":
                        map.TaxableIncome = cell.Address.ColumnNumber;
                        break;

                    case "trích thuế tncn":
                        map.Pit = cell.Address.ColumnNumber;
                        break;

                    case "trích bhxh":
                        map.Bhxh = cell.Address.ColumnNumber;
                        break;
                }
            }

            return map;
        }

        private string FindTitle(IXLWorksheet ws, int headerRow)
        {
            for (int row = 1; row < headerRow; row++)
            {
                var firstCell = ws.Row(row).FirstCellUsed();

                if (firstCell != null)
                {
                    var text = firstCell.GetString().Trim();

                    if (!string.IsNullOrWhiteSpace(text))
                        return text;
                }
            }

            return string.Empty;
        }
    }
}