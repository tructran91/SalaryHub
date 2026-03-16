using Microsoft.AspNetCore.Mvc;
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

        public IActionResult Guideline()
        {
            return View();
        }

        [HttpPost]
        [RequestSizeLimit(524288000)]
        [RequestFormLimits(MultipartBodyLengthLimit = 524288000)]
        public IActionResult Upload(int month, int year, List<IFormFile> files)
        {
            var errors = new List<object>();
            var warnings = new List<object>();
            int totalInserted = 0;

            if (files == null || files.Count == 0)
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

            foreach (var file in files)
            {
                using var stream = file.OpenReadStream();

                var result = _excelService.Import(stream, month, year);

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
