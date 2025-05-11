// Services/ExcelService.cs
using System.IO;
using DIMS.Models;
using OfficeOpenXml;

namespace DIMS.Services
{
    public class ExcelService
    {
        private readonly ILogger<ExcelService> _logger;
        private readonly string _templatePath;

        public ExcelService(ILogger<ExcelService> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            _templatePath = Path.Combine(env.ContentRootPath, "Resources", "Templates", "DIMS.xlsm");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<byte[]> GenerateExcelReport(RedmineProject project)
        {
            _logger.LogInformation($"Генерация Excel-отчета для проекта: {project.Name}");

            // Проверка на существование файла шаблона
            if (!File.Exists(_templatePath))
            {
                _logger.LogError($"Шаблон Excel не найден по пути: {_templatePath}");
                throw new FileNotFoundException($"Шаблон не найден", _templatePath);
            }

            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);

            var worksheet = package.Workbook.Worksheets[0]; // Первый лист

            int currentRow = 3;

            // Заполняем данные о задачах и проекте
            if (project.Issues != null && project.Issues.Any())
            {
                foreach (var issue in project.Issues)
                {
                    // TODO: Данные проекта (повторяются для каждой задачи)
                    worksheet.Cells[currentRow, 1].Value = project.Identifier;
                    worksheet.Cells[currentRow, 2].Value = project.Name;
                    if (project.Parent != null)
                    {
                        worksheet.Cells[currentRow, 3].Value = project.Parent.Id;
                    }
                    worksheet.Cells[currentRow, 4].Value = project.GetCustomFieldValue(43);
                    worksheet.Cells[currentRow, 5].Value = project.GetCustomFieldValue(44);
                    worksheet.Cells[currentRow, 6].Value = project.GetCustomFieldValue(38);
                    worksheet.Cells[currentRow, 7].Value = project.GetCustomFieldValue(45);
                    worksheet.Cells[currentRow, 8].Value = project.GetCustomFieldValue(39);
                    worksheet.Cells[currentRow, 9].Value = project.GetCustomFieldValue(35);
                    worksheet.Cells[currentRow, 10].Value = project.GetCustomFieldValue(36);
                    worksheet.Cells[currentRow, 11].Value = project.GetCustomFieldValue(40);
                    worksheet.Cells[currentRow, 12].Value = project.GetCustomFieldValue(41);
                    worksheet.Cells[currentRow, 13].Value = project.GetCustomFieldValue(42);

                    worksheet.Cells[currentRow, 14].Value = issue.Subject;
                    worksheet.Cells[currentRow, 15].Value = $"{issue.Tracker.Id}_{issue.Tracker.Name}";
                    worksheet.Cells[currentRow, 16].Value = issue.Subject;
                    worksheet.Cells[currentRow, 17].Value = $"{issue.Tracker.Id}_{issue.Tracker.Name}";
                    worksheet.Cells[currentRow, 18].Value = issue.Description;
                    worksheet.Cells[currentRow, 19].Value = $"{issue.AssignedTo.Id}_{issue.AssignedTo.Name}";
                    worksheet.Cells[currentRow, 20].Value = issue.StartDate;
                    worksheet.Cells[currentRow, 21].Value = issue.DueDate;
                    worksheet.Cells[currentRow, 22].Value = issue.EstimatedHours;
                    worksheet.Cells[currentRow, 23].Value = issue.GetCustomFieldValue(20);
                    worksheet.Cells[currentRow, 24].Value = issue.GetCustomFieldValue(49);
                    worksheet.Cells[currentRow, 25].Value = issue.GetCustomFieldValue(37);
                    worksheet.Cells[currentRow, 26].Value = issue.GetCustomFieldValue(47);
                    worksheet.Cells[currentRow, 27].Value = issue.GetCustomFieldValue(48);
                    worksheet.Cells[currentRow, 28].Value = issue.GetCustomFieldValue(46);
                    
                    currentRow++;
                }
            }
            
            // Сохраняем результат в byte[]
            return await package.GetAsByteArrayAsync();
        }
    }
}
