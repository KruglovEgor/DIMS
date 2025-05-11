using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using DIMS.Models;
using System.Text.Json.Serialization;
using DIMS.Services;

namespace DIMS.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DIMSController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly HttpClient _httpClient;
        private readonly ILogger<DIMSController> _logger;
        private readonly JsonSerializerOptions _jsonOptions;
        private readonly ExcelService _excelService;

        public DIMSController(
            IConfiguration configuration,
            IHttpClientFactory httpClientFactory,
            ILogger<DIMSController> logger,
            ExcelService excelService)
        {
            _configuration = configuration;
            _httpClient = httpClientFactory.CreateClient("RedmineClient");
            _logger = logger;
            _excelService = excelService;

            // Настройка базового HttpClient для Redmine
            var redmineUrl = _configuration["RedmineSettings:TasksUrl"];
            var apiKey = _configuration["RedmineSettings:TasksApiKey"];

            _httpClient.BaseAddress = new Uri(redmineUrl ?? throw new ArgumentNullException("RedmineSettings:TasksUrl не указан в конфигурации"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($"{apiKey}:")));

            _jsonOptions = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            };
        }

        [HttpGet("project/{projectName}/excel")]
        public async Task<IActionResult> GetProjectExcel(string projectName)
        {
            try
            {
                _logger.LogInformation($"Запрос Excel-отчета для проекта: {projectName}");

                // 1. Получаем информацию о проекте с custom fields
                var projectResponse = await _httpClient.GetAsync($"projects/{projectName}.json?include=custom_fields");

                if (!projectResponse.IsSuccessStatusCode)
                {
                    _logger.LogWarning($"Не удалось получить информацию о проекте {projectName}. Статус: {projectResponse.StatusCode}");
                    return StatusCode((int)projectResponse.StatusCode, $"Ошибка Redmine API: {projectResponse.ReasonPhrase}");
                }

                var projectContent = await projectResponse.Content.ReadAsStringAsync();
                _logger.LogDebug($"Ответ от Redmine API (проект): {projectContent}");

                var projectData = JsonSerializer.Deserialize<RedmineProjectResponse>(projectContent, _jsonOptions);

                if (projectData?.Project == null)
                {
                    return NotFound($"Проект '{projectName}' не найден или данные имеют неверный формат");
                }

                
                // 2. Получаем задачи проекта с включением необходимых полей
                var issuesResponse = await _httpClient.GetAsync(
                    $"projects/{projectName}/issues.json?status_id=*&include=custom_fields");

                if (issuesResponse.IsSuccessStatusCode)
                {
                    var issuesContent = await issuesResponse.Content.ReadAsStringAsync();
                    _logger.LogDebug($"Ответ от Redmine API (задачи): {issuesContent}");

                    var issuesData = JsonSerializer.Deserialize<RedmineIssuesResponse>(issuesContent, _jsonOptions);

                    if (issuesData?.Issues != null)
                    {
                        projectData.Project.Issues = issuesData.Issues;
                    }
                }
                else
                {
                    _logger.LogWarning($"Не удалось получить задачи для проекта {projectName}. Статус: {issuesResponse.StatusCode}");
                }

                // 3. Генерируем Excel-отчет
                var excelBytes = await _excelService.GenerateExcelReport(projectData.Project);

                // 4. Возвращаем файл пользователю
                return File(
                    excelBytes,
                    "application/vnd.ms-excel.sheet.macroEnabled.12",
                    $"DIMS-{projectName}-{DateTime.Now:yyyy-MM-dd}.xlsm"
                );
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogError(ex, "Шаблон Excel не найден");
                return StatusCode(500, "Шаблон для генерации Excel-отчета не найден");
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, $"Ошибка десериализации JSON: {ex.Message}");
                return StatusCode(500, $"Ошибка обработки данных Redmine: {ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Ошибка при генерации Excel-отчета для проекта {projectName}");
                return StatusCode(500, "Ошибка при генерации Excel-отчета");
            }
        }
    }
}
