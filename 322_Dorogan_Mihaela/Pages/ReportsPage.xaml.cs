using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Diagnostics;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class ReportsPage : Page
    {
        private User _currentAdmin;

        public ReportsPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            InitializeReports();
        }

        private void InitializeReports()
        {
            try
            {
                // Установка дат по умолчанию
                var endDate = DateTime.Now;
                var startDate = new DateTime(endDate.Year, endDate.Month, 1);

                DpUserReportStart.SelectedDate = startDate;
                DpUserReportEnd.SelectedDate = endDate;
                DpCategoryReportStart.SelectedDate = startDate;
                DpCategoryReportEnd.SelectedDate = endDate;
                DpSummaryReportStart.SelectedDate = startDate.AddMonths(-1);
                DpSummaryReportEnd.SelectedDate = endDate;

                // Загрузка данных для комбобоксов
                using (var db = new Entities())
                {
                    var users = db.Users.OrderBy(u => u.FIO).ToList();
                    CbUserReport.ItemsSource = users;
                    CbUserReport.SelectedIndex = -1;

                    var categories = db.Category.OrderBy(c => c.Name).ToList();
                    CbCategoryReport.ItemsSource = categories;
                    CbCategoryReport.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации отчетов: {ex.Message}");
            }
        }

        private void BtnGenerateUserReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateDates(DpUserReportStart.SelectedDate, DpUserReportEnd.SelectedDate))
                    return;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    FileName = $"Отчет_по_пользователям_{DateTime.Now:yyyyMMdd_HHmmss}.docx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    GenerateUserReport(saveDialog.FileName);
                    MessageBox.Show("Отчет успешно создан!", "Успех");

                    // Открытие файла
                    Process.Start(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания отчета: {ex.Message}");
            }
        }

        private void GenerateUserReport(string filePath)
        {
            // Здесь должна быть реализация генерации Word-отчета
            // Временная заглушка с созданием простого текстового файла
            var startDate = DpUserReportStart.SelectedDate.Value;
            var endDate = DpUserReportEnd.SelectedDate.Value;
            var selectedUser = CbUserReport.SelectedItem as User;

            using (var db = new Entities())
            {
                var paymentsQuery = db.Payment
                    .Include(p => p.User)
                    .Include(p => p.Category)
                    .Where(p => p.Date >= startDate && p.Date <= endDate);

                if (selectedUser != null)
                {
                    paymentsQuery = paymentsQuery.Where(p => p.UserID == selectedUser.ID);
                }

                var payments = paymentsQuery
                    .OrderBy(p => p.User.FIO)
                    .ThenBy(p => p.Date)
                    .ToList();

                // Создание простого текстового отчета
                var reportLines = new List<string>
                {
                    $"ОТЧЕТ ПО ПОЛЬЗОВАТЕЛЯМ",
                    $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}",
                    selectedUser != null ? $"Пользователь: {selectedUser.FIO}" : "Все пользователи",
                    $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}",
                    new string('=', 50),
                    ""
                };

                var userGroups = payments.GroupBy(p => p.User.FIO);

                foreach (var userGroup in userGroups)
                {
                    reportLines.Add($"ПОЛЬЗОВАТЕЛЬ: {userGroup.Key}");
                    reportLines.Add(new string('-', 30));

                    decimal userTotal = 0;
                    foreach (var payment in userGroup)
                    {
                        var amount = payment.Num * payment.Price;
                        userTotal += amount;
                        reportLines.Add($"{payment.Date:dd.MM.yyyy} | {payment.Category.Name} | {payment.Name} | {payment.Num} шт. × {payment.Price:N2} руб. = {amount:N2} руб.");
                    }

                    reportLines.Add($"ИТОГО: {userTotal:N2} руб.");
                    reportLines.Add("");
                }

                var totalAmount = payments.Sum(p => p.Num * p.Price);
                reportLines.Add(new string('=', 50));
                reportLines.Add($"ОБЩАЯ СУММА: {totalAmount:N2} руб.");
                reportLines.Add($"КОЛИЧЕСТВО ПЛАТЕЖЕЙ: {payments.Count}");

                System.IO.File.WriteAllLines(filePath, reportLines);
            }
        }

        private void BtnExportUserExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel("user");
        }

        private void BtnGenerateCategoryReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateDates(DpCategoryReportStart.SelectedDate, DpCategoryReportEnd.SelectedDate))
                    return;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    FileName = $"Отчет_по_категориям_{DateTime.Now:yyyyMMdd_HHmmss}.docx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    GenerateCategoryReport(saveDialog.FileName);
                    MessageBox.Show("Отчет успешно создан!", "Успех");
                    Process.Start(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания отчета: {ex.Message}");
            }
        }

        private void GenerateCategoryReport(string filePath)
        {
            var startDate = DpCategoryReportStart.SelectedDate.Value;
            var endDate = DpCategoryReportEnd.SelectedDate.Value;
            var selectedCategory = CbCategoryReport.SelectedItem as Category;

            using (var db = new Entities())
            {
                var paymentsQuery = db.Payment
                    .Include(p => p.Category)
                    .Include(p => p.User)
                    .Where(p => p.Date >= startDate && p.Date <= endDate);

                if (selectedCategory != null)
                {
                    paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
                }

                var categoryData = paymentsQuery
                    .GroupBy(p => p.Category.Name)
                    .Select(g => new
                    {
                        Category = g.Key,
                        TotalAmount = g.Sum(p => p.Num * p.Price),
                        PaymentCount = g.Count(),
                        UsersCount = g.Select(p => p.UserID).Distinct().Count()
                    })
                    .OrderByDescending(x => x.TotalAmount)
                    .ToList();

                var reportLines = new List<string>
                {
                    $"ОТЧЕТ ПО КАТЕГОРИЯМ",
                    $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}",
                    selectedCategory != null ? $"Категория: {selectedCategory.Name}" : "Все категории",
                    $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}",
                    new string('=', 50),
                    ""
                };

                foreach (var category in categoryData)
                {
                    reportLines.Add($"{category.Category}");
                    reportLines.Add($"  Количество платежей: {category.PaymentCount}");
                    reportLines.Add($"  Количество пользователей: {category.UsersCount}");
                    reportLines.Add($"  Общая сумма: {category.TotalAmount:N2} руб.");
                    reportLines.Add($"  Средний платеж: {(category.TotalAmount / category.PaymentCount):N2} руб.");
                    reportLines.Add("");
                }

                var totalAmount = categoryData.Sum(c => c.TotalAmount);
                var totalPayments = categoryData.Sum(c => c.PaymentCount);

                reportLines.Add(new string('=', 50));
                reportLines.Add($"ОБЩАЯ СУММА: {totalAmount:N2} руб.");
                reportLines.Add($"ВСЕГО ПЛАТЕЖЕЙ: {totalPayments}");
                reportLines.Add($"КОЛИЧЕСТВО КАТЕГОРИЙ: {categoryData.Count}");

                System.IO.File.WriteAllLines(filePath, reportLines);
            }
        }

        private void BtnExportCategoryExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel("category");
        }

        private void BtnGenerateSummaryReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateDates(DpSummaryReportStart.SelectedDate, DpSummaryReportEnd.SelectedDate))
                    return;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    FileName = $"Сводный_отчет_{DateTime.Now:yyyyMMdd_HHmmss}.docx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    GenerateSummaryReport(saveDialog.FileName);
                    MessageBox.Show("Сводный отчет успешно создан!", "Успех");
                    Process.Start(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания отчета: {ex.Message}");
            }
        }

        private void GenerateSummaryReport(string filePath)
        {
            var startDate = DpSummaryReportStart.SelectedDate.Value;
            var endDate = DpSummaryReportEnd.SelectedDate.Value;

            using (var db = new Entities())
            {
                // Основная статистика
                var totalPayments = db.Payments.Count(p => p.Date >= startDate && p.Date <= endDate);
                var totalAmount = db.Payments
                    .Where(p => p.Date >= startDate && p.Date <= endDate)
                    .Sum(p => (decimal?)(p.Num * p.Price)) ?? 0;

                var activeUsers = db.Payments
                    .Where(p => p.Date >= startDate && p.Date <= endDate)
                    .Select(p => p.UserID)
                    .Distinct()
                    .Count();

                var categoriesUsed = db.Payments
                    .Where(p => p.Date >= startDate && p.Date <= endDate)
                    .Select(p => p.CategoryID)
                    .Distinct()
                    .Count();

                // Топ-5 пользователей
                var topUsers = db.Payments   
                    .Include(p => p.User)
                    .Where(p => p.Date >= startDate && p.Date <= endDate)
                    .GroupBy(p => p.User.FIO)
                    .Select(g => new { User = g.Key, Total = g.Sum(p => p.Num * p.Price) })
                    .OrderByDescending(x => x.Total)
                    .Take(5)
                    .ToList();

                // Топ-5 категорий
                var topCategories = db.Payments
                    .Include(p => p.Category)
                    .Where(p => p.Date >= startDate && p.Date <= endDate)
                    .GroupBy(p => p.Category.Name)
                    .Select(g => new { Category = g.Key, Total = g.Sum(p => p.Num * p.Price) })
                    .OrderByDescending(x => x.Total)
                    .Take(5)
                    .ToList();

                var reportLines = new List<string>
                {
                    $"СВОДНЫЙ ОТЧЕТ ПО СИСТЕМЕ",
                    $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}",
                    $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}",
                    new string('=', 60),
                    "",
                    "ОСНОВНЫЕ ПОКАЗАТЕЛИ:",
                    new string('-', 30),
                    $"Общее количество платежей: {totalPayments}",
                    $"Общая сумма платежей: {totalAmount:N2} руб.",
                    $"Активных пользователей: {activeUsers}",
                    $"Использованных категорий: {categoriesUsed}",
                    $"Средний платеж: {(totalPayments > 0 ? totalAmount / totalPayments : 0):N2} руб.",
                    "",
                    "ТОП-5 ПОЛЬЗОВАТЕЛЕЙ:",
                    new string('-', 30)
                };

                foreach (var user in topUsers)
                {
                    reportLines.Add($"{user.User}: {user.Total:N2} руб.");
                }

                reportLines.Add("");
                reportLines.Add("ТОП-5 КАТЕГОРИЙ:");
                reportLines.Add(new string('-', 30));

                foreach (var category in topCategories)
                {
                    reportLines.Add($"{category.Category}: {category.Total:N2} руб.");
                }

                reportLines.Add("");
                reportLines.Add(new string('=', 60));
                reportLines.Add("СИСТЕМНАЯ ИНФОРМАЦИЯ:");
                reportLines.Add(new string('-', 30));
                reportLines.Add($"Всего пользователей в системе: {db.Users.Count()}");
                reportLines.Add($"Всего категорий в системе: {db.Categories.Count()}");
                reportLines.Add($"Всего платежей в системе: {db.Payments.Count()}");

                System.IO.File.WriteAllLines(filePath, reportLines);
            }
        }

        private void BtnExportSummaryExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel("summary");
        }

        private void BtnExportAllData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv",
                    FileName = $"Все_данные_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    ExportAllDataToCsv(saveDialog.FileName);
                    MessageBox.Show("Все данные успешно экспортированы!", "Успех");
                    Process.Start(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта данных: {ex.Message}");
            }
        }

        private void ExportAllDataToCsv(string filePath)
        {
            using (var db = new Entities())
            {
                var payments = db.Payments
                    .Include(p => p.User)
                    .Include(p => p.Category)
                    .OrderBy(p => p.Date)
                    .ToList();

                var csvLines = new List<string>
                {
                    "Дата;Пользователь;Категория;Название;Количество;Цена;Сумма"
                };

                foreach (var payment in payments)
                {
                    var amount = payment.Num * payment.Price;
                    csvLines.Add(
                        $"{payment.Date:dd.MM.yyyy};" +
                        $"{payment.User.FIO};" +
                        $"{payment.Category.Name};" +
                        $"{payment.Name};" +
                        $"{payment.Num};" +
                        $"{payment.Price:N2};" +
                        $"{amount:N2}"
                    );
                }

                System.IO.File.WriteAllLines(filePath, csvLines, System.Text.Encoding.UTF8);
            }
        }

        private void BtnGenerateSystemStats_Click(object sender, RoutedEventArgs e)
        {
            GenerateSystemStatistics();
        }

        private void GenerateSystemStatistics()
        {
            try
            {
                using (var db = new Entities())
                {
                    var stats = new List<string>
                    {
                        "📊 СТАТИСТИКА СИСТЕМЫ",
                        $"Обновлено: {DateTime.Now:dd.MM.yyyy HH:mm}",
                        "",
                        $"👥 ПОЛЬЗОВАТЕЛИ:",
                        $"  • Всего пользователей: {db.Users.Count()}",
                        $"  • Администраторов: {db.Users.Count(u => u.Role == "Admin")}",
                        $"  • Обычных пользователей: {db.Users.Count(u => u.Role == "User")}",
                        "",
                        $"📂 КАТЕГОРИИ:",
                        $"  • Всего категорий: {db.Categories.Count()}",
                        "",
                        $"💰 ПЛАТЕЖИ:",
                        $"  • Всего платежей: {db.Payments.Count()}",
                        $"  • Общая сумма: {db.Payments.Sum(p => (decimal?)(p.Num * p.Price)) ?? 0:N2} руб.",
                        $"  • Средний платеж: {db.Payments.Average(p => (decimal?)(p.Num * p.Price)) ?? 0:N2} руб.",
                        "",
                        $"📈 АКТИВНОСТЬ:",
                        $"  • Первый платеж: {db.Payments.Min(p => (DateTime?)p.Date)?.ToString("dd.MM.yyyy") ?? "нет данных"}",
                        $"  • Последний платеж: {db.Payments .Max(p => (DateTime?)p.Date)?.ToString("dd.MM.yyyy") ?? "нет данных"}",
                        $"  • Платежей за месяц: {db.Payments.Count(p => p.Date >= DateTime.Now.AddMonths(-1))}",
                        $"  • Платежей за неделю: {db.Payments.Count(p => p.Date >= DateTime.Now.AddDays(-7))}"
                    };

                    TbStats.Text = string.Join(Environment.NewLine, stats);
                }
            }
            catch (Exception ex)
            {
                TbStats.Text = $"Ошибка загрузки статистики: {ex.Message}";
            }
        }

        private void ExportToExcel(string reportType)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"{reportType}_report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Здесь должна быть реализация экспорта в Excel
                    // Временная заглушка
                    MessageBox.Show($"Функция экспорта {reportType} отчета в Excel будет реализована позже",
                        "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}");
            }
        }

        private bool ValidateDates(DateTime? startDate, DateTime? endDate)
        {
            if (startDate == null || endDate == null)
            {
                MessageBox.Show("Выберите начальную и конечную даты!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (startDate > endDate)
            {
                MessageBox.Show("Дата начала не может быть больше даты окончания!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}