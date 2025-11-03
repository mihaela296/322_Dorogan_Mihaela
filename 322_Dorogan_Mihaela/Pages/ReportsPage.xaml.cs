using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

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

                    var categories = db.Categories.OrderBy(c => c.Name).ToList();
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
            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                wordApp.Visible = false;

                var startDate = DpUserReportStart.SelectedDate.Value;
                var endDate = DpUserReportEnd.SelectedDate.Value;
                var selectedUser = CbUserReport.SelectedItem as User;

                using (var db = new Entities())
                {
                    var paymentsQuery = db.Payments
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

                    // Заголовок отчета
                    Word.Paragraph title = wordDoc.Paragraphs.Add();
                    title.Range.Text = "ОТЧЕТ ПО ПОЛЬЗОВАТЕЛЯМ";
                    title.Range.Font.Bold = 1;
                    title.Range.Font.Size = 16;
                    title.Range.InsertParagraphAfter();

                    // Информация о периоде
                    Word.Paragraph info = wordDoc.Paragraphs.Add();
                    info.Range.Text = $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                    if (selectedUser != null)
                        info.Range.Text += $"\nПользователь: {selectedUser.FIO}";
                    info.Range.Text += $"\nСгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";
                    info.Range.InsertParagraphAfter();

                    // Пустая строка
                    wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                    var userGroups = payments.GroupBy(p => p.User.FIO);

                    foreach (var userGroup in userGroups)
                    {
                        // Заголовок пользователя
                        Word.Paragraph userHeader = wordDoc.Paragraphs.Add();
                        userHeader.Range.Text = $"ПОЛЬЗОВАТЕЛЬ: {userGroup.Key}";
                        userHeader.Range.Font.Bold = 1;
                        userHeader.Range.Font.Size = 12;
                        userHeader.Range.InsertParagraphAfter();

                        // Создание таблицы для платежей пользователя
                        if (userGroup.Any())
                        {
                            Word.Table table = wordDoc.Tables.Add(
                                userHeader.Range,
                                userGroup.Count() + 1,
                                5);

                            table.Borders.Enable = 1;
                            table.Rows[1].Range.Font.Bold = 1;

                            // Заголовки таблицы
                            table.Cell(1, 1).Range.Text = "Дата";
                            table.Cell(1, 2).Range.Text = "Категория";
                            table.Cell(1, 3).Range.Text = "Название";
                            table.Cell(1, 4).Range.Text = "Количество";
                            table.Cell(1, 5).Range.Text = "Сумма";

                            int row = 2;
                            decimal userTotal = 0;

                            foreach (var payment in userGroup)
                            {
                                var amount = payment.Num * payment.Price;
                                userTotal += amount;

                                table.Cell(row, 1).Range.Text = payment.Date.ToString("dd.MM.yyyy");
                                table.Cell(row, 2).Range.Text = payment.Category.Name;
                                table.Cell(row, 3).Range.Text = payment.Name;
                                table.Cell(row, 4).Range.Text = payment.Num.ToString();
                                table.Cell(row, 5).Range.Text = amount.ToString("N2") + " руб.";
                                row++;
                            }

                            // Итоговая строка
                            Word.Paragraph total = wordDoc.Paragraphs.Add();
                            total.Range.Text = $"ИТОГО ПО ПОЛЬЗОВАТЕЛЮ: {userTotal:N2} руб.";
                            total.Range.Font.Bold = 1;
                            total.Range.InsertParagraphAfter();
                        }

                        // Пустая строка между пользователями
                        wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();
                    }

                    // Общий итог
                    var totalAmount = payments.Sum(p => p.Num * p.Price);
                    Word.Paragraph finalTotal = wordDoc.Paragraphs.Add();
                    finalTotal.Range.Text = $"ОБЩАЯ СУММА: {totalAmount:N2} руб.\nКОЛИЧЕСТВО ПЛАТЕЖЕЙ: {payments.Count}";
                    finalTotal.Range.Font.Bold = 1;
                    finalTotal.Range.Font.Size = 12;

                    // Сохранение документа
                    wordDoc.SaveAs2(filePath);
                }
            }
            finally
            {
                wordDoc?.Close();
                wordApp?.Quit();

                // Освобождение COM объектов
                if (wordDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                if (wordApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        private void BtnExportUserExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportUserToExcel();
        }

        private void ExportUserToExcel()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                if (!ValidateDates(DpUserReportStart.SelectedDate, DpUserReportEnd.SelectedDate))
                    return;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Отчет_по_пользователям_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    excelApp = new Excel.Application();
                    workbook = excelApp.Workbooks.Add();
                    worksheet = workbook.ActiveSheet;

                    var startDate = DpUserReportStart.SelectedDate.Value;
                    var endDate = DpUserReportEnd.SelectedDate.Value;
                    var selectedUser = CbUserReport.SelectedItem as User;

                    using (var db = new Entities())
                    {
                        var paymentsQuery = db.Payments
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

                        // Заголовок отчета
                        worksheet.Cells[1, 1] = "ОТЧЕТ ПО ПОЛЬЗОВАТЕЛЯМ";
                        worksheet.Cells[1, 1].Font.Bold = true;
                        worksheet.Cells[1, 1].Font.Size = 14;

                        worksheet.Cells[2, 1] = $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                        if (selectedUser != null)
                            worksheet.Cells[3, 1] = $"Пользователь: {selectedUser.FIO}";
                        worksheet.Cells[4, 1] = $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";

                        // Заголовки таблицы
                        int row = 6;
                        worksheet.Cells[row, 1] = "Дата";
                        worksheet.Cells[row, 2] = "Пользователь";
                        worksheet.Cells[row, 3] = "Категория";
                        worksheet.Cells[row, 4] = "Название";
                        worksheet.Cells[row, 5] = "Количество";
                        worksheet.Cells[row, 6] = "Цена";
                        worksheet.Cells[row, 7] = "Сумма";

                        // Стиль заголовков
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 7]];
                        headerRange.Font.Bold = true;
                        headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;

                        row++;

                        foreach (var payment in payments)
                        {
                            var amount = payment.Num * payment.Price;
                            worksheet.Cells[row, 1] = payment.Date.ToString("dd.MM.yyyy");
                            worksheet.Cells[row, 2] = payment.User.FIO;
                            worksheet.Cells[row, 3] = payment.Category.Name;
                            worksheet.Cells[row, 4] = payment.Name;
                            worksheet.Cells[row, 5] = payment.Num;
                            worksheet.Cells[row, 6] = (double)payment.Price;
                            worksheet.Cells[row, 7] = (double)amount;
                            row++;
                        }

                        // Форматирование денежных колонок
                        Excel.Range priceRange = worksheet.Range[worksheet.Cells[7, 6], worksheet.Cells[row, 7]];
                        priceRange.NumberFormat = "#,##0.00\" руб.\"";

                        // Автоподбор ширины колонок
                        worksheet.Columns.AutoFit();

                        // Итоги
                        worksheet.Cells[row + 2, 6] = "Общая сумма:";
                        worksheet.Cells[row + 2, 7] = (double)payments.Sum(p => p.Num * p.Price);
                        worksheet.Cells[row + 2, 6].Font.Bold = true;
                        worksheet.Cells[row + 2, 7].Font.Bold = true;

                        worksheet.Cells[row + 3, 6] = "Количество платежей:";
                        worksheet.Cells[row + 3, 7] = payments.Count;
                        worksheet.Cells[row + 3, 6].Font.Bold = true;

                        // Сохранение
                        workbook.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Отчет успешно экспортирован в Excel!", "Успех");

                        // Закрываем перед открытием файла
                        workbook.Close(false);
                        excelApp.Quit();

                        Process.Start(saveDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка");
            }
            finally
            {
                // Правильный порядок освобождения COM-объектов
                if (worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    worksheet = null;
                }

                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }

                // Принудительная сборка мусора для COM-объектов
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
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
            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                wordApp.Visible = false;

                var startDate = DpCategoryReportStart.SelectedDate.Value;
                var endDate = DpCategoryReportEnd.SelectedDate.Value;
                var selectedCategory = CbCategoryReport.SelectedItem as Category;

                using (var db = new Entities())
                {
                    var paymentsQuery = db.Payments
                        .Include(p => p.Category)
                        .Include(p => p.User)
                        .Where(p => p.Date >= startDate && p.Date <= endDate);

                    if (selectedCategory != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
                    }

                    var categoryData = paymentsQuery
                        .GroupBy(p => p.Category.Name)
                        .Select(g => new CategoryReportData
                        {
                            Category = g.Key,
                            TotalAmount = g.Sum(p => p.Num * p.Price),
                            PaymentCount = g.Count(),
                            UsersCount = g.Select(p => p.UserID).Distinct().Count()
                        })
                        .OrderByDescending(x => x.TotalAmount)
                        .ToList();

                    // Заголовок отчета
                    Word.Paragraph title = wordDoc.Paragraphs.Add();
                    title.Range.Text = "ОТЧЕТ ПО КАТЕГОРИЯМ";
                    title.Range.Font.Bold = 1;
                    title.Range.Font.Size = 16;
                    title.Range.InsertParagraphAfter();

                    // Информация
                    Word.Paragraph info = wordDoc.Paragraphs.Add();
                    info.Range.Text = $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                    if (selectedCategory != null)
                        info.Range.Text += $"\nКатегория: {selectedCategory.Name}";
                    info.Range.Text += $"\nСгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";
                    info.Range.InsertParagraphAfter();

                    wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Таблица с данными
                    if (categoryData.Any())
                    {
                        Word.Table table = wordDoc.Tables.Add(
                            wordDoc.Paragraphs.Add().Range,
                            categoryData.Count + 1,
                            4);

                        table.Borders.Enable = 1;
                        table.Rows[1].Range.Font.Bold = 1;

                        // Заголовки
                        table.Cell(1, 1).Range.Text = "Категория";
                        table.Cell(1, 2).Range.Text = "Кол-во платежей";
                        table.Cell(1, 3).Range.Text = "Кол-во пользователей";
                        table.Cell(1, 4).Range.Text = "Общая сумма";

                        int row = 2;
                        foreach (var category in categoryData)
                        {
                            table.Cell(row, 1).Range.Text = category.Category;
                            table.Cell(row, 2).Range.Text = category.PaymentCount.ToString();
                            table.Cell(row, 3).Range.Text = category.UsersCount.ToString();
                            table.Cell(row, 4).Range.Text = category.TotalAmount.ToString("N2") + " руб.";
                            row++;
                        }
                    }

                    // Итоги
                    wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();
                    Word.Paragraph total = wordDoc.Paragraphs.Add();
                    decimal totalSum = categoryData.Sum(c => c.TotalAmount);
                    int totalPayments = categoryData.Sum(c => c.PaymentCount);
                    total.Range.Text = $"ОБЩАЯ СУММА: {totalSum:N2} руб.\n" +
                                      $"ВСЕГО ПЛАТЕЖЕЙ: {totalPayments}\n" +
                                      $"КОЛИЧЕСТВО КАТЕГОРИЙ: {categoryData.Count}";
                    total.Range.Font.Bold = 1;

                    wordDoc.SaveAs2(filePath);
                }
            }
            finally
            {
                wordDoc?.Close();
                wordApp?.Quit();

                if (wordDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                if (wordApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        private void BtnExportCategoryExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportCategoryToExcel();
        }

        private void ExportCategoryToExcel()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                if (!ValidateDates(DpCategoryReportStart.SelectedDate, DpCategoryReportEnd.SelectedDate))
                    return;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Отчет_по_категориям_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    excelApp = new Excel.Application();
                    workbook = excelApp.Workbooks.Add();
                    worksheet = workbook.ActiveSheet;

                    var startDate = DpCategoryReportStart.SelectedDate.Value;
                    var endDate = DpCategoryReportEnd.SelectedDate.Value;
                    var selectedCategory = CbCategoryReport.SelectedItem as Category;

                    using (var db = new Entities())
                    {
                        var paymentsQuery = db.Payments
                            .Include(p => p.Category)
                            .Include(p => p.User)
                            .Where(p => p.Date >= startDate && p.Date <= endDate);

                        if (selectedCategory != null)
                        {
                            paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
                        }

                        var categoryData = paymentsQuery
                            .GroupBy(p => p.Category.Name)
                            .Select(g => new CategoryReportData
                            {
                                Category = g.Key,
                                TotalAmount = g.Sum(p => p.Num * p.Price),
                                PaymentCount = g.Count(),
                                UsersCount = g.Select(p => p.UserID).Distinct().Count()
                            })
                            .OrderByDescending(x => x.TotalAmount)
                            .ToList();

                        // Заголовок
                        worksheet.Cells[1, 1] = "ОТЧЕТ ПО КАТЕГОРИЯМ";
                        worksheet.Cells[1, 1].Font.Bold = true;
                        worksheet.Cells[1, 1].Font.Size = 14;

                        worksheet.Cells[2, 1] = $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                        if (selectedCategory != null)
                            worksheet.Cells[3, 1] = $"Категория: {selectedCategory.Name}";
                        worksheet.Cells[4, 1] = $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";

                        // Таблица
                        int row = 6;
                        worksheet.Cells[row, 1] = "Категория";
                        worksheet.Cells[row, 2] = "Кол-во платежей";
                        worksheet.Cells[row, 3] = "Кол-во пользователей";
                        worksheet.Cells[row, 4] = "Общая сумма";

                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 4]];
                        headerRange.Font.Bold = true;
                        headerRange.Interior.Color = Excel.XlRgbColor.rgbLightGray;

                        row++;

                        foreach (var category in categoryData)
                        {
                            worksheet.Cells[row, 1] = category.Category;
                            worksheet.Cells[row, 2] = category.PaymentCount;
                            worksheet.Cells[row, 3] = category.UsersCount;
                            worksheet.Cells[row, 4] = (double)category.TotalAmount;
                            row++;
                        }

                        // Форматирование
                        Excel.Range amountRange = worksheet.Range[worksheet.Cells[7, 4], worksheet.Cells[row, 4]];
                        amountRange.NumberFormat = "#,##0.00\" руб.\"";

                        worksheet.Columns.AutoFit();

                        // Итоги
                        decimal totalSum = categoryData.Sum(c => c.TotalAmount);
                        worksheet.Cells[row + 2, 3] = "Общая сумма:";
                        worksheet.Cells[row + 2, 4] = (double)totalSum;
                        worksheet.Cells[row + 2, 3].Font.Bold = true;
                        worksheet.Cells[row + 2, 4].Font.Bold = true;

                        workbook.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Отчет успешно экспортирован в Excel!", "Успех");
                        Process.Start(saveDialog.FileName);
                    }
                }
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();

                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
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
            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                wordApp.Visible = false;

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

                    // Заголовок отчета
                    Word.Paragraph title = wordDoc.Paragraphs.Add();
                    title.Range.Text = "СВОДНЫЙ ОТЧЕТ ПО СИСТЕМЕ";
                    title.Range.Font.Bold = 1;
                    title.Range.Font.Size = 16;
                    title.Range.InsertParagraphAfter();

                    // Информация
                    Word.Paragraph info = wordDoc.Paragraphs.Add();
                    info.Range.Text = $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}\n" +
                                     $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";
                    info.Range.InsertParagraphAfter();

                    wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Основные показатели
                    Word.Paragraph statsHeader = wordDoc.Paragraphs.Add();
                    statsHeader.Range.Text = "ОСНОВНЫЕ ПОКАЗАТЕЛИ:";
                    statsHeader.Range.Font.Bold = 1;
                    statsHeader.Range.Font.Size = 12;
                    statsHeader.Range.InsertParagraphAfter();

                    Word.Paragraph stats = wordDoc.Paragraphs.Add();
                    stats.Range.Text = $"Общее количество платежей: {totalPayments}\n" +
                                      $"Общая сумма платежей: {totalAmount:N2} руб.\n" +
                                      $"Активных пользователей: {activeUsers}\n" +
                                      $"Использованных категорий: {categoriesUsed}\n" +
                                      $"Средний платеж: {(totalPayments > 0 ? totalAmount / totalPayments : 0):N2} руб.";
                    stats.Range.InsertParagraphAfter();

                    wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Топ пользователей
                    Word.Paragraph topUsersHeader = wordDoc.Paragraphs.Add();
                    topUsersHeader.Range.Text = "ТОП-5 ПОЛЬЗОВАТЕЛЕЙ:";
                    topUsersHeader.Range.Font.Bold = 1;
                    topUsersHeader.Range.Font.Size = 12;
                    topUsersHeader.Range.InsertParagraphAfter();

                    foreach (var user in topUsers)
                    {
                        Word.Paragraph userItem = wordDoc.Paragraphs.Add();
                        userItem.Range.Text = $"{user.User}: {user.Total:N2} руб.";
                        userItem.Range.InsertParagraphAfter();
                    }

                    wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Топ категорий
                    Word.Paragraph topCategoriesHeader = wordDoc.Paragraphs.Add();
                    topCategoriesHeader.Range.Text = "ТОП-5 КАТЕГОРИЙ:";
                    topCategoriesHeader.Range.Font.Bold = 1;
                    topCategoriesHeader.Range.Font.Size = 12;
                    topCategoriesHeader.Range.InsertParagraphAfter();

                    foreach (var category in topCategories)
                    {
                        Word.Paragraph categoryItem = wordDoc.Paragraphs.Add();
                        categoryItem.Range.Text = $"{category.Category}: {category.Total:N2} руб.";
                        categoryItem.Range.InsertParagraphAfter();
                    }

                    wordDoc.SaveAs2(filePath);
                }
            }
            finally
            {
                wordDoc?.Close();
                wordApp?.Quit();

                if (wordDoc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                if (wordApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        private void BtnExportSummaryExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportSummaryToExcel();
        }

        private void ExportSummaryToExcel()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                if (!ValidateDates(DpSummaryReportStart.SelectedDate, DpSummaryReportEnd.SelectedDate))
                    return;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Сводный_отчет_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    excelApp = new Excel.Application();
                    workbook = excelApp.Workbooks.Add();
                    worksheet = workbook.ActiveSheet;

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

                        // Заголовок
                        worksheet.Cells[1, 1] = "СВОДНЫЙ ОТЧЕТ ПО СИСТЕМЕ";
                        worksheet.Cells[1, 1].Font.Bold = true;
                        worksheet.Cells[1, 1].Font.Size = 14;

                        worksheet.Cells[2, 1] = $"Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
                        worksheet.Cells[3, 1] = $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";

                        // Основные показатели
                        int row = 5;
                        worksheet.Cells[row, 1] = "ОСНОВНЫЕ ПОКАЗАТЕЛИ";
                        worksheet.Cells[row, 1].Font.Bold = true;
                        row++;

                        worksheet.Cells[row, 1] = "Общее количество платежей:";
                        worksheet.Cells[row, 2] = totalPayments;
                        row++;

                        worksheet.Cells[row, 1] = "Общая сумма платежей:";
                        worksheet.Cells[row, 2] = (double)totalAmount;
                        row++;

                        worksheet.Cells[row, 1] = "Активных пользователей:";
                        worksheet.Cells[row, 2] = activeUsers;
                        row++;

                        worksheet.Cells[row, 1] = "Использованных категорий:";
                        worksheet.Cells[row, 2] = categoriesUsed;
                        row++;

                        worksheet.Cells[row, 1] = "Средний платеж:";
                        worksheet.Cells[row, 2] = (double)(totalPayments > 0 ? totalAmount / totalPayments : 0);
                        row++;

                        // Форматирование денежных значений
                        Excel.Range amountRange = worksheet.Range[worksheet.Cells[7, 2], worksheet.Cells[row, 2]];
                        amountRange.NumberFormat = "#,##0.00\" руб.\"";

                        worksheet.Columns.AutoFit();

                        workbook.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Сводный отчет успешно экспортирован в Excel!", "Успех");
                        Process.Start(saveDialog.FileName);
                    }
                }
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();

                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
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
                    var totalAmount = db.Payments.Sum(p => (decimal?)(p.Num * p.Price)) ?? 0;
                    var avgPayment = db.Payments.Average(p => (decimal?)(p.Num * p.Price)) ?? 0;
                    var firstPayment = db.Payments.Min(p => (DateTime?)p.Date);
                    var lastPayment = db.Payments.Max(p => (DateTime?)p.Date);

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
                        $"  • Общая сумма: {totalAmount:N2} руб.",
                        $"  • Средний платеж: {avgPayment:N2} руб.",
                        "",
                        $"📈 АКТИВНОСТЬ:",
                        $"  • Первый платеж: {(firstPayment.HasValue ? firstPayment.Value.ToString("dd.MM.yyyy") : "нет данных")}",
                        $"  • Последний платеж: {(lastPayment.HasValue ? lastPayment.Value.ToString("dd.MM.yyyy") : "нет данных")}",
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

    // Вспомогательный класс для типизации данных отчета по категориям
    public class CategoryReportData
    {
        public string Category { get; set; }
        public decimal TotalAmount { get; set; }
        public int PaymentCount { get; set; }
        public int UsersCount { get; set; }
    }
}