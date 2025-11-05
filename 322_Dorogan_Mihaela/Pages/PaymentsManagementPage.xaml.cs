using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class PaymentsManagementPage : Page
    {
        private User _currentAdmin;

        public PaymentsManagementPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            Loaded += PaymentsManagementPage_Loaded;
        }

        private void PaymentsManagementPage_Loaded(object sender, RoutedEventArgs e)
        {
            InitializeFilters();
            LoadPayments();
        }

        private void InitializeFilters()
        {
            try
            {
                using (var db = new DEntities())
                {
                    // Загрузка пользователей
                    var users = db.Users.OrderBy(u => u.FIO).ToList();
                    CbUser.ItemsSource = users;
                    CbUser.SelectedIndex = -1;

                    // Загрузка категорий
                    var categories = db.Categories.OrderBy(c => c.Name).ToList();
                    CbCategory.ItemsSource = categories;
                    CbCategory.SelectedIndex = -1;

                    // Установка дат по умолчанию (последние 30 дней)
                    DpEndDate.SelectedDate = DateTime.Now;
                    DpStartDate.SelectedDate = DateTime.Now.AddDays(-30);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации фильтров: {ex.Message}");
            }
        }

        private void LoadPayments()
        {
            try
            {
                using (var db = new DEntities())
                {
                    var payments = GetFilteredPayments(db);
                    DgPayments.ItemsSource = payments;
                    UpdateStatistics(payments);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки платежей: {ex.Message}");
            }
        }

        private List<dynamic> GetFilteredPayments(DEntities db)
        {
            IQueryable<Payment> paymentsQuery = db.Payments
                .Include(p => p.User)
                .Include(p => p.Category);

            // Применяем фильтры
            if (DpStartDate.SelectedDate != null)
            {
                paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
            }

            if (DpEndDate.SelectedDate != null)
            {
                paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
            }

            if (CbUser.SelectedItem != null && CbUser.SelectedItem is User selectedUser)
            {
                paymentsQuery = paymentsQuery.Where(p => p.UserID == selectedUser.ID);
            }

            if (CbCategory.SelectedItem != null && CbCategory.SelectedItem is Category selectedCategory)
            {
                paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
            }

            if (!string.IsNullOrWhiteSpace(TbSearch.Text))
            {
                var searchText = TbSearch.Text.ToLower();
                paymentsQuery = paymentsQuery.Where(p => p.Name.ToLower().Contains(searchText));
            }

            return paymentsQuery
                .OrderByDescending(p => p.Date)
                .ThenByDescending(p => p.ID)
                .ToList()
                .Select(p => new
                {
                    p.ID,
                    p.Date,
                    p.User,
                    p.Category,
                    p.Name,
                    p.Num,
                    p.Price,
                    TotalAmount = p.Num * p.Price
                })
                .Cast<dynamic>()
                .ToList();
        }

        private void UpdateStatistics(List<dynamic> payments)
        {
            try
            {
                var totalCount = payments.Count;
                var totalAmount = payments.Sum(p => (decimal)p.TotalAmount);
                var avgAmount = totalCount > 0 ? totalAmount / totalCount : 0;

                TbStats.Text = $"Всего: {totalCount} платежей | Сумма: {totalAmount:N2} руб. | Средний: {avgAmount:N2} руб.";
            }
            catch (Exception ex)
            {
                TbStats.Text = "Ошибка расчета статистики";
            }
        }

        private void Filters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadPayments();
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadPayments();
        }

        private void BtnApplyFilters_Click(object sender, RoutedEventArgs e)
        {
            LoadPayments();
        }

        private void BtnClearFilters_Click(object sender, RoutedEventArgs e)
        {
            DpStartDate.SelectedDate = DateTime.Now.AddDays(-30);
            DpEndDate.SelectedDate = DateTime.Now;
            CbUser.SelectedIndex = -1;
            CbCategory.SelectedIndex = -1;
            TbSearch.Clear();
            LoadPayments();
        }

        private void BtnAddPayment_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddEditPaymentPage(null, _currentAdmin));
        }

        private void BtnEditPayment_Click(object sender, RoutedEventArgs e)
        {
            var payment = (sender as Button)?.DataContext;
            if (payment != null)
            {
                dynamic pay = payment;
                using (var db = new DEntities())
                {
                    var paymentToEdit = db.Payments.Find(pay.ID);
                    if (paymentToEdit != null)
                    {
                        NavigationService.Navigate(new AddEditPaymentPage(paymentToEdit, _currentAdmin));
                    }
                }
            }
        }

        private void BtnDeletePayment_Click(object sender, RoutedEventArgs e)
        {
            var payment = (sender as Button)?.DataContext;
            if (payment != null)
            {
                dynamic pay = payment;

                var result = MessageBox.Show(
                    $"Вы уверены, что хотите удалить платеж:\n\"{pay.Name}\" от {pay.Date:dd.MM.yyyy}?\nСумма: {pay.TotalAmount:N2} руб.",
                    "Подтверждение удаления",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var db = new DEntities())
                        {
                            var paymentToDelete = db.Payments.Find(pay.ID);
                            if (paymentToDelete != null)
                            {
                                db.Payments.Remove(paymentToDelete);
                                db.SaveChanges();
                                LoadPayments();
                                MessageBox.Show("Платеж успешно удален!", "Успех");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка удаления платежа: {ex.Message}");
                    }
                }
            }
        }

        private void BtnExportPaymentsExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel();
        }

        private void BtnExportPaymentsWord_Click(object sender, RoutedEventArgs e)
        {
            ExportToWord();
        }

        private void ExportToExcel()
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Платежи_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Всегда сохраняем как CSV для надежности
                    string csvFilePath = Path.ChangeExtension(saveDialog.FileName, ".csv");
                    ExportToCsv(csvFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка");
            }
        }

        private void ExportToCsv(string filePath)
        {
            try
            {
                var payments = GetPaymentsForExport();
                if (payments.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Информация");
                    return;
                }

                var csvLines = new List<string>
                {
                    "ID;Дата;Пользователь;Категория;Название;Количество;Цена;Сумма"
                };

                foreach (var payment in payments)
                {
                    var amount = payment.Num * payment.Price;
                    csvLines.Add(
                        $"{payment.ID};" +
                        $"{payment.Date:dd.MM.yyyy};" +
                        $"{EscapeCsvField(payment.User?.FIO ?? "")};" +
                        $"{EscapeCsvField(payment.Category?.Name ?? "")};" +
                        $"{EscapeCsvField(payment.Name)};" +
                        $"{payment.Num};" +
                        $"{payment.Price:N2};" +
                        $"{amount:N2}"
                    );
                }

                File.WriteAllLines(filePath, csvLines, System.Text.Encoding.UTF8);

                MessageBox.Show($"Данные экспортированы успешно!\nФайл: {filePath}\n\nФайл откроется в Excel автоматически.",
                    "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);

                // Открываем файл в ассоциированной программе (Excel)
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}\n\nПроверьте:\n• Доступ к папке\n• Закрыт ли файл в другой программе",
                    "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(";") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        private void ExportToWord()
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    FileName = $"Платежи_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".docx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    Word.Application wordApp = null;
                    Word.Document wordDoc = null;

                    try
                    {
                        var payments = GetPaymentsForExport();
                        if (payments.Count == 0)
                        {
                            MessageBox.Show("Нет данных для экспорта", "Информация");
                            return;
                        }

                        wordApp = new Word.Application();
                        wordDoc = wordApp.Documents.Add();
                        wordApp.Visible = false;

                        // Заголовок
                        Word.Paragraph title = wordDoc.Paragraphs.Add();
                        title.Range.Text = "ОТЧЕТ ПО ПЛАТЕЖАМ";
                        title.Range.Font.Bold = 1;
                        title.Range.Font.Size = 16;
                        title.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        title.Range.InsertParagraphAfter();

                        // Информация о фильтрах
                        Word.Paragraph info = wordDoc.Paragraphs.Add();
                        info.Range.Text = GetFilterInfo();
                        info.Range.Font.Size = 12;
                        info.Range.InsertParagraphAfter();

                        wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                        // Таблица
                        if (payments.Count > 0)
                        {
                            Word.Table table = wordDoc.Tables.Add(
                                wordDoc.Paragraphs.Add().Range,
                                payments.Count + 1,
                                8);

                            table.Borders.Enable = 1;
                            table.Rows[1].Range.Font.Bold = 1;

                            // Заголовки таблицы
                            string[] headers = { "ID", "Дата", "Пользователь", "Категория", "Название", "Кол-во", "Цена", "Сумма" };
                            for (int i = 0; i < headers.Length; i++)
                            {
                                table.Cell(1, i + 1).Range.Text = headers[i];
                                table.Cell(1, i + 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }

                            // Данные
                            int row = 2;
                            decimal totalAmount = 0;
                            foreach (var payment in payments)
                            {
                                var amount = payment.Num * payment.Price;
                                totalAmount += amount;

                                table.Cell(row, 1).Range.Text = payment.ID.ToString();
                                table.Cell(row, 2).Range.Text = payment.Date.ToString("dd.MM.yyyy");
                                table.Cell(row, 3).Range.Text = payment.User?.FIO ?? "";
                                table.Cell(row, 4).Range.Text = payment.Category?.Name ?? "";
                                table.Cell(row, 5).Range.Text = payment.Name;
                                table.Cell(row, 6).Range.Text = payment.Num.ToString();
                                table.Cell(row, 7).Range.Text = payment.Price.ToString("N2") + " руб.";
                                table.Cell(row, 8).Range.Text = amount.ToString("N2") + " руб.";
                                row++;
                            }

                            // Итоги
                            wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();
                            Word.Paragraph total = wordDoc.Paragraphs.Add();
                            total.Range.Text = $"ИТОГО: {payments.Count} платежей на сумму {totalAmount:N2} руб.";
                            total.Range.Font.Bold = 1;
                            total.Range.Font.Size = 12;

                            // Сохраняем
                            wordDoc.SaveAs2(saveDialog.FileName);
                        }

                        MessageBox.Show($"Отчет успешно создан!\n\nФайл: {saveDialog.FileName}",
                            "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);

                        // Показываем документ пользователю
                        wordApp.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при создании отчета: {ex.Message}", "Ошибка экспорта");
                    }
                    finally
                    {
                        // Не закрываем Word, чтобы пользователь увидел документ
                        // Освобождение ресурсов произойдет при закрытии Word пользователем
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private string GetFilterInfo()
        {
            var filters = new List<string>();

            if (DpStartDate.SelectedDate != null && DpEndDate.SelectedDate != null)
            {
                filters.Add($"Период: {DpStartDate.SelectedDate.Value:dd.MM.yyyy} - {DpEndDate.SelectedDate.Value:dd.MM.yyyy}");
            }

            if (CbUser.SelectedItem is User selectedUser)
            {
                filters.Add($"Пользователь: {selectedUser.FIO}");
            }

            if (CbCategory.SelectedItem is Category selectedCategory)
            {
                filters.Add($"Категория: {selectedCategory.Name}");
            }

            if (!string.IsNullOrWhiteSpace(TbSearch.Text))
            {
                filters.Add($"Поиск: {TbSearch.Text}");
            }

            filters.Add($"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}");

            return string.Join(" | ", filters);
        }

        private List<Payment> GetPaymentsForExport()
        {
            try
            {
                using (var db = new DEntities())
                {
                    IQueryable<Payment> paymentsQuery = db.Payments
                        .Include(p => p.User)
                        .Include(p => p.Category);

                    if (DpStartDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                    }

                    if (DpEndDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                    }

                    if (CbUser.SelectedItem != null && CbUser.SelectedItem is User selectedUser)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.UserID == selectedUser.ID);
                    }

                    if (CbCategory.SelectedItem != null && CbCategory.SelectedItem is Category selectedCategory)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
                    }

                    if (!string.IsNullOrWhiteSpace(TbSearch.Text))
                    {
                        var searchText = TbSearch.Text.ToLower();
                        paymentsQuery = paymentsQuery.Where(p => p.Name.ToLower().Contains(searchText));
                    }

                    return paymentsQuery
                        .OrderByDescending(p => p.Date)
                        .ThenByDescending(p => p.ID)
                        .ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка получения данных для экспорта: {ex.Message}");
                return new List<Payment>();
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadPayments();
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}