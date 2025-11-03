using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

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
                using (var db = new Entities())
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
                using (var db = new Entities())
                {
                    // Начинаем с базового запроса
                    IQueryable<Payment> paymentsQuery = db.Payments
                        .Include(p => p.User)
                        .Include(p => p.Category);

                    // Применяем фильтры только если они установлены
                    if (DpStartDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                    }

                    if (DpEndDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                    }

                    // Для комбобоксов проверяем SelectedItem
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

                    var payments = paymentsQuery
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
                        });

                    DgPayments.ItemsSource = payments;
                    UpdateStatistics(paymentsQuery);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки платежей: {ex.Message}");
            }
        }

        private void UpdateStatistics(IQueryable<Payment> paymentsQuery)
        {
            try
            {
                var totalCount = paymentsQuery.Count();
                var totalAmount = paymentsQuery.Sum(p => (decimal?)(p.Num * p.Price)) ?? 0;
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
                using (var db = new Entities())
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
                        using (var db = new Entities())
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

        private void BtnExportPayments_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Платежи_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Простая визуальная обратная связь
                    string originalText = BtnExportPayments.Content.ToString();
                    BtnExportPayments.Content = "Экспорт...";
                    BtnExportPayments.IsEnabled = false;

                    try
                    {
                        // Получаем данные
                        var payments = GetPaymentsForExport();
                        if (payments == null || payments.Count == 0)
                        {
                            MessageBox.Show("Нет данных для экспорта", "Информация");
                            return;
                        }

                        // Создаем Excel пакет
                        using (var excelPackage = new OfficeOpenXml.ExcelPackage())
                        {
                            var worksheet = excelPackage.Workbook.Worksheets.Add("Платежи");

                            // Простые заголовки
                            worksheet.Cells[1, 1].Value = "ID";
                            worksheet.Cells[1, 2].Value = "Дата";
                            worksheet.Cells[1, 3].Value = "Пользователь";
                            worksheet.Cells[1, 4].Value = "Категория";
                            worksheet.Cells[1, 5].Value = "Название";
                            worksheet.Cells[1, 6].Value = "Кол-во";
                            worksheet.Cells[1, 7].Value = "Цена";
                            worksheet.Cells[1, 8].Value = "Сумма";

                            // Делаем заголовки жирными
                            for (int i = 1; i <= 8; i++)
                            {
                                worksheet.Cells[1, i].Style.Font.Bold = true;
                            }

                            // Заполняем данные
                            int row = 2;
                            foreach (var payment in payments)
                            {
                                worksheet.Cells[row, 1].Value = payment.ID;
                                worksheet.Cells[row, 2].Value = payment.Date.ToString("dd.MM.yyyy");
                                worksheet.Cells[row, 3].Value = payment.User?.FIO ?? "";
                                worksheet.Cells[row, 4].Value = payment.Category?.Name ?? "";
                                worksheet.Cells[row, 5].Value = payment.Name;
                                worksheet.Cells[row, 6].Value = payment.Num;
                                worksheet.Cells[row, 7].Value = payment.Price;
                                worksheet.Cells[row, 8].Value = payment.Num * payment.Price;
                                row++;
                            }

                            // Автоподбор ширины столбцов
                            worksheet.Cells[1, 1, row, 8].AutoFitColumns();

                            // Сохраняем файл
                            FileInfo fileInfo = new FileInfo(saveDialog.FileName);
                            excelPackage.SaveAs(fileInfo);
                        }

                        MessageBox.Show($"Данные экспортированы успешно!\n\nФайл: {saveDialog.FileName}",
                            "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}\n\n" +
                            "Убедитесь, что:\n" +
                            "• Файл не открыт в Excel или другой программе\n" +
                            "• У вас есть права на запись в выбранную папку\n" +
                            "• Попробуйте сохранить файл в другую папку",
                            "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    finally
                    {
                        // Восстанавливаем кнопку
                        BtnExportPayments.Content = originalText;
                        BtnExportPayments.IsEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
                BtnExportPayments.Content = "📊 Экспорт в Excel";
                BtnExportPayments.IsEnabled = true;
            }
        }

        // Новый метод для получения платежей для экспорта
        private List<Payment> GetPaymentsForExport()
        {
            try
            {
                using (var db = new Entities())
                {
                    // Начинаем с базового запроса
                    IQueryable<Payment> paymentsQuery = db.Payments
                        .Include(p => p.User)
                        .Include(p => p.Category);

                    // Применяем фильтры только если они установлены
                    if (DpStartDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                    }

                    if (DpEndDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                    }

                    // Для комбобоксов проверяем SelectedItem
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