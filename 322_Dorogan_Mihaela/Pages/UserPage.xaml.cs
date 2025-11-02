using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class UserPage : Page
    {
        private User _currentUser;

        public UserPage(User user)
        {
            InitializeComponent();
            _currentUser = user;
            Loaded += UserPage_Loaded;
        }

        private void UserPage_Loaded(object sender, RoutedEventArgs e)
        {
            InitializePage();
        }

        private void InitializePage()
        {
            TbWelcome.Text = $"Добро пожаловать, {_currentUser.FIO}!";

            // Установка дат по умолчанию (текущий месяц)
            DpStartDate.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DpEndDate.SelectedDate = DateTime.Now;

            LoadCategories();
            LoadPayments();
            LoadStatistics();
        }

        private void LoadCategories()
        {
            try
            {
                using (var db = new Entities())
                {
                    var categories = db.Categories.OrderBy(c => c.Name).ToList();
                    CbCategory.ItemsSource = categories;
                    CbCategory.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки категорий: {ex.Message}");
            }
        }

        private void LoadPayments()
        {
            try
            {
                using (var db = new Entities())
                {
                    var paymentsQuery = db.Payments
                        .Include(p => p.Category)
                        .Where(p => p.UserID == _currentUser.ID);

                    // Фильтрация по дате
                    if (DpStartDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                    }

                    if (DpEndDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                    }

                    // Фильтрация по категории - только если выбрана конкретная категория
                    if (CbCategory.SelectedItem != null && CbCategory.SelectedItem is Category selectedCategory)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
                    }

                    var payments = paymentsQuery
                        .OrderByDescending(p => p.Date)
                        .ToList()
                        .Select(p => new
                        {
                            p.ID,
                            p.Date,
                            p.Name,
                            p.Category,
                            p.Num,
                            p.Price,
                            TotalAmount = p.Num * p.Price
                        });

                    DgPayments.ItemsSource = payments;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки платежей: {ex.Message}");
            }
        }

        private void LoadStatistics()
        {
            try
            {
                using (var db = new Entities())
                {
                    var totalSum = db.Payments
                        .Where(p => p.UserID == _currentUser.ID)
                        .Sum(p => (decimal?)(p.Num * p.Price)) ?? 0;

                    var statistics = db.Payments
                        .Include(p => p.Category)
                        .Where(p => p.UserID == _currentUser.ID)
                        .GroupBy(p => p.Category.Name)
                        .Select(g => new
                        {
                            CategoryName = g.Key,
                            PaymentCount = g.Count(),
                            TotalAmount = g.Sum(p => p.Num * p.Price),
                            AverageAmount = g.Average(p => p.Num * p.Price),
                            Percentage = totalSum > 0 ? g.Sum(p => p.Num * p.Price) / totalSum : 0
                        })
                        .OrderByDescending(s => s.TotalAmount)
                        .ToList();

                    DgStatistics.ItemsSource = statistics;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки статистики: {ex.Message}");
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadPayments();
            LoadStatistics();
        }

        private void CbCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadPayments();
            LoadStatistics();
        }

        private void CbChartType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Заглушка для типа диаграммы
        }

        private void BtnApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            LoadPayments();
            LoadStatistics();
        }

        private void BtnClearFilter_Click(object sender, RoutedEventArgs e)
        {
            DpStartDate.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DpEndDate.SelectedDate = DateTime.Now;
            CbCategory.SelectedIndex = -1;
            LoadPayments();
            LoadStatistics();
        }

        private void BtnAddPayment_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddEditPaymentPage(null, _currentUser));
        }

        private void BtnDeletePayment_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button?.DataContext != null)
            {
                dynamic payment = button.DataContext;
                int paymentId = payment.ID;

                var result = MessageBox.Show("Вы уверены, что хотите удалить этот платеж?",
                    "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var db = new Entities())
                        {
                            var paymentToDelete = db.Payments.Find(paymentId);
                            if (paymentToDelete != null)
                            {
                                db.Payments.Remove(paymentToDelete);
                                db.SaveChanges();
                                LoadPayments();
                                LoadStatistics();
                                MessageBox.Show("Платеж успешно удален!");
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

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Платежи_{_currentUser.FIO}_{DateTime.Now:yyyyMMdd}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    MessageBox.Show("Экспорт в Excel будет реализован позже", "Информация");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}");
            }
        }

        private void BtnChangePass_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ChangePasswordPage());
        }

        private void BtnLogout_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите выйти?", "Подтверждение выхода",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                NavigationService.Navigate(new AuthPage());
            }
        }
    }
}