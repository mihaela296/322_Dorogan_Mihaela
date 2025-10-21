using System;
using System.Linq;
using System.Net.NetworkInformation;
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
            InitializePage();
        }

        private void InitializePage()
        {
            TbWelcome.Text = $"Добро пожаловать, {_currentUser.FIO}!";

            // Установка дат по умолчанию (текущий месяц)
            DpStartDate.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DpEndDate.SelectedDate = DateTime.Now;

            LoadPayments();
            LoadStatistics();
        }

        private void LoadPayments()
        {
            try
            {
                using (var db = new Entities())
                {
                    var payments = db.Payments
                        .Include("Category")
                        .Where(p => p.UserID == _currentUser.ID)
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
                MessageBox.Show($"Ошибка загрузки платежей: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadStatistics()
        {
            try
            {
                using (var db = new Entities())
                {
                    var statistics = db.Payments
                        .Include("Category")
                        .Where(p => p.UserID == _currentUser.ID)
                        .GroupBy(p => p.Category.Name)
                        .Select(g => new
                        {
                            CategoryName = g.Key,
                            PaymentCount = g.Count(),
                            TotalAmount = g.Sum(p => p.Num * p.Price),
                            AverageAmount = g.Average(p => p.Num * p.Price)
                        })
                        .OrderByDescending(s => s.TotalAmount)
                        .ToList();

                    DgStatistics.ItemsSource = statistics;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки статистики: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            if (DpStartDate.SelectedDate == null || DpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Выберите период для фильтрации!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (DpStartDate.SelectedDate > DpEndDate.SelectedDate)
            {
                MessageBox.Show("Дата начала не может быть больше даты окончания!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using (var db = new Entities())
                {
                    var payments = db.Payments
                        .Include("Category")
                        .Where(p => p.UserID == _currentUser.ID &&
                                   p.Date >= DpStartDate.SelectedDate &&
                                   p.Date <= DpEndDate.SelectedDate)
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
                MessageBox.Show($"Ошибка фильтрации: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnClearFilter_Click(object sender, RoutedEventArgs e)
        {
            DpStartDate.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DpEndDate.SelectedDate = DateTime.Now;
            LoadPayments();
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