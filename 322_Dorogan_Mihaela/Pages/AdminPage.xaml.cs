using System.Windows;
using System.Windows.Controls;
using System.Linq;
using System.Data.Entity;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class AdminPage : Page
    {
        private User _currentUser;

        public AdminPage(User user)
        {
            InitializeComponent();
            _currentUser = user;
            TbAdminWelcome.Text = $"Администратор: {user.FIO}";
        }

        private void BtnManageUsers_Click(object sender, RoutedEventArgs e)
        {
            NavigationService?.Navigate(new UsersManagementPage(_currentUser));
        }

        private void BtnManageCategories_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new CategoriesManagementPage(_currentUser));
        }

        private void BtnManagePayments_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PaymentsManagementPage(_currentUser));
        }

        private void BtnAnalytics_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AnalyticsPage(_currentUser));
        }

        private void BtnReports_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ReportsPage(_currentUser));
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