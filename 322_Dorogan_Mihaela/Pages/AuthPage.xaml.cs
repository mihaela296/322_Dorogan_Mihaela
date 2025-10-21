using _322_Dorogan_Mihaela;
using _322_Dorogan_Mihaela.Pages;
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace YourProjectName.Pages
{
    public partial class AuthPage : Page
    {
        private int _failedAttempts = 0;
        private Random _random = new Random();

        public AuthPage()
        {
            InitializeComponent();
            TbLogin.Focus();
        }

        private string GetHash(string input)
        {
            using (var sha1 = SHA1.Create())
            {
                var hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(input));
                return string.Concat(hash.Select(b => b.ToString("X2")));
            }
        }

        private void GenerateCaptcha()
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            TbCaptcha.Text = new string(Enumerable.Repeat(chars, 6)
                .Select(s => s[_random.Next(s.Length)]).ToArray());
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            // Проверка заполнения полей
            if (string.IsNullOrWhiteSpace(TbLogin.Text))
            {
                MessageBox.Show("Введите логин!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                TbLogin.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbPassword.Password))
            {
                MessageBox.Show("Введите пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbPassword.Focus();
                return;
            }

            // Проверка капчи после 3 неудачных попыток
            if (_failedAttempts >= 3)
            {
                if (SpCaptcha.Visibility != System.Windows.Visibility.Visible)
                {
                    SpCaptcha.Visibility = System.Windows.Visibility.Visible;
                    GenerateCaptcha();
                }

                if (TbCaptchaInput.Text != TbCaptcha.Text)
                {
                    MessageBox.Show("Неверно введена капча!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    GenerateCaptcha();
                    TbCaptchaInput.Clear();
                    TbCaptchaInput.Focus();
                    return;
                }
            }

            // Проверка учетных данных
            string hashedPassword = GetHash(PbPassword.Password);

            try
            {
                using (var db = new Entities())
                {
                    var user = db.Users
                        .AsNoTracking()
                        .FirstOrDefault(u => u.Login == TbLogin.Text && u.Password == hashedPassword);

                    if (user == null)
                    {
                        _failedAttempts++;
                        MessageBox.Show("Неверный логин или пароль!", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);

                        if (_failedAttempts >= 3 && SpCaptcha.Visibility != System.Windows.Visibility.Visible)
                        {
                            SpCaptcha.Visibility = System.Windows.Visibility.Visible;
                            GenerateCaptcha();
                        }
                        return;
                    }

                    // Успешная авторизация
                    _failedAttempts = 0;
                    SpCaptcha.Visibility = System.Windows.Visibility.Collapsed;

                    MessageBox.Show($"Добро пожаловать, {user.FIO}!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    // Переход в зависимости от роли
                    switch (user.Role)
                    {
                        case "Admin":
                            NavigationService.Navigate(new AdminPage(user));
                            break;
                        case "User":
                        default:
                            NavigationService.Navigate(new UserPage(user));
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при авторизации: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new RegPage());
        }

        private void BtnChangePassword_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ChangePasswordPage());
        }
    }
}