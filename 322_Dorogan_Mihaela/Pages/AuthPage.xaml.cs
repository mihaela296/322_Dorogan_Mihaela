using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Data.Entity;

namespace _322_Dorogan_Mihaela.Pages
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

        private void ShowError(string message)
        {
            TbError.Text = message;
            TbError.Visibility = Visibility.Visible;
        }

        private void ClearError()
        {
            TbError.Text = string.Empty;
            TbError.Visibility = Visibility.Collapsed;
        }

        private void TbLogin_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearError();
        }

        private void PbPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ClearError();
        }

        private void TbCaptchaInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearError();
        }

        private void BtnRefreshCaptcha_Click(object sender, RoutedEventArgs e)
        {
            GenerateCaptcha();
            TbCaptchaInput.Clear();
            TbCaptchaInput.Focus();
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            // Валидация полей
            if (string.IsNullOrWhiteSpace(TbLogin.Text))
            {
                ShowError("Введите логин!");
                TbLogin.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbPassword.Password))
            {
                ShowError("Введите пароль!");
                PbPassword.Focus();
                return;
            }

            // Проверка капчи после 3 неудачных попыток
            if (_failedAttempts >= 3)
            {
                if (SpCaptcha.Visibility != Visibility.Visible)
                {
                    SpCaptcha.Visibility = Visibility.Visible;
                    GenerateCaptcha();
                }

                if (string.IsNullOrWhiteSpace(TbCaptchaInput.Text))
                {
                    ShowError("Введите код с картинки!");
                    TbCaptchaInput.Focus();
                    return;
                }

                if (TbCaptchaInput.Text != TbCaptcha.Text)
                {
                    _failedAttempts++;
                    ShowError($"Неверно введена капча! Осталось попыток: {5 - _failedAttempts}");
                    GenerateCaptcha();
                    TbCaptchaInput.Clear();
                    TbCaptchaInput.Focus();
                    return;
                }
            }

            // Проверка учетных данных
            try
            {
                string hashedPassword = GetHash(PbPassword.Password);

                using (var db = new Entities())
                {
                    var user = db.Users
                        .AsNoTracking()
                        .FirstOrDefault(u => u.Login == TbLogin.Text.Trim() && u.Password == hashedPassword);

                    if (user == null)
                    {
                        _failedAttempts++;

                        if (_failedAttempts >= 3 && SpCaptcha.Visibility != Visibility.Visible)
                        {
                            SpCaptcha.Visibility = Visibility.Visible;
                            GenerateCaptcha();
                            ShowError("Неверный логин или пароль! Введите капчу для продолжения.");
                        }
                        else if (_failedAttempts >= 5)
                        {
                            ShowError("Превышено количество попыток входа. Попробуйте позже.");
                            BtnLogin.IsEnabled = false;
                        }
                        else
                        {
                            ShowError($"Неверный логин или пароль! Попытка {_failedAttempts} из 5");
                        }

                        PbPassword.Clear();
                        PbPassword.Focus();
                        return;
                    }

                    // Успешная авторизация
                    _failedAttempts = 0;
                    ClearError();
                    SpCaptcha.Visibility = Visibility.Collapsed;

                    MessageBox.Show($"Добро пожаловать, {user.FIO}!", "Успешная авторизация",
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
                ShowError($"Ошибка при авторизации: {ex.Message}");
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