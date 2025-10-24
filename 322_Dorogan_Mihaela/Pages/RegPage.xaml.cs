using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Data.Entity;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class RegPage : Page
    {
        public RegPage()
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

        private bool ValidatePassword(string password)
        {
            // Минимум 6 символов
            if (password.Length < 6)
                return false;

            // Только латинские буквы и цифры
            if (!Regex.IsMatch(password, @"^[a-zA-Z0-9]+$"))
                return false;

            // Хотя бы одна цифра
            if (!password.Any(char.IsDigit))
                return false;

            // Хотя бы одна буква
            if (!password.Any(char.IsLetter))
                return false;

            return true;
        }

        private void ShowError(string message)
        {
            TbError.Text = message;
            ErrorBorder.Visibility = Visibility.Visible;
        }

        private void ClearError()
        {
            TbError.Text = string.Empty;
            ErrorBorder.Visibility = Visibility.Collapsed;
        }

        private void TbLogin_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearError();
        }

        private void PbPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ClearError();
        }

        private void PbConfirmPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ClearError();
        }

        private void TbFIO_TextChanged(object sender, TextChangedEventArgs e)
        {
            ClearError();
        }

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
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

            if (string.IsNullOrWhiteSpace(PbConfirmPassword.Password))
            {
                ShowError("Подтвердите пароль!");
                PbConfirmPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(TbFIO.Text))
            {
                ShowError("Введите ФИО!");
                TbFIO.Focus();
                return;
            }

            // Проверка логина (только латинские буквы и цифры)
            if (!Regex.IsMatch(TbLogin.Text, @"^[a-zA-Z0-9]+$"))
            {
                ShowError("Логин должен содержать только латинские буквы и цифры!");
                TbLogin.Focus();
                TbLogin.SelectAll();
                return;
            }

            // Проверка пароля
            if (!ValidatePassword(PbPassword.Password))
            {
                ShowError("Пароль не соответствует требованиям!\n\nТребования к паролю:\n• Минимум 6 символов\n• Только латинские буквы и цифры\n• Хотя бы одна цифра и одна буква");
                PbPassword.Clear();
                PbConfirmPassword.Clear();
                PbPassword.Focus();
                return;
            }

            // Проверка совпадения паролей
            if (PbPassword.Password != PbConfirmPassword.Password)
            {
                ShowError("Пароли не совпадают!");
                PbConfirmPassword.Clear();
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка уникальности логина
            try
            {
                using (var db = new Entities())
                {
                    if (db.Users.Any(u => u.Login == TbLogin.Text.Trim()))
                    {
                        ShowError("Пользователь с таким логином уже существует!");
                        TbLogin.Focus();
                        TbLogin.SelectAll();
                        return;
                    }

                    // Создание нового пользователя
                    var newUser = new User
                    {
                        Login = TbLogin.Text.Trim(),
                        Password = GetHash(PbPassword.Password),
                        FIO = TbFIO.Text.Trim(),
                        Role = "User" // По умолчанию обычный пользователь
                    };

                    db.Users.Add(newUser);
                    db.SaveChanges();

                    MessageBox.Show("Регистрация прошла успешно!\nТеперь вы можете войти в систему используя ваш логин и пароль.",
                        "Успешная регистрация", MessageBoxButton.OK, MessageBoxImage.Information);

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка при регистрации: {ex.Message}");
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}