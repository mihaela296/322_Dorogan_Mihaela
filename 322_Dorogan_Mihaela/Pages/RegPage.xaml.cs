using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

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

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            // Валидация полей
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

            if (string.IsNullOrWhiteSpace(PbConfirmPassword.Password))
            {
                MessageBox.Show("Подтвердите пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbConfirmPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(TbFIO.Text))
            {
                MessageBox.Show("Введите ФИО!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                TbFIO.Focus();
                return;
            }

            // Проверка пароля
            if (!ValidatePassword(PbPassword.Password))
            {
                MessageBox.Show("Пароль должен содержать:\n• Минимум 6 символов\n• Только латинские буквы и цифры\n• Хотя бы одну цифру и одну букву",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbPassword.Clear();
                PbConfirmPassword.Clear();
                PbPassword.Focus();
                return;
            }

            // Проверка совпадения паролей
            if (PbPassword.Password != PbConfirmPassword.Password)
            {
                MessageBox.Show("Пароли не совпадают!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbConfirmPassword.Clear();
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка уникальности логина
            try
            {
                using (var db = new Entities())
                {
                    if (db.Users.Any(u => u.Login == TbLogin.Text))
                    {
                        MessageBox.Show("Пользователь с таким логином уже существует!", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
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

                    MessageBox.Show("Регистрация прошла успешно!\nТеперь вы можете войти в систему.", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при регистрации: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}