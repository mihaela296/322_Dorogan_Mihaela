using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class ChangePasswordPage : Page
    {
        public ChangePasswordPage()
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

        private void BtnChangePassword_Click(object sender, RoutedEventArgs e)
        {
            // Валидация полей
            if (string.IsNullOrWhiteSpace(TbLogin.Text))
            {
                MessageBox.Show("Введите логин!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                TbLogin.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbCurrentPassword.Password))
            {
                MessageBox.Show("Введите текущий пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbCurrentPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbNewPassword.Password))
            {
                MessageBox.Show("Введите новый пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbNewPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbConfirmPassword.Password))
            {
                MessageBox.Show("Подтвердите новый пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка нового пароля
            if (!ValidatePassword(PbNewPassword.Password))
            {
                MessageBox.Show("Новый пароль должен содержать:\n• Минимум 6 символов\n• Только латинские буквы и цифры\n• Хотя бы одну цифру и одну букву",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbNewPassword.Clear();
                PbConfirmPassword.Clear();
                PbNewPassword.Focus();
                return;
            }

            // Проверка совпадения паролей
            if (PbNewPassword.Password != PbConfirmPassword.Password)
            {
                MessageBox.Show("Новые пароли не совпадают!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                PbConfirmPassword.Clear();
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка что новый пароль отличается от старого
            if (PbCurrentPassword.Password == PbNewPassword.Password)
            {
                MessageBox.Show("Новый пароль должен отличаться от текущего!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                PbNewPassword.Clear();
                PbConfirmPassword.Clear();
                PbNewPassword.Focus();
                return;
            }

            try
            {
                using (var db = new Entities())
                {
                    string hashedCurrentPassword = GetHash(PbCurrentPassword.Password);

                    var user = db.Users.FirstOrDefault(u => u.Login == TbLogin.Text && u.Password == hashedCurrentPassword);

                    if (user == null)
                    {
                        MessageBox.Show("Неверный логин или текущий пароль!", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        PbCurrentPassword.Clear();
                        PbCurrentPassword.Focus();
                        return;
                    }

                    // Обновление пароля
                    user.Password = GetHash(PbNewPassword.Password);
                    db.SaveChanges();

                    MessageBox.Show("Пароль успешно изменен!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при смене пароля: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}