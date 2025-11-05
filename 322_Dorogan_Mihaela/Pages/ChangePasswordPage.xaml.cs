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
    public partial class ChangePasswordPage : Page
    {
        private User _currentUser;

        // Конструктор без параметров (для общего использования)
        public ChangePasswordPage()
        {
            InitializeComponent();
            TbLogin.Focus();
        }

        // Конструктор с параметром (для использования из UserPage)
        public ChangePasswordPage(User user) : this()
        {
            _currentUser = user;

            // Автозаполняем логин текущего пользователя
            if (_currentUser != null)
            {
                TbLogin.Text = _currentUser.Login;
                TbLogin.IsEnabled = false; // Запрещаем изменение логина

                // Переводим фокус на текущий пароль
                PbCurrentPassword.Focus();
            }
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
            TbError.Visibility = Visibility.Visible;
        }

        private void ClearError()
        {
            TbError.Text = string.Empty;
            TbError.Visibility = Visibility.Collapsed;
        }

        private void BtnChangePassword_Click(object sender, RoutedEventArgs e)
        {
            // Если пользователь уже авторизован, используем его данные
            if (_currentUser != null)
            {
                ChangePasswordForCurrentUser();
            }
            else
            {
                ChangePasswordWithLogin();
            }
        }

        private void ChangePasswordForCurrentUser()
        {
            // Валидация полей для авторизованного пользователя
            if (string.IsNullOrWhiteSpace(PbCurrentPassword.Password))
            {
                ShowError("Введите текущий пароль!");
                PbCurrentPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbNewPassword.Password))
            {
                ShowError("Введите новый пароль!");
                PbNewPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbConfirmPassword.Password))
            {
                ShowError("Подтвердите новый пароль!");
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка нового пароля
            if (!ValidatePassword(PbNewPassword.Password))
            {
                ShowError("Новый пароль не соответствует требованиям!\n\nТребования к паролю:\n• Минимум 6 символов\n• Только латинские буквы и цифры\n• Хотя бы одна цифра и одна буква");
                PbNewPassword.Clear();
                PbConfirmPassword.Clear();
                PbNewPassword.Focus();
                return;
            }

            // Проверка совпадения паролей
            if (PbNewPassword.Password != PbConfirmPassword.Password)
            {
                ShowError("Новые пароли не совпадают!");
                PbConfirmPassword.Clear();
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка что новый пароль отличается от старого
            if (PbCurrentPassword.Password == PbNewPassword.Password)
            {
                ShowError("Новый пароль должен отличаться от текущего!");
                PbNewPassword.Clear();
                PbConfirmPassword.Clear();
                PbNewPassword.Focus();
                return;
            }

            try
            {
                using (var db = new DEntities())
                {
                    string hashedCurrentPassword = GetHash(PbCurrentPassword.Password);

                    // Ищем пользователя в базе данных
                    var user = db.Users.Find(_currentUser.ID);

                    if (user == null)
                    {
                        ShowError("Пользователь не найден!");
                        return;
                    }

                    // Проверяем текущий пароль
                    if (user.Password != hashedCurrentPassword)
                    {
                        ShowError("Неверный текущий пароль!");
                        PbCurrentPassword.Clear();
                        PbCurrentPassword.Focus();
                        return;
                    }

                    // Обновление пароля
                    user.Password = GetHash(PbNewPassword.Password);
                    db.SaveChanges();

                    MessageBox.Show("Пароль успешно изменен!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    // Очистка полей
                    PbCurrentPassword.Clear();
                    PbNewPassword.Clear();
                    PbConfirmPassword.Clear();
                    ClearError();

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка при смене пароля: {ex.Message}");
            }
        }

        private void ChangePasswordWithLogin()
        {
            // Валидация полей для неавторизованного пользователя
            if (string.IsNullOrWhiteSpace(TbLogin.Text))
            {
                ShowError("Введите логин!");
                TbLogin.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbCurrentPassword.Password))
            {
                ShowError("Введите текущий пароль!");
                PbCurrentPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbNewPassword.Password))
            {
                ShowError("Введите новый пароль!");
                PbNewPassword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(PbConfirmPassword.Password))
            {
                ShowError("Подтвердите новый пароль!");
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка нового пароля
            if (!ValidatePassword(PbNewPassword.Password))
            {
                ShowError("Новый пароль не соответствует требованиям!\n\nТребования к паролю:\n• Минимум 6 символов\n• Только латинские буквы и цифры\n• Хотя бы одна цифра и одна буква");
                PbNewPassword.Clear();
                PbConfirmPassword.Clear();
                PbNewPassword.Focus();
                return;
            }

            // Проверка совпадения паролей
            if (PbNewPassword.Password != PbConfirmPassword.Password)
            {
                ShowError("Новые пароли не совпадают!");
                PbConfirmPassword.Clear();
                PbConfirmPassword.Focus();
                return;
            }

            // Проверка что новый пароль отличается от старого
            if (PbCurrentPassword.Password == PbNewPassword.Password)
            {
                ShowError("Новый пароль должен отличаться от текущего!");
                PbNewPassword.Clear();
                PbConfirmPassword.Clear();
                PbNewPassword.Focus();
                return;
            }

            try
            {
                using (var db = new DEntities())
                {
                    string hashedCurrentPassword = GetHash(PbCurrentPassword.Password);

                    var user = db.Users.FirstOrDefault(u =>
                        u.Login == TbLogin.Text.Trim() &&
                        u.Password == hashedCurrentPassword);

                    if (user == null)
                    {
                        ShowError("Неверный логин или текущий пароль!");
                        PbCurrentPassword.Clear();
                        PbCurrentPassword.Focus();
                        return;
                    }

                    // Обновление пароля
                    user.Password = GetHash(PbNewPassword.Password);
                    db.SaveChanges();

                    MessageBox.Show("Пароль успешно изменен!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    // Очистка полей
                    TbLogin.Clear();
                    PbCurrentPassword.Clear();
                    PbNewPassword.Clear();
                    PbConfirmPassword.Clear();
                    ClearError();

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка при смене пароля: {ex.Message}");
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }

        // Обработчик загрузки страницы
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            // Если пользователь передан, настраиваем интерфейс соответствующим образом
            if (_currentUser != null)
            {
                TbLogin.Text = _currentUser.Login;
                TbLogin.IsEnabled = false;

                // Показываем подсказку, что логин заполнен автоматически
                var toolTip = new ToolTip();
                toolTip.Content = "Логин заполнен автоматически для текущего пользователя";
                TbLogin.ToolTip = toolTip;
            }
        }
    }
}