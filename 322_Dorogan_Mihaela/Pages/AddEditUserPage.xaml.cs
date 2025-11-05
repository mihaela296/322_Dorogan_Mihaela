using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Data.Entity;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class AddEditUserPage : Page
    {
        private User _currentUser;
        private User _editingUser;
        private bool _isNewUser;

        public AddEditUserPage(User user, User currentAdmin)
        {
            InitializeComponent();
            _currentUser = currentAdmin;
            _editingUser = user ?? new User();
            _isNewUser = user == null;

            InitializeForm();
        }

        private void InitializeForm()
        {
            if (_isNewUser)
            {
                TbTitle.Text = "ДОБАВЛЕНИЕ НОВОГО ПОЛЬЗОВАТЕЛЯ";
                LblPassword.Content = "Пароль:*";
                TbPasswordRequirements.Visibility = Visibility.Visible;
            }
            else
            {
                TbTitle.Text = $"РЕДАКТИРОВАНИЕ ПОЛЬЗОВАТЕЛЯ: {_editingUser.FIO}";
                LblPassword.Content = "Пароль:";
                TbPasswordRequirements.Visibility = Visibility.Collapsed;
                PbPassword.Password = "********"; // Заглушка для существующего пользователя
            }

            DataContext = _editingUser;

            // Установка выбранной роли в комбобоксе
            if (!string.IsNullOrEmpty(_editingUser.Role))
            {
                foreach (ComboBoxItem item in CbRole.Items)
                {
                    if (item.Content.ToString() == _editingUser.Role)
                    {
                        CbRole.SelectedItem = item;
                        break;
                    }
                }
            }
            else
            {
                CbRole.SelectedIndex = 0; // User по умолчанию
            }
        }

        private string GetHash(string input)
        {
            using (var sha1 = System.Security.Cryptography.SHA1.Create())
            {
                var hash = sha1.ComputeHash(System.Text.Encoding.UTF8.GetBytes(input));
                return string.Concat(hash.Select(b => b.ToString("X2")));
            }
        }

        private bool ValidatePassword(string password)
        {
            if (_isNewUser && string.IsNullOrEmpty(password))
                return false;

            if (!string.IsNullOrEmpty(password))
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
            }

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

        private void BtnBrowsePhoto_Click(object sender, RoutedEventArgs e)
        {
            var openDialog = new OpenFileDialog
            {
                Filter = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*",
                Title = "Выберите фото пользователя"
            };

            if (openDialog.ShowDialog() == true)
            {
                TbPhoto.Text = openDialog.FileName;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateForm())
                return;

            try
            {
                using (var db = new DEntities())
                {
                    if (_isNewUser)
                    {
                        // Проверка уникальности логина
                        if (db.Users.Any(u => u.Login == _editingUser.Login))
                        {
                            ShowError("Пользователь с таким логином уже существует!");
                            TbLogin.Focus();
                            TbLogin.SelectAll();
                            return;
                        }

                        // Установка хешированного пароля для нового пользователя
                        _editingUser.Password = GetHash(PbPassword.Password);
                        db.Users.Add(_editingUser);
                    }
                    else
                    {
                        var existingUser = db.Users.Find(_editingUser.ID);
                        if (existingUser != null)
                        {
                            // Проверка уникальности логина (исключая текущего пользователя)
                            if (db.Users.Any(u => u.Login == _editingUser.Login && u.ID != _editingUser.ID))
                            {
                                ShowError("Пользователь с таким логином уже существует!");
                                TbLogin.Focus();
                                TbLogin.SelectAll();
                                return;
                            }

                            // Обновление полей
                            existingUser.Login = _editingUser.Login;
                            existingUser.FIO = _editingUser.FIO;
                            existingUser.Role = _editingUser.Role;
                            existingUser.Photo = _editingUser.Photo;

                            // Обновление пароля только если он был изменен
                            if (PbPassword.Password != "********" && !string.IsNullOrEmpty(PbPassword.Password))
                            {
                                existingUser.Password = GetHash(PbPassword.Password);
                            }
                        }
                    }

                    db.SaveChanges();

                    MessageBox.Show(_isNewUser ? "Пользователь успешно добавлен!" : "Данные пользователя успешно обновлены!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка сохранения: {ex.Message}");
            }
        }

        private bool ValidateForm()
        {
            ClearError();

            if (string.IsNullOrWhiteSpace(_editingUser.Login))
            {
                ShowError("Введите логин!");
                TbLogin.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(_editingUser.FIO))
            {
                ShowError("Введите ФИО!");
                TbFIO.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(_editingUser.Role))
            {
                ShowError("Выберите роль!");
                CbRole.Focus();
                return false;
            }

            // Валидация пароля
            if (_isNewUser)
            {
                if (string.IsNullOrWhiteSpace(PbPassword.Password))
                {
                    ShowError("Введите пароль!");
                    PbPassword.Focus();
                    return false;
                }

                if (!ValidatePassword(PbPassword.Password))
                {
                    ShowError("Пароль не соответствует требованиям!");
                    PbPassword.Focus();
                    return false;
                }
            }
            else if (PbPassword.Password != "********" && !string.IsNullOrEmpty(PbPassword.Password))
            {
                // Проверка пароля при его изменении для существующего пользователя
                if (!ValidatePassword(PbPassword.Password))
                {
                    ShowError("Пароль не соответствует требованиям!");
                    PbPassword.Focus();
                    return false;
                }
            }

            // Проверка формата логина
            if (!Regex.IsMatch(_editingUser.Login, @"^[a-zA-Z0-9]+$"))
            {
                ShowError("Логин должен содержать только латинские буквы и цифры!");
                TbLogin.Focus();
                return false;
            }

            return true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}