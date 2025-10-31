﻿using System;
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

            // Проверяем и создаем базу данных при загрузке
            this.Loaded += (s, e) =>
            {
                CheckAndCreateDatabase();
            };
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

        private void CheckAndCreateDatabase()
        {
            try
            {
                using (var db = new Entities())
                {
                    // Проверяем существование базы данных и создаем при необходимости
                    if (!db.Database.Exists())
                    {
                        MessageBox.Show("База данных не существует. Создаем новую базу...", "Информация",
                            MessageBoxButton.OK, MessageBoxImage.Information);

                        // Создаем базу данных и таблицы
                        db.Database.Create();

                        // Добавляем начальные данные
                        AddInitialData(db);

                        MessageBox.Show("База данных успешно создана!", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Проверяем существование таблиц
                        if (!db.Users.Any())
                        {
                            AddInitialData(db);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации базы данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddInitialData(Entities db)
        {
            // Добавляем администратора по умолчанию
            if (!db.Users.Any(u => u.Login == "admin"))
            {
                var adminUser = new User
                {
                    Login = "admin",
                    Password = GetHash("admin123"),
                    FIO = "Администратор Системы",
                    Role = "Admin"
                };
                db.Users.Add(adminUser);
            }

            // Добавляем базовые категории
            if (!db.Categories.Any())
            {
                var categories = new[]
                {
                    new Category { Name = "Продукты питания" },
                    new Category { Name = "Коммунальные услуги" },
                    new Category { Name = "Транспорт" },
                    new Category { Name = "Развлечения" },
                    new Category { Name = "Одежда" }
                };

                foreach (var category in categories)
                {
                    db.Categories.Add(category);
                }
            }

            db.SaveChanges();
        }

        private void VerifyDataSaved()
        {
            try
            {
                using (var db = new Entities())
                {
                    var userCount = db.Users.Count();
                    var categoryCount = db.Categories.Count();

                    MessageBox.Show($"В базе данных:\nПользователей: {userCount}\nКатегорий: {categoryCount}",
                        "Проверка данных", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка проверки данных: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

            // Проверка выбора роли
            if (CbRole.SelectedItem == null)
            {
                ShowError("Выберите роль пользователя!");
                CbRole.Focus();
                return;
            }

            // Получаем выбранную роль
            string selectedRole = ((ComboBoxItem)CbRole.SelectedItem).Content.ToString();

            // Проверка логина
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

            // Проверка уникальности логина и сохранение
            try
            {
                using (var db = new Entities())
                {
                    // Проверяем подключение к базе
                    if (!db.Database.Exists())
                    {
                        ShowError("Не удалось подключиться к базе данных. Проверьте настройки подключения.");
                        return;
                    }

                    // Проверяем уникальность логина
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
                        Role = selectedRole
                    };

                    db.Users.Add(newUser);

                    // Сохраняем изменения
                    int result = db.SaveChanges();

                    if (result > 0)
                    {
                        MessageBox.Show("Регистрация прошла успешно!\nТеперь вы можете войти в систему используя ваш логин и пароль.",
                            "Успешная регистрация", MessageBoxButton.OK, MessageBoxImage.Information);

                        // Проверяем, что данные действительно сохранились
                        VerifyDataSaved();

                        // Очищаем форму
                        TbLogin.Clear();
                        PbPassword.Clear();
                        PbConfirmPassword.Clear();
                        TbFIO.Clear();
                        CbRole.SelectedIndex = 0;
                        ClearError();

                        NavigationService.GoBack();
                    }
                    else
                    {
                        ShowError("Не удалось сохранить пользователя в базу данных.");
                    }
                }
            }
            catch (System.Data.Entity.Infrastructure.DbUpdateException dbEx)
            {
                ShowError($"Ошибка обновления базы данных: {dbEx.InnerException?.Message ?? dbEx.Message}");
            }
            catch (System.Data.SqlClient.SqlException sqlEx)
            {
                ShowError($"Ошибка SQL Server: {sqlEx.Message}\nНомер ошибки: {sqlEx.Number}");
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка при регистрации: {ex.Message}");
            }
        }

        private void BtnCheckDB_Click(object sender, RoutedEventArgs e)
        {
            CheckAndCreateDatabase();
            VerifyDataSaved();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}